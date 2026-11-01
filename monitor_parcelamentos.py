"""
Monitor de parcelamentos tributários (dados fictícios).

Lê a planilha de controle, calcula a situação de cada parcelamento e monta
UM e-mail consolidado com todos eles, agrupados por gravidade.
Por padrão o envio é simulado: o e-mail é salvo na pasta 'saida'.
Para enviar de verdade, use --enviar e configure as variáveis de ambiente.
"""
import argparse
import calendar
import json
import os
import smtplib
import ssl
from datetime import date, datetime
from email.message import EmailMessage
from html import escape
from pathlib import Path

from openpyxl import load_workbook

PLANILHA_PADRAO = Path("dados/parcelamentos.xlsx")
PASTA_SAIDA = Path("saida")
ARQUIVO_ESTADO = PASTA_SAIDA / "estado_ultimo_envio.json"

NIVEIS = ["rompido", "urgencia", "alerta", "atencao", "em_dia", "quitado"]
GRAVIDADE = {"quitado": 0, "em_dia": 0, "atencao": 1, "alerta": 2, "urgencia": 3, "rompido": 4}

TITULOS = {
    "rompido": "ROMPIDOS - verificar possibilidade de reparcelamento",
    "urgencia": "URGÊNCIA - se a próxima parcela vencer sem pagamento, o parcelamento é rompido",
    "alerta": "ALERTA - parcelas em atraso",
    "atencao": "ATENÇÃO - uma parcela em atraso",
    "em_dia": "EM DIA",
    "quitado": "QUITADOS",
}
ACOES = {
    "rompido": "Verificar possibilidade de reparcelamento",
    "urgencia": "Pagar já: a próxima parcela vencida rompe o parcelamento",
    "alerta": "Regularizar com prioridade",
    "atencao": "Regularizar a parcela em atraso",
    "em_dia": "",
    "quitado": "Quitado",
}
CORES = {
    "rompido": "#C00000", "urgencia": "#E36C09", "alerta": "#BF9000",
    "atencao": "#2E75B6", "em_dia": "#548235", "quitado": "#7F7F7F",
}


def add_meses(d, n):
    m = d.month - 1 + n
    ano, mes = d.year + m // 12, m % 12 + 1
    return date(ano, mes, min(d.day, calendar.monthrange(ano, mes)[1]))


def parcelas_vencidas(primeira, qtd, referencia):
    return sum(1 for k in range(qtd) if add_meses(primeira, k) < referencia)


def classificar(faltantes, atrasadas, limites):
    if faltantes == 0:
        return "quitado"
    if atrasadas >= limites["rompimento"]:
        return "rompido"
    if atrasadas >= limites["urgencia"]:
        return "urgencia"
    if atrasadas >= limites["alerta"]:
        return "alerta"
    if atrasadas >= limites["atencao"]:
        return "atencao"
    return "em_dia"


def para_data(valor):
    if isinstance(valor, datetime):
        return valor.date()
    return valor if isinstance(valor, date) else None


def carregar_config(wb):
    ws = wb["config"]
    referencia = para_data(ws["B1"].value) or date.today()
    limites = {}
    for ambito, atencao, alerta, urgencia, rompimento in ws.iter_rows(
            min_row=5, max_col=5, values_only=True):
        if ambito:
            limites[ambito] = {"atencao": atencao, "alerta": alerta,
                               "urgencia": urgencia, "rompimento": rompimento}
    return referencia, limites


def carregar_analise(caminho, referencia=None):
    """Devolve (lista de parcelamentos analisados, data de referência)."""
    wb = load_workbook(caminho, data_only=True)
    ref_planilha, limites = carregar_config(wb)
    referencia = referencia or ref_planilha

    ws = wb["parcelamentos"]
    cabecalho = [c.value for c in ws[1]]
    analise = []
    for linha in ws.iter_rows(min_row=2, values_only=True):
        d = dict(zip(cabecalho, linha))
        if not d.get("id_parcelamento"):
            continue
        qtd, pagas = int(d["qtd_parcelas"]), int(d["parcelas_pagas"] or 0)
        primeira = para_data(d["vencimento_1a_parcela"])
        faltantes = qtd - pagas
        vencidas = parcelas_vencidas(primeira, qtd, referencia)
        atrasadas = max(0, vencidas - pagas)
        nivel = classificar(faltantes, atrasadas, limites[d["ambito"]])
        analise.append({
            "id": d["id_parcelamento"], "ambito": d["ambito"], "tributo": d["tributo"],
            "cnpj": d["cnpj"], "qtd": qtd, "pagas": pagas, "faltantes": faltantes,
            "atrasadas": atrasadas, "nivel": nivel, "valor": float(d["valor_parcela"]),
            "proxima": add_meses(primeira, pagas) if faltantes > 0 else None,
        })
    analise.sort(key=lambda p: (-GRAVIDADE[p["nivel"]], -p["atrasadas"], p["id"]))
    return analise, referencia


def carregar_estado():
    if ARQUIVO_ESTADO.exists():
        return json.loads(ARQUIVO_ESTADO.read_text(encoding="utf-8"))
    return {}


def salvar_estado(analise):
    PASTA_SAIDA.mkdir(exist_ok=True)
    ARQUIVO_ESTADO.write_text(
        json.dumps({p["id"]: p["nivel"] for p in analise}, ensure_ascii=False, indent=2),
        encoding="utf-8")


def piorou(p, anteriores):
    antes = anteriores.get(p["id"])
    return antes if antes and GRAVIDADE[p["nivel"]] > GRAVIDADE[antes] else None


def brl(v):
    return "R$ " + f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def contagens(analise):
    return {n: sum(1 for p in analise if p["nivel"] == n) for n in NIVEIS}


def assunto(analise, referencia):
    c = contagens(analise)
    return (f"Parcelamentos {referencia:%d/%m/%Y}: {c['rompido']} rompido(s), "
            f"{c['urgencia']} em urgência, {c['alerta']} em alerta, {c['atencao']} em atenção")


def observacao(p, anteriores):
    texto = ACOES[p["nivel"]]
    antes = piorou(p, anteriores)
    if antes:
        texto += f" (piorou: era {antes})"
    return texto


def montar_texto(analise, referencia, anteriores):
    c = contagens(analise)
    linhas = [f"MONITOR DE PARCELAMENTOS - {referencia:%d/%m/%Y}", "",
              f"Rompidos: {c['rompido']} | Urgência: {c['urgencia']} | Alerta: {c['alerta']} | "
              f"Atenção: {c['atencao']} | Em dia: {c['em_dia']} | Quitados: {c['quitado']}"]
    for nivel in NIVEIS:
        grupo = [p for p in analise if p["nivel"] == nivel]
        if not grupo:
            continue
        linhas += ["", f"== {TITULOS[nivel]} ({len(grupo)}) =="]
        for p in grupo:
            prox = f"{p['proxima']:%d/%m/%Y}" if p["proxima"] else "-"
            linhas.append(f"- {p['id']} | {p['ambito']} | {p['tributo']} | "
                          f"{p['pagas']}/{p['qtd']} pagas | {p['atrasadas']} em atraso | "
                          f"próxima: {prox} | {observacao(p, anteriores)}")
    linhas += ["", "Dados fictícios, para demonstração."]
    return "\n".join(linhas)


LARGURAS = [15, 8, 11, 15, 6, 6, 10, 11, 18]


def th(txt, largura):
    return (f'<th style="width:{largura}%;text-align:left;padding:6px 8px;background:#F2F2F2;'
            f'border-bottom:1px solid #D9D9D9;font-size:12px">{txt}</th>')


def td(txt, estilo=""):
    return f'<td style="padding:6px 8px;border-bottom:1px solid #EEE;font-size:12px;{estilo}">{txt}</td>'


def montar_html(analise, referencia, anteriores):
    c = contagens(analise)
    caixas = ""
    for nivel, rotulo in [("rompido", "Rompidos"), ("urgencia", "Urgência"), ("alerta", "Alerta"),
                          ("atencao", "Atenção"), ("em_dia", "Em dia"), ("quitado", "Quitados")]:
        caixas += (f'<td style="text-align:center;padding:10px 14px;background:{CORES[nivel]};'
                   f'color:#FFF;border-radius:4px"><div style="font-size:22px;font-weight:bold">'
                   f'{c[nivel]}</div><div style="font-size:11px">{rotulo}</div></td>'
                   '<td style="width:6px"></td>')

    blocos = ""
    for nivel in NIVEIS:
        grupo = [p for p in analise if p["nivel"] == nivel]
        if not grupo:
            continue
        linhas = ""
        for p in grupo:
            vencida = p["proxima"] and p["proxima"] < referencia
            prox = f"{p['proxima']:%d/%m/%Y}" if p["proxima"] else "-"
            estilo_prox = "color:#C00000;font-weight:bold;" if vencida else ""
            obs = escape(observacao(p, anteriores))
            if piorou(p, anteriores):
                obs = f'<b style="color:#C00000">{obs}</b>'
            linhas += ("<tr>" + td(escape(p["id"]), "font-weight:bold;") + td(escape(p["ambito"]))
                       + td(escape(str(p["tributo"]))) + td(escape(str(p["cnpj"])))
                       + td(f"{p['pagas']}/{p['qtd']}") + td(str(p["atrasadas"]))
                       + td(prox, estilo_prox) + td(brl(p["valor"])) + td(obs) + "</tr>")
        titulos_colunas = ["Parcelamento", "Âmbito", "Tributo", "CNPJ", "Pagas",
                           "Em atraso", "Próxima a pagar", "Valor da parcela", "Observação"]
        cabecalho = "".join(th(t, w) for t, w in zip(titulos_colunas, LARGURAS))
        blocos += (f'<h3 style="margin:22px 0 6px;color:{CORES[nivel]};font-size:14px;'
                   f'border-left:5px solid {CORES[nivel]};padding-left:8px">'
                   f'{escape(TITULOS[nivel])} ({len(grupo)})</h3>'
                   f'<table style="border-collapse:collapse;width:100%;table-layout:fixed"><tr>{cabecalho}</tr>'
                   f'{linhas}</table>')

    return (f'<html><head><meta charset="utf-8"></head>'
            f'<body style="font-family:Arial,sans-serif;color:#222;max-width:980px;margin:auto">'
            f'<h2 style="margin-bottom:4px">Monitor de parcelamentos</h2>'
            f'<div style="color:#666;font-size:12px;margin-bottom:12px">Situação em '
            f'{referencia:%d/%m/%Y}</div><table style="border-collapse:collapse"><tr>{caixas}</tr></table>'
            f'{blocos}<p style="color:#888;font-size:11px;margin-top:24px">'
            f'Dados fictícios, para demonstração.</p></body></html>')


def enviar_email(assunto_email, texto, html):
    faltando = [v for v in ("SMTP_HOST", "SMTP_USER", "SMTP_PASSWORD", "EMAIL_DESTINATARIOS")
                if not os.environ.get(v)]
    if faltando:
        raise SystemExit(f"Defina as variáveis de ambiente: {', '.join(faltando)}")
    msg = EmailMessage()
    msg["Subject"] = assunto_email
    msg["From"] = os.environ.get("EMAIL_REMETENTE", os.environ["SMTP_USER"])
    msg["To"] = os.environ["EMAIL_DESTINATARIOS"]
    msg.set_content(texto)
    msg.add_alternative(html, subtype="html")
    with smtplib.SMTP(os.environ["SMTP_HOST"], int(os.environ.get("SMTP_PORT", 587))) as s:
        s.starttls(context=ssl.create_default_context())
        s.login(os.environ["SMTP_USER"], os.environ["SMTP_PASSWORD"])
        s.send_message(msg)


def main():
    ap = argparse.ArgumentParser(description="E-mail diário de parcelamentos")
    ap.add_argument("--planilha", type=Path, default=PLANILHA_PADRAO)
    ap.add_argument("--data", help="data de referência AAAA-MM-DD (padrão: a da aba config)")
    ap.add_argument("--hoje", action="store_true", help="usar a data de hoje")
    ap.add_argument("--enviar", action="store_true", help="enviar por SMTP (padrão: simulado)")
    args = ap.parse_args()

    referencia = None
    if args.hoje:
        referencia = date.today()
    if args.data:
        referencia = date.fromisoformat(args.data)

    analise, referencia = carregar_analise(args.planilha, referencia)
    anteriores = carregar_estado()
    texto = montar_texto(analise, referencia, anteriores)
    html = montar_html(analise, referencia, anteriores)
    titulo = assunto(analise, referencia)

    PASTA_SAIDA.mkdir(exist_ok=True)
    (PASTA_SAIDA / f"email_{referencia:%Y-%m-%d}.html").write_text(html, encoding="utf-8")
    (PASTA_SAIDA / f"email_{referencia:%Y-%m-%d}.txt").write_text(texto, encoding="utf-8")
    salvar_estado(analise)

    print(titulo)
    if args.enviar:
        enviar_email(titulo, texto, html)
        print("E-mail enviado.")
    else:
        print(f"Modo simulado: e-mail salvo em {PASTA_SAIDA}/. Nada foi enviado.")


if __name__ == "__main__":
    main()
