"""
Gera os dados FICTÍCIOS do monitor de parcelamentos, no modelo da planilha de controle:
a pessoa digita só a coluna 'parcelas_pagas'; o restante é calculado por fórmulas do Excel.
Semente fixa: os dados são sempre os mesmos.
"""
import csv
import random
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation


random.seed(2027)


def dv_cnpj(base12):
    pesos1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
    pesos2 = [6] + pesos1
    d1 = sum(int(n) * p for n, p in zip(base12, pesos1)) % 11
    d1 = 0 if d1 < 2 else 11 - d1
    d2 = sum(int(n) * p for n, p in zip(base12 + str(d1), pesos2)) % 11
    d2 = 0 if d2 < 2 else 11 - d2
    return f"{d1}{d2}"


def gerar_cnpj():
    base8 = "".join(random.choice("0123456789") for _ in range(8))
    base12 = base8 + "0001"
    return base12 + dv_cnpj(base12)


def fmt_cnpj(n):
    return f"{n[:2]}.{n[2:5]}.{n[5:8]}/{n[8:12]}-{n[12:]}"

PASTA = Path(".")
DATA_REFERENCIA = date(2026, 10, 31)
LIMITES = {"Federal": (1, 2, 3, 4), "Estadual": (1, 2, 3, 4), "Municipal": (1, 2, 3, 4)}

AMBITOS = {
    "Federal": ("Receita Federal (fictício)", "PARC-FED",
                ["IRPJ", "CSLL", "PIS", "COFINS", "IPI", "Contribuição previdenciária", "PIS"]),
    "Estadual": ("Secretaria da Fazenda Estadual (fictícia)", "PARC-EST",
                 ["ICMS", "ICMS-ST", "IPVA", "ICMS", "ITCMD", "ICMS", "ICMS-ST"]),
    "Municipal": ("Prefeitura Municipal (fictícia)", "PARC-MUN",
                  ["ISS", "IPTU", "ISS", "Taxa de licença", "ISS", "IPTU", "ISS"]),
}
CENARIOS = [0, 1, 2, 3, 4, "quitado", "nao_iniciado"]


def add_meses(d, n):
    m = d.month - 1 + n
    return date(d.year + m // 12, m % 12 + 1, d.day)


def vencidas(primeira, qtd, ref):
    """Mesma lógica da fórmula do Excel: parcelas com vencimento anterior à data de referência."""
    if ref <= primeira:
        return 0
    m = (ref.year - primeira.year) * 12 + ref.month - primeira.month
    return min(qtd, m + (1 if primeira.day < ref.day else 0))


def nivel(faltantes, atrasadas, lim):
    atencao, alerta, urgencia, rompimento = lim
    if faltantes == 0:
        return "quitado"
    if atrasadas >= rompimento:
        return "rompido"
    if atrasadas >= urgencia:
        return "urgencia"
    if atrasadas >= alerta:
        return "alerta"
    if atrasadas >= atencao:
        return "atencao"
    return "em_dia"


ACOES = {
    "quitado": "Parcelamento quitado",
    "em_dia": "Nenhuma ação",
    "atencao": "Regularizar a parcela em atraso",
    "alerta": "Regularizar as parcelas em atraso com prioridade",
    "urgencia": "Pagar já: se a próxima parcela vencer sem pagamento, o parcelamento é rompido",
    "rompido": "Parcelamento rompido: verificar situação e possibilidade de reparcelamento",
}


def montar_linhas():
    linhas = []
    primeiro = True
    for ambito, (orgao, prefixo, tributos) in AMBITOS.items():
        for i, cen in enumerate(CENARIOS):
            id_parc = f"{prefixo}-2026-{i + 1:03d}"
            valor = round(random.randint(450, 9800) + random.choice([0, 0.5]), 2)
            dia = random.choice([1, 10, 15, 20, 25, 28])
            if primeiro:                       # exemplo igual ao da explicação
                primeira, qtd, primeiro = date(2026, 3, 1), 30, False
            elif cen == "quitado":
                primeira, qtd = date(2026, random.randint(1, 3), dia), 6
            elif cen == "nao_iniciado":
                primeira, qtd = date(2026, random.choice([11, 12]), dia), 24
            else:
                primeira = date(2026, random.randint(1, 5), dia)
                qtd = None
            if cen == "nao_iniciado":
                venc, pagas = 0, 0
            elif cen == "quitado":
                venc, pagas = qtd, qtd
            else:
                venc = vencidas(primeira, 999, DATA_REFERENCIA)
                if qtd is None:
                    qtd = venc + random.randint(6, 36)
                pagas = venc - cen
            linhas.append({
                "id": id_parc, "ambito": ambito, "orgao": orgao,
                "tributo": tributos[i], "cnpj": fmt_cnpj(gerar_cnpj()),
                "data_parcelamento": add_meses(primeira, -1), "qtd": qtd,
                "primeira": primeira, "valor": valor, "pagas": pagas,
            })
    return linhas


def main():
    PASTA.mkdir(parents=True, exist_ok=True)
    linhas = montar_linhas()
    wb = Workbook()
    ws = wb.active
    ws.title = "parcelamentos"
    cfg = wb.create_sheet("config")
    leg = wb.create_sheet("legenda")

    arial = Font(name="Arial", size=10)
    azul = Font(name="Arial", size=10, color="0000FF")
    cab = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    fundo_cab = PatternFill("solid", fgColor="305496")
    amarelo = PatternFill("solid", fgColor="FFFF00")
    cinza = PatternFill("solid", fgColor="D9D9D9")

    # ---------------- config ----------------
    cfg["A1"], cfg["B1"] = "data_referencia", DATA_REFERENCIA
    cfg["B1"].number_format = "dd/mm/yyyy"
    cfg["C1"] = ("Data fixa para a demonstração e para o gabarito bater. "
                 "No uso real, troque por =HOJE().")
    cfg["A3"] = "Limites de parcelas em atraso por âmbito (valores ILUSTRATIVOS; as regras reais variam por modalidade)"
    for j, t in enumerate(["ambito", "atencao", "alerta", "urgencia", "rompimento"], 1):
        c = cfg.cell(row=4, column=j, value=t)
        c.font, c.fill = cab, fundo_cab
    for k, (amb, lim) in enumerate(LIMITES.items(), 5):
        cfg.cell(row=k, column=1, value=amb)
        for j, v in enumerate(lim, 2):
            cfg.cell(row=k, column=j, value=v).font = azul
    cfg["B1"].font = azul

    # ---------------- parcelamentos ----------------
    cols = ["id_parcelamento", "ambito", "orgao", "tributo", "cnpj", "data_parcelamento",
            "qtd_parcelas", "vencimento_1a_parcela", "valor_parcela", "parcelas_pagas",
            "parcelas_faltantes", "parcelas_vencidas", "parcelas_em_atraso", "nivel",
            "proxima_parcela", "mensagem_alerta", "ultimo_nivel_alertado", "enviar_alerta"]
    for j, t in enumerate(cols, 1):
        c = ws.cell(row=1, column=j, value=t)
        c.font, c.fill = cab, fundo_cab
        c.alignment = Alignment(horizontal="left")

    ref = "config!$B$1"
    tabela = "config!$A$5:$A$7"
    for r, d in enumerate(linhas, 2):
        entrada = [d["id"], d["ambito"], d["orgao"], d["tributo"], d["cnpj"],
                   d["data_parcelamento"], d["qtd"], d["primeira"], d["valor"], d["pagas"]]
        for j, v in enumerate(entrada, 1):
            c = ws.cell(row=r, column=j, value=v)
            c.font = azul if j >= 6 else arial
        ws.cell(row=r, column=10).fill = amarelo

        lim = lambda col: f"INDEX(config!${col}$5:${col}$7,MATCH(B{r},{tabela},0))"
        f = {
            11: f"=G{r}-J{r}",
            12: (f"=IF({ref}<=H{r},0,MIN(G{r},(YEAR({ref})-YEAR(H{r}))*12+MONTH({ref})-MONTH(H{r})"
                 f"+IF(DAY(H{r})<DAY({ref}),1,0)))"),
            13: f"=MAX(0,L{r}-J{r})",
            14: (f'=IF(K{r}=0,"quitado",IF(M{r}>={lim("E")},"rompido",IF(M{r}>={lim("D")},"urgencia",'
                 f'IF(M{r}>={lim("C")},"alerta",IF(M{r}>={lim("B")},"atencao","em_dia")))))'),
            15: f'=IF(K{r}=0,"",EDATE(H{r},J{r}))',
            16: (f'=IF(N{r}="rompido","ROMPIDO: "&M{r}&" parcelas em atraso. Verificar situação e possibilidade de reparcelamento.",'
                 f'IF(N{r}="urgencia","URGENTE: "&M{r}&" parcelas em atraso. Se a próxima parcela vencer sem pagamento, o parcelamento será rompido.",'
                 f'IF(N{r}="alerta","ALERTA: "&M{r}&" parcelas em atraso. Regularizar com prioridade.",'
                 f'IF(N{r}="atencao","Atenção: 1 parcela em atraso.",IF(N{r}="quitado","Parcelamento quitado.","Em dia.")))))'),
            18: f'=IF(AND(OR(N{r}="alerta",N{r}="urgencia",N{r}="rompido"),N{r}<>Q{r}),"SIM","NAO")',
        }
        for j, formula in f.items():
            ws.cell(row=r, column=j, value=formula).font = arial
        ws.cell(row=r, column=17).font = arial
        ws.cell(row=r, column=17).fill = cinza
        for j in (6, 8, 15):
            ws.cell(row=r, column=j).number_format = "dd/mm/yyyy"
        ws.cell(row=r, column=9).number_format = "#,##0.00"

    ultima = len(linhas) + 1
    dv = DataValidation(type="whole", operator="between", formula1="0", formula2="G2",
                        showErrorMessage=True, errorTitle="Valor inválido",
                        error="Digite um número inteiro entre 0 e a quantidade de parcelas.")
    ws.add_data_validation(dv)
    dv.add(f"J2:J{ultima}")

    faixa = f"N2:N{ultima}"
    cores = {"rompido": "FF7C80", "urgencia": "FFB366", "alerta": "FFE699",
             "atencao": "FFF2CC", "em_dia": "C6E0B4", "quitado": "D9D9D9"}
    for nome, cor in cores.items():
        ws.conditional_formatting.add(
            faixa, CellIsRule(operator="equal", formula=[f'"{nome}"'],
                              fill=PatternFill("solid", bgColor=cor, fgColor=cor)))

    larg = {"A": 20, "B": 11, "C": 38, "D": 24, "E": 20, "F": 18, "G": 13, "H": 22, "I": 14,
            "J": 15, "K": 18, "L": 17, "M": 18, "N": 11, "O": 16, "P": 95, "Q": 22, "R": 14}
    for k, v in larg.items():
        ws.column_dimensions[k].width = v
    ws.freeze_panes = "B2"
    cfg.column_dimensions["A"].width = 20
    for k in "BCDE":
        cfg.column_dimensions[k].width = 14

    # ---------------- legenda ----------------
    texto = [
        "COMO USAR A PLANILHA (dados fictícios)",
        "",
        "Cor azul = dado de entrada. Cor preta = fórmula (não editar).",
        "Coluna J (amarela, parcelas_pagas): é a ÚNICA que a pessoa digita no dia a dia. Coloque quantas parcelas já foram pagas (1, 2, 3...).",
        "parcelas_faltantes = qtd_parcelas - parcelas_pagas.",
        "parcelas_vencidas = quantas parcelas já passaram do vencimento na data de referência (aba config).",
        "parcelas_em_atraso = parcelas_vencidas - parcelas_pagas (nunca menor que zero).",
        "nivel: em_dia, atencao (1 em atraso), alerta (2), urgencia (3), rompido (4 ou mais), ou quitado.",
        "proxima_parcela: vencimento da próxima parcela a pagar.",
        "ultimo_nivel_alertado (cinza): preenchida pelo programa de alertas, não pela pessoa. Serve para não enviar o mesmo alerta todo dia.",
        "enviar_alerta: SIM quando o nível é alerta, urgência ou rompido e ainda não foi avisado.",
        "",
        "Premissas: as parcelas vencem todo mês no mesmo dia da primeira; uma parcela só é considerada em atraso depois do dia do vencimento;",
        "a regra de antecipar as primeiras parcelas (comum em parcelamentos federais) foi ignorada; os limites são ilustrativos.",
        "Exemplo: a linha PARC-FED-2026-001 tem parcelamento em 01/02/2026, 30 parcelas e primeira parcela em 01/03/2026.",
    ]
    for i, t in enumerate(texto, 1):
        c = leg.cell(row=i, column=1, value=t)
        c.font = Font(name="Arial", size=10, bold=(i == 1))
    leg.column_dimensions["A"].width = 150

    wb.save(PASTA / "parcelamentos.xlsx")

    # ---------------- gabarito ----------------
    with open(PASTA / "gabarito.csv", "w", newline="", encoding="utf-8") as fh:
        wr = csv.writer(fh)
        wr.writerow(["id_parcelamento", "ambito", "parcelas_pagas", "parcelas_faltantes",
                     "parcelas_vencidas", "parcelas_em_atraso", "nivel_esperado",
                     "proxima_parcela", "acao_esperada"])
        for d in linhas:
            venc = vencidas(d["primeira"], d["qtd"], DATA_REFERENCIA)
            falt = d["qtd"] - d["pagas"]
            atr = max(0, venc - d["pagas"])
            niv = nivel(falt, atr, LIMITES[d["ambito"]])
            prox = add_meses(d["primeira"], d["pagas"]).isoformat() if falt > 0 else ""
            wr.writerow([d["id"], d["ambito"], d["pagas"], falt, venc, atr, niv, prox, ACOES[niv]])
    print(f"{len(linhas)} parcelamentos gerados")


if __name__ == "__main__":
    main()
