# Monitor de Parcelamentos Tributários

Programa em Python que lê a planilha de controle de parcelamentos, calcula quais estão em risco de rompimento e monta **um e-mail diário consolidado**, com todos os parcelamentos agrupados por gravidade.

> Projeto de portfólio com **dados 100% fictícios**. Reproduz, de forma limpa e segura, uma rotina de compliance fiscal que automatizei na prática.

![Exemplo do e-mail gerado](docs/exemplo_email.png)

## O problema

Parcelamentos tributários (federais, estaduais e municipais) são rompidos quando as parcelas em atraso se acumulam, e isso pode passar despercebido quando o controle é uma planilha conferida à mão. O programa avisa antes que o parcelamento rompa.

## Como funciona

1. A pessoa mantém a planilha de controle e digita **só uma coluna**: `parcelas_pagas`.
2. O programa lê a planilha, calcula as parcelas vencidas e em atraso de cada parcelamento e classifica o nível de risco.
3. Monta um e-mail consolidado (HTML e texto), com as caixinhas de resumo no topo e os parcelamentos agrupados do mais grave ao menos grave.
4. Marca em vermelho quem **piorou desde o último envio**, para destacar o que mudou.

### Regra de classificação

| Parcelas em atraso | Nível | O que o e-mail diz |
|---|---|---|
| 0 | Em dia | Nenhuma ação |
| 1 | Atenção | Regularizar a parcela em atraso |
| 2 | Alerta | Regularizar com prioridade |
| 3 | Urgência | Pagar já: se a próxima parcela vencer sem pagamento, o parcelamento é rompido |
| 4 ou mais | Rompido | Verificar possibilidade de reparcelamento |
| (nenhuma faltando) | Quitado | Parcelamento quitado |

Os limites ficam na aba `config` da planilha e são **ilustrativos**: as regras reais de rescisão variam conforme a modalidade do parcelamento.

## Estrutura do repositório

```
monitor-parcelamentos/
├── monitor_parcelamentos.py        # programa principal
├── test_monitor_parcelamentos.py   # testes automatizados
├── gerar_dados.py                  # gera os dados fictícios
├── dados/
│   ├── parcelamentos.xlsx          # planilha de controle (fictícia)
│   └── gabarito.csv                # resultado esperado, usado nos testes
├── docs/                           # exemplo do e-mail gerado
└── requirements.txt
```

## A planilha

A aba `parcelamentos` tem 21 parcelamentos (7 federais, 7 estaduais e 7 municipais). Os dados em azul são de entrada; a coluna amarela (`parcelas_pagas`) é a única que a pessoa digita no dia a dia; o restante são fórmulas do Excel (`parcelas_faltantes`, `parcelas_vencidas`, `parcelas_em_atraso`, `nivel`, `proxima_parcela` e `mensagem_alerta`). A aba `legenda` explica cada coluna e as premissas.

## Como rodar

Requer Python 3.9 ou superior.

```
pip install -r requirements.txt
python monitor_parcelamentos.py
```

Por padrão o envio é **simulado**: o e-mail é salvo na pasta `saida/` (em `.html` e `.txt`) e nada é enviado. A data de referência vem da aba `config` (31/10/2026, fixa, para o resultado bater com o gabarito).

Outras opções:

```
python monitor_parcelamentos.py --hoje                 # usa a data de hoje
python monitor_parcelamentos.py --data 2026-11-15      # usa uma data específica
python monitor_parcelamentos.py --hoje --enviar        # envia por SMTP
```

### Envio real por e-mail

O envio lê as credenciais de variáveis de ambiente (nunca coloque senha no código):

```
SMTP_HOST=smtp.exemplo.com
SMTP_PORT=587
SMTP_USER=usuario@exemplo.com
SMTP_PASSWORD=sua-senha
EMAIL_DESTINATARIOS=gestor1@exemplo.com,gestor2@exemplo.com
```

### Agendamento diário (9h)

O programa não se agenda sozinho. Use o Agendador de Tarefas do Windows (gatilho diário às 9h) ou o cron no Linux/Mac:

```
0 9 * * 1-5 cd /caminho/do/projeto && python monitor_parcelamentos.py --hoje --enviar
```

## Testes

```
python -m unittest -v
```

Os testes conferem a contagem de parcelas vencidas, a classificação de cada nível e, principalmente, se o resultado dos 21 parcelamentos bate com o `gabarito.csv`.

## Dados fictícios

O `gerar_dados.py` recria a planilha e o gabarito (semente fixa: sempre os mesmos dados). Os dígitos verificadores dos CNPJs são válidos, mas os números foram sorteados; qualquer coincidência com uma empresa real é casual. Se regenerar a planilha, ela só terá os valores calculados das fórmulas depois de aberta e salva no Excel; o programa não depende disso, pois recalcula tudo em Python a partir dos dados de entrada.

## Limitações e próximos passos

- Os limites de alerta são ilustrativos e iguais nos três âmbitos.
- Parcelas vencem sempre no mesmo dia da primeira; feriados e prazos de carência não são tratados.
- A antecipação das primeiras parcelas (comum em parcelamentos federais) não foi modelada.
- O envio real depende de credenciais SMTP configuradas pelo usuário.
- Ideia futura: registrar a data de cada pagamento, em vez de só a quantidade paga, para detectar pagamentos feitos com atraso.

## Como usei IA neste projeto

Eu defini o problema, o modelo da planilha e as regras de alerta, vindos da rotina que automatizei na prática. O código foi escrito com apoio de IA (Claude) e validado por testes automatizados.
