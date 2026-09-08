import pandas as pd
import os
from datetime import datetime

# ==========================
# CONFIGURAÇÕES
# ==========================

ARQUIVO_ATUAL = r"C:\Users\OLIVDT\Desktop\Carteira\CARTEIRA.xlsx"
HISTORICO = r"C:\Users\OLIVDT\Desktop\Carteira\HISTORICO\historico_carteira.xlsx"

ABA = "Sheet1"

HOJE = datetime.now().strftime("%d/%m/%Y")

# ==========================
# LEITURA DA BASE ATUAL
# ==========================

df_atual = pd.read_excel(
    ARQUIVO_ATUAL,
    sheet_name=ABA
)

# F=Pedido G=Item I=Centro K=DataDesejada
df_atual = df_atual.iloc[:, [5, 6, 8, 10]]

df_atual.columns = [
    "Pedido",
    "Item",
    "Centro",
    "DataDesejada"
]

# Remove linhas vazias
df_atual = df_atual.dropna(subset=["Pedido"])

# Remove .0 dos números
df_atual["Pedido"] = pd.to_numeric(
    df_atual["Pedido"],
    errors="coerce"
).fillna(0).astype(int)

df_atual["Item"] = pd.to_numeric(
    df_atual["Item"],
    errors="coerce"
).fillna(0).astype(int)

# Padroniza data
df_atual["DataDesejada"] = pd.to_datetime(
    df_atual["DataDesejada"],
    errors="coerce"
).dt.strftime("%d/%m/%Y")

# Cria chave
df_atual["Chave"] = (
    df_atual["Pedido"].astype(str)
    + "|"
    + df_atual["Item"].astype(str)
    + "|"
    + df_atual["Centro"].astype(str)
)

# ==========================
# PRIMEIRA EXECUÇÃO
# ==========================

if not os.path.exists(HISTORICO):

    historico = df_atual.copy()

    historico["DataEntrada"] = HOJE
    historico["DataSaida"] = ""
    historico["Status"] = "ATIVO"

    historico.to_excel(
        HISTORICO,
        index=False
    )

    print("Histórico criado com sucesso.")
    quit()

# ==========================
# CARREGA HISTÓRICO
# ==========================

historico = pd.read_excel(
    HISTORICO,
    dtype=str
).fillna("")

# Padroniza datas do histórico
historico["DataDesejada"] = pd.to_datetime(
    historico["DataDesejada"],
    errors="coerce"
).dt.strftime("%d/%m/%Y")

ativos = historico[
    historico["Status"] == "ATIVO"
].copy()

# ==========================
# NOVOS
# ==========================

novos = df_atual[
    ~df_atual["Chave"].isin(
        ativos["Chave"]
    )
].copy()

# ==========================
# REMOVIDOS
# ==========================

removidos = ativos[
    ~ativos["Chave"].isin(
        df_atual["Chave"]
    )
].copy()

# ==========================
# ALTERAÇÕES
# ==========================

alteracoes = []

ativos_dict = ativos.set_index("Chave")

for _, linha in df_atual.iterrows():

    chave = linha["Chave"]

    if chave in ativos_dict.index:

        data_atual = str(linha["DataDesejada"])
        data_antiga = str(
            ativos_dict.loc[
                chave,
                "DataDesejada"
            ]
        )

        if data_atual != data_antiga:

            alteracoes.append({
                "Chave": chave,
                "DataAnterior": data_antiga,
                "DataAtual": data_atual
            })

            historico.loc[
                historico["Chave"] == chave,
                "DataDesejada"
            ] = data_atual

# ==========================
# MARCA REMOVIDOS
# ==========================

for chave in removidos["Chave"]:

    historico.loc[
        historico["Chave"] == chave,
        "Status"
    ] = "REMOVIDO"

    historico.loc[
        historico["Chave"] == chave,
        "DataSaida"
    ] = HOJE

# ==========================
# INSERE NOVOS
# ==========================

if len(novos) > 0:

    novos["DataEntrada"] = HOJE
    novos["DataSaida"] = ""
    novos["Status"] = "ATIVO"

    historico = pd.concat(
        [historico, novos],
        ignore_index=True
    )

# ==========================
# SALVA HISTÓRICO
# ==========================

historico.to_excel(
    HISTORICO,
    index=False
)

# ==========================
# RELATÓRIO
# ==========================

nome_relatorio = (
    f"Relatorio_{datetime.now():%Y%m%d}.xlsx"
)

with pd.ExcelWriter(
    nome_relatorio,
    engine="openpyxl"
) as writer:

    resumo = pd.DataFrame({
        "Evento": [
            "Novos",
            "Removidos",
            "Alterações",
            "Ativos"
        ],
        "Quantidade": [
            len(novos),
            len(removidos),
            len(alteracoes),
            len(
                historico[
                    historico["Status"] == "ATIVO"
                ]
            )
        ]
    })

    resumo.to_excel(
        writer,
        sheet_name="Resumo",
        index=False
    )

    novos.to_excel(
        writer,
        sheet_name="Novos",
        index=False
    )

    removidos.to_excel(
        writer,
        sheet_name="Removidos",
        index=False
    )

    pd.DataFrame(
        alteracoes
    ).to_excel(
        writer,
        sheet_name="Alteracoes",
        index=False
    )

print("")
print("================================")
print("ANÁLISE CONCLUÍDA")
print("================================")
print(f"Novos: {len(novos)}")
print(f"Removidos: {len(removidos)}")
print(f"Alterações: {len(alteracoes)}")
print(f"Ativos: {len(historico[historico['Status']=='ATIVO'])}")
print("================================")
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    PageBreak
)

from reportlab.lib.styles import getSampleStyleSheet

pdf = SimpleDocTemplate(
    f"Dashboard_{datetime.now():%Y%m%d}.pdf"
)

styles = getSampleStyleSheet()

elementos = []

elementos.append(
    Paragraph(
        "Dashboard de Carteira",
        styles["Title"]
    )
)

elementos.append(Spacer(1,20))

elementos.append(
    Paragraph(
        f"Novos: {len(novos)}",
        styles["Normal"]
    )
)

elementos.append(
    Paragraph(
        f"Removidos: {len(removidos)}",
        styles["Normal"]
    )
)

elementos.append(
    Paragraph(
        f"Alterações: {len(alteracoes)}",
        styles["Normal"]
    )
)

pdf.build(elementos)