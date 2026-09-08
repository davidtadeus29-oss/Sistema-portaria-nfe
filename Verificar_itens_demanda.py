import os
import re
from pathlib import Path
from datetime import datetime

import pandas as pd

# =========================================================
# CONFIGURAÇÃO
# =========================================================
PASTA_RELATORIOS = Path(r"C:\Users\OLIVDT\Desktop\Relatorios_transferencia")

# Informe aqui os códigos que deseja verificar.
# Pode deixar 750520 e 750521 como exemplo ou substituir pelos seus códigos.
CODIGOS_PARA_VERIFICAR = ["750520", "750521"]

# Nome sugerido do relatório de demanda.
# Se não encontrar, o programa exibirá os arquivos disponíveis.
NOME_DEMANDA = "DEMANDA CENTROS.XLSX"

# =========================================================
# FUNÇÕES
# =========================================================
def normalizar_codigo(valor):
    """
    Normaliza códigos somente para comparação.

    Exemplos:
        000000000000750520 -> 750520
        000000000000750521 -> 750521
        750520.0           -> 750520
        750521.0           -> 750521

    Importante: 750520 e 750521 continuam diferentes.
    """
    if pd.isna(valor):
        return ""

    texto = str(valor).strip().upper()

    # Remove apenas decimal formado por zeros.
    texto = re.sub(r"^(\d+)\.0+$", r"\1", texto)

    # Remove zeros à esquerda somente se o código for totalmente numérico.
    if texto.isdigit():
        texto = texto.lstrip("0") or "0"

    return texto


def listar_relatorios(pasta):
    extensoes = (".xlsx", ".xls", ".csv")
    arquivos = []

    for arquivo in pasta.iterdir():
        nome = arquivo.name

        # Ignora arquivos temporários criados pelo Excel, como ~$DEMANDA...
        if nome.startswith("~$"):
            continue

        if arquivo.is_file() and nome.lower().endswith(extensoes):
            arquivos.append(arquivo)

    return sorted(arquivos, key=lambda p: p.name.lower())


def localizar_demanda(pasta):
    caminho_exato = pasta / NOME_DEMANDA

    if caminho_exato.exists() and not caminho_exato.name.startswith("~$"):
        return caminho_exato

    candidatos = [
        arquivo for arquivo in listar_relatorios(pasta)
        if "demanda" in arquivo.name.lower()
    ]

    if len(candidatos) == 1:
        return candidatos[0]

    return None


def carregar_arquivo(caminho):
    try:
        if caminho.suffix.lower() == ".csv":
            return pd.read_csv(caminho)

        return pd.read_excel(caminho, sheet_name=0)

    except PermissionError:
        print(
            f"\nERRO: o arquivo está aberto ou bloqueado pelo Excel:\n{caminho}\n"
            "Feche o relatório e execute novamente."
        )
        raise SystemExit(1)

    except Exception as erro:
        print(f"\nERRO ao ler o arquivo {caminho.name}: {erro}")
        raise SystemExit(1)


def escolher_coluna_codigo(df):
    candidatos = [
        "PRODUTO",
        "Material",
        "Produto",
        "Item",
        "Código",
        "Codigo",
    ]

    for coluna in candidatos:
        if coluna in df.columns:
            return coluna

    print("\nNão foi encontrada automaticamente a coluna de código.")
    print("Colunas disponíveis:")
    for indice, coluna in enumerate(df.columns):
        print(f"{indice}: {coluna}")

    valor = input("Digite o nome exato da coluna de código: ").strip()

    if valor not in df.columns:
        print(f"Coluna não encontrada: {valor}")
        raise SystemExit(1)

    return valor


def escolher_coluna(df, candidatos, descricao):
    for coluna in candidatos:
        if coluna in df.columns:
            return coluna

    print(f"Aviso: coluna de {descricao} não encontrada.")
    return None


def converter_numero(serie):
    return pd.to_numeric(serie, errors="coerce").fillna(0)


# =========================================================
# EXECUÇÃO
# =========================================================
print("=" * 80)
print("VERIFICAÇÃO INDEPENDENTE DE ITENS - DEMANDA CENTROS")
print("=" * 80)

if not PASTA_RELATORIOS.exists():
    print(f"Pasta não encontrada:\n{PASTA_RELATORIOS}")
    raise SystemExit(1)

arquivo_demanda = localizar_demanda(PASTA_RELATORIOS)

if arquivo_demanda is None:
    print("\nNão foi possível localizar o relatório de demanda.")
    print("Arquivos encontrados:")
    for arquivo in listar_relatorios(PASTA_RELATORIOS):
        print(f"- {arquivo.name}")
    raise SystemExit(1)

print(f"\nArquivo analisado: {arquivo_demanda.name}")

# Normaliza a lista digitada pelo usuário.
codigos = {
    normalizar_codigo(codigo)
    for codigo in CODIGOS_PARA_VERIFICAR
    if normalizar_codigo(codigo)
}

print(f"Códigos pesquisados: {', '.join(sorted(codigos))}")

df = carregar_arquivo(arquivo_demanda)
print(f"Linhas carregadas: {len(df):,}")
print(f"Colunas carregadas: {len(df.columns)}")

col_codigo = escolher_coluna_codigo(df)
col_centro = escolher_coluna(
    df,
    ["Centro", "Centro Destino", "Local Exped."],
    "centro"
)
col_quantidade = escolher_coluna(
    df,
    ["Saldo Pendente", "Qtde OV", "Quantidade", "Qtd"],
    "quantidade"
)
col_data = escolher_coluna(
    df,
    ["Data Pedido", "Data do Pedido", "Dt Pedido", "Data"],
    "data do pedido"
)
col_valor = escolher_coluna(
    df,
    ["Valor líquido", "Valor liquido", "Valor Líquido", "Valor"],
    "valor líquido"
)

print(f"\nColuna de código usada: {col_codigo}")
print(f"Coluna de centro usada: {col_centro}")

# Cria coluna auxiliar sem alterar o arquivo original.
df["_Codigo_Normalizado"] = df[col_codigo].map(normalizar_codigo)

resultado = df[df["_Codigo_Normalizado"].isin(codigos)].copy()

if resultado.empty:
    print("\nNenhum dos códigos pesquisados foi encontrado.")
else:
    print(f"\nTotal de linhas encontradas: {len(resultado):,}")

    colunas_exibicao = [
        col_codigo,
        "_Codigo_Normalizado",
        col_centro,
        col_quantidade,
        col_data,
        col_valor,
    ]

    colunas_exibicao = [
        coluna for coluna in colunas_exibicao
        if coluna is not None and coluna in resultado.columns
    ]

    print("\nResumo por código e centro:")

    resumo = (
        resultado
        .groupby(["_Codigo_Normalizado", col_centro], as_index=False)
        .agg(
            Linhas=("_Codigo_Normalizado", "size"),
            Saldo_Pendente=(col_quantidade, lambda x: converter_numero(x).sum()),
            Valor_Liquido=(col_valor, lambda x: converter_numero(x).sum())
            if col_valor else ("_Codigo_Normalizado", "size")
        )
        .sort_values(["_Codigo_Normalizado", col_centro])
    )

    print(resumo.to_string(index=False))

    print("\nLinhas detalhadas:")
    print(resultado[colunas_exibicao].to_string(index=False))

    # Verificação explícita dos códigos esperados.
    print("\nVerificação individual:")
    for codigo in sorted(codigos):
        subconjunto = resultado[
            resultado["_Codigo_Normalizado"] == codigo
        ]

        centros = []
        if col_centro and not subconjunto.empty:
            centros = sorted(
                subconjunto[col_centro]
                .dropna()
                .astype(str)
                .str.strip()
                .unique()
                .tolist()
            )

        print(
            f"- {codigo}: {len(subconjunto)} linha(s); "
            f"centros: {', '.join(centros) if centros else 'nenhum'}"
        )

    # Exporta resultado para conferência humana.
    data_hora = datetime.now().strftime("%Y%m%d_%H%M%S")
    arquivo_saida = PASTA_RELATORIOS / f"verificacao_itens_{data_hora}.xlsx"

    with pd.ExcelWriter(arquivo_saida, engine="openpyxl") as writer:
        resultado.to_excel(writer, sheet_name="Linhas_Encontradas", index=False)
        resumo.to_excel(writer, sheet_name="Resumo_Codigo_Centro", index=False)

    print(f"\nRelatório de verificação criado: {arquivo_saida.name}")

print("\nVerificação concluída.")
input("Pressione ENTER para fechar...")
