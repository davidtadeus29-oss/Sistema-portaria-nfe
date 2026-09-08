import os
import io
import re
import json
import hashlib
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# =========================================================
# CONFIGURAÇÃO
# =========================================================
st.set_page_config(page_title="Painel de Transferências", layout="wide")

PASTA_RELATORIOS = r"C:\Users\OLIVDT\Desktop\Relatorios_transferencia"

if not os.path.exists(PASTA_RELATORIOS):
    st.error(f"Pasta não encontrada: {PASTA_RELATORIOS}")
    st.stop()

if "criterios" not in st.session_state:
    st.session_state.criterios = {
        "criterio_principal": "FIFO - data do pedido",
        "estoque_minimo": 500,
        "estoque_maximo": 10000,
        "centro_origem": None,
        "centro_destino": None,
        "depositos_origem": [],
        "codigos_avaliar": [],
        "esvaziar_deposito": False,
    }

if "modelo_hash" not in st.session_state:
    st.session_state.modelo_hash = None

if "resultados" not in st.session_state:
    st.session_state.resultados = {}

if "centro_clicado" not in st.session_state:
    st.session_state.centro_clicado = None

# =========================================================
# INTERFACE PADRÃO STREAMLIT
# =========================================================
st.title("Painel de Transferências")
st.caption("FIFO global • demanda líquida • saldo de origem e destino")

# =========================================================
# FUNÇÕES BÁSICAS
# =========================================================
def buscar_arquivos(pasta):
    """Lista somente relatórios reais; ignora arquivos temporários ~$ do Excel."""
    extensoes = (".xlsx", ".xls", ".csv")
    arquivos = []
    for nome in os.listdir(pasta):
        if nome.lower().startswith("~$"):
            continue
        if not nome.lower().endswith(extensoes):
            continue
        caminho = os.path.join(pasta, nome)
        if os.path.isfile(caminho):
            arquivos.append(nome)
    return sorted(arquivos)

def buscar_historicos_consumo(pasta):
    """
    Procura arquivos de histórico de consumo em:
    1) subpastas usuais (Historico_consumo / historico_consumo / HISTORICO_CONSUMO)
    2) pasta raiz de relatórios (arquivos contendo 'historico' e 'consumo' no nome)

    Ignora arquivos temporários do Excel (~$...).
    Retorna caminhos relativos à pasta base.
    """
    extensoes = (".xlsx", ".xls", ".csv")
    encontrados = set()

    # 1) Pastas candidatas (variações de caixa).
    candidatos = [
        os.path.join(pasta, "Historico_consumo"),
        os.path.join(pasta, "historico_consumo"),
        os.path.join(pasta, "HISTORICO_CONSUMO"),
    ]

    # Remove duplicados preservando ordem.
    vistos = set()
    pastas_busca = []
    for c in candidatos:
        k = os.path.normcase(os.path.normpath(c))
        if k not in vistos:
            vistos.add(k)
            pastas_busca.append(c)

    for raiz in pastas_busca:
        if not os.path.isdir(raiz):
            continue
        for atual, _, nomes in os.walk(raiz):
            for nome in nomes:
                if nome.lower().startswith("~$"):
                    continue
                if not nome.lower().endswith(extensoes):
                    continue
                caminho = os.path.join(atual, nome)
                if os.path.isfile(caminho):
                    encontrados.add(os.path.relpath(caminho, pasta))

    # 2) Fallback: arquivos na raiz com padrão de nome do histórico.
    for nome in os.listdir(pasta):
        caminho = os.path.join(pasta, nome)
        nome_l = nome.lower()
        if not os.path.isfile(caminho):
            continue
        if nome_l.startswith("~$"):
            continue
        if not nome_l.endswith(extensoes):
            continue
        if ("historico" in nome_l and "consumo" in nome_l) or ("hist" in nome_l and "cons" in nome_l):
            encontrados.add(nome)

    return sorted(encontrados)


def historico_mais_recente(pasta, arquivos_relativos):
    """Retorna o histórico mais recentemente modificado (caminho relativo)."""
    candidatos = []
    for rel in arquivos_relativos:
        caminho = os.path.join(pasta, rel)
        if os.path.isfile(caminho):
            candidatos.append((os.path.getmtime(caminho), rel))
    if not candidatos:
        return None
    candidatos.sort(reverse=True)
    return candidatos[0][1]

def assinatura_arquivo(caminho):
    """Retorna uma assinatura que muda quando o relatório é alterado."""
    try:
        info = os.stat(caminho)
        return (info.st_mtime_ns, info.st_size)
    except OSError:
        return (None, None)


def data_hora_modificacao(caminho):
    """Retorna datetime da última modificação do arquivo, ou None."""
    try:
        return datetime.fromtimestamp(os.path.getmtime(caminho))
    except OSError:
        return None


def fmt_data_hora(dt):
    if dt is None:
        return "N/A"
    return dt.strftime("%d/%m/%Y %H:%M:%S")

@st.cache_data(show_spinner=False)
def carregar_arquivo(caminho, assinatura):
    """
    Carrega o relatório usando data de alteração e tamanho como parte da chave
    de cache. Assim, substituir o Excel no mesmo caminho força uma nova leitura.
    """
    try:
        if caminho.lower().endswith(".csv"):
            return pd.read_csv(caminho)
        return pd.read_excel(caminho, sheet_name=0)
    except PermissionError as erro:
        nome = os.path.basename(caminho)
        raise PermissionError(
            f"O arquivo '{nome}' está aberto ou bloqueado pelo Excel. "
            "Feche o arquivo e atualize o app."
        ) from erro
    except Exception as erro:
        nome = os.path.basename(caminho)
        raise RuntimeError(
            f"Não foi possível ler o arquivo '{nome}': {erro}"
        ) from erro

def num(df, col):
    return pd.to_numeric(df[col], errors="coerce").fillna(0)

def normalizar_codigo(valor):
    """
    Normaliza código sem alterar identidade do material:
    - remove sufixo decimal .0/.00 (ex.: 750520.0 -> 750520)
    - remove zeros à esquerda APENAS quando o código é totalmente numérico
      (ex.: 000000750520 -> 750520)
    - mantém códigos alfanuméricos como texto (apenas strip/upper)
    """
    if pd.isna(valor):
        return ""

    texto = str(valor).strip()
    if not texto:
        return ""

    texto = re.sub(r"^(\d+)\.0+$", r"\1", texto)

    if re.fullmatch(r"\d+", texto):
        texto = texto.lstrip("0") or "0"

    return texto.upper()

def normalizar_serie_codigo(serie):
    return serie.map(normalizar_codigo)

def parse_codigos(texto):
    if not texto or not texto.strip():
        return []
    return [x.strip() for x in re.split(r"[,;\n\t ]+", texto.strip()) if x.strip()]

def escolher_coluna(label, cols, sugestoes):
    lower = [str(c).lower() for c in cols]
    idx = 0
    for s in sugestoes:
        if s.lower() in lower:
            idx = lower.index(s.lower())
            break
    return st.selectbox(label, cols, index=idx)


def padronizar_historico_consumo(df_historico_bruto):
    """
    Converte o histórico para formato longo padronizado:
    Codigo | Centro_Destino | Consumo_Historico

    Suporta:
    - formato longo (Código/Centro/Consumo)
    - formato matriz/pivô (centros em colunas ES01, ES02, ...)
    - cabeçalho interno na primeira linha de dados (colunas Unnamed)
    """
    if df_historico_bruto is None or df_historico_bruto.empty:
        return pd.DataFrame(columns=["Codigo", "Centro_Destino", "Consumo_Historico"])

    df = df_historico_bruto.copy().dropna(axis=1, how="all")
    if df.empty:
        return pd.DataFrame(columns=["Codigo", "Centro_Destino", "Consumo_Historico"])

    # Detecta cabeçalho interno (ex.: linha com "Código", "ES01", ...)
    header_idx = None
    for i in range(min(20, len(df))):
        vals = [str(v).strip() for v in df.iloc[i].tolist() if pd.notna(v)]
        vals_up = [v.upper() for v in vals]
        has_codigo = any(v in ["CÓDIGO", "CODIGO", "MATERIAL", "PRODUTO", "ITEM"] for v in vals_up)
        has_centro = any(re.fullmatch(r"ES\d{2}|DPEX", v) for v in vals_up)
        if has_codigo and has_centro:
            header_idx = i
            break

    if header_idx is not None:
        headers = [str(x).strip() if pd.notna(x) else "" for x in df.iloc[header_idx].tolist()]
        df = df.iloc[header_idx + 1:].copy()
        df.columns = headers
        df = df[[c for c in df.columns if str(c).strip() not in ["", "nan", "None"]]]

    cols = [str(c).strip() for c in df.columns]
    cols_lower = {c.lower(): c for c in cols}

    cand_cod = ["código", "codigo", "material", "produto", "item"]
    cand_centro = ["centro", "centro destino", "local", "centro_consumidor"]
    cand_cons = ["consumo", "quantidade", "qtd", "qtde", "quantidade consumida", "consumo_historico"]

    col_codigo = next((cols_lower[k] for k in cand_cod if k in cols_lower), None)
    col_centro = next((cols_lower[k] for k in cand_centro if k in cols_lower), None)
    col_consumo = next((cols_lower[k] for k in cand_cons if k in cols_lower), None)

    # Formato longo
    if col_codigo and col_centro and col_consumo:
        out = pd.DataFrame({
            "Codigo": df[col_codigo].map(normalizar_codigo),
            "Centro_Destino": df[col_centro].astype(str).str.strip().str.upper(),
            "Consumo_Historico": pd.to_numeric(df[col_consumo], errors="coerce").fillna(0),
        })
        out = out[(out["Codigo"] != "") & (out["Centro_Destino"] != "") & (out["Consumo_Historico"] > 0)]
        return out

    # Formato matriz/pivô
    col_codigo = None
    for c in df.columns:
        if str(c).strip().upper() in ["CÓDIGO", "CODIGO", "MATERIAL", "PRODUTO", "ITEM"]:
            col_codigo = c
            break
    if col_codigo is None:
        col_codigo = df.columns[0]

    centros_cols = [c for c in df.columns if re.fullmatch(r"ES\d{2}|DPEX", str(c).strip().upper())]
    if not centros_cols:
        return pd.DataFrame(columns=["Codigo", "Centro_Destino", "Consumo_Historico"])

    base = df[[col_codigo] + centros_cols].copy()
    base[col_codigo] = base[col_codigo].map(normalizar_codigo)

    out = base.melt(
        id_vars=[col_codigo],
        value_vars=centros_cols,
        var_name="Centro_Destino",
        value_name="Consumo_Historico",
    ).rename(columns={col_codigo: "Codigo"})

    out["Centro_Destino"] = out["Centro_Destino"].astype(str).str.strip().str.upper()
    out["Consumo_Historico"] = pd.to_numeric(out["Consumo_Historico"], errors="coerce").fillna(0)
    out = out[(out["Codigo"] != "") & (out["Centro_Destino"] != "") & (out["Consumo_Historico"] > 0)]
    return out[["Codigo", "Centro_Destino", "Consumo_Historico"]]


def indices_por_nome(arquivos):
    i_est = i_dem = i_sep = 0
    for i, nome in enumerate(arquivos):
        n = nome.lower()
        if "estoque" in n:
            i_est = i
        elif "demanda" in n:
            i_dem = i
        elif "separa" in n or "picking" in n:
            i_sep = i
    return i_est, i_dem, i_sep

def hash_modelo(obj):
    return hashlib.md5(json.dumps(obj, sort_keys=True, default=str).encode()).hexdigest()

def destacar(df, coluna):
    if df is None or df.empty or coluna not in df.columns or len(df) > 3000:
        return df
    def estilo(v):
        try:
            return "background-color:#C6EFCE;color:#006100;font-weight:bold;" if float(v) > 0 else ""
        except Exception:
            return ""
    styler = df.style
    return styler.map(estilo, subset=[coluna]) if hasattr(styler, "map") else styler.applymap(estilo, subset=[coluna])

def filtros(df, prefixo, habilitado):
    if df.empty or not habilitado:
        return df
    out = df.copy()
    with st.expander("🔎 Filtros por coluna", expanded=False):
        for col in out.columns:
            valor = st.text_input(str(col), key=f"{prefixo}_{col}")
            if valor:
                out = out[out[col].astype(str).str.contains(valor, case=False, na=False)]
    return out

# =========================================================
# CÁLCULOS DE SALDO
# =========================================================
def saldo_origem(df_est, df_dem, df_sep, origem, depositos, codigos,
                  ecod, ecentro, edep, eqtd, dcod, dcentro, dqtd,
                  scod, scentro, sqtd, estoque_minimo):
    est = df_est[df_est[ecentro].astype(str) == str(origem)].copy()
    if depositos:
        est = est[est[edep].astype(str).isin([str(x) for x in depositos])]
    dem = df_dem[df_dem[dcentro].astype(str) == str(origem)].copy()
    sep = df_sep[df_sep[scentro].astype(str) == str(origem)].copy()
    if codigos:
        est = est[normalizar_serie_codigo(est[ecod]).isin(codigos)]
        dem = dem[normalizar_serie_codigo(dem[dcod]).isin(codigos)]
        sep = sep[normalizar_serie_codigo(sep[scod]).isin(codigos)]

    a = est.groupby(normalizar_serie_codigo(est[ecod]))[eqtd].apply(lambda x: pd.to_numeric(x, errors="coerce").fillna(0).sum()).rename("Estoque_Origem")
    b = sep.groupby(normalizar_serie_codigo(sep[scod]))[sqtd].apply(lambda x: pd.to_numeric(x, errors="coerce").fillna(0).sum()).rename("Em_Separacao_Origem")
    c = dem.groupby(normalizar_serie_codigo(dem[dcod]))[dqtd].apply(lambda x: pd.to_numeric(x, errors="coerce").fillna(0).sum()).rename("Demanda_Origem")
    base = pd.concat([a, b, c], axis=1).fillna(0).reset_index()
    base.columns = ["Codigo", "Estoque_Origem", "Em_Separacao_Origem", "Demanda_Origem"]
    base["Estoque_Minimo_Origem"] = estoque_minimo
    base["Saldo_Transferivel_Origem"] = (base["Estoque_Origem"] - base["Em_Separacao_Origem"] - base["Demanda_Origem"] - base["Estoque_Minimo_Origem"]).clip(lower=0)
    return base

def saldo_destinos(df_est, df_sep, origem, codigos, ecod, ecentro, eqtd, scod, scentro, sqtd):
    est = df_est[df_est[ecentro].astype(str) != str(origem)].copy()
    sep = df_sep[df_sep[scentro].astype(str) != str(origem)].copy()
    if codigos:
        est = est[normalizar_serie_codigo(est[ecod]).isin(codigos)]
        sep = sep[normalizar_serie_codigo(sep[scod]).isin(codigos)]

    a = est.groupby([normalizar_serie_codigo(est[ecod]), est[ecentro].astype(str)])[eqtd].apply(lambda x: pd.to_numeric(x, errors="coerce").fillna(0).sum()).rename("Estoque_Destino").reset_index()
    a.columns = ["Codigo", "Centro", "Estoque_Destino"]
    b = sep.groupby([normalizar_serie_codigo(sep[scod]), sep[scentro].astype(str)])[sqtd].apply(lambda x: pd.to_numeric(x, errors="coerce").fillna(0).sum()).rename("Em_Separacao_Destino").reset_index()
    b.columns = ["Codigo", "Centro", "Em_Separacao_Destino"]
    out = a.merge(b, on=["Codigo", "Centro"], how="outer").fillna(0)
    out["Saldo_Destino_Disponivel"] = (out["Estoque_Destino"] - out["Em_Separacao_Destino"]).clip(lower=0)
    return out

def demanda_liquida(df_dem, saldo_dest, origem, codigos, dcod, dcentro, dqtd, ddata, dvalor):
    dem = df_dem[df_dem[dcentro].astype(str) != str(origem)].copy()
    if codigos:
        dem = dem[normalizar_serie_codigo(dem[dcod]).isin(codigos)]
    dem = dem[[dcod, dcentro, dqtd, ddata, dvalor]].copy()
    dem.columns = ["Codigo", "Centro", "Demanda_Linha", "Data_Pedido", "Valor_Linha"]
    dem["Codigo"] = normalizar_serie_codigo(dem["Codigo"])
    dem["Centro"] = dem["Centro"].astype(str)
    dem["Linha_ID"] = np.arange(len(dem))
    dem["Demanda_Linha"] = pd.to_numeric(dem["Demanda_Linha"], errors="coerce").fillna(0)
    dem["Valor_Linha"] = pd.to_numeric(dem["Valor_Linha"], errors="coerce").fillna(0)
    dem["Valor_Unitario"] = np.where(dem["Demanda_Linha"] > 0, dem["Valor_Linha"] / dem["Demanda_Linha"], 0)
    dem["Data_Pedido"] = pd.to_datetime(dem["Data_Pedido"], errors="coerce")
    dem = dem.sort_values(["Codigo", "Centro", "Data_Pedido"])

    saldo_map = {(str(r.Codigo), str(r.Centro)): float(r.Saldo_Destino_Disponivel) for r in saldo_dest.itertuples()}
    linhas = []
    for (codigo, centro), grupo in dem.groupby(["Codigo", "Centro"], dropna=False):
        saldo = saldo_map.get((str(codigo), str(centro)), 0)
        for _, r in grupo.iterrows():
            original = float(r["Demanda_Linha"])
            usado = min(original, saldo)
            saldo -= usado
            linhas.append({
                "Linha_ID": int(r["Linha_ID"]),
                "Codigo": str(codigo), "Centro": str(centro), "Data_Pedido": r["Data_Pedido"],
                "Demanda_Original_Linha": original, "Saldo_Destino_Usado": usado,
                "Demanda_Liquida_Linha": max(original - usado, 0),
                "Valor_Linha_Original": float(r["Valor_Linha"]), "Valor_Unitario": float(r["Valor_Unitario"]),
                "Saldo_Destino_Remanescente": max(saldo, 0)
            })
    return pd.DataFrame(linhas)

def aplicar_fifo(dem_liquida, saldo_map):
    if dem_liquida.empty:
        return pd.DataFrame()
    dem_liquida = dem_liquida.sort_values(["Codigo", "Data_Pedido", "Centro"])
    linhas = []
    for codigo, grupo in dem_liquida.groupby("Codigo", dropna=False):
        saldo = float(saldo_map.get(str(codigo), 0))
        for _, r in grupo.iterrows():
            demanda = float(r["Demanda_Liquida_Linha"])
            transferida = min(demanda, saldo) if saldo > 0 else 0
            saldo -= transferida
            linhas.append({
                "Linha_ID": int(r["Linha_ID"]),
                "Codigo": str(codigo), "Centro": str(r["Centro"]), "Data_Pedido": r["Data_Pedido"],
                "Demanda_Original_Linha": float(r["Demanda_Original_Linha"]),
                "Saldo_Destino_Usado": float(r["Saldo_Destino_Usado"]),
                "Demanda_Liquida_Linha": demanda, "Qtd_Transferida_FIFO": transferida,
                "Backlog_Linha": max(demanda - transferida, 0),
                "Valor_Transferido_Linha": transferida * float(r["Valor_Unitario"]),
                "Saldo_Origem_Remanescente": max(saldo, 0)
            })
    return pd.DataFrame(linhas)

def resumo_fifo(fifo):
    if fifo.empty:
        return pd.DataFrame()
    out = fifo.groupby(["Codigo", "Centro"], as_index=False).agg(
        Data_FIFO_Mais_Antiga=("Data_Pedido", "min"),
        Demanda_Original_Total=("Demanda_Original_Linha", "sum"),
        Saldo_Destino_Usado_Total=("Saldo_Destino_Usado", "sum"),
        Demanda_Liquida_Total=("Demanda_Liquida_Linha", "sum"),
        Qtd_Transferida_FIFO=("Qtd_Transferida_FIFO", "sum"),
        Backlog_Total=("Backlog_Linha", "sum"),
        Valor_Transferido_Centro=("Valor_Transferido_Linha", "sum")
    )
    out = out.sort_values(["Codigo", "Data_FIFO_Mais_Antiga", "Centro"])
    # Datas inválidas ou vazias ficam no final do FIFO.
    # O preenchimento evita IntCastingNaNError ao converter a ordem para inteiro.
    data_fifo = pd.to_datetime(
        out["Data_FIFO_Mais_Antiga"],
        errors="coerce"
    )
    data_fifo_para_rank = data_fifo.fillna(pd.Timestamp.max)

    out["Ordem_FIFO"] = (
        data_fifo_para_rank
        .groupby(out["Codigo"])
        .rank(method="dense", ascending=True)
        .fillna(0)
        .astype(int)
    )
    return out

def top5_centro(fifo, centro, origem):
    if fifo.empty:
        return pd.DataFrame()
    d = fifo[fifo["Centro"].astype(str) == str(centro)]
    if d.empty:
        return pd.DataFrame()
    out = d.groupby("Codigo", as_index=False).agg(
        Demanda_Liquida=("Demanda_Liquida_Linha", "sum"),
        Qtd_Transferida=("Qtd_Transferida_FIFO", "sum"),
        Backlog=("Backlog_Linha", "sum"),
        Valor_Transferido=("Valor_Transferido_Linha", "sum")
    ).sort_values("Valor_Transferido", ascending=False).head(5)
    out["Centro_Origem"] = origem
    out["Centro_Destino"] = centro
    return out

def construir_oportunidades(fifo, centro_origem_fallback=""):
    """
    Consolida as oportunidades por Centro Origem + Centro Destino + Código.
    A origem não é perdida no agrupamento, especialmente no cenário geral.
    """
    colunas = [
        "Centro_Origem", "Centro_Destino", "Codigo",
        "Data_Pedido_Mais_Antiga", "Demanda_Liquida",
        "Qtd_Transferida", "Backlog", "Valor_Transferido"
    ]

    if fifo is None or fifo.empty:
        return pd.DataFrame(columns=colunas)

    dados = fifo.copy()

    if "Centro_Origem" not in dados.columns:
        dados["Centro_Origem"] = str(centro_origem_fallback)
    else:
        dados["Centro_Origem"] = dados["Centro_Origem"].fillna(
            centro_origem_fallback
        ).astype(str)
        vazias = dados["Centro_Origem"].isin(["", "nan", "None"])
        dados.loc[vazias, "Centro_Origem"] = str(centro_origem_fallback)

    dados["Centro"] = dados["Centro"].astype(str)
    dados["Codigo"] = normalizar_serie_codigo(dados["Codigo"])

    agrupadores = ["Centro_Origem", "Centro", "Codigo"]
    out = (
        dados.groupby(agrupadores, as_index=False)
        .agg(
            Data_Pedido_Mais_Antiga=("Data_Pedido", "min"),
            Demanda_Liquida=("Demanda_Liquida_Linha", "sum"),
            Qtd_Transferida=("Qtd_Transferida_FIFO", "sum"),
            Backlog=("Backlog_Linha", "sum"),
            Valor_Transferido=("Valor_Transferido_Linha", "sum"),
        )
        .rename(columns={"Centro": "Centro_Destino"})
    )

    out = out[out["Qtd_Transferida"] > 0].copy()

    if out.empty:
        return pd.DataFrame(columns=colunas)

    return out[colunas].sort_values(
        ["Centro_Origem", "Valor_Transferido", "Qtd_Transferida"],
        ascending=[True, False, False]
    ).reset_index(drop=True)


def top5_oportunidades_por_centro(oportunidades, centro=None):
    """Top 5 geral ou Top 5 do centro selecionado."""
    if oportunidades is None or oportunidades.empty:
        return pd.DataFrame(columns=oportunidades.columns if oportunidades is not None else [])

    dados = oportunidades.copy()
    if centro:
        dados = dados[dados["Centro_Destino"].astype(str) == str(centro)].copy()

    return dados.sort_values(
        ["Valor_Transferido", "Qtd_Transferida"],
        ascending=[False, False]
    ).head(5)

def escrever_aba_excel(writer, dados, nome_aba):
    """
    Escreve uma aba mesmo quando o DataFrame está vazio.
    Isso evita gerar um workbook sem nenhuma aba visível.
    """
    if dados is None:
        dados = pd.DataFrame({"Informacao": ["Sem dados para este cenário."]})
    elif not isinstance(dados, pd.DataFrame):
        dados = pd.DataFrame(dados)

    # Garante pelo menos uma coluna em DataFrames vazios sem colunas.
    if dados.empty and len(dados.columns) == 0:
        dados = pd.DataFrame({"Informacao": ["Sem dados para este cenário."]})

    dados.to_excel(writer, sheet_name=nome_aba[:31], index=False)

def montar_descricoes(df_estoque, df_demanda, col_est_codigo, col_dem_codigo):
    """Monta uma descrição por código usando as colunas disponíveis."""
    mapa = {}
    candidatos = [
        (df_estoque, col_est_codigo, ["Texto breve material", "Descrição", "Descricao", "Texto"]),
        (df_demanda, col_dem_codigo, ["Texto item da ordem", "Descrição", "Descricao", "Texto"]),
    ]
    for df, col_codigo, possiveis in candidatos:
        if df is None or col_codigo not in df.columns:
            continue
        col_desc = next((c for c in possiveis if c in df.columns), None)
        if col_desc is None:
            continue
        for _, row in df[[col_codigo, col_desc]].dropna(subset=[col_codigo]).iterrows():
            codigo = normalizar_codigo(row[col_codigo])
            descricao = str(row[col_desc]).strip() if pd.notna(row[col_desc]) else ""
            if codigo and descricao and codigo not in mapa:
                mapa[codigo] = descricao
    return mapa


def obter_serie_coluna(df, coluna):
    """
    Retorna sempre uma Series, mesmo quando o arquivo possui cabeçalhos
    duplicados. Quando há duas colunas com o mesmo nome, utiliza a primeira.
    """
    selecionada = df.loc[:, coluna]

    if isinstance(selecionada, pd.DataFrame):
        return selecionada.iloc[:, 0]

    return selecionada

def distribuir_saldo_restante_por_historico(base_origem, df_historico, oportunidades, saldo_dest, origem, codigos,
                                             col_hist_codigo, col_hist_centro, col_hist_consumo):
    """
    MODO ESVAZIAR DEPÓSITO (checkbox habilitado):
    - ignora a lógica de saldo/demanda do destino para a alocação extra;
    - usa centros com histórico de consumo (preferencialmente do próprio material);
    - distribui todo o saldo restante buscando equalizar estoque entre centros.

    Se um material não existir no histórico por código, aplica fallback para
    os centros do histórico global (todos os materiais), garantindo esvaziamento.
    """
    colunas = [
        "Centro_Origem", "Centro_Destino", "Codigo",
        "Data_Pedido_Mais_Antiga", "Demanda_Liquida",
        "Qtd_Transferida", "Backlog", "Valor_Transferido"
    ]

    if base_origem is None or base_origem.empty:
        return pd.DataFrame(columns=colunas)

    if df_historico is None or df_historico.empty:
        return pd.DataFrame(columns=colunas)

    hist = pd.DataFrame({
        "Codigo": obter_serie_coluna(df_historico, col_hist_codigo).map(normalizar_codigo),
        "Centro_Destino": obter_serie_coluna(df_historico, col_hist_centro).astype(str).str.strip(),
        "Consumo_Historico": pd.to_numeric(
            obter_serie_coluna(df_historico, col_hist_consumo),
            errors="coerce"
        ).fillna(0),
    })

    hist = hist[
        (hist["Centro_Destino"] != str(origem))
        & (hist["Consumo_Historico"] > 0)
    ]

    if codigos:
        hist = hist[hist["Codigo"].isin(codigos)]

    if hist.empty:
        return pd.DataFrame(columns=colunas)

    hist = hist.groupby(["Codigo", "Centro_Destino"], as_index=False)["Consumo_Historico"].sum()
    hist_global = hist.groupby(["Centro_Destino"], as_index=False)["Consumo_Historico"].sum()

    base_disp = base_origem.copy()
    base_disp["Codigo"] = normalizar_serie_codigo(base_disp["Codigo"])
    if codigos:
        base_disp = base_disp[base_disp["Codigo"].isin(codigos)]

    disponivel_por_codigo = (
        base_disp.groupby("Codigo")["Saldo_Transferivel_Origem"]
        .sum()
        .astype(float)
    )

    if oportunidades is not None and not oportunidades.empty:
        op = oportunidades.copy()
        op["Codigo"] = normalizar_serie_codigo(op["Codigo"])
        ja_transferido_codigo = op.groupby("Codigo")["Qtd_Transferida"].sum().astype(float)
        ja_transferido_centro = (
            op.groupby(["Codigo", "Centro_Destino"])["Qtd_Transferida"]
            .sum()
            .astype(float)
            .reset_index(name="Ja_Transferido_Centro")
        )
    else:
        ja_transferido_codigo = pd.Series(dtype=float)
        ja_transferido_centro = pd.DataFrame(columns=["Codigo", "Centro_Destino", "Ja_Transferido_Centro"])

    if saldo_dest is not None and not saldo_dest.empty:
        dest = saldo_dest.copy()
        dest["Codigo"] = normalizar_serie_codigo(dest["Codigo"])
        dest["Centro"] = dest["Centro"].astype(str).str.strip()
        estoque_destino = (
            dest.groupby(["Codigo", "Centro"])["Saldo_Destino_Disponivel"]
            .sum()
            .astype(float)
            .reset_index(name="Estoque_Destino_Atual")
            .rename(columns={"Centro": "Centro_Destino"})
        )
    else:
        estoque_destino = pd.DataFrame(columns=["Codigo", "Centro_Destino", "Estoque_Destino_Atual"])

    def alocar_equalizando(estoques_base, saldo):
        arr = np.array(estoques_base, dtype=float)
        n = len(arr)
        if n == 0 or saldo <= 0:
            return np.zeros(n, dtype=float)

        ordem = np.argsort(arr, kind="mergesort")
        s = arr[ordem].copy()
        a = np.zeros(n, dtype=float)
        restante = float(saldo)

        for i in range(n - 1):
            gap = float(s[i + 1] - s[i])
            if gap <= 0:
                continue
            custo = gap * (i + 1)
            if restante >= custo:
                a[:i + 1] += gap
                s[:i + 1] += gap
                restante -= custo
            else:
                inc = restante / (i + 1)
                a[:i + 1] += inc
                restante = 0.0
                break

        if restante > 0:
            a += (restante / n)

        out = np.zeros(n, dtype=float)
        out[ordem] = a
        return out

    linhas = []

    for codigo in disponivel_por_codigo.index.tolist():
        disponivel_total = float(disponivel_por_codigo.get(codigo, 0))
        transferido_total = float(ja_transferido_codigo.get(codigo, 0))
        saldo_restante = max(disponivel_total - transferido_total, 0)

        if saldo_restante <= 0:
            continue

        grupo = hist[hist["Codigo"].astype(str) == str(codigo)].copy()

        # Fallback: sem histórico por código -> usa centros do histórico global.
        if grupo.empty:
            grupo = hist_global.copy()
            grupo["Codigo"] = str(codigo)
            grupo = grupo[["Codigo", "Centro_Destino", "Consumo_Historico"]]

        if grupo.empty:
            continue

        grupo = grupo.merge(
            estoque_destino[estoque_destino["Codigo"].astype(str) == str(codigo)][["Centro_Destino", "Estoque_Destino_Atual"]],
            on="Centro_Destino",
            how="left"
        )
        grupo = grupo.merge(
            ja_transferido_centro[ja_transferido_centro["Codigo"].astype(str) == str(codigo)][["Centro_Destino", "Ja_Transferido_Centro"]],
            on="Centro_Destino",
            how="left"
        )

        grupo["Estoque_Destino_Atual"] = pd.to_numeric(grupo["Estoque_Destino_Atual"], errors="coerce").fillna(0)
        grupo["Ja_Transferido_Centro"] = pd.to_numeric(grupo["Ja_Transferido_Centro"], errors="coerce").fillna(0)
        grupo["Estoque_Base"] = grupo["Estoque_Destino_Atual"] + grupo["Ja_Transferido_Centro"]

        aloc = alocar_equalizando(grupo["Estoque_Base"].tolist(), saldo_restante)
        grupo["Qtd_Transferida"] = aloc

        residual = float(saldo_restante - grupo["Qtd_Transferida"].sum())
        if residual > 1e-9:
            soma_hist = float(grupo["Consumo_Historico"].sum())
            if soma_hist > 0:
                grupo["Qtd_Transferida"] += residual * (grupo["Consumo_Historico"] / soma_hist)

        for _, linha in grupo.iterrows():
            qtd = float(linha["Qtd_Transferida"])
            if qtd <= 0:
                continue
            linhas.append({
                "Centro_Origem": str(origem),
                "Centro_Destino": str(linha["Centro_Destino"]),
                "Codigo": str(codigo),
                "Data_Pedido_Mais_Antiga": pd.NaT,
                "Demanda_Liquida": qtd,
                "Qtd_Transferida": qtd,
                "Backlog": 0.0,
                "Valor_Transferido": 0.0,
            })

    return pd.DataFrame(linhas, columns=colunas)


def reconciliar_esvaziamento_total(oportunidades, base_origem, df_historico, origem,
                                   col_hist_codigo, col_hist_centro, col_hist_consumo,
                                   tolerancia=1e-6):
    """
    Garantia final do modo 'esvaziar depósito':
    para cada código da origem, força total transferido == saldo disponível,
    distribuindo eventual diferença positiva entre centros com histórico.
    """
    if oportunidades is None:
        oportunidades = pd.DataFrame(columns=[
            "Centro_Origem", "Centro_Destino", "Codigo", "Data_Pedido_Mais_Antiga",
            "Demanda_Liquida", "Qtd_Transferida", "Backlog", "Valor_Transferido"
        ])
    if base_origem is None or base_origem.empty:
        return oportunidades
    if df_historico is None or df_historico.empty:
        return oportunidades

    op = oportunidades.copy()
    if op.empty:
        op = pd.DataFrame(columns=[
            "Centro_Origem", "Centro_Destino", "Codigo", "Data_Pedido_Mais_Antiga",
            "Demanda_Liquida", "Qtd_Transferida", "Backlog", "Valor_Transferido"
        ])

    op["Codigo"] = normalizar_serie_codigo(op.get("Codigo", pd.Series(dtype=str)))
    op["Centro_Destino"] = op.get("Centro_Destino", pd.Series(dtype=str)).astype(str)

    hist = pd.DataFrame({
        "Codigo": obter_serie_coluna(df_historico, col_hist_codigo).map(normalizar_codigo),
        "Centro_Destino": obter_serie_coluna(df_historico, col_hist_centro).astype(str).str.strip(),
        "Consumo_Historico": pd.to_numeric(obter_serie_coluna(df_historico, col_hist_consumo), errors="coerce").fillna(0),
    })
    hist = hist[(hist["Centro_Destino"] != str(origem)) & (hist["Consumo_Historico"] > 0)]
    if hist.empty:
        return op

    hist_cod = hist.groupby(["Codigo", "Centro_Destino"], as_index=False)["Consumo_Historico"].sum()
    hist_global = hist.groupby(["Centro_Destino"], as_index=False)["Consumo_Historico"].sum()

    base = base_origem.copy()
    base["Codigo"] = normalizar_serie_codigo(base["Codigo"])
    saldo_map = base.groupby("Codigo")["Saldo_Transferivel_Origem"].sum().astype(float).to_dict()

    linhas_extra = []
    for codigo, saldo in saldo_map.items():
        saldo = float(saldo)
        if saldo <= tolerancia:
            continue

        atual = float(op.loc[op["Codigo"].astype(str) == str(codigo), "Qtd_Transferida"].sum()) if not op.empty else 0.0
        diff = saldo - atual
        if diff <= tolerancia:
            continue

        elegiveis = hist_cod[hist_cod["Codigo"].astype(str) == str(codigo)]["Centro_Destino"].astype(str).tolist()
        if not elegiveis:
            elegiveis = hist_global["Centro_Destino"].astype(str).tolist()
        elegiveis = [c for c in elegiveis if str(c) != str(origem)]
        if not elegiveis:
            continue

        # centro com menor transferência atual para o código (reduz concentração)
        atual_centro = (
            op[op["Codigo"].astype(str) == str(codigo)]
            .groupby("Centro_Destino")["Qtd_Transferida"].sum()
            .to_dict()
        ) if not op.empty else {}

        centro_escolhido = sorted(elegiveis, key=lambda c: float(atual_centro.get(c, 0.0)))[0]

        linhas_extra.append({
            "Centro_Origem": str(origem),
            "Centro_Destino": str(centro_escolhido),
            "Codigo": str(codigo),
            "Data_Pedido_Mais_Antiga": pd.NaT,
            "Demanda_Liquida": float(diff),
            "Qtd_Transferida": float(diff),
            "Backlog": 0.0,
            "Valor_Transferido": 0.0,
        })

    if linhas_extra:
        op = pd.concat([op, pd.DataFrame(linhas_extra)], ignore_index=True)
        op = op.groupby(["Centro_Origem", "Centro_Destino", "Codigo"], as_index=False).agg(
            Data_Pedido_Mais_Antiga=("Data_Pedido_Mais_Antiga", "min"),
            Demanda_Liquida=("Demanda_Liquida", "sum"),
            Qtd_Transferida=("Qtd_Transferida", "sum"),
            Backlog=("Backlog", "sum"),
            Valor_Transferido=("Valor_Transferido", "sum"),
        )

    return op


def montar_cenario_modelo(oportunidades, bases_origem, df_estoque, df_demanda,
                          col_est_codigo, col_dem_codigo, col_dem_centro,
                          centros_preferenciais, centro_foco=None):
    """Gera o layout do modelo: origem, disponibilidade e destinos em colunas."""
    colunas = ["Centro Origem", "Qtd Disponivel origem", "Código", "Descrição", "Sobra no centro de origem"]
    preferencia = ["DPEX", "ES02", "ES03", "ES04", "ES05", "ES07", "ES09", "ES14"]
    encontrados = sorted(set(df_demanda[col_dem_centro].dropna().astype(str).str.strip())) if col_dem_centro in df_demanda.columns else []
    centros = []
    for c in preferencia + encontrados + list(centros_preferenciais or []):
        c = str(c).strip()
        if c and c not in centros:
            centros.append(c)
    colunas.extend(centros)

    dados = oportunidades.copy() if oportunidades is not None else pd.DataFrame()
    if centro_foco:
        dados = dados[dados["Centro_Destino"].astype(str) == str(centro_foco)].copy()

    if dados.empty:
        vazio = pd.DataFrame(columns=colunas)
        vazio.loc[0, "Código"] = "Sem oportunidades para o cenário"
        return vazio

    descricoes = montar_descricoes(df_estoque, df_demanda, col_est_codigo, col_dem_codigo)
    if isinstance(bases_origem, list):
        base = pd.concat(bases_origem, ignore_index=True) if bases_origem else pd.DataFrame()
    else:
        base = bases_origem.copy() if bases_origem is not None else pd.DataFrame()
    if not base.empty:
        base["Codigo"] = base["Codigo"].map(normalizar_codigo)
        if "Centro_Origem" not in base.columns:
            # No cenário de critérios, a base contém uma única origem.
            # Recupera essa origem pelas oportunidades para não deixar a coluna vazia.
            origens_dados = (
                dados["Centro_Origem"]
                .dropna()
                .astype(str)
                .replace(["", "nan", "None"], pd.NA)
                .dropna()
                .unique()
                .tolist()
            ) if "Centro_Origem" in dados.columns else []
            base["Centro_Origem"] = origens_dados[0] if len(origens_dados) == 1 else ""
        base["Centro_Origem"] = base["Centro_Origem"].fillna("").astype(str)

    linhas = []
    for (origem, codigo), grupo in dados.groupby(["Centro_Origem", "Codigo"], as_index=False):
        origem = str(origem)
        codigo = normalizar_codigo(codigo)
        saldo = 0.0
        if not base.empty:
            mask = base["Centro_Origem"].astype(str).eq(origem) & base["Codigo"].astype(str).eq(codigo)
            saldo = float(base.loc[mask, "Saldo_Transferivel_Origem"].sum())
        transferencias = grupo.groupby("Centro_Destino")["Qtd_Transferida"].sum().to_dict()
        total = float(grupo["Qtd_Transferida"].sum())

        # Segurança de exibição:
        # se houve transferência mas a base não retornou saldo (ex.: mismatch pontual),
        # não deixar "Qtd Disponivel origem" zerada no relatório.
        saldo_exibicao = max(saldo, total)

        sobra = max(saldo_exibicao - total, 0)
        if sobra < 1e-6:
            sobra = 0.0
        linha = {"Centro Origem": origem, "Qtd Disponivel origem": saldo_exibicao, "Código": codigo, "Descrição": descricoes.get(codigo, ""), "Sobra no centro de origem": sobra}
        for centro in centros:
            linha[centro] = float(transferencias.get(centro, 0))
        linhas.append(linha)
    resultado = pd.DataFrame(linhas, columns=colunas)

    # Garantia final para o Excel: origem e quantidade sempre aparecem.
    if "Centro Origem" not in resultado.columns:
        resultado["Centro Origem"] = ""
    if "Qtd Disponivel origem" not in resultado.columns:
        resultado["Qtd Disponivel origem"] = 0.0

    resultado["Centro Origem"] = resultado["Centro Origem"].fillna("").astype(str)
    resultado["Qtd Disponivel origem"] = pd.to_numeric(
        resultado["Qtd Disponivel origem"],
        errors="coerce"
    ).fillna(0)

    return resultado.sort_values(
        ["Centro Origem", "Código"]
    ).reset_index(drop=True)

def calcular_geral(df_estoque, df_demanda, df_separacao, centros,
                   col_est_codigo, col_est_centro, col_est_deposito, col_est_quantidade,
                   col_dem_codigo, col_dem_centro, col_dem_quantidade, col_dem_data, col_dem_valor,
                   col_sep_codigo, col_sep_centro, col_sep_quantidade, estoque_minimo):
    """Calcula o cenário geral, sem usar os critérios selecionados na interface."""
    dummy = "__GERAL__"
    saldo_dest = saldo_destinos(df_estoque, df_separacao, dummy, set(), col_est_codigo, col_est_centro, col_est_quantidade, col_sep_codigo, col_sep_centro, col_sep_quantidade)
    demanda_base = demanda_liquida(df_demanda, saldo_dest, dummy, set(), col_dem_codigo, col_dem_centro, col_dem_quantidade, col_dem_data, col_dem_valor)
    residual = demanda_base.copy()
    bases, fifos = [], []
    for origem in sorted(set(centros)):
        base = saldo_origem(df_estoque, df_demanda, df_separacao, origem, [], set(), col_est_codigo, col_est_centro, col_est_deposito, col_est_quantidade, col_dem_codigo, col_dem_centro, col_dem_quantidade, col_sep_codigo, col_sep_centro, col_sep_quantidade, estoque_minimo)
        base["Centro_Origem"] = origem
        bases.append(base)
        saldo_map = dict(zip(normalizar_serie_codigo(base["Codigo"]), base["Saldo_Transferivel_Origem"]))
        demanda_origem = residual[residual["Centro"].astype(str) != str(origem)].copy()
        fifo = aplicar_fifo(demanda_origem, saldo_map)
        if fifo.empty:
            continue
        fifo["Centro_Origem"] = origem
        fifos.append(fifo)
        atendido = fifo.groupby("Linha_ID")["Qtd_Transferida_FIFO"].sum()
        residual = residual.copy()
        residual["Demanda_Liquida_Linha"] -= residual["Linha_ID"].map(atendido).fillna(0)
        residual["Demanda_Liquida_Linha"] = residual["Demanda_Liquida_Linha"].clip(lower=0)
        residual = residual[residual["Demanda_Liquida_Linha"] > 0].copy()
        if residual.empty:
            break
    base_final = pd.concat(bases, ignore_index=True) if bases else pd.DataFrame()
    fifo_final = pd.concat(fifos, ignore_index=True) if fifos else pd.DataFrame()
    oportunidades = construir_oportunidades(fifo_final, "") if not fifo_final.empty else pd.DataFrame()
    return base_final, oportunidades

def arredondar_quantidades_relatorio(df):
    """
    Arredonda somente as quantidades exibidas no Excel.
    Os cálculos internos continuam usando os valores originais.
    """
    if df is None:
        return pd.DataFrame()

    resultado = df.copy()

    # Nome final usado no relatório.
    if "Faturar no centro de origem" in resultado.columns:
        resultado = resultado.rename(
            columns={
                "Faturar no centro de origem":
                "Sobra no centro de origem"
            }
        )

    colunas_texto = {
        "Centro Origem",
        "Código",
        "Descrição",
        "Sobra no centro de origem"
    }

    for coluna in resultado.columns:
        # As colunas dos centros e a quantidade da origem são quantidades.
        eh_quantidade = (
            coluna == "Qtd Disponivel origem"
            or coluna == "Sobra no centro de origem"
            or coluna not in colunas_texto
        )

        if not eh_quantidade:
            continue

        serie = pd.to_numeric(resultado[coluna], errors="coerce")

        # Só arredonda colunas que possuem valores numéricos.
        if serie.notna().any():
            resultado[coluna] = serie.fillna(0).round(0).astype(int)

    return resultado

def formatar_aba_modelo(workbook, nome_aba):
    """Aplica cabeçalho na linha 2 e filtros no padrão do arquivo-modelo."""
    ws = workbook[nome_aba]
    ws.freeze_panes = "B3"
    ws.auto_filter.ref = f"A2:{get_column_letter(ws.max_column)}{ws.max_row}"
    fill = PatternFill("solid", fgColor="D9EAF7")
    thin = Side(style="thin", color="B7C9D6")
    for cell in ws[2]:
        cell.font = Font(bold=True, color="1F2937")
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=thin)
    for row in ws.iter_rows(min_row=3):
        for cell in row:
            cell.alignment = Alignment(vertical="center")
    larguras = {1: 18, 2: 22, 3: 16, 4: 42, 5: 25}
    for i in range(5, ws.max_column + 1):
        larguras[i] = 14
    for i, largura in larguras.items():
        ws.column_dimensions[get_column_letter(i)].width = largura
    ws.row_dimensions[2].height = 32

def consulta(df, prefixo):
    if df is None or df.empty:
        return pd.DataFrame()
    out = df.copy()
    busca = st.text_input("Busca global", key=f"{prefixo}_busca")
    if busca:
        mask = pd.Series(False, index=out.index)
        for c in out.columns:
            mask |= out[c].astype(str).str.contains(busca, case=False, na=False)
        out = out[mask]
    if st.toggle("Ativar filtros por coluna", key=f"{prefixo}_toggle"):
        with st.expander("Filtros", expanded=False):
            for c in out.columns:
                v = st.text_input(str(c), key=f"{prefixo}_{c}")
                if v:
                    out = out[out[c].astype(str).str.contains(v, case=False, na=False)]
    return out

# =========================================================
# ARQUIVOS E MAPEAMENTO
# =========================================================
st.sidebar.header("📂 Arquivos")
arquivos = buscar_arquivos(PASTA_RELATORIOS)
if not arquivos:
    st.error("Nenhum arquivo Excel, XLS ou CSV encontrado.")
    st.stop()

i_est, i_dem, i_sep = indices_por_nome(arquivos)
arq_est = st.sidebar.selectbox("Relatório de estoque", arquivos, index=i_est)
arq_dem = st.sidebar.selectbox("Relatório de demanda", arquivos, index=i_dem)
arq_sep = st.sidebar.selectbox("Relatório de separações", arquivos, index=i_sep)

caminho_estoque = os.path.join(PASTA_RELATORIOS, arq_est)
caminho_demanda = os.path.join(PASTA_RELATORIOS, arq_dem)
caminho_separacao = os.path.join(PASTA_RELATORIOS, arq_sep)

df_estoque = carregar_arquivo(
    caminho_estoque, assinatura_arquivo(caminho_estoque)
)
df_demanda = carregar_arquivo(
    caminho_demanda, assinatura_arquivo(caminho_demanda)
)
df_separacao = carregar_arquivo(
    caminho_separacao, assinatura_arquivo(caminho_separacao)
)

historicos_disponiveis = buscar_historicos_consumo(PASTA_RELATORIOS)
st.sidebar.markdown("---")
st.sidebar.header("📚 4️⃣ Histórico de consumo")

# Modo diagnóstico (oculto por padrão)
modo_diagnostico = st.sidebar.toggle("Modo diagnóstico", value=False)

# Seleção automática do histórico mais recente quando possível.
auto_historico = historico_mais_recente(PASTA_RELATORIOS, historicos_disponiveis)
opcoes_historico = ["(Nenhum selecionado)"] + historicos_disponiveis

# Se não houver valor salvo, valor inválido, ou estiver em "Nenhum", aplica auto.
valor_atual_hist = st.session_state.get("arquivo_historico")
if valor_atual_hist not in opcoes_historico or valor_atual_hist == "(Nenhum selecionado)":
    st.session_state["arquivo_historico"] = auto_historico if auto_historico else "(Nenhum selecionado)"

arq_historico = st.sidebar.selectbox(
    "Arquivo do histórico",
    opcoes_historico,
    key="arquivo_historico"
)

if arq_historico == "(Nenhum selecionado)":
    caminho_historico = None
    df_historico = pd.DataFrame()
    if historicos_disponiveis:
        st.sidebar.warning(
            "Histórico encontrado, mas não selecionado. Selecione um arquivo para habilitar 'Esvaziar depósito'."
        )
    else:
        st.sidebar.warning(
            "Histórico não localizado automaticamente. Verifique a pasta 'Historico_consumo' (ou variações) e o nome do arquivo."
        )
else:
    caminho_historico = os.path.join(
        PASTA_RELATORIOS,
        arq_historico
    )
    df_historico_bruto = carregar_arquivo(
        caminho_historico,
        assinatura_arquivo(caminho_historico)
    )
    df_historico = padronizar_historico_consumo(df_historico_bruto)

    if df_historico.empty:
        st.sidebar.error(
            "Histórico carregado, mas não foi possível interpretar o layout. "
            "Verifique se existe coluna de código e centros (ES01, ES02, ...), ou colunas Código/Centro/Consumo."
        )
    else:
        st.sidebar.success(
            f"Histórico selecionado automaticamente: {arq_historico}" if arq_historico == auto_historico else f"Histórico selecionado: {arq_historico}"
        )
        if modo_diagnostico:
            st.sidebar.caption(
                f"Histórico padronizado: {len(df_historico):,} linhas úteis em formato Código/Centro/Consumo."
            )

# Painel de atualização (arquivos + página)
ultima_pagina = datetime.now()
ultima_est = data_hora_modificacao(caminho_estoque)
ultima_dem = data_hora_modificacao(caminho_demanda)
ultima_sep = data_hora_modificacao(caminho_separacao)
ultima_hist = data_hora_modificacao(caminho_historico) if caminho_historico else None
ultimas_validas = [d for d in [ultima_est, ultima_dem, ultima_sep, ultima_hist] if d is not None]
ultima_arquivos = max(ultimas_validas) if ultimas_validas else None

with st.sidebar.expander("Arquivos carregados", expanded=False):
    st.write(f"**Última atualização da página:** {fmt_data_hora(ultima_pagina)}")
    st.write(f"**Última atualização dos arquivos (mais recente):** {fmt_data_hora(ultima_arquivos)}")
    st.write("**Detalhe por arquivo:**")
    st.write(f"- Estoque ({arq_est}): {fmt_data_hora(ultima_est)}")
    st.write(f"- Demanda ({arq_dem}): {fmt_data_hora(ultima_dem)}")
    st.write(f"- Separações ({arq_sep}): {fmt_data_hora(ultima_sep)}")
    if caminho_historico:
        st.write(f"- Histórico ({arq_historico}): {fmt_data_hora(ultima_hist)}")
    else:
        st.write("- Histórico: não selecionado")

if modo_diagnostico:
    st.sidebar.caption("Códigos são normalizados apenas para remover .0; 750520 e 750521 continuam distintos.")
    st.sidebar.caption(
        "Relatórios são recarregados quando o tamanho ou a data de alteração mudam."
    )
    with st.sidebar.expander("🔍 Verificar códigos 750520 / 750521", expanded=False):
        if "PRODUTO" in df_demanda.columns:
            diagnostico = df_demanda.copy()
            diagnostico["_Codigo_Normalizado"] = normalizar_serie_codigo(diagnostico["PRODUTO"])
            diag = diagnostico[diagnostico["_Codigo_Normalizado"].isin(["750520", "750521"])]
            st.write("Linhas encontradas:", len(diag))
            st.dataframe(
                diag[[c for c in ["PRODUTO", "Centro", "Saldo Pendente", "Pedido", "Data Pedido"] if c in diag.columns]],
                width="stretch",
                hide_index=True
            )
        else:
            st.info("A coluna PRODUTO não existe no arquivo selecionado.")
    st.sidebar.caption("Códigos numéricos são normalizados apenas para remover .0; 750520 e 750521 permanecem distintos.")

st.sidebar.markdown("---")
st.sidebar.header("🧭 Mapeamento de colunas")

with st.sidebar.expander("Estoque", expanded=True):
    cols = df_estoque.columns.tolist()
    col_est_codigo = escolher_coluna("Código", cols, ["Material", "PRODUTO", "Produto", "Item"])
    col_est_centro = escolher_coluna("Centro", cols, ["Centro"])
    col_est_deposito = escolher_coluna("Depósito", cols, ["Depósito", "Deposito", "Armazém", "Armazem"])
    col_est_quantidade = escolher_coluna("Quantidade disponível", cols, ["Utilização livre", "Utilizacao livre", "Quantidade", "Qtd"])

with st.sidebar.expander("Demanda", expanded=True):
    cols = df_demanda.columns.tolist()
    col_dem_codigo = escolher_coluna("Código", cols, ["PRODUTO", "Material", "Produto", "Item"])
    col_dem_centro = escolher_coluna("Centro", cols, ["Centro"])
    col_dem_quantidade = escolher_coluna("Quantidade", cols, ["Saldo Pendente", "Qtde OV", "Quantidade", "Qtd"])
    sugestoes = ["Data Pedido", "Data do Pedido", "Dt Pedido", "Data", "Remessa Desejada"]
    if len(cols) >= 10:
        sugestoes.append(cols[9])
    col_dem_data = escolher_coluna("Data do pedido FIFO", cols, sugestoes)
    col_dem_valor = escolher_coluna("Valor líquido", cols, ["Valor líquido", "Valor liquido", "Valor Líquido", "Valor"])

with st.sidebar.expander("Separações", expanded=True):
    cols = df_separacao.columns.tolist()
    col_sep_codigo = escolher_coluna("Código", cols, ["PRODUTO", "Material", "Produto", "Item", "Descrição"])
    col_sep_centro = escolher_coluna("Centro", cols, ["Local Exped.", "Centro", "Centro_Origem"])
    col_sep_quantidade = escolher_coluna("Quantidade em separação", cols, ["Qtde. Conf. Pick.", "Qtde. Fornec.", "Quantidade", "Qtd"])

if not df_historico.empty:
    # Após padronização, o histórico sempre usa colunas fixas.
    col_hist_codigo = "Codigo"
    col_hist_centro = "Centro_Destino"
    col_hist_consumo = "Consumo_Historico"

    if modo_diagnostico:
        with st.sidebar.expander("Histórico de consumo - estrutura detectada", expanded=False):
            st.write("Código:", col_hist_codigo)
            st.write("Centro:", col_hist_centro)
            st.write("Consumo:", col_hist_consumo)
            st.dataframe(df_historico.head(10), width="stretch", hide_index=True)
else:
    col_hist_codigo = col_hist_centro = col_hist_consumo = None

# =========================================================
# CRITÉRIOS E MODELO
# =========================================================
criterios = st.session_state.criterios
origem = criterios.get("centro_origem")
destino_foco = criterios.get("centro_destino")
depositos = criterios.get("depositos_origem", [])
codigos = {normalizar_codigo(x) for x in criterios.get("codigos_avaliar", []) if normalizar_codigo(x)}
estoque_minimo = float(criterios.get("estoque_minimo", 0))

parametros = {
    "arquivos": [arq_est, arq_dem, arq_sep],
    "assinatura_estoque": assinatura_arquivo(caminho_estoque),
    "assinatura_demanda": assinatura_arquivo(caminho_demanda),
    "assinatura_separacao": assinatura_arquivo(caminho_separacao),
    "origem": origem, "destino": destino_foco,
    "depositos": sorted(depositos), "codigos": sorted(codigos),
    "estoque_minimo": estoque_minimo,
    "esvaziar_deposito": bool(criterios.get("esvaziar_deposito", False)),
    "arquivo_historico": arq_historico,
    "assinatura_historico": assinatura_arquivo(caminho_historico) if caminho_historico else None,
    "colunas": [col_est_codigo, col_est_centro, col_est_deposito, col_est_quantidade,
                 col_dem_codigo, col_dem_centro, col_dem_quantidade, col_dem_data, col_dem_valor,
                 col_sep_codigo, col_sep_centro, col_sep_quantidade]
}

centros_geral = sorted(set(df_estoque[col_est_centro].dropna().astype(str)) | set(df_demanda[col_dem_centro].dropna().astype(str)))
h_geral = hash_modelo({
    "arquivos": [arq_est, arq_dem, arq_sep],
    "assinatura_estoque": assinatura_arquivo(caminho_estoque),
    "assinatura_demanda": assinatura_arquivo(caminho_demanda),
    "assinatura_separacao": assinatura_arquivo(caminho_separacao),
    "centros": centros_geral,
    "estoque_minimo": estoque_minimo
})
if st.session_state.get("modelo_geral_hash") != h_geral:
    base_geral, oportunidades_geral = calcular_geral(
        df_estoque, df_demanda, df_separacao, centros_geral,
        col_est_codigo, col_est_centro, col_est_deposito, col_est_quantidade,
        col_dem_codigo, col_dem_centro, col_dem_quantidade, col_dem_data, col_dem_valor,
        col_sep_codigo, col_sep_centro, col_sep_quantidade, estoque_minimo
    )
    st.session_state["modelo_geral"] = {"base": base_geral, "oportunidades": oportunidades_geral}
    st.session_state["modelo_geral_hash"] = h_geral

geral = st.session_state.get("modelo_geral", {})
base_geral = geral.get("base", pd.DataFrame())
oportunidades_geral = geral.get("oportunidades", pd.DataFrame())

h = hash_modelo(parametros)
if origem and st.session_state.modelo_hash != h:
    base_origem = saldo_origem(
        df_estoque, df_demanda, df_separacao, origem, depositos, codigos,
        col_est_codigo, col_est_centro, col_est_deposito, col_est_quantidade,
        col_dem_codigo, col_dem_centro, col_dem_quantidade,
        col_sep_codigo, col_sep_centro, col_sep_quantidade, estoque_minimo
    )
    saldo_dest = saldo_destinos(
        df_estoque, df_separacao, origem, codigos,
        col_est_codigo, col_est_centro, col_est_quantidade,
        col_sep_codigo, col_sep_centro, col_sep_quantidade
    )
    dem_liq = demanda_liquida(
        df_demanda, saldo_dest, origem, codigos,
        col_dem_codigo, col_dem_centro, col_dem_quantidade,
        col_dem_data, col_dem_valor
    )
    saldo_map = dict(zip(normalizar_serie_codigo(base_origem["Codigo"]), base_origem["Saldo_Transferivel_Origem"]))
    fifo = aplicar_fifo(dem_liq, saldo_map)
    centros_fifo = resumo_fifo(fifo)
    resultado_destino = centros_fifo[centros_fifo["Centro"].astype(str) == str(destino_foco)].copy() if destino_foco else centros_fifo.copy()
    valor_centro = fifo.groupby("Centro", as_index=False)["Valor_Transferido_Linha"].sum().rename(columns={"Valor_Transferido_Linha": "Valor_Transferido_Centro"}).sort_values("Valor_Transferido_Centro", ascending=False) if not fifo.empty else pd.DataFrame()
    if not valor_centro.empty:
        valor_centro = valor_centro[valor_centro["Valor_Transferido_Centro"] > 0]
    oportunidades = construir_oportunidades(fifo, origem)
    if criterios.get("esvaziar_deposito", False) and not df_historico.empty:
        extras = distribuir_saldo_restante_por_historico(
            base_origem, df_historico, oportunidades, saldo_dest, origem, codigos,
            col_hist_codigo, col_hist_centro, col_hist_consumo
        )
        if not extras.empty:
            oportunidades = pd.concat([oportunidades, extras], ignore_index=True)
            oportunidades = oportunidades.groupby(["Centro_Origem", "Centro_Destino", "Codigo"], as_index=False).agg(
                Data_Pedido_Mais_Antiga=("Data_Pedido_Mais_Antiga", "min"),
                Demanda_Liquida=("Demanda_Liquida", "sum"),
                Qtd_Transferida=("Qtd_Transferida", "sum"),
                Backlog=("Backlog", "sum"),
                Valor_Transferido=("Valor_Transferido", "sum")
            )

        # Garantia final: zera sobra por código na origem quando o modo esvaziar está ativo.
        oportunidades = reconciliar_esvaziamento_total(
            oportunidades, base_origem, df_historico, origem,
            col_hist_codigo, col_hist_centro, col_hist_consumo
        )
    top5 = top5_oportunidades_por_centro(oportunidades)
    top5_por_centro = oportunidades
    st.session_state.resultados = {"base_origem": base_origem, "saldo_destino": saldo_dest, "demanda_liquida": dem_liq, "fifo": fifo, "centros_fifo": centros_fifo, "resultado_destino": resultado_destino, "valor_centro": valor_centro, "top5": top5, "top5_por_centro": top5_por_centro, "oportunidades": oportunidades}
    st.session_state.modelo_hash = h

res = st.session_state.resultados
base_origem = res.get("base_origem", pd.DataFrame())
saldo_dest = res.get("saldo_destino", pd.DataFrame())
dem_liq = res.get("demanda_liquida", pd.DataFrame())
fifo = res.get("fifo", pd.DataFrame())
centros_fifo = res.get("centros_fifo", pd.DataFrame())
resultado_destino = res.get("resultado_destino", pd.DataFrame())
valor_centro = res.get("valor_centro", pd.DataFrame())
top5 = res.get("top5", pd.DataFrame())
top5_por_centro = res.get("top5_por_centro", pd.DataFrame())
oportunidades = res.get("oportunidades", pd.DataFrame())

# =========================================================
# ABAS
# =========================================================
tab_dash, tab_crit, tab_rec, tab_cons, tab_rel = st.tabs(["📊 Dashboard", "⚙️ Critérios", "🎯 Recomendações", "❓ Consultas", "📈 Relatório"])

with tab_dash:
    st.subheader("Visão geral")
    c1, c2, c3, c4 = st.columns(4)
    total_est = num(df_estoque, col_est_quantidade).sum()
    total_dem = num(df_demanda, col_dem_quantidade).sum()
    total_sep = num(df_separacao, col_sep_quantidade).sum()
    c1.metric("Estoque total", f"{total_est:,.0f}")
    c2.metric("Demanda total", f"{total_dem:,.0f}")
    c3.metric("Em separação", f"{total_sep:,.0f}")
    c4.metric("Disponível geral", f"{total_est-total_sep:,.0f}")

    g1, g2 = st.columns(2)
    with g1:
        s = num(df_estoque, col_est_quantidade).groupby(df_estoque[col_est_centro].astype(str)).sum().sort_values(ascending=False).head(10)
        st.bar_chart(s, height=300)
        st.caption("Estoque por centro")
    with g2:
        s = num(df_demanda, col_dem_quantidade).groupby(df_demanda[col_dem_centro].astype(str)).sum().sort_values(ascending=False).head(10)
        st.bar_chart(s, height=300)
        st.caption("Demanda bruta por centro")

    st.markdown("---")
    st.subheader("💰 Valor líquido transferível por centro")
    if valor_centro.empty:
        st.info("Não há valor transferível no cenário atual.")
    else:
        esquerda, direita = st.columns([2, 1])
        with esquerda:
            fig = px.bar(valor_centro, x="Centro", y="Valor_Transferido_Centro", text_auto=".2s", title="Valor após abatimento do saldo do destino")
            fig.update_layout(clickmode="event+select", xaxis_title="Centro destino", yaxis_title="Valor transferido")
            evento = st.plotly_chart(
                fig,
                use_container_width=True,
                key="grafico_valor",
                on_select="rerun",
                selection_mode=("points",)
            )

            # O retorno do Streamlit pode ser PlotlyState, dict ou objeto de ponto.
            # O tratamento abaixo torna o clique na barra compatível com esses formatos.
            centro_clicado = None
            try:
                selecao = getattr(evento, "selection", None)
                pontos = getattr(selecao, "points", None) if selecao is not None else None

                if pontos is None and isinstance(evento, dict):
                    selecao = evento.get("selection", {})
                    pontos = selecao.get("points", []) if isinstance(selecao, dict) else []

                if pontos:
                    ponto = pontos[0]
                    if isinstance(ponto, dict):
                        centro_clicado = ponto.get("x") or ponto.get("label")
                    else:
                        centro_clicado = getattr(ponto, "x", None) or getattr(ponto, "label", None)
            except Exception:
                centro_clicado = None

            if centro_clicado is not None:
                st.session_state.centro_clicado = str(centro_clicado)

            st.caption("Clique em uma barra para ver o cenário do centro selecionado.")
        with direita:
            st.metric("Valor total", f"{valor_centro['Valor_Transferido_Centro'].sum():,.2f}")
            centros_grafico = ["(Nenhum)"] + valor_centro["Centro"].astype(str).tolist()
            selecao_manual = st.selectbox(
                "Centro para detalhar",
                centros_grafico,
                index=(centros_grafico.index(st.session_state.centro_clicado)
                       if st.session_state.centro_clicado in centros_grafico else 0),
                key="selecao_manual_centro"
            )

            if selecao_manual != "(Nenhum)":
                st.session_state.centro_clicado = selecao_manual
            elif st.session_state.centro_clicado not in valor_centro["Centro"].astype(str).tolist():
                st.session_state.centro_clicado = None

            centro = st.session_state.centro_clicado
            if st.button("Limpar seleção"):
                st.session_state.centro_clicado = None
                centro = None
                st.rerun()
            if centro:
                d = fifo[fifo["Centro"].astype(str) == str(centro)]
                st.success(f"Centro selecionado: {centro}")
                st.metric("Demanda original", f"{d['Demanda_Original_Linha'].sum():,.2f}")
                st.metric("Saldo do destino abatido", f"{d['Saldo_Destino_Usado'].sum():,.2f}")
                st.metric("Demanda líquida", f"{d['Demanda_Liquida_Linha'].sum():,.2f}")
                st.metric("Qtd transferida", f"{d['Qtd_Transferida_FIFO'].sum():,.2f}")
                st.metric("Valor transferido", f"{d['Valor_Transferido_Linha'].sum():,.2f}")
                st.dataframe(destacar(d.groupby("Codigo", as_index=False).agg(Qtd_Transferida=("Qtd_Transferida_FIFO", "sum"), Valor_Transferido=("Valor_Transferido_Linha", "sum")).sort_values("Valor_Transferido", ascending=False).head(5), "Qtd_Transferida"), width="stretch", hide_index=True)
            else:
                st.markdown("#### Top 5 oportunidades gerais")
                st.dataframe(
                    destacar(top5, "Qtd_Transferida"),
                    width="stretch",
                    hide_index=True
                )

            st.markdown("#### Oportunidades de transferência")
            centro_para_top5 = st.session_state.get("centro_clicado", None)
            quadro_top5 = top5_oportunidades_por_centro(
                oportunidades,
                centro_para_top5
            )
            if centro_para_top5:
                st.success(
                    f"Top 5 do centro selecionado: {centro_para_top5}"
                )
            else:
                st.info(
                    "Sem centro selecionado: exibindo o Top 5 geral. "
                    "Clique em uma barra para filtrar."
                )
            st.dataframe(
                destacar(quadro_top5, "Qtd_Transferida"),
                width="stretch",
                hide_index=True
            )

with tab_crit:
    st.subheader("Configuração de critérios")
    centros = sorted(set(df_estoque[col_est_centro].dropna().astype(str)) | set(df_demanda[col_dem_centro].dropna().astype(str)))
    origem_nova = st.selectbox("Centro de origem", centros, index=0 if centros else None)
    destinos = [x for x in centros if x != origem_nova]
    destino_novo = st.selectbox("Centro de destino para foco", ["(Todos)"] + destinos, index=0)
    deps = sorted(df_estoque[df_estoque[col_est_centro].astype(str) == str(origem_nova)][col_est_deposito].dropna().astype(str).unique())
    deps_novos = st.multiselect("Depósitos de retirada", deps, default=[x for x in criterios.get("depositos_origem", []) if x in deps])
    texto_codigos = st.text_area("Cole os códigos dos materiais", value="\n".join(criterios.get("codigos_avaliar", [])), height=120)
    codigos_novos = parse_codigos(texto_codigos)
    c1, c2 = st.columns(2)
    minimo = c1.number_input("Estoque mínimo na origem", min_value=0, value=int(criterios.get("estoque_minimo", 500)), step=100)
    maximo = c2.number_input("Estoque máximo", min_value=0, value=int(criterios.get("estoque_maximo", 10000)), step=100)

    esvaziar_novo = st.checkbox(
        "Esvaziar depósito selecionado após aplicar os demais critérios",
        value=bool(criterios.get("esvaziar_deposito", False)),
        disabled=df_historico.empty,
        help="Ignora saldo/demanda do destino para o saldo restante e distribui pelos centros com histórico, buscando equalizar os estoques."
    )
    if df_historico.empty:
        st.caption("Carregue um arquivo em Historico_consumo para habilitar esta opção.")
    if st.button("💾 Salvar critérios"):
        st.session_state.criterios = {"criterio_principal": "FIFO - data do pedido", "estoque_minimo": minimo, "estoque_maximo": maximo, "centro_origem": origem_nova, "centro_destino": None if destino_novo == "(Todos)" else destino_novo, "depositos_origem": deps_novos, "codigos_avaliar": codigos_novos, "esvaziar_deposito": esvaziar_novo}
        st.session_state.modelo_hash = None
        st.session_state.centro_clicado = None
        st.rerun()

with tab_rec:
    st.subheader("Recomendações")
    if not origem:
        st.warning("Defina a origem na aba Critérios.")
    elif fifo.empty:
        st.info("Não há demanda líquida para processar.")
    else:
        ativar = st.toggle("Ativar filtros avançados", value=False)
        st.markdown("### Prioridades antes do destino (FIFO)")
        q1 = filtros(centros_fifo, "fifo", ativar)
        st.dataframe(destacar(q1, "Qtd_Transferida_FIFO"), width="stretch", hide_index=True)
        st.markdown("### Resultado para o centro destino")
        q2 = filtros(resultado_destino, "destino", ativar)
        st.dataframe(destacar(q2, "Qtd_Transferida_FIFO"), width="stretch", hide_index=True)
        st.markdown("### Detalhe FIFO linha a linha")
        q3 = filtros(fifo, "detalhe", ativar)
        st.dataframe(destacar(q3, "Qtd_Transferida_FIFO"), width="stretch", hide_index=True)
        st.markdown("### Saldo considerado nos destinos")
        st.dataframe(saldo_dest, width="stretch", hide_index=True)

        st.markdown("### Top 5 oportunidades de transferência por centro")
        if top5_por_centro.empty:
            st.info("Não há oportunidades de transferência efetiva por centro.")
        else:
            st.dataframe(
                destacar(top5_por_centro, "Qtd_Transferida"),
                width="stretch",
                hide_index=True
            )

with tab_cons:
    st.subheader("Consultas funcionais")
    fontes = {"Demanda original": df_demanda, "Estoque": df_estoque, "Separações": df_separacao, "Saldo dos destinos": saldo_dest, "Demanda líquida": dem_liq, "FIFO detalhe": fifo, "FIFO por centro": centros_fifo, "Resultado destino": resultado_destino, "Valor por centro": valor_centro, "Top 5 oportunidades por centro": top5_por_centro}
    fonte = st.selectbox("Fonte", list(fontes.keys()))
    resultado = consulta(fontes[fonte], "consulta")
    st.dataframe(resultado, width="stretch", hide_index=True)
    st.download_button("Exportar consulta CSV", resultado.to_csv(index=False).encode("utf-8-sig"), file_name="consulta.csv", mime="text/csv")

with tab_rel:
    st.subheader("Exportar relatório")
    st.info("O Excel terá duas abas: Cenário Geral e Cenário dos Critérios.")
    if st.button("📥 Gerar Excel", key="gerar_excel"):
        try:
            centros_export = sorted(set(df_demanda[col_dem_centro].dropna().astype(str)))
            cenario_geral = montar_cenario_modelo(oportunidades_geral, base_geral, df_estoque, df_demanda, col_est_codigo, col_dem_codigo, col_dem_centro, centros_export, centro_foco=None)
            cenario_criterios = montar_cenario_modelo(oportunidades, base_origem, df_estoque, df_demanda, col_est_codigo, col_dem_codigo, col_dem_centro, centros_export, centro_foco=destino_foco)
            # Arredonda somente a apresentação no relatório.
            # O modelo continua calculando com os valores originais.
            cenario_geral = arredondar_quantidades_relatorio(cenario_geral)
            cenario_criterios = arredondar_quantidades_relatorio(cenario_criterios)

            arquivo = io.BytesIO()
            with pd.ExcelWriter(arquivo, engine="openpyxl") as writer:
                cenario_geral.to_excel(writer, sheet_name="Cenario Geral", index=False, startrow=1)
                cenario_criterios.to_excel(writer, sheet_name="Cenario Criterios", index=False, startrow=1)
            arquivo.seek(0)
            workbook = load_workbook(arquivo)
            formatar_aba_modelo(workbook, "Cenario Geral")
            formatar_aba_modelo(workbook, "Cenario Criterios")
            workbook.active = 0
            final = io.BytesIO()
            workbook.save(final)
            final.seek(0)
            st.download_button("Baixar Excel com dois cenários", data=final.getvalue(), file_name=f"transferencias_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="download_dois_cenarios")
            st.success("Relatório gerado com Cenário Geral e Cenário dos Critérios.")
        except Exception as erro:
            st.error(f"Erro ao gerar o Excel: {erro}")

# =========================================================
# FIM
# =========================================================
print("Aplicação Streamlit carregada com abatimento do saldo do destino.")
