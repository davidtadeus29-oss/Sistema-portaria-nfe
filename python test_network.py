import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io
import os

st.set_page_config(page_title="IA Análise de Transferências", layout="wide")

st.title("🤖 IA de Análise de Transferências de Estoque")
st.markdown("Analisa demanda, estoque e separações para otimizar transferências entre centros")

# ============= CONFIGURAÇÃO DE PASTA =============
PASTA_RELATORIOS = r"C:\Users\OLIVDT\Desktop\Relatorios_transferencia"

if not os.path.exists(PASTA_RELATORIOS):
    PASTA_RELATORIOS = os.path.expanduser("~/Desktop")
    st.warning(f"⚠️ Pasta configurada não encontrada. Usando: {PASTA_RELATORIOS}")

st.sidebar.header("📁 Localização dos Arquivos")
st.sidebar.info(f"📂 Pasta: `{PASTA_RELATORIOS}`")

# ============= BUSCAR ARQUIVOS =============
def buscar_arquivos(pasta, extensoes=['.xlsx', '.xls', '.csv']):
    """Busca todos os arquivos Excel/CSV na pasta"""
    arquivos = []
    if os.path.exists(pasta):
        for arquivo in os.listdir(pasta):
            if any(arquivo.lower().endswith(ext) for ext in extensoes):
                arquivos.append(arquivo)
    return sorted(arquivos)

arquivos_disponiveis = buscar_arquivos(PASTA_RELATORIOS)

if not arquivos_disponiveis:
    st.error(f"❌ Nenhum arquivo encontrado em:\n`{PASTA_RELATORIOS}`")
    st.stop()

st.sidebar.markdown("### Arquivos Encontrados:")
for arquivo in arquivos_disponiveis:
    st.sidebar.text(f"✓ {arquivo}")

# ============= SELEÇÃO DE ARQUIVOS =============
st.sidebar.markdown("---")
st.sidebar.header("🔗 Selecionar Arquivos")

rel_estoque_nome = st.sidebar.selectbox("📦 Estoque", arquivos_disponiveis, index=0)
rel_demanda_nome = st.sidebar.selectbox("📥 Demanda", arquivos_disponiveis, index=1 if len(arquivos_disponiveis) > 1 else 0)
rel_separacao_nome = st.sidebar.selectbox("📦 Separações", arquivos_disponiveis, index=2 if len(arquivos_disponiveis) > 2 else 0)

# ============= CARREGANDO OS DADOS =============
def carregar_arquivo(caminho):
    """Carrega arquivo Excel ou CSV"""
    try:
        if caminho.lower().endswith('.csv'):
            return pd.read_csv(caminho)
        else:
            return pd.read_excel(caminho, sheet_name=0)
    except Exception as e:
        st.error(f"❌ Erro ao carregar `{os.path.basename(caminho)}`:\n{e}")
        return None

df_estoque = carregar_arquivo(os.path.join(PASTA_RELATORIOS, rel_estoque_nome))
df_demanda = carregar_arquivo(os.path.join(PASTA_RELATORIOS, rel_demanda_nome))
df_separacao = carregar_arquivo(os.path.join(PASTA_RELATORIOS, rel_separacao_nome))

if df_estoque is None or df_demanda is None or df_separacao is None:
    st.error("❌ Erro ao carregar arquivos")
    st.stop()

st.success(f"✅ Arquivos carregados com sucesso!")

# ============= VISUALIZAR DADOS =============
with st.expander("📋 Visualizar Dados Carregados"):
    col1, col2, col3 = st.columns(3)
    with col1:
        st.subheader("Estoque")
        st.dataframe(df_estoque.head(5), width='stretch', height=200)
    with col2:
        st.subheader("Demanda")
        st.dataframe(df_demanda.head(5), width='stretch', height=200)
    with col3:
        st.subheader("Separações")
        st.dataframe(df_separacao.head(5), width='stretch', height=200)

# ============= ANÁLISE =============
st.markdown("---")
st.header("🔍 Análise de Transferências")

tab1, tab2, tab3, tab4 = st.tabs(["📊 Dashboard", "🎯 Recomendações", "❓ Consultas", "📈 Relatório"])

# ============= TAB 1: DASHBOARD =============
with tab1:
    try:
        # MAPEAMENTO DE COLUNAS REAIS
        col1, col2, col3, col4 = st.columns(4)
        
        # ESTOQUE
        estoque_total = pd.to_numeric(df_estoque['Utilização livre'], errors='coerce').sum()
        with col1:
            st.metric("📦 Total Estoque", f"{estoque_total:,.0f}")
        
        # DEMANDA
        demanda_total = pd.to_numeric(df_demanda['Saldo Pendente'], errors='coerce').sum()
        with col2:
            st.metric("📥 Total Demanda", f"{demanda_total:,.0f}")
        
        # SEPARAÇÃO
        separacao_total = pd.to_numeric(df_separacao['Qtde. Conf. Pick.'], errors='coerce').sum()
        with col3:
            st.metric("📦 Em Separação", f"{separacao_total:,.0f}")
        
        # DISPONÍVEL
        disponivel = estoque_total - separacao_total
        with col4:
            st.metric("✅ Disponível", f"{disponivel:,.0f}")
        
        st.markdown("---")
        
        # GRÁFICOS
        col1, col2 = st.columns(2)
        
        with col1:
            estoque_por_centro = pd.to_numeric(df_estoque['Utilização livre'], errors='coerce').groupby(df_estoque['Centro']).sum().sort_values(ascending=False).head(10)
            st.bar_chart(estoque_por_centro, height=300)
            st.caption("Top 10 - Estoque por Centro")
        
        with col2:
            demanda_por_centro = pd.to_numeric(df_demanda['Saldo Pendente'], errors='coerce').groupby(df_demanda['Centro']).sum().sort_values(ascending=False).head(10)
            st.bar_chart(demanda_por_centro, height=300)
            st.caption("Top 10 - Demanda por Centro")
    
    except Exception as e:
        st.error(f"❌ Erro no dashboard: {e}")

# ============= TAB 2: RECOMENDAÇÕES =============
with tab2:
    st.subheader("🎯 Recomendações de Transferência")
    
    try:
        # Preparar dados
        estoque_centro = pd.to_numeric(df_estoque['Utilização livre'], errors='coerce').groupby(df_estoque['Centro']).sum().reset_index()
        estoque_centro.columns = ['Centro', 'Estoque_Total']
        
        demanda_centro = pd.to_numeric(df_demanda['Saldo Pendente'], errors='coerce').groupby(df_demanda['Centro']).sum().reset_index()
        demanda_centro.columns = ['Centro', 'Demanda_Total']
        
        separacao_centro = pd.to_numeric(df_separacao['Qtde. Conf. Pick.'], errors='coerce').groupby(df_separacao['Local Exped.']).sum().reset_index()
        separacao_centro.columns = ['Centro', 'Em_Separacao']
        
        # Merge
        analise = estoque_centro.merge(demanda_centro, on='Centro', how='outer').fillna(0)
        analise = analise.merge(separacao_centro, on='Centro', how='outer').fillna(0)
        
        analise['Disponivel'] = analise['Estoque_Total'] - analise['Em_Separacao']
        analise['Deficit'] = analise['Demanda_Total'] - analise['Disponivel']
        analise['Excesso'] = analise['Disponivel'] - analise['Demanda_Total']
        
        st.markdown("### Centros com EXCESSO (Origem)")
        excesso = analise[analise['Excesso'] > 0].sort_values('Excesso', ascending=False)
        if len(excesso) > 0:
            st.dataframe(excesso[['Centro', 'Estoque_Total', 'Demanda_Total', 'Disponivel', 'Excesso']], width='stretch', hide_index=True)
        else:
            st.info("Nenhum centro com excesso")
        
        st.markdown("### Centros com DÉFICIT (Destino)")
        deficit = analise[analise['Deficit'] > 0].sort_values('Deficit', ascending=False)
        if len(deficit) > 0:
            st.dataframe(deficit[['Centro', 'Demanda_Total', 'Disponivel', 'Deficit']], width='stretch', hide_index=True)
        else:
            st.success("✅ Todos os centros têm estoque suficiente!")
    
    except Exception as e:
        st.error(f"❌ Erro: {e}")
        st.info(f"Colunas disponíveis:\n- Estoque: {df_estoque.columns.tolist()[:10]}")

# ============= TAB 3: CONSULTAS =============
with tab3:
    st.subheader("❓ Faça suas Perguntas")
    pergunta = st.text_area(label="Sua pergunta", placeholder="Ex: Qual centro tem maior demanda?", height=100)
    
    if st.button("🔍 Analisar"):
        if pergunta:
            st.info("💡 Análise em progresso...")

# ============= TAB 4: RELATÓRIO =============
with tab4:
    st.subheader("📈 Gerar Relatório Consolidado")
    
    if st.button("📊 Gerar Relatório"):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_estoque.to_excel(writer, sheet_name='Estoque', index=False)
            df_demanda.to_excel(writer, sheet_name='Demanda', index=False)
            df_separacao.to_excel(writer, sheet_name='Separações', index=False)
        
        output.seek(0)
        st.download_button(
            label="📥 Baixar Relatório Excel",
            data=output.getvalue(),
            file_name=f"relatorio_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.success("✅ Relatório pronto!")

st.markdown("---")
st.caption("🤖 IA de Transferências - Streamlit")
