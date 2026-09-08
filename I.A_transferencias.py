import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io
import os
from pathlib import Path

st.set_page_config(page_title="IA Análise de Transferências", layout="wide")

st.title("🤖 IA de Análise de Transferências de Estoque")
st.markdown("Analisa demanda, estoque e separações para otimizar transferências entre centros")

# ============= CONFIGURAÇÃO DE PASTA =============
# ALTERE AQUI PARA SUA PASTA DE RELATÓRIOS
PASTA_RELATORIOS = r"C:\Users\OLIVDT\Desktop\Relatorios_transferencia"

# Se a pasta não existe, use Desktop como padrão
if not os.path.exists(PASTA_RELATORIOS):
    PASTA_RELATORIOS = os.path.expanduser("~/Desktop")
    st.warning(f"⚠️ Pasta configurada não encontrada. Usando: {PASTA_RELATORIOS}")

st.sidebar.header("📁 Localização dos Arquivos")
st.sidebar.info(f"📂 Pasta: `{PASTA_RELATORIOS}`\n\nAbaixo estão os arquivos encontrados:")

# ============= BUSCAR ARQUIVOS AUTOMATICAMENTE =============
def buscar_arquivos(pasta, extensoes=['.xlsx', '.csv', '.xls', '.xlsm']):
    """Busca todos os arquivos Excel/CSV na pasta - VERSÃO CORRIGIDA"""
    arquivos = []
    if os.path.exists(pasta):
        for arquivo in os.listdir(pasta):
            if any(arquivo.lower().endswith(ext) for ext in extensoes):
                arquivos.append(arquivo)
    return sorted(arquivos)

# Listar arquivos disponíveis
arquivos_disponiveis = buscar_arquivos(PASTA_RELATORIOS)

if not arquivos_disponiveis:
    st.error(f"❌ Nenhum arquivo Excel ou CSV encontrado em:\n`{PASTA_RELATORIOS}`")
    st.info("📝 **Passos para resolver:**\n1. Coloque seus 3 relatórios (xlsx, xls ou csv) na pasta acima\n2. Atualize a página (F5)")
    st.stop()

st.sidebar.markdown("### Arquivos Encontrados:")
for arquivo in arquivos_disponiveis:
    st.sidebar.text(f"✓ {arquivo}")

# ============= SELEÇÃO DE ARQUIVOS =============
st.sidebar.markdown("---")
st.sidebar.header("🔗 Selecionar Arquivos")

rel_estoque_nome = st.sidebar.selectbox("📦 Relatório 1: Estoque Disponível", arquivos_disponiveis)
rel_demanda_nome = st.sidebar.selectbox("📥 Relatório 2: Demanda por Centro", arquivos_disponiveis, index=1 if len(arquivos_disponiveis) > 1 else 0)
rel_separacao_nome = st.sidebar.selectbox("📦 Relatório 3: Itens em Separação", arquivos_disponiveis, index=2 if len(arquivos_disponiveis) > 2 else 0)

# ============= CARREGANDO OS DADOS =============
def carregar_arquivo(caminho):
    """Carrega arquivo Excel ou CSV com tratamento de erro"""
    try:
        if caminho.lower().endswith('.csv'):
            return pd.read_csv(caminho)
        else:
            # Para arquivos .xls e .xlsx
            return pd.read_excel(caminho, sheet_name=0)
    except Exception as e:
        st.error(f"❌ Erro ao carregar `{os.path.basename(caminho)}`:\n{e}")
        return None

# Carrega os 3 relatórios
df_estoque = carregar_arquivo(os.path.join(PASTA_RELATORIOS, rel_estoque_nome))
df_demanda = carregar_arquivo(os.path.join(PASTA_RELATORIOS, rel_demanda_nome))
df_separacao = carregar_arquivo(os.path.join(PASTA_RELATORIOS, rel_separacao_nome))

# ============= VALIDAÇÃO DOS DADOS =============
if df_estoque is None or df_demanda is None or df_separacao is None:
    st.error("❌ Erro ao carregar um ou mais arquivos. Verifique o formato e estrutura.")
    st.stop()

st.success(f"✅ Todos os arquivos carregados com sucesso!")

# ============= EXIBIR DADOS CARREGADOS =============
with st.expander("📋 Visualizar Dados Carregados"):
    col1, col2, col3 = st.columns(3)
    with col1:
        st.subheader("Estoque Disponível")
        st.dataframe(df_estoque.head(10), use_container_width=True, height=300)
    with col2:
        st.subheader("Demanda")
        st.dataframe(df_demanda.head(10), use_container_width=True, height=300)
    with col3:
        st.subheader("Separações")
        st.dataframe(df_separacao.head(10), use_container_width=True, height=300)

# ============= ANÁLISE E RECOMENDAÇÕES =============
st.markdown("---")
st.header("🔍 Análise de Transferências")

# Tab para diferentes análises
tab1, tab2, tab3, tab4 = st.tabs(["📊 Dashboard", "🎯 Recomendações", "❓ Consultas Livres", "📈 Relatório"])

# ============= TAB 1: DASHBOARD =============
with tab1:
    try:
        col1, col2, col3, col4 = st.columns(4)
        
        # Detectar automaticamente as colunas
        col_quantidade_estoque = [col for col in df_estoque.columns if 'quantidade' in col.lower() or 'qtd' in col.lower()]
        col_quantidade_demanda = [col for col in df_demanda.columns if 'quantidade' in col.lower() or 'qtd' in col.lower()]
        col_quantidade_separacao = [col for col in df_separacao.columns if 'quantidade' in col.lower() or 'qtd' in col.lower()]
        
        col_centro_estoque = [col for col in df_estoque.columns if 'centro' in col.lower() or 'local' in col.lower()]
        col_centro_demanda = [col for col in df_demanda.columns if 'centro' in col.lower() or 'local' in col.lower()]
        
        if col_quantidade_estoque:
            total_estoque = pd.to_numeric(df_estoque[col_quantidade_estoque[0]], errors='coerce').sum()
            with col1:
                st.metric("📦 Total Estoque", f"{total_estoque:,.0f}")
        
        if col_quantidade_demanda:
            total_demanda = pd.to_numeric(df_demanda[col_quantidade_demanda[0]], errors='coerce').sum()
            with col2:
                st.metric("📥 Total Demanda", f"{total_demanda:,.0f}")
        
        if col_quantidade_separacao:
            total_sep = pd.to_numeric(df_separacao[col_quantidade_separacao[0]], errors='coerce').sum()
            with col3:
                st.metric("📦 Em Separação", f"{total_sep:,.0f}")
        
        if col_quantidade_estoque and col_quantidade_separacao:
            disponivel = total_estoque - total_sep
            with col4:
                st.metric("✅ Disponível", f"{disponivel:,.0f}")
        
        st.markdown("---")
        
        # Gráficos por centro
        col1, col2 = st.columns(2)
        
        with col1:
            if col_centro_estoque and col_quantidade_estoque:
                estoque_por_centro = pd.to_numeric(df_estoque[col_quantidade_estoque[0]], errors='coerce').groupby(df_estoque[col_centro_estoque[0]]).sum().sort_values(ascending=False)
                st.bar_chart(estoque_por_centro, height=300)
                st.caption("Estoque por Centro")
        
        with col2:
            if col_centro_demanda and col_quantidade_demanda:
                demanda_por_centro = pd.to_numeric(df_demanda[col_quantidade_demanda[0]], errors='coerce').groupby(df_demanda[col_centro_demanda[0]]).sum().sort_values(ascending=False)
                st.bar_chart(demanda_por_centro, height=300)
                st.caption("Demanda por Centro")
    
    except Exception as e:
        st.error(f"❌ Erro ao criar dashboard: {e}")

# ============= TAB 2: RECOMENDAÇÕES =============
with tab2:
    st.subheader("🎯 Recomendações de Transferência")
    
    try:
        # Detectar colunas
        col_centro = [col for col in df_estoque.columns if 'centro' in col.lower() or 'local' in col.lower()][0]
        col_qtd = [col for col in df_estoque.columns if 'quantidade' in col.lower() or 'qtd' in col.lower()][0]
        col_data = [col for col in df_estoque.columns if 'data' in col.lower() or 'entrada' in col.lower()]
        
        estoque_por_centro = pd.to_numeric(df_estoque[col_qtd], errors='coerce').groupby(df_estoque[col_centro]).sum().reset_index()
        estoque_por_centro.columns = ['Centro', 'Estoque_Total']
        
        demanda_col_qtd = [col for col in df_demanda.columns if 'quantidade' in col.lower() or 'qtd' in col.lower()][0]
        demanda_col_centro = [col for col in df_demanda.columns if 'centro' in col.lower() or 'local' in col.lower()][0]
        
        demanda_por_centro = pd.to_numeric(df_demanda[demanda_col_qtd], errors='coerce').groupby(df_demanda[demanda_col_centro]).sum().reset_index()
        demanda_por_centro.columns = ['Centro', 'Demanda_Total']
        
        separacao_col_qtd = [col for col in df_separacao.columns if 'quantidade' in col.lower() or 'qtd' in col.lower()][0]
        separacao_col_centro = [col for col in df_separacao.columns if 'origem' in col.lower() or 'centro' in col.lower()][0]
        
        separacao_por_centro = pd.to_numeric(df_separacao[separacao_col_qtd], errors='coerce').groupby(df_separacao[separacao_col_centro]).sum().reset_index()
        separacao_por_centro.columns = ['Centro', 'Em_Separacao']
        
        # Merge
        analise = estoque_por_centro.merge(demanda_por_centro, on='Centro', how='outer').fillna(0)
        analise = analise.merge(separacao_por_centro, on='Centro', how='outer').fillna(0)
        
        analise['Disponivel'] = analise['Estoque_Total'] - analise['Em_Separacao']
        analise['Deficit'] = analise['Demanda_Total'] - analise['Disponivel']
        analise['Excesso'] = analise['Disponivel'] - analise['Demanda_Total']
        
        st.markdown("### Centros com EXCESSO de Estoque")
        analise_excesso = analise[analise['Excesso'] > 0].sort_values('Excesso', ascending=False)
        if len(analise_excesso) > 0:
            st.dataframe(analise_excesso, use_container_width=True, hide_index=True)
        else:
            st.info("Nenhum centro com excesso.")
        
        st.markdown("### Centros com DÉFICIT")
        analise_deficit = analise[analise['Deficit'] > 0].sort_values('Deficit', ascending=False)
        if len(analise_deficit) > 0:
            st.dataframe(analise_deficit, use_container_width=True, hide_index=True)
        else:
            st.success("✅ Todos têm estoque suficiente!")
    
    except Exception as e:
        st.error(f"❌ Erro na análise: {e}")
        st.info(f"Colunas disponíveis:\n- Estoque: {df_estoque.columns.tolist()}\n- Demanda: {df_demanda.columns.tolist()}\n- Separação: {df_separacao.columns.tolist()}")

# ============= TAB 3: CONSULTAS =============
with tab3:
    st.subheader("❓ Faça suas Perguntas")
    pergunta = st.text_area("O que você gostaria de saber?", placeholder="Ex: Qual centro tem maior demanda?", height=100)
    
    if st.button("🔍 Analisar"):
        if pergunta:
            st.info("💡 Análise em progresso...")

# ============= TAB 4: RELATÓRIO =============
with tab4:
    st.subheader("📈 Gerar Relatório")
    
    if st.button("📊 Gerar Relatório Completo"):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_estoque.to_excel(writer, sheet_name='Estoque', index=False)
            df_demanda.to_excel(writer, sheet_name='Demanda', index=False)
            df_separacao.to_excel(writer, sheet_name='Separação', index=False)
        
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
