import streamlit as st
import pandas as pd
import numpy as np
import io
import xlsxwriter 
import base64

# --- 1. Configurações e Variáveis ---

DATE_COL_NAME = 'data' 
ID_COL_NAME = 'pedido'
GROUP_CODE_COL = 'codigo grupo' 
EMP_COL_NAME = 'empresa' 
GROUP_COL_NAME = 'nome grupo'
BASE_FILE = 'base.csv' 
SEPARATOR = ',' 

# ----------------------------------------------------
# Funções de Criação do Arquivo Excel Interativo (Mantida)
# ----------------------------------------------------

def to_excel(df):
    """Converte o DataFrame para um buffer de memória XLSX (Dados Brutos)."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Dados', index=False)
    return output.getvalue()

# ----------------------------------------------------
# Leitura e Pré-Processamento (Cache Otimizado) (Mantida)
# ----------------------------------------------------

@st.cache_data
def load_and_clean_data():
    """Lê, limpa, e pré-processa os dados base, gerando a base de pedidos únicos."""
    try:
        df = pd.read_csv(
            BASE_FILE, 
            sep=SEPARATOR, 
            header=None,  
            skiprows=1,   
            names=[DATE_COL_NAME, ID_COL_NAME, GROUP_CODE_COL, EMP_COL_NAME, GROUP_COL_NAME],
            encoding='utf-8'
        )
        
        df[ID_COL_NAME] = df[ID_COL_NAME].astype(str).str.strip()
        df[EMP_COL_NAME] = df[EMP_COL_NAME].astype(str).str.strip()
        df[GROUP_COL_NAME] = df[GROUP_COL_NAME].astype(str).str.strip().replace(['', 'nan', 'NaN'], np.nan) 
        df[DATE_COL_NAME] = pd.to_datetime(df[DATE_COL_NAME], errors='coerce', dayfirst=True)
        df.dropna(subset=[DATE_COL_NAME], inplace=True)

        if df.empty:
            return None

        df['Entidade de Consolidação'] = df[GROUP_COL_NAME].fillna(df[EMP_COL_NAME])
        df['Mês/Ano'] = df[DATE_COL_NAME].dt.strftime('%m/%Y')
        
        df_pedidos_unicos = df.groupby(ID_COL_NAME).agg(
            {'Entidade de Consolidação': 'first', 'Mês/Ano': 'first'}
        ).reset_index()

        df_pedidos_unicos['PKI Pedidos'] = 1 
        
        df_base_pivot = df_pedidos_unicos[['Entidade de Consolidação', 'Mês/Ano', 'PKI Pedidos']]
        
        return df_base_pivot

    except FileNotFoundError:
        st.error(f"❌ ERRO FATAL: O arquivo '{BASE_FILE}' não foi encontrado.")
        return None
    except Exception as e:
        st.error(f"❌ ERRO FATAL ao processar o arquivo. Detalhe: {e}")
        return None

# ----------------------------------------------------
# --- 2. Interface Streamlit ---
# ----------------------------------------------------

st.set_page_config(layout="wide", page_title="Dashboard Pedidos Reserve")

df_base_pivot = load_and_clean_data()

if df_base_pivot is not None:
    
    # Geração do Título Dinâmico
    min_date = df_base_pivot['Mês/Ano'].min()
    max_date = df_base_pivot['Mês/Ano'].max()
    dashboard_title = f"Pedidos Reserve - Período {min_date} a {max_date}"
    
    st.title("📊 Dashboard de Pedidos - Visão Matriz")
    st.markdown(f"### {dashboard_title}")
    st.markdown("---")
    
    # --- FILTROS STREAMLIT NATIVOS ---
    col1, col2, col3 = st.columns([1, 1, 1])

    entidades = ['Todas'] + sorted(df_base_pivot['Entidade de Consolidação'].unique().tolist())
    entidade_selecionada = col1.selectbox('Selecione a Entidade', entidades)
    
    meses = ['Todos'] + sorted(df_base_pivot['Mês/Ano'].unique().tolist(), key=lambda x: pd.to_datetime(x, format='%m/%Y'))
    mes_selecionado = col2.selectbox('Selecione o Mês/Ano', meses)

    # Lógica de Filtragem
    df_filtrado = df_base_pivot.copy() 

    if entidade_selecionada != 'Todas':
        df_filtrado = df_filtrado[df_filtrado['Entidade de Consolidação'] == entidade_selecionada]
    
    if mes_selecionado != 'Todos':
        df_filtrado = df_filtrado[df_filtrado['Mês/Ano'] == mes_selecionado]

    # Recalcula os totais (KPI Principal)
    total_pedidos = df_filtrado['PKI Pedidos'].sum()
    col3.metric(label="Total de Pedidos Únicos", value=f"{total_pedidos:,.0f}".replace(",", "#").replace(".", ",").replace("#", "."))

    st.markdown("---")
    
    # ====================================================
    # NOVO BLOCO 1: FRAMES DE TOTAIS POR MÊS (Mantido)
    # ====================================================

    if not df_filtrado.empty:
        st.subheader("🚀 Total de Pedidos por Mês (KPIs Dinâmicos)")

        df_monthly_totals = df_filtrado.groupby('Mês/Ano')['PKI Pedidos'].sum().reset_index()
        df_monthly_totals.columns = ['Mês/Ano', 'Total Pedidos']
        
        df_monthly_totals['Data Ordenacao'] = pd.to_datetime(df_monthly_totals['Mês/Ano'], format='%m/%Y')
        df_monthly_totals = df_monthly_totals.sort_values('Data Ordenacao').drop(columns='Data Ordenacao')
        
        num_months = len(df_monthly_totals)
        cols_per_row = 6 
        
        for i in range(0, num_months, cols_per_row):
            current_months = df_monthly_totals.iloc[i:i + cols_per_row]
            cols = st.columns(len(current_months))
            
            for j, row in current_months.iterrows():
                month = row['Mês/Ano']
                total = row['Total Pedidos']
                
                cols[current_months.index.get_loc(j)].metric(
                    label=f"Total em {month}",
                    value=f"{total:,.0f}".replace(",", "#").replace(".", ",").replace("#", ".")
                )

        st.markdown("---")

        # ====================================================
        # ✅ NOVO BLOCO 2: TOP 3 ENTIDADES POR MÊS (COLUNAS)
        # ====================================================
        
        st.subheader("🏆 Top 3 Entidades por Mês")

        # 1. Agrupar dados por Mês e Entidade
        df_monthly_entity = df_filtrado.groupby(['Mês/Ano', 'Entidade de Consolidação'])['PKI Pedidos'].sum().reset_index()
        df_monthly_entity.columns = ['Mês/Ano', 'Entidade', 'Total Pedidos']

        month_order = df_monthly_totals['Mês/Ano'].tolist() # Usa a ordem de meses já calculada

        # 2. Definir o layout de colunas
        # Usaremos no máximo 4 colunas por linha para o Top 3 ficar legível
        cols_per_row_top3 = 4
        num_months_top3 = len(month_order)
        
        for i in range(0, num_months_top3, cols_per_row_top3):
            # Seleciona os meses para a linha atual
            current_month_batch = month_order[i:i + cols_per_row_top3]
            
            # Cria as colunas Streamlit
            cols = st.columns(len(current_month_batch))
            
            for index, month in enumerate(current_month_batch):
                
                # Filtra dados para o mês atual
                df_month = df_monthly_entity[df_monthly_entity['Mês/Ano'] == month]
                
                # Ordena e pega o Top 3
                df_top3 = df_month.sort_values(by='Total Pedidos', ascending=False).head(3)
                
                # Formata o DataFrame para o display
                df_top3_display = df_top3[['Entidade', 'Total Pedidos']].copy()
                df_top3_display['Total Pedidos'] = df_top3_display['Total Pedidos'].apply(lambda x: f"{x:,.0f}".replace(",", "#").replace(".", ",").replace("#", "."))
                
                # Exibir na coluna atual
                with cols[index]:
                    st.markdown(f"**{month}**")
                    st.dataframe(df_top3_display, 
                                 use_container_width=True, 
                                 hide_index=True,
                                 # Define altura fixa para que colunas com menos de 3 itens não desequilibrem
                                 height=180) 

        st.markdown("---")
        
    # Geração e Exibição da Tabela Pivotada (Pandas Nativo) (Mantida)
    # [Restante do código da Tabela Pivotada]

    if df_filtrado.empty:
        st.warning("Nenhum dado encontrado para a combinação de filtros selecionada.")
        df_pivot_final = pd.DataFrame() 
    else:
        df_pivot_final = pd.pivot_table(
            df_filtrado,
            index=['Entidade de Consolidação'], 
            columns=['Mês/Ano'], 
            values=['PKI Pedidos'], 
            aggfunc='sum',
            fill_value=0, 
            margins=True, 
            margins_name='Total Geral'
        )

        df_pivot_final.columns = df_pivot_final.columns.get_level_values(1)

        st.subheader("Tabela de Pedidos - Entidades por Mês/Ano")
        
        st.dataframe(
            df_pivot_final.style.format("{:,.0f}").background_gradient(cmap='Blues'), 
            use_container_width=True
        )


    st.markdown("---")
    
    # Botão de Download NATIVO XLSX (Dados Brutos) (Mantido)
    st.markdown("### 💾 Exportar Dados Brutos (Para Criar a Tabela Dinâmica no Excel)")
    
    xlsx_data = to_excel(df_base_pivot)

    st.download_button(
        label="Download Dados Brutos (Excel XLSX)",
        data=xlsx_data,
        file_name='relatorio_pedidos_dados_brutos.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )