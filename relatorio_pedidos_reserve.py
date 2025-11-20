import streamlit as st
import pandas as pd
import numpy as np
import io
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
# Funções de Criação do Arquivo Excel (Apenas Dados)
# ----------------------------------------------------

def to_excel(df):
    """Converte o DataFrame para um buffer de memória XLSX."""
    output = io.BytesIO()
    # Usando o Pandas para exportar para XLSX é mais estável
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Exporta apenas os dados únicos (fonte para a Pivot)
        df.to_excel(writer, sheet_name='Dados', index=False)
    
    # Retorna o conteúdo binário
    return output.getvalue()

# ----------------------------------------------------
# Leitura e Pré-Processamento (Cache Otimizado)
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
        
        # Limpeza e criação de chaves
        df[ID_COL_NAME] = df[ID_COL_NAME].astype(str).str.strip()
        df[EMP_COL_NAME] = df[EMP_COL_NAME].astype(str).str.strip()
        df[GROUP_COL_NAME] = df[GROUP_COL_NAME].astype(str).str.strip().replace(['', 'nan', 'NaN'], np.nan) 
        df[DATE_COL_NAME] = pd.to_datetime(df[DATE_COL_NAME], errors='coerce', dayfirst=True)
        df.dropna(subset=[DATE_COL_NAME], inplace=True)

        if df.empty:
            return None

        df['Entidade de Consolidação'] = df[GROUP_COL_NAME].fillna(df[EMP_COL_NAME])
        df['Mês/Ano'] = df[DATE_COL_NAME].dt.strftime('%m/%Y')
        
        # Contagem Única (PKI)
        df_pedidos_unicos = df.groupby(ID_COL_NAME).agg(
            {'Entidade de Consolidação': 'first', 'Mês/Ano': 'first'}
        ).reset_index()

        df_pedidos_unicos['PKI Pedidos'] = 1 
        
        # DataFrame final de base (somente colunas relevantes para a pivot)
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
    
    # Geração do Título Dinâmico (omissão para brevidade)
    # [Restante do código de filtros e exibição na tela...]
    
    min_date = df_base_pivot['Mês/Ano'].min()
    max_date = df_base_pivot['Mês/Ano'].max()
    dashboard_title = f"Pedidos Reserve - Período {min_date} a {max_date}"
    
    st.title("📊 Dashboard de Pedidos - Visão Matriz")
    st.markdown(f"### {dashboard_title}")
    st.markdown("---")
    
    # --- FILTROS STREAMLIT NATIVOS (OMITIDO PARA BREVIDADE, MAS EXISTE) ---
    col1, col2, col3 = st.columns([1, 1, 1])

    entidades = ['Todas'] + sorted(df_base_pivot['Entidade de Consolidação'].unique().tolist())
    entidade_selecionada = col1.selectbox('Selecione a Entidade', entidades)
    
    meses = ['Todos'] + sorted(df_base_pivot['Mês/Ano'].unique().tolist(), key=lambda x: pd.to_datetime(x, format='%m/%Y'))
    mes_selecionado = col2.selectbox('Selecione o Mês/Ano', meses)

    df_filtrado = df_base_pivot.copy() 

    if entidade_selecionada != 'Todas':
        df_filtrado = df_filtrado[df_filtrado['Entidade de Consolidação'] == entidade_selecionada]
    
    if mes_selecionado != 'Todos':
        df_filtrado = df_filtrado[df_filtrado['Mês/Ano'] == mes_selecionado]

    total_pedidos = df_filtrado['PKI Pedidos'].sum()
    col3.metric(label="Total de Pedidos Únicos", value=f"{total_pedidos:,.0f}".replace(",", "#").replace(".", ",").replace("#", "."))

    st.markdown("---")
    
    # Geração e Exibição da Tabela Pivotada (Pandas Nativo)
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
    
    # ----------------------------------------------------
    # Botão de Download NATIVO XLSX (Dados Brutos)
    # ----------------------------------------------------
    st.markdown("### 💾 Exportar Dados Brutos (Para Criar a Tabela Dinâmica no Excel)")
    
    # Cria o arquivo XLSX com a função simplificada
    xlsx_data = to_excel(df_base_pivot)

    st.download_button(
        label="Download Dados Brutos (Excel XLSX)",
        data=xlsx_data,
        file_name='relatorio_pedidos_dados_brutos.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )