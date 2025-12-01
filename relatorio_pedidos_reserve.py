# VERSÃO 10.1 - DASHBOARD COMPLETO COM INCREMENTO E FILTRO DE SISTEMA CORRIGIDO (FINAL)

import streamlit as st
import pandas as pd
import numpy as np
import io
import xlsxwriter
import base64
import os
from datetime import datetime

# --- 1. Configurações e Variáveis ---

# Nomes de Colunas PADRÕES para unificação
DATE_COL_NAME = 'data'
ID_COL_NAME = 'pedido' 
GROUP_CODE_COL = 'codigo grupo'
EMP_COL_NAME = 'empresa'
GROUP_COL_NAME = 'nome grupo'
SYSTEM_COL_NAME = 'Sistema'

# Arquivos de Entrada
BASE_RESERVE_FILE = 'base.xlsx'
LOGO_FILE = 'logo.png' 
MAX_LOGO_HEIGHT = '80px'

# Arquivo de Saída Consolidado
CONSOLIDATED_FILE = 'base_consolidada.xlsx'

# Arquivos ARGOIT (ASSUMIR que estão na mesma pasta do script)
ARGOIT_FILES = {
    '07/2025': 'ARGO-JULHO-25.xlsx',
    '08/2025': 'ARGO-AGOSTO-25.xlsx',
    '09/2025': 'ARGO-SETEMBRO-25.xlsx',
    '10/2025': 'ARGO-OUTUBRO-25.xlsx'
}

# Constantes para o mapeamento de Grupos (Reserve)
GRUPO_SHEET_NAME = 'GRUPOS'
GRUPO_MAPPING_CODE_COL = 'Codigo'
GRUPO_MAPPING_NAME_COL = 'Nome do Grupo'

# --- DEFINIÇÃO DE CORES ---
ORANGE_COLOR = '#ff8c00' # Laranja, usado para Reserve e para o estilo principal
RESERVE_COLOR = ORANGE_COLOR # Cor específica para Reserve
ARGOIT_COLOR = '#FFD700' # Amarelo Ouro, para ARGOIT
BACKGROUND_COLOR_DARK_BLUE = '#131B36'
CONTRAST_BACKGROUND_COLOR = '#1D2A4A'
DARK_BACKGROUND_COLOR = CONTRAST_BACKGROUND_COLOR

# ----------------------------------------------------
# Funções Auxiliares de Exportação e Imagem
# ----------------------------------------------------

def to_excel(df):
    """Converte o DataFrame para um buffer de memória XLSX (Dados Brutos)."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Consolidado', index=False)
    return output.getvalue()

def to_excel_styled(df_pivot):
    """Converte o DataFrame Pivotado para um buffer de memória XLSX aplicando estilos de totais."""
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    sheet_name = 'Tabela_Pivotada'
    df_pivot.to_excel(writer, sheet_name=sheet_name, index=True)

    workbook = writer.book
    worksheet = writer.sheets[sheet_name]

    # Formatos de cor
    header_format = workbook.add_format({
        'bold': True, 'text_wrap': True, 'valign': 'top', 
        'fg_color': ORANGE_COLOR, 'border': 1, 'font_color': 'white'
    })

    total_format = workbook.add_format({
        'bold': True, 'fg_color': ORANGE_COLOR, 'border': 1, 
        'font_color': 'white', 'num_format': '#,##0' 
    })
    
    content_format_even = workbook.add_format({
        'fg_color': 'white', 'border': 1, 'font_color': 'black', 'num_format': '#,##0'
    })
    
    content_format_odd = workbook.add_format({
        'fg_color': '#f0f2f6', 'border': 1, 'font_color': 'black', 'num_format': '#,##0'
    })

    # Obter dimensões e índice
    num_rows, num_cols = df_pivot.shape
    index_cols = len(df_pivot.index.names)

    # 1. Aplicar estilo ao cabeçalho (Colunas)
    for col_num, value in enumerate(df_pivot.columns.values):
        worksheet.write(index_cols, col_num + index_cols, value, header_format) 
    
    # 2. Aplicar estilo às células de DADOS, LINHA DE TOTAIS e Índice
    for row_num, (index, row) in enumerate(df_pivot.iterrows()):
        is_total_row = (row_num == num_rows - 1)
        
        # Células de Dados (Meses/Ano - Exceto a última coluna de Total Geral)
        for col_num in range(num_cols - 1): 
            cell_format = total_format if is_total_row else (content_format_even if row_num % 2 == 0 else content_format_odd)
            worksheet.write(row_num + index_cols + 1, col_num + index_cols, row.iloc[col_num], cell_format)
            
        # Célula de Total GERAL (Última Coluna)
        worksheet.write(row_num + index_cols + 1, num_cols + index_cols - 1, row.iloc[-1], total_format)
        
        # Células de Índice (Linhas)
        for i in range(index_cols):
            index_value = index[i] if index_cols > 1 else index
            worksheet.write(row_num + index_cols + 1, i, index_value, header_format)
        
    # Aplicar o formato de cabeçalho ao nome do índice (canto superior esquerdo)
    for i in range(index_cols):
        worksheet.write(i, i, df_pivot.index.names[i], header_format)

    # Definir formato de Total Geral no canto inferior direito
    if num_rows > 0 and num_cols > 0:
        worksheet.write(num_rows + index_cols, num_cols + index_cols - 1, df_pivot.iloc[-1, -1], total_format)

    writer.close()
    return output.getvalue()


def image_to_base64(file_path, file_type="png"):
    """Lê um arquivo de imagem (PNG) e codifica em Base64 para HTML."""
    if not os.path.exists(file_path):
        return None, f"O arquivo {file_path} não foi encontrado."
    try:
        with open(file_path, "rb") as f:
            image_bytes = f.read()
        b64_encoded = base64.b64encode(image_bytes).decode('utf-8')
        return f"data:image/{file_type};base64,{b64_encoded}", None
    except Exception as e:
        return None, f"Erro ao processar a imagem: {e}"

# ----------------------------------------------------
# Leitura e Padronização das Bases de ORIGEM
# ----------------------------------------------------

def load_reserve_data(file_path):
    """Lê a base Reserve e a tabela de grupos, e atribui o nome do sistema."""
    try:
        df = pd.read_excel(
            file_path, sheet_name='base', header=None, skiprows=1,
            names=[DATE_COL_NAME, ID_COL_NAME, GROUP_CODE_COL, EMP_COL_NAME, GROUP_COL_NAME],
            engine='openpyxl'
        )
        df_grupos = pd.read_excel(
            file_path, sheet_name=GRUPO_SHEET_NAME,
            usecols=[GRUPO_MAPPING_CODE_COL, GRUPO_MAPPING_NAME_COL],
            engine='openpyxl'
        )
        df_grupos.rename(
            columns={GRUPO_MAPPING_CODE_COL: 'merge_key', GRUPO_MAPPING_NAME_COL: 'Nome_Grupo_Mapeado'},
            inplace=True
        )
        df_grupos['merge_key'] = df_grupos['merge_key'].apply(
            lambda x: str(int(x)) if pd.notna(x) and str(x).replace('.', '', 1).isdigit() else str(x)
        ).str.strip()
        df['merge_key'] = df[GROUP_CODE_COL].apply(
            lambda x: str(int(x)) if pd.notna(x) and str(x).replace('.', '', 1).isdigit() else str(x)
        ).str.strip()
        df = pd.merge(df, df_grupos[['merge_key', 'Nome_Grupo_Mapeado']], on='merge_key', how='left')
        df[GROUP_COL_NAME] = df['Nome_Grupo_Mapeado'].fillna(df[GROUP_COL_NAME])
        df[SYSTEM_COL_NAME] = 'Reserve'
        df = df[[DATE_COL_NAME, ID_COL_NAME, EMP_COL_NAME, GROUP_COL_NAME, SYSTEM_COL_NAME]].copy()
        return df, None
    except FileNotFoundError:
        return pd.DataFrame(), f"O arquivo '{file_path}' (Reserve) não foi encontrado."
    except Exception as e:
        return pd.DataFrame(), f"Erro grave ao processar base Reserve: {e}"

def load_argoit_data(file_map):
    """Lê e concatena os arquivos mensais do ARGOIT."""
    all_argoit_data = []
    ARGOIT_MAPPING = {
        'Data Inclusao': DATE_COL_NAME, 'Numero da Solicitacao': ID_COL_NAME, 
        'Empresa de Débito': EMP_COL_NAME, 'Cliente': GROUP_COL_NAME,
    }
    COLS_TO_READ = list(ARGOIT_MAPPING.keys())

    for month_year, file_path in file_map.items():
        if not os.path.exists(file_path):
            st.warning(f"⚠️ Aviso: Arquivo ARGOIT '{file_path}' para {month_year} não encontrado. Pulando.")
            continue
        try:
            # header=1 pois a linha 1 (índice 0) é vazia e a linha 2 (índice 1) contém o cabeçalho
            df_month = pd.read_excel(file_path, header=1, usecols=COLS_TO_READ, engine='openpyxl')
            df_month.rename(columns=ARGOIT_MAPPING, inplace=True)
            # A data precisa ser lida corretamente, garantindo o formato dia/mês/ano
            df_month[DATE_COL_NAME] = pd.to_datetime(df_month[DATE_COL_NAME], errors='coerce', dayfirst=True) 
            df_month.dropna(subset=[DATE_COL_NAME], inplace=True)
            if df_month.empty: 
                st.info(f"O arquivo '{file_path}' (ARGOIT) foi lido, mas está vazio após a limpeza de datas. Pulando.")
                continue
            df_month[SYSTEM_COL_NAME] = 'ARGOIT'
            required_cols = [DATE_COL_NAME, ID_COL_NAME, EMP_COL_NAME, GROUP_COL_NAME, SYSTEM_COL_NAME]
            df_month = df_month[required_cols].copy()
            all_argoit_data.append(df_month)
        except Exception as e:
            st.error(f"❌ Erro ao ler arquivo ARGOIT '{file_path}': {type(e).__name__} - {e}")
            continue

    if not all_argoit_data:
        return pd.DataFrame(), "Nenhum arquivo ARGOIT válido foi carregado após a tentativa de leitura de todos os arquivos."

    df_argoit_combined = pd.concat(all_argoit_data, ignore_index=True)
    return df_argoit_combined, None

# ----------------------------------------------------
# CRIAÇÃO DA BASE CONSOLIDADA COM INCREMENTO
# ----------------------------------------------------

def create_and_save_consolidated_base():
    """Implementa a lógica de incremento: lê existente, adiciona só os novos pedidos, e salva."""
    
    st.info("🔄 Criando e limpando a base consolidada (base_consolidada.xlsx). Isso pode levar alguns segundos...")
    
    # 1. CARREGAR A BASE CONSOLIDADA EXISTENTE
    df_existing = pd.DataFrame()
    initial_rows_existing = 0
    try:
        if os.path.exists(CONSOLIDATED_FILE):
            df_existing = pd.read_excel(CONSOLIDATED_FILE, engine='openpyxl', sheet_name='Consolidado')
            initial_rows_existing = len(df_existing)
            st.success(f"✅ Base consolidada existente carregada com sucesso. ({initial_rows_existing} linhas iniciais)")
            if not df_existing.empty:
                df_existing[ID_COL_NAME] = df_existing[ID_COL_NAME].astype(str).str.strip()
        else:
            st.info("ℹ️ Arquivo consolidado não encontrado. Será criado do zero a partir dos dados de origem.")
    except Exception as e:
        st.error(f"❌ Erro ao ler base consolidada existente. Será tratada como nova. Detalhe: {e}")
        df_existing = pd.DataFrame()

    # 2. CARREGAR NOVOS DADOS (RAW)
    df_reserve, error_r = load_reserve_data(BASE_RESERVE_FILE)
    df_argoit, error_a = load_argoit_data(ARGOIT_FILES)
    if error_r: st.warning(f"Aviso Reserve: {error_r}")
    if error_a: st.warning(f"Aviso ARGOIT: {error_a}")
    
    st.write(f"Linhas carregadas do Reserve: **{len(df_reserve):,.0f}**")
    st.write(f"Linhas carregadas do ARGOIT: **{len(df_argoit):,.0f}**")
    
    df_new_raw_combined = pd.concat([df_reserve, df_argoit], ignore_index=True)

    if df_new_raw_combined.empty:
        st.warning("Nenhuma linha válida encontrada nos arquivos de origem.")
        df_final_consolidated = df_existing 
    
    else:
        # 3. LIMPEZA E DEDUPLICAÇÃO INTERNA DO NOVO RAW
        df_new_raw_combined[DATE_COL_NAME] = pd.to_datetime(df_new_raw_combined[DATE_COL_NAME], errors='coerce', dayfirst=True)
        df_new_raw_combined.dropna(subset=[DATE_COL_NAME], inplace=True)
        df_new_raw_combined[ID_COL_NAME] = df_new_raw_combined[ID_COL_NAME].astype(str).str.strip()
        df_new_raw_combined[EMP_COL_NAME] = df_new_raw_combined[EMP_COL_NAME].astype(str).str.strip()
        df_new_raw_combined[GROUP_COL_NAME] = df_new_raw_combined[GROUP_COL_NAME].astype(str).str.strip().replace(['', 'nan', 'NaN'], np.nan)
        
        df_new_unique = df_new_raw_combined.drop_duplicates(subset=[ID_COL_NAME], keep='first')
        
        # 4. IDENTIFICAR PEDIDOS FALTANTES (INCREMENTO)
        existing_ids = set(df_existing[ID_COL_NAME].unique()) if not df_existing.empty else set()
        df_to_append = df_new_unique[~df_new_unique[ID_COL_NAME].isin(existing_ids)]
        
        # 5. CONSOLIDAR (JUNTAR)
        df_final_consolidated = pd.concat([df_existing, df_to_append], ignore_index=True)
        
        st.write(f"Linhas carregadas dos arquivos de origem (Raw Data, após deduplicação): **{len(df_new_unique):,.0f}**")
        st.write(f"Pedidos **NOVOS** para adicionar à base existente: **{len(df_to_append):,.0f}**")


    # 6. SALVAR O ARQUIVO CONSOLIDADO (Base Bruta FINAL)
    total_rows_final = len(df_final_consolidated)
    
    # Salva apenas se o arquivo foi incrementado ou se foi criado do zero
    if total_rows_final > initial_rows_existing or initial_rows_existing == 0:
        try:
            # Garantimos que o DF final, antes de salvar, esteja limpo de duplicatas
            df_to_save = df_final_consolidated.drop_duplicates(subset=[ID_COL_NAME], keep='first').copy()
            df_to_save.to_excel(CONSOLIDATED_FILE, index=False, engine='xlsxwriter', sheet_name='Consolidado')
            st.success(f"✅ Base consolidada **ATUALIZADA** salva com sucesso em **`{CONSOLIDATED_FILE}`**. Total de pedidos: {len(df_to_save):,.0f}")
            df_final_consolidated = df_to_save # Atualiza a variável com a versão salva e limpa
        except Exception as e:
            st.error(f"❌ Erro ao salvar o arquivo consolidado. Verifique se ele não está aberto. Detalhe: {e}")
            return pd.DataFrame()
    else:
        st.info(f"ℹ️ Base consolidada não foi alterada. Nenhum pedido novo encontrado. Total de pedidos: {total_rows_final:,.0f}")
            
    # 7. Retorna o DF final (LIMPO E ÚNICO POR PEDIDO)
    df_unique_final = df_final_consolidated.drop_duplicates(subset=[ID_COL_NAME], keep='first').copy()
    
    return df_unique_final # Retorna a base final limpa e pronta para o dashboard

# ----------------------------------------------------
# Leitura e Pré-processamento (Cache Otimizado)
# ----------------------------------------------------

@st.cache_data
def load_and_clean_data():
    """
    Tenta carregar a base consolidada e realiza o pré-processamento para o pivotamento.
    Retorna o DF final (limpo) e um DF pronto para pivotar.
    """
    df_final_consolidated = create_and_save_consolidated_base()

    if df_final_consolidated.empty:
        return None, None

    # 1. PRÉ-PROCESSAMENTO PARA O DASHBOARD (Criamos as colunas de Entidade e Mês/Ano)
    df_final_consolidated['Entidade de Consolidação'] = df_final_consolidated[GROUP_COL_NAME].fillna(df_final_consolidated[EMP_COL_NAME])
    df_final_consolidated['Mês/Ano'] = df_final_consolidated[DATE_COL_NAME].dt.strftime('%m/%Y')
    df_final_consolidated['PKI Pedidos'] = 1
    
    # df_base_pivot é a base que será usada para todos os cálculos e visualizações
    df_base_pivot = df_final_consolidated[['Entidade de Consolidação', 'Mês/Ano', 'PKI Pedidos', SYSTEM_COL_NAME, ID_COL_NAME]]
    
    return df_final_consolidated, df_base_pivot # Retorna a base completa e a base para pivotar

# ----------------------------------------------------
# --- 2. Interface Streamlit ---
# ----------------------------------------------------

st.set_page_config(layout="wide", page_title="Dashboard Pedidos Consolidado")

# Aplica o CSS (Mantido)
st.markdown(
    f"""
    <style>
    .stApp {{ background-color: {BACKGROUND_COLOR_DARK_BLUE}; color: white; }}
    h1, h2, h3, h4, h5, h6, .stMarkdown, label, [data-testid="stMetricLabel"] {{ color: white !important; }}
    [data-testid="column"] {{ display: flex; flex-direction: column; justify-content: center; }}
    h1 {{ margin-top: 0px !important; }}
    .custom-logo-img {{ width: auto !important; height: 100% !important; max-height: {MAX_LOGO_HEIGHT} !important; object-fit: contain; margin: 0px auto; }}
    .logo-container {{ display: flex; align-items: center; justify-content: center; height: {MAX_LOGO_HEIGHT}; }}
    
    /* CSS de Filtros e KPI */
    div[data-testid="stVerticalBlock"]:nth-of-type(1) > div:nth-child(1) {{ background-color: {CONTRAST_BACKGROUND_COLOR}; padding: 15px 20px 5px 20px; border-radius: 10px; color: white; margin-bottom: 20px; }}
    div[data-testid="stVerticalBlock"]:nth-of-type(1) > div:nth-child(1) [data-testid="stMetricLabel"] {{ color: white !important; text-align: center; width: 100%; display: block; }}
    div[data-testid="stVerticalBlock"]:nth-of-type(1) > div:nth-child(1) [data-testid="stMetricValue"] {{ color: {ORANGE_COLOR} !important; font-size: 3em !important; text-align: center; width: 100%; display: block; }}
    
    /* Novo estilo para KPIs menores */
    .metric-small {{ background-color: {BACKGROUND_COLOR_DARK_BLUE}; border: 1px solid {CONTRAST_BACKGROUND_COLOR}; padding: 10px; border-radius: 8px; margin-bottom: 10px;}}
    .metric-small [data-testid="stMetricLabel"] {{ color: #a0a0a0 !important; font-size: 0.9em !important; }}
    .metric-small [data-testid="stMetricValue"] {{ color: white !important; font-size: 1.5em !important; }}

    /* Estilos para os novos quadros mensais particionados */
    .kpi-box-reserve {{ 
        background-color: {RESERVE_COLOR}; border-radius: 10px; padding: 10px; text-align: center;
        box-shadow: 0 1px 2px rgba(0,0,0,0.1); border: 2px solid {BACKGROUND_COLOR_DARK_BLUE}; margin-bottom: 10px;
    }}
    .kpi-box-argoit {{ 
        background-color: {ARGOIT_COLOR}; border-radius: 10px; padding: 10px; text-align: center;
        box-shadow: 0 1px 2px rgba(0,0,0,0.1); border: 2px solid {BACKGROUND_COLOR_DARK_BLUE}; margin-bottom: 10px;
    }}
    .kpi-box-reserve p, .kpi-box-argoit p {{ color: white; margin: 0; font-size: 1.0em; font-weight: bold;}}
    .kpi-box-reserve h2, .kpi-box-argoit h2 {{ color: {BACKGROUND_COLOR_DARK_BLUE}; margin: 5px 0 0 0; font-size: 2.0em;}}

    </style>
    """, 
    unsafe_allow_html=True
)

# Carrega a base completa e a base limpa para pivotar/dashboard
df_final_consolidated, df_base_pivot = load_and_clean_data()

# --- INÍCIO DO DASHBOARD ---
if df_base_pivot is not None and not df_base_pivot.empty:
    
    min_date = df_base_pivot['Mês/Ano'].min()
    max_date = df_base_pivot['Mês/Ano'].max()
    dashboard_title = f"Pedidos Consolidado (Reserve + ARGOIT) - Período {min_date} a {max_date}"
    
    # Cabeçalho (Mantido)
    logo_col, title_col = st.columns([1, 4])
    
    try:
        img_base64_data, error = image_to_base64(LOGO_FILE, file_type="png")
        if img_base64_data:
            with logo_col:
                st.markdown(
                    f"""
                    <div class="logo-container">
                        <img src="{img_base64_data}" class="custom-logo-img" alt="Logomarca">
                    </div>
                    """, unsafe_allow_html=True
                )
    except:
        with logo_col: st.warning("Logo não carregada.")
            
    with title_col:
        st.markdown(f"<h1>📊 Dashboard de Pedidos - Visão Consolidada (V10.1)</h1>", unsafe_allow_html=True)
        st.markdown(f"### {dashboard_title}")
    
    st.markdown("---")
    
    # ====================================================
    # BLOCO 2: FILTROS E KPI PRINCIPAL
    # ====================================================
    
    with st.container():
        col1, col2, col3, col4_total, col5_reserve, col6_argoit = st.columns([1, 1, 1, 1, 1, 1])

        entidades = ['Todas'] + sorted(df_base_pivot['Entidade de Consolidação'].unique().tolist())
        entidade_selecionada = col1.selectbox('Selecione a Entidade', entidades, key='entidade_filtro')
        
        meses = ['Todos'] + sorted(df_base_pivot['Mês/Ano'].unique().tolist(), key=lambda x: pd.to_datetime(x, format='%m/%Y'))
        mes_selecionado = col2.selectbox('Selecione o Mês/Ano', meses, key='mes_filtro')
        
        sistemas = ['Todos'] + sorted(df_base_pivot[SYSTEM_COL_NAME].unique().tolist())
        sistema_selecionado = col3.selectbox('Selecione o Sistema', sistemas, key='sistema_filtro')

        # DF BASE: Aplicar filtros de Entidade e Mês/Ano
        df_base_filtrada = df_base_pivot.copy()
        
        if entidade_selecionada != 'Todas':
            df_base_filtrada = df_base_filtrada[df_base_filtrada['Entidade de Consolidação'] == entidade_selecionada]
        
        if mes_selecionado != 'Todos':
            df_base_filtrada = df_base_filtrada[df_base_filtrada['Mês/Ano'] == mes_selecionado]
            
        
        # DF VISUAL: Aplicar filtro de Sistema (Este DF é usado no KPI principal, Pivot, Leaderboard)
        if sistema_selecionado != 'Todos':
            df_visual_filtrada = df_base_filtrada[df_base_filtrada[SYSTEM_COL_NAME] == sistema_selecionado]
        else:
            df_visual_filtrada = df_base_filtrada
            
        
        # --- CÁLCULO DOS KPIS ---
        total_pedidos = df_visual_filtrada['PKI Pedidos'].sum()
        
        # O cálculo de Reserve e ARGOIT usa o DF filtrado apenas por Entidade e Mês/Ano (para mostrar o total real consolidado)
        total_reserve = df_base_filtrada[df_base_filtrada[SYSTEM_COL_NAME] == 'Reserve']['PKI Pedidos'].sum()
        total_argoit = df_base_filtrada[df_base_filtrada[SYSTEM_COL_NAME] == 'ARGOIT']['PKI Pedidos'].sum()
        
        
        # --- EXIBIÇÃO DOS KPIS ---
        
        # Função para formatar números
        def format_number(value):
            return f"{value:,.0f}".replace(",", "#").replace(".", ",").replace("#", ".")

        with col4_total:
            # Mostra o total do que está VISÍVEL após todos os filtros (incluindo o Sistema)
            st.metric(label="Total de Pedidos Únicos", value=format_number(total_pedidos))

        with col5_reserve:
            # Mostra o total de Reserve, APENAS pelos filtros de Entidade e Mês/Ano
            st.markdown(
                f"""
                <div class="metric-small">
                    <p data-testid="stMetricLabel">Total Reserve</p>
                    <p data-testid="stMetricValue">{format_number(total_reserve)}</p>
                </div>
                """, unsafe_allow_html=True
            )

        with col6_argoit:
            # Mostra o total de ARGOIT, APENAS pelos filtros de Entidade e Mês/Ano
            st.markdown(
                f"""
                <div class="metric-small">
                    <p data-testid="stMetricLabel">Total ARGOIT</p>
                    <p data-testid="stMetricValue">{format_number(total_argoit)}</p>
                </div>
                """, unsafe_allow_html=True
            )

    st.markdown("---")
    
    # ====================================================
    # BLOCO 1: FRAMES DE TOTAIS POR MÊS (PARTICIONADO POR SISTEMA)
    # ====================================================

    if not df_visual_filtrada.empty:
        st.subheader("🚀 Total de Pedidos por Mês (KPIs Dinâmicos)")

        # Usamos df_visual_filtrada (já filtrado por sistema, se aplicável)
        df_monthly_systems = df_visual_filtrada.groupby(['Mês/Ano', SYSTEM_COL_NAME])['PKI Pedidos'].sum().unstack(fill_value=0).reset_index()
        
        if 'Reserve' not in df_monthly_systems.columns: df_monthly_systems['Reserve'] = 0
        if 'ARGOIT' not in df_monthly_systems.columns: df_monthly_systems['ARGOIT'] = 0
        
        df_monthly_systems['Data Ordenacao'] = pd.to_datetime(df_monthly_systems['Mês/Ano'], format='%m/%Y')
        df_monthly_systems = df_monthly_systems.sort_values('Data Ordenacao').drop(columns='Data Ordenacao')
        
        month_order = df_monthly_systems['Mês/Ano'].tolist()
        cols_per_row = 4
        num_months = len(month_order)
        
        # Exibe Reserve apenas se o filtro de sistema não for "ARGOIT"
        if sistema_selecionado != 'ARGOIT':
            st.markdown("#### Total Reserve")
            for i in range(0, num_months, cols_per_row):
                current_months = df_monthly_systems[df_monthly_systems['Mês/Ano'].isin(month_order[i:i + cols_per_row])]
                cols = st.columns(len(current_months))
                for j, row in current_months.iterrows():
                    formatted_value = format_number(row['Reserve'])
                    with cols[current_months.index.get_loc(j)]:
                        st.markdown(
                            f"""
                            <div class="kpi-box-reserve">
                                <p>{row['Mês/Ano']}</p>
                                <h2>{formatted_value}</h2>
                            </div>
                            """, unsafe_allow_html=True
                        )

        # Exibe ARGOIT apenas se o filtro de sistema não for "Reserve"
        if sistema_selecionado != 'Reserve':
            st.markdown("#### Total ARGOIT")
            for i in range(0, num_months, cols_per_row):
                current_months = df_monthly_systems[df_monthly_systems['Mês/Ano'].isin(month_order[i:i + cols_per_row])]
                cols = st.columns(len(current_months))
                for j, row in current_months.iterrows():
                    formatted_value = format_number(row['ARGOIT'])
                    with cols[current_months.index.get_loc(j)]:
                        st.markdown(
                            f"""
                            <div class="kpi-box-argoit">
                                <p>{row['Mês/Ano']}</p>
                                <h2>{formatted_value}</h2>
                            </div>
                            """, unsafe_allow_html=True
                        )

        st.markdown("---")

        # ====================================================
        # BLOCO 2: TOP 3 ENTIDADES POR MÊS 
        # ====================================================
        
        st.subheader("🏆 Top 3 Entidades (Leaderboard Mensal por Quantidade)")

        # df_visual_filtrada está filtrado por entidade, mês e sistema (se aplicável)
        df_monthly_entity = df_visual_filtrada.groupby(['Mês/Ano', 'Entidade de Consolidação', SYSTEM_COL_NAME])['PKI Pedidos'].sum().reset_index()
        df_monthly_entity.columns = ['Mês/Ano', 'Entidade', 'Sistema', 'Total Pedidos']
        
        # O ranking deve ser sempre baseado no total da entidade (soma dos sistemas)
        df_rank = df_monthly_entity.groupby(['Mês/Ano', 'Entidade'])['Total Pedidos'].sum().reset_index()
        df_rank.columns = ['Mês/Ano', 'Entidade', 'Total Rank']


        cols_per_row_top3 = 4
        num_months_top3 = len(month_order)
        
        for i in range(0, num_months_top3, cols_per_row_top3):
            current_month_batch = month_order[i:i + cols_per_row_top3]
            cols = st.columns(len(current_month_batch))
            
            for index, month in enumerate(current_month_batch):
                
                with cols[index]:
                    st.markdown(
                        f"""
                        <div style="background-color: {CONTRAST_BACKGROUND_COLOR}; border: 2px solid {BACKGROUND_COLOR_DARK_BLUE}; border-radius: 8px; padding: 15px; margin-bottom: 20px; box-shadow: 0 1px 2px rgba(0,0,0,0.05);">
                            <h4 style="margin-top: 0; color: white; text-align: center;">{month}</h4>
                        """, unsafe_allow_html=True
                    )
                    
                    df_month_rank = df_rank[df_rank['Mês/Ano'] == month]
                    df_top3_rank = df_month_rank.sort_values(by='Total Rank', ascending=False).head(3)
                    
                    if df_top3_rank.empty:
                        st.markdown("<p style='text-align: center; color: #888;'>S/Dados</p>", unsafe_allow_html=True)
                    else:
                        max_pedidos_visual = df_top3_rank['Total Rank'].max() # Máximo entre as 3 entidades
                        
                        for rank_num, (idx, row) in enumerate(df_top3_rank.iterrows()):
                            entity_name = row['Entidade']
                            total_pedidos_rank = row['Total Rank']
                            
                            df_entity_systems = df_monthly_entity[
                                (df_monthly_entity['Mês/Ano'] == month) & 
                                (df_monthly_entity['Entidade'] == entity_name)
                            ]
                            
                            formatted_value = format_number(total_pedidos_rank)
                            
                            st.markdown(
                                f"""
                                <div style="margin-bottom: 5px; font-weight: bold; color: white;">
                                    {rank_num + 1}º {entity_name} ({formatted_value})
                                </div>
                                <div style="display: flex; align-items: center; gap: 0px; margin-bottom: 10px;">
                                """, unsafe_allow_html=True
                            )
                            
                            # Verifica a distribuição entre sistemas (se o filtro 'Todos' estiver ativo)
                            for _, sys_row in df_entity_systems.iterrows():
                                system = sys_row['Sistema']
                                total_sys = sys_row['Total Pedidos']
                                
                                # A barra de progresso usa a proporção do total da Entidade / Pedidos Totais MÁXIMOS
                                if max_pedidos_visual > 0:
                                    # Largura é (total do sistema / total da Entidade) * (total da Entidade / max_pedidos) * 100
                                    # Simplificando, é a proporção do sistema em relação ao máximo do top 3.
                                    width_percent = (total_sys / max_pedidos_visual) * 100 
                                else:
                                    width_percent = 0
                                
                                bar_color = ARGOIT_COLOR if system == 'ARGOIT' else RESERVE_COLOR
                                
                                if width_percent > 0:
                                    st.markdown(
                                        f"""
                                        <div title="{system}: {format_number(total_sys)}" style="width: {width_percent}%; height: 16px; background-color: {bar_color};"></div>
                                        """, unsafe_allow_html=True
                                    )
                            
                            st.markdown("</div>", unsafe_allow_html=True)
                            
                    st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("---")
        
    # ====================================================
    # BLOCO 3: TABELA PIVOTADA CUSTOMIZADA
    # ====================================================
    
    df_pivot_final = pd.DataFrame() 

    if df_visual_filtrada.empty:
        st.warning("Nenhum dado encontrado para a combinação de filtros selecionada.")
    else:
        # CORREÇÃO: Indexação da pivot table
        if sistema_selecionado == 'Todos':
            # Se 'Todos' estiver selecionado, detalha por Sistema
            pivot_index = ['Entidade de Consolidação', SYSTEM_COL_NAME]
            st.subheader("Tabela de Pedidos - Entidades, Sistemas por Mês/Ano")
        else:
            # Se um sistema específico estiver selecionado, agrupa apenas por Entidade
            pivot_index = ['Entidade de Consolidação']
            st.subheader(f"Tabela de Pedidos - Entidades ({sistema_selecionado}) por Mês/Ano")
            
        df_pivot_final = pd.pivot_table(
            df_visual_filtrada,
            index=pivot_index, 
            columns=['Mês/Ano'], 
            values=['PKI Pedidos'], 
            aggfunc='sum',
            fill_value=0, 
            margins=True, 
            margins_name='Total Geral'
        )

        df_pivot_final.columns = df_pivot_final.columns.get_level_values(1)

        # --- FUNÇÃO DE ESTILO PARA O CONTEÚDO (APENAS DADOS) ---
        def highlight_content(data, color):
            is_content = pd.DataFrame('', index=data.index, columns=data.columns)
            background_attr = f'background-color: white; color: black;'
            background_attr_alt = f'background-color: #f0f2f6; color: black;'
            for i in range(len(data)):
                if i < len(data) - 1:
                    is_content.iloc[i, :-1] = background_attr if i % 2 == 0 else background_attr_alt
            is_content.iloc[:-1, :-1] = is_content.iloc[:-1, :-1].apply(lambda x: f'{x} color: black;')
            return is_content

        # --- APLICAÇÃO DO ESTILO ---
        header_totals_css = f'background-color: {ORANGE_COLOR}; color: white; font-weight: bold;'
        
        styled_df = df_pivot_final.style \
            .format("{:,.0f}") \
            .apply(highlight_content, color=ORANGE_COLOR, axis=None)

        styled_df = styled_df.set_table_styles(
            [
                {'selector': 'th.col_heading', 'props': header_totals_css},
                {'selector': 'th.row_heading', 'props': header_totals_css},
                {'selector': 'th.index_name', 'props': header_totals_css},
                {'selector': 'tbody tr:last-child td', 'props': header_totals_css},
                {'selector': 'td:last-child', 'props': header_totals_css},
                {'selector': 'tbody tr:last-child td:last-child', 'props': header_totals_css},
            ], overwrite=True
        )

        st.dataframe(
            styled_df, 
            use_container_width=True
        )


    st.markdown("---")
    
    # ====================================================
    # BLOCO 4: EXPORTAÇÃO (Com Download Pivotado e Bruto)
    # ====================================================
    st.markdown("### 💾 Exportar Dados")
    
    col_bruta, col_pivot = st.columns(2)

    # --- DOWNLOAD BASE BRUTA CONSOLIDADA (INCREMENTAL) ---
    with col_bruta:
        st.markdown("#### Base Bruta (Consolidada e Incremental)")
        try:
            # Garante que o DF para download seja o final, limpo e salvo
            df_to_download_consolidated = df_final_consolidated.drop_duplicates(subset=[ID_COL_NAME], keep='first').copy()
            
            if not df_to_download_consolidated.empty:
                xlsx_data = to_excel(df_to_download_consolidated)

                st.download_button(
                    label="📥 Download Base Consolidada",
                    data=xlsx_data,
                    file_name=f"INCREMENTAL_V10.1_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
            else:
                st.warning(f"O arquivo consolidado está vazio.")
            
        except Exception as e:
            st.error(f"Não foi possível gerar o link de download da Base Consolidada. Detalhe: {e}")

    # --- DOWNLOAD TABELA PIVOTADA FILTRADA ---
    with col_pivot:
        st.markdown("#### Tabela Pivotada (Filtrada e Formatada)")
        try:
            if not df_pivot_final.empty:
                xlsx_pivot_data = to_excel_styled(df_pivot_final)
                
                entidade_tag = entidade_selecionada.replace('Todas', 'ALL').replace(' ', '_').replace('/', '')
                mes_tag = mes_selecionado.replace('Todos', 'ALL').replace('/', '')
                sistema_tag = sistema_selecionado.replace('Todos', 'ALL').replace(' ', '_')
                
                file_name_pivot = f"PIVOT_{entidade_tag}_{mes_tag}_{sistema_tag}.xlsx"

                st.download_button(
                    label="📥 Download Tabela Pivotada",
                    data=xlsx_pivot_data,
                    file_name=file_name_pivot,
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
            else:
                st.info("Gere a tabela pivotada filtrando os dados primeiro.")
            
        except Exception as e:
            st.error(f"Não foi possível gerar o link de download da Tabela Pivotada. Detalhe: {e}")
            
else:
    st.error("❌ Falha crítica: Não foi possível processar ou carregar os dados. Verifique os arquivos de origem e os logs de erro acima.")