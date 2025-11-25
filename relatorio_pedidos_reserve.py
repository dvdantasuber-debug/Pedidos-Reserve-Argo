#VERSÃO 7.4

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

# --- DEFINIÇÃO DE CORES (ATUALIZADAS) ---
ORANGE_COLOR = '#ff8c00' # Laranja, usado para Reserve e para o estilo principal
RESERVE_COLOR = ORANGE_COLOR # Cor específica para Reserve
ARGOIT_COLOR = '#FFD700'  # Amarelo Ouro, para ARGOIT
BACKGROUND_COLOR_DARK_BLUE = '#131B36' 
CONTRAST_BACKGROUND_COLOR = '#1D2A4A' 
DARK_BACKGROUND_COLOR = CONTRAST_BACKGROUND_COLOR
DARK_FONT_COLOR = 'white'
BACKGROUND_BAR_COLOR = '#e0e0e0' 

# ----------------------------------------------------
# Funções Auxiliares (Não Alteradas)
# ----------------------------------------------------

def to_excel(df):
    """Converte o DataFrame para um buffer de memória XLSX (Dados Brutos)."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Dados', index=False)
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
# Leitura e Padronização das Bases de ORIGEM (Não Alteradas)
# ----------------------------------------------------

def load_reserve_data(file_path):
    """Lê a base Reserve e a tabela de grupos, e atribui o nome do sistema."""
    try:
        # 1. LEITURA DA BASE PRINCIPAL (5 colunas)
        df = pd.read_excel(
            file_path,
            sheet_name='base',
            header=None,
            skiprows=1,
            names=[DATE_COL_NAME, ID_COL_NAME, GROUP_CODE_COL, EMP_COL_NAME, GROUP_COL_NAME],
            engine='openpyxl'
        )
        
        # 2. LEITURA DA TABELA DE GRUPOS
        df_grupos = pd.read_excel(
            file_path,
            sheet_name=GRUPO_SHEET_NAME,
            usecols=[GRUPO_MAPPING_CODE_COL, GRUPO_MAPPING_NAME_COL],
            engine='openpyxl'
        )
        
        # --- PREPARAÇÃO DA CHAVE DE MERGE ---
        df_grupos.rename(
            columns={
                GRUPO_MAPPING_CODE_COL: 'merge_key',
                GRUPO_MAPPING_NAME_COL: 'Nome_Grupo_Mapeado'
            },
            inplace=True
        )
        df_grupos['merge_key'] = df_grupos['merge_key'].apply(
            lambda x: str(int(x)) if pd.notna(x) and str(x).replace('.', '', 1).isdigit() else str(x)
        ).str.strip()

        # 3. PREPARAÇÃO DA BASE PRINCIPAL
        df['merge_key'] = df[GROUP_CODE_COL].apply(
            lambda x: str(int(x)) if pd.notna(x) and str(x).replace('.', '', 1).isdigit() else str(x)
        ).str.strip()
        
        # 4. REALIZAR O MERGE (VLOOKUP)
        df = pd.merge(
            df,
            df_grupos[['merge_key', 'Nome_Grupo_Mapeado']],
            on='merge_key',
            how='left'
        )
        
        # 5. CONSOLIDAÇÃO DO NOME DO GRUPO
        df[GROUP_COL_NAME] = df['Nome_Grupo_Mapeado'].fillna(df[GROUP_COL_NAME])
        
        # 6. ATRIBUIÇÃO DO SISTEMA
        df[SYSTEM_COL_NAME] = 'Reserve'
        
        # Seleciona apenas as colunas padrão antes de retornar
        df = df[[DATE_COL_NAME, ID_COL_NAME, EMP_COL_NAME, GROUP_COL_NAME, SYSTEM_COL_NAME]].copy()

        return df, None

    except FileNotFoundError:
        return pd.DataFrame(), f"O arquivo '{file_path}' (Reserve) não foi encontrado."
    except Exception as e:
        return pd.DataFrame(), f"Erro grave ao processar base Reserve: {e}"

def load_argoit_data(file_map):
    """Lê e concatena os arquivos mensais do ARGOIT, usando header=1 (Linha 2) e mapeamento por nome."""
    all_argoit_data = []
    
    # Mapeamento dos nomes de coluna **REAIS** no Excel (Linha 2) para os nomes **PADRONIZADOS**
    ARGOIT_MAPPING = {
        'Data Inclusao': DATE_COL_NAME, 
        'Numero da Solicitacao': ID_COL_NAME,
        'Empresa de Débito': EMP_COL_NAME,
        'Cliente': GROUP_COL_NAME, # Corrigido para a forma mais comum
    }
    
    COLS_TO_READ = list(ARGOIT_MAPPING.keys())

    for month_year, file_path in file_map.items():
        if not os.path.exists(file_path):
            st.warning(f"⚠️ Aviso: Arquivo ARGOIT '{file_path}' para {month_year} não encontrado. Pulando.")
            continue
            
        try:
            # 1. LEITURA: Usa a Linha 2 (índice 1) como cabeçalho.
            df_month = pd.read_excel(
                file_path,
                header=1,  # CHAVE: Usa a Linha 2 (índice 1) como cabeçalho
                usecols=COLS_TO_READ, # Lê apenas as colunas mapeadas
                engine='openpyxl'
            )
            
            # 2. RENOMEAR AS COLUNAS (do nome real para o nome padrão)
            df_month.rename(columns=ARGOIT_MAPPING, inplace=True)

            # 3. LIMPEZA INICIAL
            df_month[DATE_COL_NAME] = pd.to_datetime(df_month[DATE_COL_NAME], errors='coerce', dayfirst=True)
            df_month.dropna(subset=[DATE_COL_NAME], inplace=True)
            
            if df_month.empty:
                 st.warning(f"O arquivo '{file_path}' (ARGOIT) foi lido, mas está vazio após a limpeza de datas. Pulando.")
                 continue

            # 4. ATRIBUIÇÃO DO SISTEMA E SELEÇÃO FINAL
            df_month[SYSTEM_COL_NAME] = 'ARGOIT'
            
            required_cols = [DATE_COL_NAME, ID_COL_NAME, EMP_COL_NAME, GROUP_COL_NAME, SYSTEM_COL_NAME]
            df_month = df_month[required_cols].copy()
            
            all_argoit_data.append(df_month)

        except KeyError as e:
            st.error(f"❌ Erro de Mapeamento no ARGOIT '{file_path}': Coluna '{e}' não encontrada. Por favor, ajuste o nome no dicionário ARGOIT_MAPPING (Linha 2).")
            st.warning("Se o erro persistir, o nome da coluna no Excel é diferente do que está no código. Por favor, verifique letras maiúsculas, minúsculas ou espaços.")
            continue
        except Exception as e:
            st.error(f"❌ Erro grave ao ler arquivo ARGOIT '{file_path}' para {month_year}: {type(e).__name__} - {e}")
            st.warning("Pulando este arquivo.")
            continue

    if not all_argoit_data:
        return pd.DataFrame(), "Nenhum arquivo ARGOIT válido foi carregado após a tentativa de leitura de todos os arquivos."

    df_argoit_combined = pd.concat(all_argoit_data, ignore_index=True)
    return df_argoit_combined, None

# ----------------------------------------------------
# CRIAÇÃO DA BASE CONSOLIDADA (Não Alterada)
# ----------------------------------------------------

def create_and_save_consolidated_base():
    """Carrega, unifica, limpa e salva as bases em um único arquivo XLSX."""
    
    st.info("🔄 Criando e limpando a base consolidada (`base_consolidada.xlsx`). Isso pode levar alguns segundos...")
    
    # 1. CARREGAR BASES DE ORIGEM
    df_reserve, error_r = load_reserve_data(BASE_RESERVE_FILE)
    df_argoit, error_a = load_argoit_data(ARGOIT_FILES)
    
    # Exibe avisos se houver
    if error_r:
        st.warning(f"Aviso Reserve: {error_r}")
    if error_a:
        st.warning(f"Aviso ARGOIT: {error_a}")
        
    # Exibe contagem de linhas para debug
    st.write(f"Linhas carregadas do Reserve: **{len(df_reserve)}**")
    st.write(f"Linhas carregadas do ARGOIT: **{len(df_argoit)}**")


    if df_reserve.empty and df_argoit.empty:
        st.error("❌ Falha crítica: Nenhuma base de dados (Reserve ou ARGOIT) pôde ser carregada para consolidação.")
        return pd.DataFrame()

    # 2. CONCATENAR AS DUAS BASES
    df_combined = pd.concat([df_reserve, df_argoit], ignore_index=True)
    
    # 3. LIMPEZA E PREPARAÇÃO FINAL
    
    df_combined[DATE_COL_NAME] = pd.to_datetime(df_combined[DATE_COL_NAME], errors='coerce', dayfirst=True)
    df_combined.dropna(subset=[DATE_COL_NAME], inplace=True)
    
    # Limpeza de strings
    df_combined[ID_COL_NAME] = df_combined[ID_COL_NAME].astype(str).str.strip()
    df_combined[EMP_COL_NAME] = df_combined[EMP_COL_NAME].astype(str).str.strip()
    df_combined[GROUP_COL_NAME] = df_combined[GROUP_COL_NAME].astype(str).str.strip().replace(['', 'nan', 'NaN'], np.nan)
    
    if df_combined.empty:
        st.warning("A base de dados consolidada está vazia após a limpeza e data dropna.")
        return pd.DataFrame()

    # 4. SALVAR O ARQUIVO CONSOLIDADO
    try:
        df_combined.to_excel(CONSOLIDATED_FILE, index=False, engine='xlsxwriter', sheet_name='Consolidado')
        st.success(f"✅ Base consolidada salva com sucesso em **`{CONSOLIDATED_FILE}`**.")
    except Exception as e:
        st.error(f"❌ Erro ao salvar o arquivo consolidado. Certifique-se de que ele não está aberto em outro programa. Detalhe: {e}")
        return pd.DataFrame()

    return df_combined

# ----------------------------------------------------
# Leitura e Pré-processamento (Cache Otimizado com Mapeamento) (Não Alterada)
# ----------------------------------------------------

@st.cache_data
def load_and_clean_data():
    """
    Tenta carregar a base consolidada. Se não existir, a cria.
    Em seguida, realiza o pré-processamento para o pivotamento.
    """
    df_combined = pd.DataFrame()
    
    # 1. Tentar carregar o arquivo CONSOLIDADO
    try:
        if os.path.exists(CONSOLIDATED_FILE):
            df_combined = pd.read_excel(CONSOLIDATED_FILE, engine='openpyxl', sheet_name='Consolidado')
            st.success(f"✅ Base carregada de **`{CONSOLIDATED_FILE}`**.")
        else:
            df_combined = create_and_save_consolidated_base()

    except Exception as e:
        st.error(f"❌ Erro ao carregar base consolidada existente ({CONSOLIDATED_FILE}): {e}. Tentando criar novamente.")
        df_combined = create_and_save_consolidated_base()

    if df_combined.empty:
        st.error("Não foi possível carregar ou criar a base consolidada. Verifique os arquivos de origem.")
        return None

    # 2. PRÉ-PROCESSAMENTO PARA O DASHBOARD 
    df_combined['Entidade de Consolidação'] = df_combined[GROUP_COL_NAME].fillna(df_combined[EMP_COL_NAME])
    df_combined['Mês/Ano'] = df_combined[DATE_COL_NAME].dt.strftime('%m/%Y')
    
    # 2.2. PKI Pedidos (Únicos por ID)
    df_pedidos_unicos = df_combined.groupby(ID_COL_NAME).agg(
        {'Entidade de Consolidação': 'first', 'Mês/Ano': 'first', SYSTEM_COL_NAME: 'first'}
    ).reset_index()

    df_pedidos_unicos['PKI Pedidos'] = 1
    
    df_base_pivot = df_pedidos_unicos[['Entidade de Consolidação', 'Mês/Ano', 'PKI Pedidos', SYSTEM_COL_NAME]]
    
    return df_base_pivot

# ----------------------------------------------------
# --- 2. Interface Streamlit ---
# ----------------------------------------------------

st.set_page_config(layout="wide", page_title="Dashboard Pedidos Consolidado")

# Aplica o CSS
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
        background-color: {RESERVE_COLOR}; 
        border-radius: 10px;
        padding: 10px;
        text-align: center;
        box-shadow: 0 1px 2px rgba(0,0,0,0.1); 
        border: 2px solid {BACKGROUND_COLOR_DARK_BLUE}; 
        margin-bottom: 10px;
    }}
    .kpi-box-argoit {{ 
        background-color: {ARGOIT_COLOR}; 
        border-radius: 10px;
        padding: 10px;
        text-align: center;
        box-shadow: 0 1px 2px rgba(0,0,0,0.1); 
        border: 2px solid {BACKGROUND_COLOR_DARK_BLUE}; 
        margin-bottom: 10px;
    }}
    .kpi-box-reserve p, .kpi-box-argoit p {{ color: white; margin: 0; font-size: 1.0em; font-weight: bold;}}
    .kpi-box-reserve h2, .kpi-box-argoit h2 {{ color: {BACKGROUND_COLOR_DARK_BLUE}; margin: 5px 0 0 0; font-size: 2.0em;}}

    </style>
    """, 
    unsafe_allow_html=True
)


df_base_pivot = load_and_clean_data()

# --- CABEÇALHO COM LOGO E TÍTULO ---
if df_base_pivot is not None:
    
    min_date = df_base_pivot['Mês/Ano'].min()
    max_date = df_base_pivot['Mês/Ano'].max()
    dashboard_title = f"Pedidos Consolidado (Reserve + ARGOIT) - Período {min_date} a {max_date}"
    
    logo_col, title_col = st.columns([1, 4])
    
    # Lógica do Logo e Título 
    try:
        img_base64_data, error = image_to_base64(LOGO_FILE, file_type="png")
        if img_base64_data:
            with logo_col:
                st.markdown(
                    f"""
                    <div class="logo-container">
                        <img src="{img_base64_data}" class="custom-logo-img" alt="Logomarca">
                    </div>
                    """,
                    unsafe_allow_html=True
                )
        else:
            with logo_col:
                st.markdown(f"<p style='color: red; font-size: 0.8em;'>Erro ao carregar logo: {error}</p>", unsafe_allow_html=True)
    except:
        with logo_col:
            st.warning("Logo não carregada.")
            
    with title_col:
        st.markdown(f"<h1>📊 Dashboard de Pedidos - Visão Consolidada</h1>", unsafe_allow_html=True)
        st.markdown(f"### {dashboard_title}")
    
    st.markdown("---")
    
    # ====================================================
    # BLOCO 2: FILTROS E KPI PRINCIPAL
    # ====================================================
    
    with st.container():
        # Aumentamos o número de colunas para 6 para incluir os KPIs do sistema
        col1, col2, col3, col4_total, col5_reserve, col6_argoit = st.columns([1, 1, 1, 1, 1, 1])

        entidades = ['Todas'] + sorted(df_base_pivot['Entidade de Consolidação'].unique().tolist())
        entidade_selecionada = col1.selectbox('Selecione a Entidade', entidades, key='entidade_filtro')
        
        meses = ['Todos'] + sorted(df_base_pivot['Mês/Ano'].unique().tolist(), key=lambda x: pd.to_datetime(x, format='%m/%Y'))
        mes_selecionado = col2.selectbox('Selecione o Mês/Ano', meses, key='mes_filtro')
        
        # O filtro de sistema continua aqui, mas afeta o TOTAL GERAL e as tabelas
        sistemas = ['Todos'] + sorted(df_base_pivot[SYSTEM_COL_NAME].unique().tolist())
        sistema_selecionado = col3.selectbox('Selecione o Sistema', sistemas, key='sistema_filtro')

        df_filtrado = df_base_pivot.copy()
        
        if entidade_selecionada != 'Todas':
            df_filtrado = df_filtrado[df_filtrado['Entidade de Consolidação'] == entidade_selecionada]
        
        if mes_selecionado != 'Todos':
            df_filtrado = df_filtrado[df_filtrado['Mês/Ano'] == mes_selecionado]
            
        if sistema_selecionado != 'Todos':
            # df_filtrado_sistema é a base que alimenta a maioria dos gráficos e a tabela final
            df_filtrado_sistema = df_filtrado[df_filtrado[SYSTEM_COL_NAME] == sistema_selecionado]
        else:
            df_filtrado_sistema = df_filtrado # Se for "Todos", usa o DataFrame filtrado por Entidade/Mês
            
        
        # --- CÁLCULO DOS KPIS ---
        total_pedidos = df_filtrado_sistema['PKI Pedidos'].sum()
        
        # Particionamento por Sistema (afetado por Entidade e Mês, mas não pelo Filtro Sistema)
        total_reserve = df_filtrado[df_filtrado[SYSTEM_COL_NAME] == 'Reserve']['PKI Pedidos'].sum()
        total_argoit = df_filtrado[df_filtrado[SYSTEM_COL_NAME] == 'ARGOIT']['PKI Pedidos'].sum()
        
        
        # --- EXIBIÇÃO DOS KPIS ---
        
        # KPI Total Geral (afetado por todos os filtros)
        with col4_total:
            st.metric(label="Total de Pedidos Únicos", value=f"{total_pedidos:,.0f}".replace(",", "#").replace(".", ",").replace("#", "."))

        # KPI Reserve (afetado por Entidade e Mês, mas não pelo Filtro Sistema)
        with col5_reserve:
            st.markdown(
                f"""
                <div class="metric-small">
                    <p data-testid="stMetricLabel">Total Reserve</p>
                    <p data-testid="stMetricValue">{f"{total_reserve:,.0f}".replace(",", "#").replace(".", ",").replace("#", ".")}</p>
                </div>
                """,
                unsafe_allow_html=True
            )

        # KPI ARGOIT (afetado por Entidade e Mês, mas não pelo Filtro Sistema)
        with col6_argoit:
            st.markdown(
                f"""
                <div class="metric-small">
                    <p data-testid="stMetricLabel">Total ARGOIT</p>
                    <p data-testid="stMetricValue">{f"{total_argoit:,.0f}".replace(",", "#").replace(".", ",").replace("#", ".")}</p>
                </div>
                """,
                unsafe_allow_html=True
            )

    st.markdown("---")
    
    # ====================================================
    # BLOCO 1: FRAMES DE TOTAIS POR MÊS (PARTICIONADO POR SISTEMA)
    # ====================================================

    if not df_filtrado_sistema.empty:
        st.subheader("🚀 Total de Pedidos por Mês (KPIs Dinâmicos)")

        # Agrupamento para Reserve e ARGOIT
        df_monthly_systems = df_filtrado_sistema.groupby(['Mês/Ano', SYSTEM_COL_NAME])['PKI Pedidos'].sum().unstack(fill_value=0).reset_index()
        
        # Garante que as colunas Reserve e ARGOIT existam, mesmo que vazias após o filtro
        if 'Reserve' not in df_monthly_systems.columns:
            df_monthly_systems['Reserve'] = 0
        if 'ARGOIT' not in df_monthly_systems.columns:
            df_monthly_systems['ARGOIT'] = 0
        
        # Calcula a ordem correta dos meses
        df_monthly_systems['Data Ordenacao'] = pd.to_datetime(df_monthly_systems['Mês/Ano'], format='%m/%Y')
        df_monthly_systems = df_monthly_systems.sort_values('Data Ordenacao').drop(columns='Data Ordenacao')
        
        month_order = df_monthly_systems['Mês/Ano'].tolist()
        num_months = len(month_order)
        cols_per_row = 4
        
        st.markdown("#### Total Reserve")
        
        for i in range(0, num_months, cols_per_row):
            current_months = df_monthly_systems[df_monthly_systems['Mês/Ano'].isin(month_order[i:i + cols_per_row])]
            cols = st.columns(len(current_months))
            
            for j, row in current_months.iterrows():
                month = row['Mês/Ano']
                total = row['Reserve']
                
                formatted_value = f"{total:,.0f}".replace(",", "#").replace(".", ",").replace("#", ".")
                
                with cols[current_months.index.get_loc(j)]:
                    st.markdown(
                        f"""
                        <div class="kpi-box-reserve">
                            <p>{month}</p>
                            <h2>{formatted_value}</h2>
                        </div>
                        """, unsafe_allow_html=True
                    )

        st.markdown("#### Total ARGOIT")

        for i in range(0, num_months, cols_per_row):
            current_months = df_monthly_systems[df_monthly_systems['Mês/Ano'].isin(month_order[i:i + cols_per_row])]
            cols = st.columns(len(current_months))
            
            for j, row in current_months.iterrows():
                month = row['Mês/Ano']
                total = row['ARGOIT']
                
                formatted_value = f"{total:,.0f}".replace(",", "#").replace(".", ",").replace("#", ".")
                
                with cols[current_months.index.get_loc(j)]:
                    st.markdown(
                        f"""
                        <div class="kpi-box-argoit">
                            <p>{month}</p>
                            <h2>{formatted_value}</h2>
                        </div>
                        """, unsafe_allow_html=True
                    )

        st.markdown("---")

        # ====================================================
        # BLOCO 2: TOP 3 ENTIDADES POR MÊS (COM COR DE SISTEMA)
        # ====================================================
        
        st.subheader("🏆 Top 3 Entidades (Leaderboard Mensal por Quantidade)")

        # df_filtrado_sistema é usado para garantir que respeita todos os filtros
        df_monthly_entity = df_filtrado_sistema.groupby(['Mês/Ano', 'Entidade de Consolidação', SYSTEM_COL_NAME])['PKI Pedidos'].sum().reset_index()
        df_monthly_entity.columns = ['Mês/Ano', 'Entidade', 'Sistema', 'Total Pedidos']
        
        # Calcular o total MENSAL para obter o ranking
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
                        <div style="
                            background-color: {CONTRAST_BACKGROUND_COLOR}; 
                            border: 2px solid {BACKGROUND_COLOR_DARK_BLUE}; 
                            border-radius: 8px;
                            padding: 15px;
                            margin-bottom: 20px;
                            box-shadow: 0 1px 2px rgba(0,0,0,0.05);
                        ">
                            <h4 style="margin-top: 0; color: white; text-align: center;">{month}</h4>
                        """, unsafe_allow_html=True
                    )
                    
                    df_month_rank = df_rank[df_rank['Mês/Ano'] == month]
                    df_top3_rank = df_month_rank.sort_values(by='Total Rank', ascending=False).head(3)
                    
                    if df_top3_rank.empty:
                        st.markdown("<p style='text-align: center; color: #888;'>S/Dados</p>", unsafe_allow_html=True)
                    else:
                        max_pedidos = df_top3_rank['Total Rank'].max()
                        
                        for rank_num, (idx, row) in enumerate(df_top3_rank.iterrows()):
                            entity_name = row['Entidade']
                            total_pedidos_rank = row['Total Rank']
                            
                            # Obter a distribuição por sistema para esta entidade (para colorir)
                            df_entity_systems = df_monthly_entity[
                                (df_monthly_entity['Mês/Ano'] == month) & 
                                (df_monthly_entity['Entidade'] == entity_name)
                            ]
                            
                            # Formatação
                            formatted_value = f"{total_pedidos_rank:,.0f}".replace(",", "#").replace(".", ",").replace("#", ".")
                            
                            st.markdown(
                                f"""
                                <div style="margin-bottom: 5px; font-weight: bold; color: white;">
                                    {rank_num + 1}º {entity_name} ({formatted_value})
                                </div>
                                <div style="display: flex; align-items: center; gap: 0px; margin-bottom: 10px;">
                                """, unsafe_allow_html=True
                            )
                            
                            # Desenhar barras particionadas por sistema
                            for _, sys_row in df_entity_systems.iterrows():
                                system = sys_row['Sistema']
                                total_sys = sys_row['Total Pedidos']
                                
                                # A largura é proporcional ao total do mês para aquela entidade, em relação ao maior do top 3
                                if max_pedidos > 0:
                                    width_percent = (total_sys / total_pedidos_rank) * (total_pedidos_rank / max_pedidos) * 100
                                else:
                                    width_percent = 0
                                
                                bar_color = ARGOIT_COLOR if system == 'ARGOIT' else RESERVE_COLOR
                                
                                if width_percent > 0:
                                    st.markdown(
                                        f"""
                                        <div title="{system}: {total_sys}" style="width: {width_percent}%; height: 16px; background-color: {bar_color};"></div>
                                        """, unsafe_allow_html=True
                                    )
                            
                            st.markdown("</div>", unsafe_allow_html=True)
                            
                    st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("---")
        
    # ====================================================
    # BLOCO 3: TABELA PIVOTADA CUSTOMIZADA (Não Alterada)
    # ====================================================

    # Usa df_filtrado_sistema
    if df_filtrado_sistema.empty:
        st.warning("Nenhum dado encontrado para a combinação de filtros selecionada.")
    else:
        if sistema_selecionado == 'Todos':
            pivot_index = ['Entidade de Consolidação', SYSTEM_COL_NAME]
            st.subheader("Tabela de Pedidos - Entidades, Sistemas por Mês/Ano")
        else:
            pivot_index = ['Entidade de Consolidação']
            st.subheader("Tabela de Pedidos - Entidades por Mês/Ano")
            
        df_pivot_final = pd.pivot_table(
            df_filtrado_sistema,
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
    
    # Botão de Download NATIVO XLSX (Dados Brutos) 
    st.markdown("### 💾 Exportar Base Consolidada (`base_consolidada.xlsx`)")
    
    # Tenta ler a base consolidada para o download
    try:
        if os.path.exists(CONSOLIDATED_FILE):
            df_consolidada_raw = pd.read_excel(CONSOLIDATED_FILE, engine='openpyxl', sheet_name='Consolidado')
            xlsx_data = to_excel(df_consolidada_raw)

            st.download_button(
                label="Download Base Consolidada (Excel XLSX)",
                data=xlsx_data,
                file_name=CONSOLIDATED_FILE,
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        else:
            st.warning(f"O arquivo '{CONSOLIDATED_FILE}' ainda não existe. Por favor, recarregue o dashboard para criá-lo.")
            
    except Exception as e:
        st.error(f"Não foi possível gerar o link de download para '{CONSOLIDATED_FILE}'. Detalhe: {e}")