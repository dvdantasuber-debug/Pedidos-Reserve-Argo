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
# Define o nome do arquivo principal como EXCEL
BASE_FILE = 'base.xlsx'
SEPARATOR = ','

# Constantes para o mapeamento de Grupos
GRUPO_SHEET_NAME = 'GRUPOS'
GRUPO_MAPPING_CODE_COL = 'Codigo'
GRUPO_MAPPING_NAME_COL = 'Nome do Grupo'

# Define a cor laranja para a barra e texto
ORANGE_COLOR = '#ff8c00' # Laranja escuro (Principal)
# Cor da barra de fundo (cinza claro)
BACKGROUND_BAR_COLOR = '#e0e0e0'

# Cores e estilos para a tabela customizada
HEADER_COLOR = ORANGE_COLOR
TOTALS_COLOR = ORANGE_COLOR
HEADER_FONT_COLOR = 'white'
TOTALS_FONT_COLOR = 'white'
CONTENT_FONT_COLOR = ORANGE_COLOR # Conteúdo (valores) em laranja

# Cor para o fundo da seção de Filtros/KPIs
DARK_BACKGROUND_COLOR = '#333333' # Cinza escuro
DARK_FONT_COLOR = 'white'

# ----------------------------------------------------
# Funções de Criação do Arquivo Excel Interativo
# ----------------------------------------------------

def to_excel(df):
    """Converte o DataFrame para um buffer de memória XLSX (Dados Brutos)."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Dados', index=False)
    return output.getvalue()

# ----------------------------------------------------
# Leitura e Pré-processamento (Cache Otimizado com Mapeamento)
# ----------------------------------------------------

@st.cache_data
def load_and_clean_data():
    """
    Lê a base e a tabela de grupos do arquivo Excel, realiza o MERGE 
    (VLOOKUP) com tratamento de tipos e pré-processa os dados.
    """
    try:
        # 1. LEITURA DA BASE PRINCIPAL (Assumindo que a aba se chama 'base')
        df = pd.read_excel(
            BASE_FILE,
            sheet_name='base', # Assumindo que a aba de dados se chama 'base'
            header=None,
            skiprows=1,
            names=[DATE_COL_NAME, ID_COL_NAME, GROUP_CODE_COL, EMP_COL_NAME, GROUP_COL_NAME],
            engine='openpyxl'
        )
        
        # 2. LEITURA DA TABELA DE GRUPOS
        df_grupos = pd.read_excel(
            BASE_FILE,
            sheet_name=GRUPO_SHEET_NAME, # Lendo a aba 'GRUPOS'
            usecols=[GRUPO_MAPPING_CODE_COL, GRUPO_MAPPING_NAME_COL],
            engine='openpyxl'
        )
        
        # --- PREPARAÇÃO DA CHAVE DE MERGE (Garantindo Consistência) ---
        
        # 2.1. Preparo da Tabela de Mapeamento (df_grupos)
        df_grupos.rename(
            columns={
                GRUPO_MAPPING_CODE_COL: 'merge_key',
                GRUPO_MAPPING_NAME_COL: 'Nome_Grupo_Mapeado'
            },
            inplace=True
        )
        
        # Tenta converter a chave de mapeamento para string limpa, tratando inteiros corretamente
        df_grupos['merge_key'] = df_grupos['merge_key'].apply(
            lambda x: str(int(x)) if pd.notna(x) and str(x).replace('.', '', 1).isdigit() else str(x)
        ).str.strip()

        # 3. LIMPEZA E PREPARAÇÃO DA BASE PRINCIPAL (df)
        df[ID_COL_NAME] = df[ID_COL_NAME].astype(str).str.strip()
        df[EMP_COL_NAME] = df[EMP_COL_NAME].astype(str).str.strip()
        df[GROUP_COL_NAME] = df[GROUP_COL_NAME].astype(str).str.strip().replace(['', 'nan', 'NaN'], np.nan)
        
        # Tenta converter a chave da base principal para string limpa, tratando inteiros corretamente
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
        # Preenche a coluna 'nome grupo' original com o valor mapeado (Nome_Grupo_Mapeado).
        df[GROUP_COL_NAME] = df['Nome_Grupo_Mapeado'].fillna(df[GROUP_COL_NAME])
        
        # 6. LIMPEZA FINAL E GERAÇÃO DA PKI
        df[DATE_COL_NAME] = pd.to_datetime(df[DATE_COL_NAME], errors='coerce', dayfirst=True)
        df.dropna(subset=[DATE_COL_NAME], inplace=True)

        if df.empty:
            return None

        # Definição da Entidade de Consolidação FINAL
        df['Entidade de Consolidação'] = df[GROUP_COL_NAME].fillna(df[EMP_COL_NAME])
        df['Mês/Ano'] = df[DATE_COL_NAME].dt.strftime('%m/%Y')
        
        df_pedidos_unicos = df.groupby(ID_COL_NAME).agg(
            {'Entidade de Consolidação': 'first', 'Mês/Ano': 'first'}
        ).reset_index()

        df_pedidos_unicos['PKI Pedidos'] = 1
        
        df_base_pivot = df_pedidos_unicos[['Entidade de Consolidação', 'Mês/Ano', 'PKI Pedidos']]
        
        return df_base_pivot

    except FileNotFoundError:
        st.error(f"❌ ERRO FATAL: O arquivo '{BASE_FILE}' não foi encontrado. Verifique se ele se chama 'base.xlsx'.")
        return None
    except ValueError as e:
        if "worksheet named" in str(e):
             st.error(f"❌ ERRO FATAL: Não foi possível encontrar a aba principal ou a aba '{GRUPO_SHEET_NAME}' no arquivo '{BASE_FILE}'.")
             return None
        st.error(f"❌ ERRO FATAL ao processar o arquivo. Detalhe: {e}")
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
    
    # ====================================================
    # ✅ BLOCO 2: FILTROS E KPI PRINCIPAL (FUNDO ESCURO E KPI CENTRALIZADO)
    # ====================================================
    
    st.markdown("---")
    
    # Adiciona o estilo CSS para o container dos filtros e KPI
    st.markdown(
        f"""
        <style>
        /* Procura o primeiro div que contém os widgets após o markdown "---" */
        /* e aplica o estilo ao seu container pai, que é o st.container (ou similar) */
        div[data-testid="stVerticalBlock"]:nth-of-type(1) > div:nth-child(1) {{
            background-color: {DARK_BACKGROUND_COLOR};
            padding: 15px 20px 5px 20px;
            border-radius: 10px;
            color: {DARK_FONT_COLOR};
            margin-bottom: 20px;
        }}
        /* Estiliza os labels e o metric dentro do bloco escuro */
        div[data-testid="stVerticalBlock"]:nth-of-type(1) > div:nth-child(1) label,
        div[data-testid="stVerticalBlock"]:nth-of-type(1) > div:nth-child(1) [data-testid="stMetricLabel"] {{
            color: {DARK_FONT_COLOR} !important;
            text-align: center; /* NOVO: Centraliza o label */
            width: 100%; /* NOVO: Garante a largura total para centralização */
            display: block; /* NOVO: Garante que o label se comporte como um bloco */
        }}
        /* Estiliza o valor do metric para destaque em laranja */
        div[data-testid="stVerticalBlock"]:nth-of-type(1) > div:nth-child(1) [data-testid="stMetricValue"] {{
            color: {ORANGE_COLOR} !important;
            font-size: 3em !important; /* NOVO: Aumenta o tamanho da fonte */
            text-align: center; /* NOVO: Centraliza o valor */
            width: 100%; /* NOVO: Garante a largura total para centralização */
            display: block; /* NOVO: Garante que o valor se comporte como um bloco */
        }}
        </style>
        """, 
        unsafe_allow_html=True
    )
    
    with st.container():
        col1, col2, col3 = st.columns([1, 1, 1])

        entidades = ['Todas'] + sorted(df_base_pivot['Entidade de Consolidação'].unique().tolist())
        entidade_selecionada = col1.selectbox('Selecione a Entidade', entidades, key='entidade_filtro')
        
        meses = ['Todos'] + sorted(df_base_pivot['Mês/Ano'].unique().tolist(), key=lambda x: pd.to_datetime(x, format='%m/%Y'))
        mes_selecionado = col2.selectbox('Selecione o Mês/Ano', meses, key='mes_filtro')

        # Lógica de Filtragem
        df_filtrado = df_base_pivot.copy()

        if entidade_selecionada != 'Todas':
            df_filtrado = df_filtrado[df_filtrado['Entidade de Consolidação'] == entidade_selecionada]
        
        if mes_selecionado != 'Todos':
            df_filtrado = df_filtrado[df_filtrado['Mês/Ano'] == mes_selecionado]

        # Recalcula os totais (KPI Principal)
        total_pedidos = df_filtrado['PKI Pedidos'].sum()
        
        # O metric é colocado na terceira coluna
        with col3:
            st.metric(label="Total de Pedidos Únicos", value=f"{total_pedidos:,.0f}".replace(",", "#").replace(".", ",").replace("#", "."))

    st.markdown("---")
    
    # ====================================================
    # BLOCO 1: FRAMES DE TOTAIS POR MÊS 
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
                
                # Formatação para o valor (tratando a pontuação)
                formatted_value = f"{total:,.0f}".replace(",", "#").replace(".", ",").replace("#", ".")
                
                with cols[current_months.index.get_loc(j)]:
                    st.markdown(
                        f"""
                        <div style="
                            background-color: {ORANGE_COLOR};
                            border-radius: 10px;
                            padding: 10px;
                            text-align: center;
                            /* Sombra ainda mais sutil */
                            box-shadow: 0 1px 2px rgba(0,0,0,0.1); 
                            /* Borda preta um pouco mais grossa */
                            border: 2px solid #333333; /* Usando DARK_BACKGROUND_COLOR para a borda */
                            margin-bottom: 10px;
                        ">
                            <p style="
                                color: white;
                                margin: 0;
                                font-size: 1.1em;
                                font-weight: bold;
                            ">{month}</p>
                            <h2 style="
                                color: white;
                                margin: 5px 0 0 0;
                                font-size: 2.2em;
                            ">{formatted_value}</h2>
                        </div>
                        """, unsafe_allow_html=True
                    )


        st.markdown("---")

        # ====================================================
        # BLOCO 2: TOP 3 ENTIDADES POR MÊS 
        # ====================================================
        
        st.subheader("🏆 Top 3 Entidades (Leaderboard Mensal por Quantidade)")

        # 1. Agrupar dados por Mês e Entidade
        df_monthly_entity = df_filtrado.groupby(['Mês/Ano', 'Entidade de Consolidação'])['PKI Pedidos'].sum().reset_index()
        df_monthly_entity.columns = ['Mês/Ano', 'Entidade', 'Total Pedidos']

        month_order = df_monthly_totals['Mês/Ano'].tolist()

        # 2. Definir o layout de colunas para os meses (4 por linha)
        cols_per_row_top3 = 4
        num_months_top3 = len(month_order)
        
        for i in range(0, num_months_top3, cols_per_row_top3):
            current_month_batch = month_order[i:i + cols_per_row_top3]
            cols = st.columns(len(current_month_batch))
            
            for index, month in enumerate(current_month_batch):
                
                # Inicia o CARD (Retângulo) para o mês
                with cols[index]:
                    st.markdown(
                        f"""
                        <div style="
                            background-color: #f0f2f6;
                            /* Borda preta um pouco mais grossa */
                            border: 2px solid #333333; /* Usando DARK_BACKGROUND_COLOR para a borda */
                            border-radius: 8px;
                            padding: 15px;
                            margin-bottom: 20px;
                            /* Sombra ainda mais sutil */
                            box-shadow: 0 1px 2px rgba(0,0,0,0.05);
                        ">
                            <h4 style="margin-top: 0; color: #1e81b0; text-align: center;">{month}</h4>
                        """, unsafe_allow_html=True
                    )
                    
                    # Filtra dados para o mês atual
                    df_month = df_monthly_entity[df_monthly_entity['Mês/Ano'] == month]
                    df_top3 = df_month.sort_values(by='Total Pedidos', ascending=False).head(3)
                    
                    if df_top3.empty:
                        st.markdown("<p style='text-align: center; color: #888;'>S/Dados</p>", unsafe_allow_html=True)
                    else:
                        max_pedidos = df_top3['Total Pedidos'].max()
                        
                        for rank, (idx, row) in enumerate(df_top3.iterrows()):
                            entity_name = row['Entidade']
                            total_pedidos_entity = row['Total Pedidos']
                            
                            ratio = total_pedidos_entity / max_pedidos if max_pedidos > 0 else 0
                            
                            # Formatação para o valor (tratando a pontuação)
                            formatted_value = f"{total_pedidos_entity:,.0f}".replace(",", "#").replace(".", ",").replace("#", ".")
                            
                            # --- ESTRUTURA FLEXBOX COM flex-grow: 1 NA BARRA ---
                            st.markdown(
                                f"""
                                <div style="
                                    margin-bottom: 5px;
                                    font-weight: bold;
                                    color: #333;
                                ">
                                    {rank + 1}º {entity_name}
                                </div>
                                <div style="
                                    display: flex;
                                    align-items: center;
                                    gap: 10px; /* Espaço entre a barra e o valor */
                                    margin-bottom: 10px;
                                ">
                                    <div style="
                                        height: 16px;
                                        background-color: {BACKGROUND_BAR_COLOR};
                                        border-radius: 5px;
                                        overflow: hidden;
                                        flex-grow: 1;
                                        position: relative;
                                    ">
                                        <div style="
                                            width: {ratio * 100}%;
                                            height: 100%;
                                            background-color: {ORANGE_COLOR};
                                            border-radius: 5px;
                                            min-width: 5px;
                                        "></div>
                                    </div>
                                    <span style="
                                        color: {ORANGE_COLOR};
                                        font-size: 0.9em;
                                        font-weight: bold;
                                        white-space: nowrap;
                                        flex-shrink: 0;
                                    ">{formatted_value}</span>
                                </div>
                                """, unsafe_allow_html=True
                            )
                            
                    # Fecha o CARD (Retângulo)
                    st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("---")
        
    # ====================================================
    # BLOCO 3: TABELA PIVOTADA CUSTOMIZADA 
    # ====================================================

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
        
        # --- FUNÇÃO DE ESTILO PARA O CONTEÚDO (APENAS DADOS) ---
        
        def highlight_content(data, color):
            """Aplica a cor do texto Laranja em todas as células, exceto a última linha e coluna."""
            attr = f'color: {color}'
            
            # Cria um DataFrame de estilos do mesmo tamanho, inicializado com strings vazias
            is_content = pd.DataFrame('', index=data.index, columns=data.columns)
            
            # Define o estilo para todas as linhas, exceto a última (Total Geral)
            # E todas as colunas, exceto a última (Total Geral)
            is_content.iloc[:-1, :-1] = attr 
            
            return is_content

        # --- DEFINIÇÃO DE ESTILOS CSS ---
        
        # Estilo para Cabeçalhos e Totais (Fundo Laranja, Texto Branco)
        header_totals_css = f'background-color: {TOTALS_COLOR}; color: {TOTALS_FONT_COLOR}; font-weight: bold;'
        
        # --- APLICAÇÃO DO ESTILO ---
        
        # 1. Aplica o formato de número e a cor laranja ao conteúdo
        styled_df = df_pivot_final.style \
            .format("{:,.0f}") \
            .apply(highlight_content, color=CONTENT_FONT_COLOR, axis=None)


        # 2. Aplica o estilo Laranja/Branco para TODOS os Cabeçalhos e Células de Totais
        styled_df = styled_df.set_table_styles(
            [
                # Cabeçalhos de Coluna (Mês/Ano e Total Geral)
                {'selector': 'th.col_heading', 'props': header_totals_css},
                # Cabeçalhos de Linha (Entidade de Consolidação)
                {'selector': 'th.row_heading', 'props': header_totals_css},
                # Célula superior esquerda ('Entidade de Consolidação' no canto)
                {'selector': 'th.index_name', 'props': header_totals_css},
                
                # Linha de Totais (Valores na última linha)
                {'selector': 'tbody tr:last-child td', 'props': header_totals_css},
                # Coluna de Totais (Valores na última coluna)
                {'selector': 'td:last-child', 'props': header_totals_css},
                # Célula de interseção no canto inferior direito
                {'selector': 'tbody tr:last-child td:last-child', 'props': header_totals_css},

            ], overwrite=True
        )

        # O Streamlit renderiza o DataFrame estilizado
        st.dataframe(
            styled_df, 
            use_container_width=True
        )


    st.markdown("---")
    
    # Botão de Download NATIVO XLSX (Dados Brutos) 
    st.markdown("### 💾 Exportar Dados Brutos (Para Criar a Tabela Dinâmica no Excel)")
    
    xlsx_data = to_excel(df_base_pivot)

    st.download_button(
        label="Download Dados Brutos (Excel XLSX)",
        data=xlsx_data,
        file_name='relatorio_pedidos_dados_brutos.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )