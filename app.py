import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import StringIO
from datetime import datetime
import io
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from PIL import Image
import base64
from pptx.dml.color import RGBColor

# Configurações do Streamlit para permitir upload de arquivos
st.set_option('deprecation.showfileUploaderEncoding', False)

# Configuração da página
st.set_page_config(
    page_title='Dashboard Cubo Itaú',
    layout='wide',
    initial_sidebar_state='collapsed'
)

# Cores do Itaú
CORES_ITAU = {
    'laranja': '#EC7000',
    'azul_escuro': '#003366',
    'azul_claro': '#0057FF',
    'branco': '#FFFFFF',
    'cinza_claro': '#F5F6FA'
}

# =====================
# Função para carregar e pré-processar os dados
# =====================
def carregar_dados():
    st.sidebar.header('Carregar Dados')
    
    # Configuração do file_uploader com todos os tipos de Excel
    uploaded_file = st.sidebar.file_uploader(
        'Faça upload do arquivo Excel',
        type=['xlsx', 'xls', 'xlsm', 'xlsb'],
        accept_multiple_files=False,
        help='Formatos aceitos: .xlsx, .xls, .xlsm, .xlsb'
    )
    
    df = None
    if uploaded_file is not None:
        try:
            # Detecta o tipo de arquivo
            file_type = uploaded_file.name.split('.')[-1].lower()
            
            if file_type == 'xls':
                # Para arquivos Excel 97-2003
                df = pd.read_excel(uploaded_file, engine='xlrd')
            else:
                # Para outros formatos
                try:
                    df = pd.read_excel(uploaded_file, engine='openpyxl')
                except:
                    df = pd.read_excel(uploaded_file, engine='xlrd')
            
            st.sidebar.success('Arquivo carregado com sucesso!')
            
        except Exception as e:
            st.sidebar.error(f'Erro ao ler o arquivo: {str(e)}')
            st.sidebar.info('Dica: Se o arquivo for Excel 97-2003 (.xls), tente salvá-lo como Excel 2007 ou superior (.xlsx)')
            return None
    else:
        st.sidebar.write('Ou cole os dados da planilha (Ctrl+V)')
        clipboard_data = st.sidebar.text_area('Cole aqui os dados copiados da planilha')
        if clipboard_data:
            try:
                df = pd.read_csv(StringIO(clipboard_data), sep='\t')
                st.sidebar.success('Dados colados com sucesso!')
            except Exception as e:
                st.sidebar.error(f'Erro ao ler os dados colados: {str(e)}')
                return None
    
    if df is not None:
        df = preprocessar_dados(df)
    return df

# [Mantenha todas as outras funções como estão até a função main()]

def main():
    st.markdown(f"""
        <style>
        .modern-card {{
            background: {CORES_ITAU['cinza_claro']};
            border-radius: 18px;
            padding: 24px 16px 16px 16px;
            margin-bottom: 12px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.04);
            text-align: center;
            min-width: 250px;
            max-width: 250px;
            height: 140px;
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
        }}
        .modern-card .big-number {{
            font-size: 2.5em;
            font-weight: bold;
            color: {CORES_ITAU['azul_escuro']};
        }}
        .modern-card .card-label {{
            font-size: 1.2em;
            color: {CORES_ITAU['laranja']};
            font-weight: 600;
        }}
        .main-title {{
            font-size: 2.5em;
            font-weight: bold;
            color: {CORES_ITAU['azul_escuro']};
            margin-bottom: 1em;
        }}
        </style>
    """, unsafe_allow_html=True)
    
    st.markdown('<div class="main-title">Dashboard de Visitas - Cubo Itaú</div>', unsafe_allow_html=True)
    
    df = carregar_dados()
    if df is None:
        st.info('Por favor, carregue um arquivo Excel ou cole os dados para iniciar a análise.')
        return

    # Cards em linha horizontal usando st.columns, igualmente espaçados
    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        st.markdown(f'<div class="modern-card"><div class="card-label">Total de Convites</div><div class="big-number">{total_convites(df)}</div></div>', unsafe_allow_html=True)
    with col2:
        st.markdown(f'<div class="modern-card"><div class="card-label">Anfitriões Notificados</div><div class="big-number">{anfitrioes_notificados(df)}</div></div>', unsafe_allow_html=True)
    with col3:
        st.markdown(f'<div class="modern-card"><div class="card-label">Não Notificados</div><div class="big-number">{anfitrioes_nao_notificados(df)}</div></div>', unsafe_allow_html=True)
    with col4:
        st.markdown(f'<div class="modern-card"><div class="card-label">Convidados Cubo</div><div class="big-number">{total_convidados_cubo(df)}</div></div>', unsafe_allow_html=True)
    with col5:
        st.markdown(f'<div class="modern-card"><div class="card-label">Média por Dia Útil</div><div class="big-number">{media_convidados_dia_util(df)}</div></div>', unsafe_allow_html=True)

    # Filtro de período
    st.sidebar.header('Filtro de Período')
    if df.empty or 'Ano' not in df.columns or 'Mês' not in df.columns:
        st.error('Dados inválidos ou incompletos. Verifique se o arquivo contém as colunas necessárias.')
        return

    anos = sorted(df['Ano'].dropna().unique(), reverse=True)
    meses = sorted(df['Mês'].dropna().unique())
    if not anos or not meses:
        st.error('Não há dados de período disponíveis.')
        return

    ano_sel = st.sidebar.selectbox('Ano', anos)
    mes_sel = st.sidebar.selectbox('Mês', meses)
    df_filtro = df[(df['Ano'] == ano_sel) & (df['Mês'] == mes_sel)]
    if df_filtro.empty:
        st.warning('Não há dados para o período selecionado.')
        return

    # Primeira linha de gráficos (2 colunas)
    col1, col2 = st.columns(2)
    with col1:
        st.plotly_chart(grafico_top_empresas(df_filtro), use_container_width=True)
    with col2:
        st.plotly_chart(grafico_convidados_por_data(df_filtro), use_container_width=True)

    # Segunda linha de gráficos (2 colunas)
    col1, col2 = st.columns(2)
    with col1:
        st.plotly_chart(grafico_convidados_por_dia_semana(df_filtro), use_container_width=True)
    with col2:
        st.subheader('Visitantes Frequentes por Empresa (>4 visitas no mês)')
        tabela_frequentes = visitantes_frequentes(df_filtro)
        if not tabela_frequentes.empty:
            st.dataframe(tabela_frequentes, height=420)

    # Terceira seção (consolidado e painel)
    st.markdown('---')
    col1, col2 = st.columns(2)
    with col1:
        st.subheader('Consolidado de Empresas com Visitantes Frequentes')
        st.dataframe(consolidado_frequentes(df_filtro), height=200)
        fig_consolidado = consolidado_frequentes_grafico(df_filtro)
        if fig_consolidado:
            st.plotly_chart(fig_consolidado, use_container_width=True)
    with col2:
        st.subheader('Painel de Empresas com Visitantes Frequentes')
        st.markdown(painel_empresas_frequentes(df_filtro), unsafe_allow_html=True)

    # Botão para download em PPTX
    if st.button('Baixar visualização em PPTX'):
        pptx_bytes = gerar_pptx(df, df_filtro)
        st.download_button('Clique aqui para baixar o PPTX', pptx_bytes, file_name='dashboard_cubo.pptx')

if __name__ == '__main__':
    main()