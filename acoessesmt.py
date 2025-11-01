import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import openpyxl

# Configuração da página
st.set_page_config(
    page_title="BI SESMT - Rezende Energia",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Cores da empresa
COR_PRINCIPAL = "#000000"
COR_SECUNDARIA = "#F7931E"
COR_FUNDO = "#FFFFFF"
COR_TEXTO = "#333333"

# CSS Customizado
st.markdown(f"""
<style>
    /* Importar fonte moderna */
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;600;700&display=swap');

    /* Resetar fonte padrão */
    html, body, [class*="css"] {{
        font-family: 'Poppins', sans-serif;
    }}

    /* Estilização do sidebar */
    [data-testid="stSidebar"] {{
        background: linear-gradient(180deg, {COR_PRINCIPAL} 0%, #1a1a1a 100%);
    }}

    [data-testid="stSidebar"] * {{
        color: white !important;
    }}

    /* Título principal */
    .main-title {{
        background: linear-gradient(90deg, {COR_PRINCIPAL} 0%, {COR_SECUNDARIA} 100%);
        padding: 30px;
        border-radius: 15px;
        text-align: center;
        margin-bottom: 30px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }}

    .main-title h1 {{
        color: white !important;
        font-size: 2.5rem;
        font-weight: 700;
        margin: 0;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
    }}

    .main-title p {{
        color: white !important;
        font-size: 1.1rem;
        margin: 10px 0 0 0;
        opacity: 0.9;
    }}

    /* Cards de métricas */
    [data-testid="stMetricValue"] {{
        font-size: 2rem;
        font-weight: 700;
        color: {COR_SECUNDARIA};
    }}

    [data-testid="stMetricLabel"] {{
        font-size: 1rem;
        font-weight: 600;
        color: {COR_TEXTO};
    }}

    /* Melhorar aparência das métricas */
    [data-testid="metric-container"] {{
        background: white;
        padding: 20px;
        border-radius: 10px;
        border-left: 5px solid {COR_SECUNDARIA};
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }}

    /* Botões */
    .stButton > button {{
        background: linear-gradient(90deg, {COR_SECUNDARIA} 0%, #ff9d3d 100%);
        color: white;
        border: none;
        border-radius: 8px;
        padding: 12px 24px;
        font-weight: 600;
        transition: all 0.3s;
    }}

    .stButton > button:hover {{
        transform: translateY(-2px);
        box-shadow: 0 4px 8px rgba(247, 147, 30, 0.3);
    }}

    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {{
        gap: 8px;
        background-color: #f8f9fa;
        padding: 10px;
        border-radius: 10px;
    }}

    .stTabs [data-baseweb="tab"] {{
        background-color: white;
        border-radius: 8px;
        padding: 10px 20px;
        font-weight: 600;
        border: 2px solid transparent;
    }}

    .stTabs [aria-selected="true"] {{
        background: {COR_SECUNDARIA};
        color: white;
        border-color: {COR_SECUNDARIA};
    }}

    /* Upload box */
    [data-testid="stFileUploader"] {{
        background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
        padding: 30px;
        border-radius: 15px;
        border: 2px dashed {COR_SECUNDARIA};
    }}

    /* Headers */
    h1, h2, h3 {{
        color: {COR_PRINCIPAL};
        font-weight: 700;
    }}

    /* Dataframes */
    [data-testid="stDataFrame"] {{
        border-radius: 10px;
        overflow: hidden;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }}

    /* Selectbox e outros inputs */
    [data-baseweb="select"] {{
        border-radius: 8px;
    }}

    /* Divisor customizado */
    hr {{
        margin: 30px 0;
        border: none;
        height: 2px;
        background: linear-gradient(90deg, transparent 0%, {COR_SECUNDARIA} 50%, transparent 100%);
    }}

    /* Tooltips e info */
    [data-testid="stMarkdownContainer"] p {{
        color: {COR_TEXTO};
        line-height: 1.6;
    }}
</style>
""", unsafe_allow_html=True)


# Funções auxiliares
def criar_layout_cores():
    """Retorna configuração de layout para gráficos Plotly"""
    return dict(
        plot_bgcolor='white',
        paper_bgcolor='white',
        font=dict(family="Poppins, sans-serif", color=COR_TEXTO),
        title_font=dict(size=20, color=COR_PRINCIPAL, family="Poppins, sans-serif"),
        hoverlabel=dict(bgcolor="white", font_size=12, font_family="Poppins"),
    )


def processar_dados(df):
    """Processa os dados do arquivo"""
    import locale

    # Tentar configurar locale para português
    try:
        locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
    except:
        try:
            locale.setlocale(locale.LC_TIME, 'Portuguese_Brazil.1252')
        except:
            pass

    # Converter data para datetime
    df['Data'] = pd.to_datetime(df['Data'])
    df['Mês'] = df['Data'].dt.to_period('M')
    df['Mês_Ordenacao'] = df['Data'].dt.to_period('M').astype(str)

    # Tradução manual dos meses para garantir que funcione
    meses_pt = {
        1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril',
        5: 'Maio', 6: 'Junho', 7: 'Julho', 8: 'Agosto',
        9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'
    }

    df['Mês_Nome'] = df['Data'].dt.month.map(meses_pt) + '/' + df['Data'].dt.year.astype(str)

    return df


# Header Principal
st.markdown(f"""
<div class="main-title">
    <h1>⚡ BI SESMT - Rezende Energia</h1>
    <p>Acompanhamento de Ações de Segurança do Trabalho</p>
</div>
""", unsafe_allow_html=True)

# Sidebar com upload
with st.sidebar:
    st.markdown("### 📤 Upload de Dados")
    uploaded_file = st.file_uploader(
        "Carregar planilha de acompanhamento",
        type=['xlsx', 'xls'],
        help="Faça upload da planilha de Acompanhamento de Ações SESMT"
    )

    st.markdown("---")
    st.markdown("### 📊 Navegação")
    st.info("Selecione uma aba acima para visualizar diferentes análises")

    st.markdown("---")
    st.markdown("### ℹ️ Sobre")
    st.markdown("**Rezende Energia**")
    st.markdown("Sistema de Business Intelligence para monitoramento de ações do SESMT")

# Verificar se há arquivo carregado
if uploaded_file is not None:
    # Carregar dados
    try:
        df = pd.read_excel(uploaded_file)
        df = processar_dados(df)

        # Sidebar - Filtros
        with st.sidebar:
            st.markdown("---")
            st.markdown("### 🔍 Filtros")

            # Filtro de período
            min_date = df['Data'].min()
            max_date = df['Data'].max()

            date_range = st.date_input(
                "Período",
                value=(min_date, max_date),
                min_value=min_date,
                max_value=max_date
            )

            # Filtro de tipo
            tipos = ['Todos'] + list(df['Tipo'].unique())
            tipo_selecionado = st.selectbox("Tipo de Ação", tipos)

            # Filtro de contrato
            contratos = ['Todos'] + list(df['Contrato'].unique())
            contrato_selecionado = st.selectbox("Contrato/Região", contratos)

            # Aplicar filtros
            df_filtrado = df.copy()

            if len(date_range) == 2:
                df_filtrado = df_filtrado[
                    (df_filtrado['Data'] >= pd.Timestamp(date_range[0])) &
                    (df_filtrado['Data'] <= pd.Timestamp(date_range[1]))
                    ]

            if tipo_selecionado != 'Todos':
                df_filtrado = df_filtrado[df_filtrado['Tipo'] == tipo_selecionado]

            if contrato_selecionado != 'Todos':
                df_filtrado = df_filtrado[df_filtrado['Contrato'] == contrato_selecionado]

        # Criar tabs
        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            "📊 Visão Geral",
            "📈 Análise por Categoria",
            "🗺️ Análise Regional",
            "👥 Ações Comunitárias",
            "📋 Dados Detalhados"
        ])

        # TAB 1 - VISÃO GERAL
        with tab1:
            # KPIs principais
            col1, col2, col3 = st.columns(3)

            with col1:
                st.metric(
                    "Total de Ações",
                    f"{len(df_filtrado)}",
                    delta=f"{len(df_filtrado) - len(df)} ações" if len(df_filtrado) != len(df) else None
                )

            with col2:
                total_pessoas = df_filtrado['Pessoas Impactadas'].sum()
                st.metric(
                    "Pessoas Impactadas",
                    f"{total_pessoas:,}".replace(',', '.')
                )

            with col3:
                media_participantes = df_filtrado['Pessoas Impactadas'].mean()
                st.metric(
                    "Média de Participantes",
                    f"{media_participantes:.0f}"
                )

            st.markdown("---")

            # Gráficos
            col1, col2 = st.columns(2)

            with col1:
                # Evolução de ações ao longo do tempo
                acoes_por_mes = df_filtrado.groupby(['Mês_Ordenacao', 'Mês_Nome']).size().reset_index(name='Quantidade')
                acoes_por_mes = acoes_por_mes.sort_values('Mês_Ordenacao')

                fig1 = px.line(
                    acoes_por_mes,
                    x='Mês_Nome',
                    y='Quantidade',
                    title='Evolução de Ações ao Longo do Tempo',
                    markers=True
                )
                fig1.update_traces(
                    line_color=COR_SECUNDARIA,
                    line_width=3,
                    marker=dict(size=10, color=COR_SECUNDARIA)
                )
                fig1.update_layout(**criar_layout_cores())
                st.plotly_chart(fig1, use_container_width=True)

            with col2:
                # Pessoas impactadas por mês
                pessoas_por_mes = df_filtrado.groupby(['Mês_Ordenacao', 'Mês_Nome'])[
                    'Pessoas Impactadas'].sum().reset_index()
                pessoas_por_mes = pessoas_por_mes.sort_values('Mês_Ordenacao')

                fig2 = px.bar(
                    pessoas_por_mes,
                    x='Mês_Nome',
                    y='Pessoas Impactadas',
                    title='Pessoas Impactadas por Mês',
                    color_discrete_sequence=[COR_SECUNDARIA]
                )
                fig2.update_layout(**criar_layout_cores())
                st.plotly_chart(fig2, use_container_width=True)

            # Distribuição por tipo
            st.markdown("### Distribuição por Tipo de Ação")
            tipo_dist = df_filtrado['Tipo'].value_counts().reset_index()
            tipo_dist.columns = ['Tipo', 'Quantidade']

            fig3 = px.pie(
                tipo_dist,
                values='Quantidade',
                names='Tipo',
                title='Distribuição de Ações por Tipo',
                color_discrete_sequence=[COR_SECUNDARIA, '#ff9d3d', '#ffb366', '#ffc999']
            )
            fig3.update_layout(**criar_layout_cores())
            fig3.update_traces(textposition='inside', textinfo='percent+label')
            st.plotly_chart(fig3, use_container_width=True)

        # TAB 2 - ANÁLISE POR CATEGORIA
        with tab2:
            st.markdown("### 📊 Performance por Tipo de Evento")

            col1, col2 = st.columns(2)

            with col1:
                # Ranking de eventos por quantidade
                eventos_ranking = df_filtrado['Evento'].value_counts().reset_index()
                eventos_ranking.columns = ['Evento', 'Quantidade']
                eventos_ranking = eventos_ranking.head(10)

                fig4 = px.bar(
                    eventos_ranking,
                    y='Evento',
                    x='Quantidade',
                    orientation='h',
                    title='Top 10 Eventos Mais Realizados',
                    color='Quantidade',
                    color_continuous_scale=[[0, COR_PRINCIPAL], [1, COR_SECUNDARIA]]
                )
                fig4.update_layout(**criar_layout_cores())
                st.plotly_chart(fig4, use_container_width=True)

            with col2:
                # Pessoas impactadas por tipo de evento
                pessoas_por_evento = df_filtrado.groupby('Evento')['Pessoas Impactadas'].sum().reset_index()
                pessoas_por_evento = pessoas_por_evento.sort_values('Pessoas Impactadas', ascending=False).head(10)

                fig5 = px.bar(
                    pessoas_por_evento,
                    y='Evento',
                    x='Pessoas Impactadas',
                    orientation='h',
                    title='Top 10 Eventos com Maior Alcance',
                    color='Pessoas Impactadas',
                    color_continuous_scale=[[0, COR_PRINCIPAL], [1, COR_SECUNDARIA]]
                )
                fig5.update_layout(**criar_layout_cores())
                st.plotly_chart(fig5, use_container_width=True)

            st.markdown("---")
            st.markdown("### 📋 Tabela Resumo por Evento")

            # Tabela dinâmica
            tabela_eventos = df_filtrado.groupby('Evento').agg({
                'Pessoas Impactadas': ['sum', 'mean', 'count']
            }).round(1)
            tabela_eventos.columns = ['Total Pessoas', 'Média Pessoas', 'Qtd Ações']
            tabela_eventos = tabela_eventos.sort_values('Total Pessoas', ascending=False)
            tabela_eventos = tabela_eventos.reset_index()

            st.dataframe(
                tabela_eventos,
                use_container_width=True,
                hide_index=True,
                height=400
            )

        # TAB 3 - ANÁLISE REGIONAL
        with tab3:
            st.markdown("### 🗺️ Comparativo Regional")

            col1, col2, col3 = st.columns(3)

            contratos = df_filtrado['Contrato'].unique()

            for idx, contrato in enumerate(contratos):
                with [col1, col2, col3][idx % 3]:
                    df_contrato = df_filtrado[df_filtrado['Contrato'] == contrato]
                    acoes_contrato = len(df_contrato)
                    pessoas_contrato = df_contrato['Pessoas Impactadas'].sum()

                    st.markdown(f"""
                    <div style='background: linear-gradient(135deg, {COR_PRINCIPAL} 0%, #333333 100%); 
                                padding: 20px; border-radius: 10px; color: white; text-align: center;'>
                        <h2 style='color: {COR_SECUNDARIA}; margin: 0;'>{contrato}</h2>
                        <p style='font-size: 1.2rem; margin: 10px 0;'><b>{acoes_contrato}</b> ações</p>
                        <p style='font-size: 1.2rem; margin: 10px 0;'><b>{pessoas_contrato:,}</b> pessoas</p>
                    </div>
                    """.replace(',', '.'), unsafe_allow_html=True)

            st.markdown("---")

            col1, col2 = st.columns(2)

            with col1:
                # Ações por região
                acoes_regiao = df_filtrado.groupby('Contrato').size().reset_index(name='Quantidade')
                fig6 = px.bar(
                    acoes_regiao,
                    x='Contrato',
                    y='Quantidade',
                    title='Ações Realizadas por Região',
                    color='Quantidade',
                    color_continuous_scale=[[0, COR_PRINCIPAL], [1, COR_SECUNDARIA]]
                )
                fig6.update_layout(**criar_layout_cores())
                st.plotly_chart(fig6, use_container_width=True)

            with col2:
                # Pessoas impactadas por região
                pessoas_regiao = df_filtrado.groupby('Contrato')['Pessoas Impactadas'].sum().reset_index()
                fig7 = px.bar(
                    pessoas_regiao,
                    x='Contrato',
                    y='Pessoas Impactadas',
                    title='Pessoas Impactadas por Região',
                    color='Pessoas Impactadas',
                    color_continuous_scale=[[0, COR_PRINCIPAL], [1, COR_SECUNDARIA]]
                )
                fig7.update_layout(**criar_layout_cores())
                st.plotly_chart(fig7, use_container_width=True)

            # Análise por colaborador e região
            st.markdown("### 👥 Performance por Colaborador e Região")

            tabela_colaborador = df_filtrado.groupby(['Contrato', 'Colaborador']).agg({
                'Pessoas Impactadas': 'sum',
                'Evento': 'count'
            }).round(1)
            tabela_colaborador.columns = ['Total Pessoas', 'Qtd Ações']
            tabela_colaborador = tabela_colaborador.sort_values('Total Pessoas', ascending=False)
            tabela_colaborador = tabela_colaborador.reset_index()

            st.dataframe(
                tabela_colaborador,
                use_container_width=True,
                hide_index=True
            )

        # TAB 4 - AÇÕES COMUNITÁRIAS
        with tab4:
            st.markdown("### 🤝 Impacto Comunitário")

            # Filtrar apenas ações comunitárias
            df_comunidade = df_filtrado[df_filtrado['Tipo'] == 'Comunidade']

            if len(df_comunidade) > 0:
                col1, col2, col3 = st.columns(3)

                with col1:
                    st.metric(
                        "Ações Comunitárias",
                        f"{len(df_comunidade)}"
                    )

                with col2:
                    st.metric(
                        "Pessoas da Comunidade",
                        f"{df_comunidade['Pessoas Impactadas'].sum():,}".replace(',', '.')
                    )

                with col3:
                    st.metric(
                        "Média por Ação",
                        f"{df_comunidade['Pessoas Impactadas'].mean():.0f}"
                    )

                st.markdown("---")

                # Timeline de campanhas
                st.markdown("### 📅 Timeline de Ações Comunitárias")

                df_comunidade_sorted = df_comunidade.sort_values('Data')

                fig8 = go.Figure()

                for idx, row in df_comunidade_sorted.iterrows():
                    fig8.add_trace(go.Scatter(
                        x=[row['Data']],
                        y=[row['Pessoas Impactadas']],
                        mode='markers+text',
                        marker=dict(size=15, color=COR_SECUNDARIA),
                        text=[row['Evento'][:30] + '...'],
                        textposition='top center',
                        name=row['Evento'],
                        hovertemplate=f"<b>{row['Evento']}</b><br>" +
                                      f"Data: {row['Data'].strftime('%d/%m/%Y')}<br>" +
                                      f"Pessoas: {row['Pessoas Impactadas']}<br>" +
                                      "<extra></extra>"
                    ))

                fig8.update_layout(
                    title='Timeline de Ações Comunitárias',
                    xaxis_title='Data',
                    yaxis_title='Pessoas Impactadas',
                    showlegend=False,
                    **criar_layout_cores()
                )
                st.plotly_chart(fig8, use_container_width=True)

                st.markdown("---")

                # Detalhes das ações comunitárias
                st.markdown("### 📋 Detalhes das Ações Comunitárias")

                for idx, row in df_comunidade_sorted.iterrows():
                    with st.expander(f"📍 {row['Data'].strftime('%d/%m/%Y')} - {row['Evento']}"):
                        col1, col2 = st.columns([2, 1])

                        with col1:
                            st.markdown(f"**Observações:**")
                            st.write(row['Observações'])

                        with col2:
                            st.markdown(f"**Pessoas Impactadas:** {row['Pessoas Impactadas']}")
                            st.markdown(f"**Responsável:** {row['Colaborador']}")
                            st.markdown(f"**Região:** {row['Contrato']}")
            else:
                st.info("Nenhuma ação comunitária encontrada no período selecionado.")

        # TAB 5 - DADOS DETALHADOS
        with tab5:
            st.markdown("### 📋 Tabela Completa de Ações")

            # Preparar dados para exibição
            df_exibicao = df_filtrado[
                ['Data', 'Evento', 'Pessoas Impactadas', 'Colaborador', 'Contrato', 'Tipo', 'Observações']].copy()
            df_exibicao['Data'] = df_exibicao['Data'].dt.strftime('%d/%m/%Y')

            # Mostrar estatísticas
            col1, col2, col3, col4 = st.columns(4)

            with col1:
                st.metric("Total de Registros", len(df_exibicao))

            with col2:
                st.metric("Tipos Diferentes", df_exibicao['Tipo'].nunique())

            with col3:
                st.metric("Eventos Diferentes", df_exibicao['Evento'].nunique())

            with col4:
                st.metric("Colaboradores", df_exibicao['Colaborador'].nunique())

            st.markdown("---")

            # Barra de pesquisa
            search = st.text_input("🔍 Pesquisar na tabela", "")

            if search:
                df_exibicao = df_exibicao[
                    df_exibicao.apply(lambda row: row.astype(str).str.contains(search, case=False).any(), axis=1)
                ]

            # Exibir tabela
            st.dataframe(
                df_exibicao,
                use_container_width=True,
                hide_index=True,
                height=600
            )

            # Botão de download
            csv = df_exibicao.to_csv(index=False).encode('utf-8-sig')
            st.download_button(
                label="📥 Baixar dados filtrados (CSV)",
                data=csv,
                file_name=f'acoes_sesmt_{datetime.now().strftime("%Y%m%d")}.csv',
                mime='text/csv',
            )

    except Exception as e:
        st.error(f"Erro ao processar arquivo: {str(e)}")
        st.info("Por favor, verifique se o arquivo está no formato correto.")

else:
    # Página inicial sem dados
    st.markdown("""
    <div style='text-align: center; padding: 50px;'>
        <h2 style='color: #666;'>👈 Faça o upload da planilha para começar</h2>
        <p style='color: #999; font-size: 1.1rem;'>
            Carregue o arquivo Excel com os dados de acompanhamento do SESMT na barra lateral.
        </p>
    </div>
    """, unsafe_allow_html=True)

    # Mostrar exemplo de estrutura esperada
    with st.expander("📖 Estrutura esperada do arquivo"):
        st.markdown("""
        A planilha deve conter as seguintes colunas:

        - **Data**: Data da ação
        - **Evento**: Nome/tipo do evento realizado
        - **Pessoas Impactadas**: Número de participantes
        - **Observações**: Detalhes da ação
        - **Colaborador**: Responsável pela ação
        - **Cargo**: Cargo do responsável
        - **Contrato**: Região/contrato (Oeste, Nordeste, etc.)
        - **Tipo**: Tipo de ação (Interno, Treinamento, Comunidade, EQTL)
        """)