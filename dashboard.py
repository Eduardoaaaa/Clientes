import streamlit as st
import pandas as pd
from sqlalchemy import create_engine
import plotly.express as px
import io 
from datetime import datetime, timedelta, timezone 

from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# 1. Configuração da página
st.set_page_config(page_title="Painel ABS Distribuidora", layout="wide", page_icon="📊")

# 2. Conexão com o Banco de Dados
SUPABASE_DB_URL = st.secrets["SUPABASE_DB_URL"]
FUSO_BR = timezone(timedelta(hours=-3))

@st.cache_data(ttl=60) 
def buscar_data_atualizacao(tabela):
    try:
        engine = create_engine(SUPABASE_DB_URL)
        df = pd.read_sql(f"SELECT data_atualizacao FROM log_atualizacoes WHERE tabela = '{tabela}'", engine)
        if not df.empty:
            ultima_data = df['data_atualizacao'].max()
            return pd.to_datetime(ultima_data).strftime('%d/%m/%Y às %H:%M')
        return "Aguardando envio..."
    except:
        return "Aguardando envio..."

def limpar_colunas_tarefas(df):
    if df.empty:
        return df
    for col in df.columns:
        col_lower = str(col).strip().lower()
        if col_lower in ['data visita', 'data_visita']:
            df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%d/%m/%Y').fillna('-')
        elif col_lower in ['setor', 'gv', 'qtd solicitada', 'qtd já comprada', 'qtd_solicitada', 'qtd_comprada']:
            df[col] = df[col].apply(lambda x: str(x)[:-2] if str(x).endswith('.0') else str(x))
            df[col] = df[col].replace(['nan', 'None', 'NaN', 'NaT'], '-')
    return df

@st.cache_data(ttl=300)
def buscar_dados_cliente(codigo):
    try:
        engine = create_engine(SUPABASE_DB_URL)
        query_vendas = f"SELECT * FROM vendas_consolidadas WHERE CAST(codigo_cliente AS TEXT) = '{codigo}'"
        df_vendas = pd.read_sql(query_vendas, engine)
        
        query_equip = f"SELECT * FROM equipamentos_clientes WHERE TRIM(CAST(codigo_cliente AS TEXT)) = '{codigo.strip()}'"
        df_equip = pd.read_sql(query_equip, engine)
        
        query_tarefas = f"SELECT * FROM tarefas_clientes WHERE CAST(codigo_cliente AS TEXT) = '{codigo}'"
        try:
            df_tarefas = pd.read_sql(query_tarefas, engine)
            df_tarefas.rename(columns=lambda x: str(x).strip(), inplace=True) # GARANTIA DE NOME EXATO
            df_tarefas = limpar_colunas_tarefas(df_tarefas)
        except:
            df_tarefas = pd.DataFrame()
            
        return df_vendas, df_equip, df_tarefas
    except Exception as e:
        st.error(f"Erro ao ligar à base de dados: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

@st.cache_data(ttl=300)
def buscar_dados_macro():
    try:
        engine = create_engine(SUPABASE_DB_URL)
        query = """
            SELECT TO_CHAR(data_venda, 'YYYY-MM') AS "Mes_Ano",
                   nome_cliente,
                   nome_produto,
                   codigo_produto,
                   categoria_produto,
                   equipamento,
                   SUM(faturamento_reais) AS faturamento_reais,
                   SUM(volume_hl) AS volume_hl,
                   SUM(volume_caixas) AS volume_caixas
            FROM vendas_consolidadas
            GROUP BY TO_CHAR(data_venda, 'YYYY-MM'), nome_cliente, nome_produto, codigo_produto, categoria_produto, equipamento
        """
        df_macro = pd.read_sql(query, engine)
        return df_macro
    except Exception as e:
        st.error(f"Erro ao carregar dados macro: {e}")
        return pd.DataFrame()

@st.cache_data(ttl=300)
def buscar_todas_tarefas():
    try:
        engine = create_engine(SUPABASE_DB_URL)
        df_todas = pd.read_sql("SELECT * FROM tarefas_clientes", engine)
        df_todas.rename(columns=lambda x: str(x).strip(), inplace=True) # GARANTIA DE NOME EXATO
        df_todas = limpar_colunas_tarefas(df_todas)
        return df_todas
    except Exception as e:
        st.error(f"⚠️ Erro ao buscar a tabela global de tarefas: {e}")
        return pd.DataFrame()

def gerar_excel_formatado(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Tarefas Exportadas')
        workbook = writer.book
        worksheet = writer.sheets['Tarefas Exportadas']
        num_rows, num_cols = df.shape
        if num_rows > 0 and num_cols > 0:
            col_settings = [{'header': str(c)} for c in df.columns]
            worksheet.add_table(0, 0, num_rows, num_cols - 1, {'columns': col_settings, 'style': 'Table Style Medium 9'})
            for i in range(num_cols):
                worksheet.set_column(i, i, 20)
    return output.getvalue()

def gerar_pdf_formatado(df):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
    elements = []
    
    styles = getSampleStyleSheet()
    
    style_normal = ParagraphStyle('TabelaNormal', parent=styles['Normal'], fontSize=7)
    style_header = ParagraphStyle('TabelaHeader', parent=styles['Normal'], fontSize=8, alignment=1)
    
    headers = [Paragraph(f"<font color='white'><b>{c}</b></font>", style_header) for c in df.columns]
    data = [headers]
    
    for _, row in df.iterrows():
        linha_formatada = []
        for item in row:
            texto_limpo = str(item).replace("<", "&lt;").replace(">", "&gt;")
            linha_formatada.append(Paragraph(f"<font color='black'>{texto_limpo}</font>", style_normal))
        data.append(linha_formatada)
        
    total_width = 800
    col_widths = []
    for col in df.columns:
        c_name = str(col).strip().lower()
        if c_name in ['texto da tarefa', 'texto_da_tarefa']:
            col_widths.append(total_width * 0.35)
        elif c_name in ['nome fantasia', 'nome_fantasia']:
            col_widths.append(total_width * 0.15)
        elif c_name in ['qtd solicitada', 'qtd já comprada', 'gv', 'setor', 'operação']:
            col_widths.append(total_width * 0.05)
        else:
            col_widths.append(total_width * 0.08)
            
    factor = total_width / sum(col_widths)
    col_widths = [w * factor for w in col_widths]
    
    t = Table(data, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#004A99")),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.lightgrey),
        ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.whitesmoke, colors.white])
    ]))
    
    elements.append(t)
    doc.build(elements)
    return buffer.getvalue()

# ---> NOVA ESTRUTURA VISUAL (Sem Mês/Ano e com Nome Fantasia após o código) <---
colunas_ordem_tarefas = [
    'Data Visita', 'Operação', 'codigo_cliente', 'Nome Fantasia', 'GV', 
    'Setor', 'Cluster Primário', 'Categoria', 'QTD Solicitada', 
    'QTD Já Comprada', 'Texto da Tarefa'
]

# --- MENU LATERAL DE NAVEGAÇÃO E CONFIGURAÇÃO ---
st.sidebar.image("https://cdn-icons-png.flaticon.com/512/2666/2666505.png", width=120)
st.sidebar.markdown("## Menu de Navegação")

menu = st.sidebar.radio(
    "Selecione o nível de análise:",
    ["🎯 Visão Macro (Distribuidora)", "👤 Visão Micro (Por Cliente)", "📋 Planificador de Tarefas"]
)

st.sidebar.markdown("---")
st.sidebar.markdown("## Configurações de Visualização")

opcao_metrica = st.sidebar.radio("Visualizar dados por:", ["Faturamento (R$)", "Volume (HL)"])

if opcao_metrica == "Faturamento (R$)":
    col_metrica = "faturamento_reais"
    prefixo_unidade = "R$ "
    sufixo_unidade = ""
    formato_num = ",.2f"
    label_kpi = "Faturamento"
else:
    col_metrica = "volume_hl"
    prefixo_unidade = ""
    sufixo_unidade = " HL"
    formato_num = ",.2f"
    label_kpi = "Volume"

# =====================================================================
# PÁGINA 1: VISÃO MACRO (GERAL)
# =====================================================================
if menu == "🎯 Visão Macro (Distribuidora)":
    st.title("🎯 Visão Macro - ABS Distribuidora")
    
    with st.spinner('A carregar base consolidada da nuvem...'):
        df_macro = buscar_dados_macro()
        
    st.markdown(f"Acompanhamento global configurado para a análise de **{opcao_metrica}**.")
    st.caption(f"🔄 Nuvem atualizada com Vendas em: **{buscar_data_atualizacao('vendas')}**")

    if not df_macro.empty:
        st.write("---")
        st.markdown("### 📅 Selecione o Período de Análise")
        
        meses_disponiveis = sorted(df_macro['Mes_Ano'].dropna().unique())
        mes_inicio, mes_fim = st.select_slider(
            "Arraste as extremidades para definir o intervalo:",
            options=meses_disponiveis,
            value=(meses_disponiveis[0], meses_disponiveis[-1]),
            label_visibility="collapsed",
            key="slider_macro"
        )
        
        df_filtrado_macro = df_macro[(df_macro['Mes_Ano'] >= mes_inicio) & (df_macro['Mes_Ano'] <= mes_fim)]

        st.write("---")
        st.markdown("### 🔍 Filtro de Categoria")
        categorias_disp = sorted(df_filtrado_macro['categoria_produto'].dropna().unique().tolist())
        col_filtro, _ = st.columns([1, 2])
        with col_filtro:
            cat_selecionadas = st.multiselect("Filtrar por Categoria:", options=categorias_disp, default=categorias_disp, key="cat_macro")
        
        if cat_selecionadas:
            df_filtrado_macro = df_filtrado_macro[df_filtrado_macro['categoria_produto'].isin(cat_selecionadas)]

        st.write("")
        valor_global = df_filtrado_macro[col_metrica].sum()
        total_clientes = df_filtrado_macro['nome_cliente'].nunique()

        c1, c2 = st.columns(2)
        c1.metric(f"{label_kpi} Global", f"{prefixo_unidade}{valor_global:{formato_num}}".replace(",", "X").replace(".", ",").replace("X", ".") + sufixo_unidade)
        c2.metric("Clientes Positivados", f"{total_clientes}")

        st.write("---")
        col_esq, col_dir = st.columns([2, 1])

        with col_esq:
            st.markdown(f"#### 📈 Evolução Mensal de {label_kpi}")
            df_evo = df_filtrado_macro.groupby('Mes_Ano')[col_metrica].sum().reset_index().sort_values('Mes_Ano')
            if not df_evo.empty:
                df_evo['Rotulo'] = df_evo[col_metrica].apply(lambda x: f"{prefixo_unidade}{x:{formato_num}}".replace(",", "X").replace(".", ",").replace("X", ".") + sufixo_unidade)
                fig_evo = px.bar(df_evo, x='Mes_Ano', y=col_metrica, text='Rotulo')
                fig_evo.update_traces(textposition='outside', textfont_size=14, marker_color="#004A99", cliponaxis=False)
                fig_evo.update_layout(plot_bgcolor='rgba(0,0,0,0)', xaxis_title=None, yaxis_title=None, margin=dict(t=30), xaxis=dict(type='category'))
                st.plotly_chart(fig_evo, use_container_width=True)

        with col_dir:
            st.markdown(f"#### 🍕 Mix de Categorias ({label_kpi})")
            df_mix_cat = df_filtrado_macro.groupby('categoria_produto')[col_metrica].sum().reset_index()
            if not df_mix_cat.empty:
                fig_mix = px.pie(df_mix_cat, values=col_metrica, names='categoria_produto', hole=0.4, color_discrete_sequence=px.colors.qualitative.Pastel)
                st.plotly_chart(fig_mix, use_container_width=True)

        st.divider()
        st.subheader(f"🏆 Top 15 Clientes (Maior {label_kpi})")
        
        df_top_cli = df_filtrado_macro.groupby('nome_cliente')[col_metrica].sum().reset_index()
        df_top_cli = df_top_cli.sort_values(col_metrica, ascending=False).head(15)
        
        if not df_top_cli.empty:
            df_top_cli = df_top_cli.iloc[::-1]
            df_top_cli['Rotulo'] = df_top_cli[col_metrica].apply(lambda x: f"{prefixo_unidade}{x:{formato_num}}".replace(",", "X").replace(".", ",").replace("X", ".") + sufixo_unidade)
            
            fig_cli = px.bar(df_top_cli, x=col_metrica, y='nome_cliente', orientation='h', text='Rotulo', color=col_metrica, color_continuous_scale='Blues')
            fig_cli.update_traces(textposition='outside', textfont_size=14, cliponaxis=False)
            fig_cli.update_layout(xaxis=dict(showticklabels=False, showgrid=False), yaxis_title=None, coloraxis_showscale=False, plot_bgcolor='rgba(0,0,0,0)', margin=dict(l=0, r=100, t=10, b=0), height=500)
            st.plotly_chart(fig_cli, use_container_width=True, config={'displayModeBar': False})


# =====================================================================
# PÁGINA 2: VISÃO MICRO (POR CLIENTE)
# =====================================================================
elif menu == "👤 Visão Micro (Por Cliente)":
    st.title("👤 Portal de Autoatendimento - Cliente")
    codigo_input = st.text_input("🔎 Digite o Código do Cliente para iniciar:", placeholder="Ex: 9528")

    if codigo_input:
        df_cliente, df_equip, df_tarefas = buscar_dados_cliente(codigo_input)

        if df_cliente.empty:
            st.warning(f"⚠️ Nenhum histórico encontrado para o código {codigo_input}.")
        else:
            df_cliente['data_venda'] = pd.to_datetime(df_cliente['data_venda'])
            df_cliente['Mes_Ano'] = df_cliente['data_venda'].dt.to_period('M').astype(str)
            nome_cliente = df_cliente['nome_cliente'].iloc[0]
            
            st.subheader(f"👤 Cliente: {codigo_input} - {nome_cliente}")
            st.caption(f"🔄 Dados em Nuvem -> Vendas: **{buscar_data_atualizacao('vendas')}** | Tarefas: **{buscar_data_atualizacao('tarefas')}**")
            
            tab_resumo, tab_tarefas = st.tabs(["📊 Resumo Financeiro e Mix", "📋 Planificador do Cliente"])
            
            with tab_resumo:
                st.write("---")
                st.markdown("### 📅 Selecione o Período de Análise")
                meses_disponiveis = sorted(df_cliente['Mes_Ano'].unique())
                
                if len(meses_disponiveis) > 1:
                    mes_inicio, mes_fim = st.select_slider(
                        "Arraste as extremidades para definir o intervalo:",
                        options=meses_disponiveis,
                        value=(meses_disponiveis[0], meses_disponiveis[-1]),
                        label_visibility="collapsed",
                        key="slider_micro"
                    )
                    df_filtrado = df_cliente[(df_cliente['Mes_Ano'] >= mes_inicio) & (df_cliente['Mes_Ano'] <= mes_fim)]
                else:
                    st.info(f"Este cliente possui apenas um mês de histórico: {meses_disponiveis[0]}")
                    df_filtrado = df_cliente

                st.write("---")
                st.markdown("### 🔍 Filtro de Categoria")
                categorias_disponiveis = sorted(df_filtrado['categoria_produto'].dropna().unique().tolist())
                col_filtro, _ = st.columns([1, 2])
                with col_filtro:
                    categorias_selecionadas = st.multiselect("Selecione o Tipo de Produto:", options=categorias_disponiveis, default=categorias_disponiveis, key="cat_micro")
                    
                if categorias_selecionadas:
                    df_filtrado = df_filtrado[df_filtrado['categoria_produto'].isin(categorias_selecionadas)]
                else:
                    df_filtrado = df_filtrado.copy()

                st.write("")
                valor_kpi_micro = df_filtrado[col_metrica].sum()
                total_vol_caixas = df_filtrado['volume_caixas'].sum()
                dias_compra = df_filtrado['data_venda'].nunique()

                c1, c2, c3 = st.columns(3)
                c1.metric(label_kpi, f"{prefixo_unidade}{valor_kpi_micro:{formato_num}}".replace(",", "X").replace(".", ",").replace("X", ".") + sufixo_unidade)
                c2.metric("Volume Físico", f"{total_vol_caixas:,.0f} cx".replace(",", "."))
                c3.metric("Frequência de Pedidos", f"{dias_compra} dias")

                st.write("---")
                col_esq, col_dir = st.columns([2, 1])

                with col_esq:
                    st.markdown(f"#### 📈 Evolução Mensal ({label_kpi})")
                    resumo_grafico = df_filtrado.groupby('Mes_Ano').agg({col_metrica: 'sum'}).reset_index().sort_values('Mes_Ano')
                    
                    if not resumo_grafico.empty:
                        resumo_grafico['Rotulo_Valor'] = resumo_grafico[col_metrica].apply(lambda x: f"{prefixo_unidade}{x:{formato_num}}".replace(",", "X").replace(".", ",").replace("X", ".") + sufixo_unidade)
                        fig = px.bar(resumo_grafico, x='Mes_Ano', y=col_metrica, text='Rotulo_Valor')
                        fig.update_traces(textposition='outside', marker_color="#004A99", textfont_size=16, textfont_color="white", cliponaxis=False)
                        fig.update_layout(xaxis_title=None, yaxis_title=None, plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', margin=dict(t=50, b=0, l=0, r=0), xaxis=dict(type='category', tickangle=0))
                        st.plotly_chart(fig, use_container_width=True)
                    else:
                        st.info("Selecione um período para visualizar.")

                with col_dir:
                    st.markdown("#### 🧊 Giro de Equipamentos")
                    if not df_equip.empty:
                        df_equip = df_equip.groupby('tipo_equipamento')['quantidade'].sum().reset_index()
                        num_meses = (df_filtrado['Mes_Ano'].nunique()) or 1
                        
                        for _, row in df_equip.iterrows():
                            tipo_eq = str(row['tipo_equipamento'])
                            qtd = int(row['quantidade'])
                            
                            if tipo_eq == 'VISA':
                                meta = 1200 * qtd * num_meses
                                fat_realizado = df_filtrado[df_filtrado['equipamento'] == 'Visa']['faturamento_reais'].sum()
                                titulo = f"🥤 {qtd}x VISA (NAB)"
                            elif tipo_eq == 'SOPI':
                                meta = 2000 * qtd * num_meses
                                fat_realizado = df_filtrado[df_filtrado['equipamento'] == 'Sopi']['faturamento_reais'].sum()
                                titulo = f"🍺 {qtd}x SOPI (Cerveja)"
                            elif tipo_eq == 'CHOPEIRA':
                                meta = 3870 * qtd * num_meses
                                codigos_chopp = ['838', '8037']
                                fat_realizado = df_filtrado[df_filtrado['codigo_produto'].astype(str).isin(codigos_chopp)]['faturamento_reais'].sum()
                                titulo = f"🍻 {qtd}x CHOPEIRA (Chopp)"
                            else:
                                continue
                                
                            st.write(f"**{titulo}**")
                            st.caption(f"Meta período: R$ {meta:,.0f}".replace(",", "."))
                            pct = min(fat_realizado / meta, 1.0) if meta > 0 else 0
                            st.progress(pct)
                            
                            realizado_str = f"R$ {fat_realizado:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                            if fat_realizado >= meta:
                                st.markdown(f"Realizado: **{realizado_str}**")
                                st.success("✅ **Meta Atingida!** Equipamento justificado.")
                            else:
                                gap = meta - fat_realizado
                                gap_str = f"R$ {gap:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                                st.markdown(f"Realizado: **{realizado_str}**")
                                st.error(f"⚠️ **GAP: Faltam {gap_str}**")
                            st.write("---")
                    else:
                        st.info("✅ Cliente sem equipamentos comodatados na base.")

                st.divider()
                st.subheader("📦 Análise de Volume por Produto")
                produtos_disponiveis = sorted(df_filtrado['nome_produto'].unique())
                produtos_selecionados = st.multiselect("Selecione um ou mais produtos para analisar:", options=produtos_disponiveis)

                if produtos_selecionados:
                    df_mix = df_filtrado[df_filtrado['nome_produto'].isin(produtos_selecionados)]
                    resumo_mix = df_mix.groupby(['Mes_Ano', 'nome_produto']).agg({'volume_caixas': 'sum'}).reset_index().sort_values('Mes_Ano')

                    if not resumo_mix.empty:
                        resumo_mix['volume_caixas'] = resumo_mix['volume_caixas'].round(0)
                        resumo_mix['Rotulo'] = resumo_mix['volume_caixas'].apply(lambda x: f"{x:,.0f} cx".replace(',', '.'))

                        fig_mix_prod = px.bar(resumo_mix, x='Mes_Ano', y='volume_caixas', color='nome_produto', text='Rotulo')
                        max_vol = resumo_mix['volume_caixas'].max()
                        fig_mix_prod.update_traces(texttemplate='%{text}', textposition='outside', textfont_size=16, textfont_color="white", cliponaxis=False)
                        fig_mix_prod.update_layout(barmode='group', xaxis=dict(type='category'), yaxis=dict(range=[0, max_vol * 1.2], showticklabels=False, showgrid=False), xaxis_title=None, yaxis_title=None, plot_bgcolor='rgba(0,0,0,0)', legend=dict(title="", orientation="h", yanchor="bottom", y=-0.3, xanchor="center", x=0.5), margin=dict(t=40, b=80))
                        st.plotly_chart(fig_mix_prod, use_container_width=True, config={'displayModeBar': False})
                        
                        with st.expander("Ver detalhamento em tabela"):
                            tabela_mix = df_mix.groupby(['codigo_produto', 'nome_produto', 'categoria_produto']).agg({'volume_caixas': 'sum'}).reset_index().sort_values('volume_caixas', ascending=False)
                            tabela_mix['volume_caixas'] = tabela_mix['volume_caixas'].map('{:,.0f} cx'.format).str.replace(',', '.')
                            st.dataframe(tabela_mix, use_container_width=True, hide_index=True)
                else:
                    st.info("👆 Selecione um ou mais produtos no filtro acima para gerar o gráfico de evolução.")

                st.divider()
                st.subheader("🏆 Ranking Geral do Mix de Produtos (Caixas)")
                
                ranking_mix = df_filtrado.groupby('nome_produto').agg({'volume_caixas': 'sum'}).reset_index().sort_values('volume_caixas', ascending=False)
                
                if not ranking_mix.empty:
                    ranking_mix = ranking_mix.iloc[::-1]
                    altura_grafico = max(450, len(ranking_mix) * 35)
                    
                    fig_ranking = px.bar(
                        ranking_mix, 
                        x='volume_caixas', 
                        y='nome_produto', 
                        orientation='h', 
                        text='volume_caixas', 
                        color='volume_caixas', 
                        color_continuous_scale='RdYlGn'
                    )
                    
                    fig_ranking.update_traces(texttemplate='%{text:,.0f} cx', textposition='outside', textfont_size=14, cliponaxis=False)
                    fig_ranking.update_layout(xaxis=dict(showticklabels=False, showgrid=False), yaxis_title=None, showlegend=False, coloraxis_showscale=False, plot_bgcolor='rgba(0,0,0,0)', margin=dict(l=0, r=80, t=30, b=0), height=altura_grafico)
                    st.plotly_chart(fig_ranking, use_container_width=True, config={'displayModeBar': False})
                else:
                    st.info("Dados insuficientes para gerar o ranking no período selecionado.")

            with tab_tarefas:
                st.markdown("### 📋 Planificador de Execução (Tabela Original)")
                if not df_tarefas.empty:
                    colunas_presentes = [c for c in colunas_ordem_tarefas if c in df_tarefas.columns]
                    df_tarefas_limpo = df_tarefas[colunas_presentes]
                    
                    col_espaco, col_excel, col_pdf = st.columns([2, 1, 1])
                    
                    with col_excel:
                        excel_cliente = gerar_excel_formatado(df_tarefas_limpo)
                        st.download_button(
                            label="📥 Baixar Excel",
                            data=excel_cliente,
                            file_name=f"Tarefas_Cliente_{codigo_input}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        
                    with col_pdf:
                        pdf_cliente = gerar_pdf_formatado(df_tarefas_limpo)
                        st.download_button(
                            label="📄 Baixar PDF",
                            data=pdf_cliente,
                            file_name=f"Tarefas_Cliente_{codigo_input}.pdf",
                            mime="application/pdf"
                        )
                    
                    st.dataframe(df_tarefas_limpo, use_container_width=True, hide_index=True)
                else:
                    st.info("Nenhuma tarefa mapeada para este cliente no ficheiro Excel.")
    else:
        st.info("👋 Bem-vindo! Utilize o campo acima para buscar as informações de um cliente pelo código.")


# =====================================================================
# PÁGINA 3: PLANIFICADOR GLOBAL (A TABELA COM FILTROS DA SUA IMAGEM)
# =====================================================================
elif menu == "📋 Planificador de Tarefas":
    st.title("📋 Planificador Global de Tarefas")
    
    with st.spinner("A carregar base de tarefas global..."):
        df_todas_tarefas = buscar_todas_tarefas()
        
    st.markdown("Visão completa de execução e missões com filtros ativos da base de dados.")
    st.caption(f"🔄 Nuvem atualizada com Tarefas em: **{buscar_data_atualizacao('tarefas')}**")

    if not df_todas_tarefas.empty:
        # Função à prova de balas para mapear os filtros
        def obter_coluna(df, possiveis_nomes):
            col_map = {str(c).lower().strip(): c for c in df.columns}
            for n in possiveis_nomes:
                if str(n).lower().strip() in col_map:
                    return col_map[str(n).lower().strip()]
            return None
            
        c_cliente = obter_coluna(df_todas_tarefas, ['codigo_cliente', 'codigo cliente', 'cliente'])
        c_cat = obter_coluna(df_todas_tarefas, ['Categoria', 'categoria'])
        c_clust = obter_coluna(df_todas_tarefas, ['Cluster Primário', 'Cluster Primario', 'cluster_primario'])
        c_setor = obter_coluna(df_todas_tarefas, ['Setor', 'setor'])
        c_data_visita = obter_coluna(df_todas_tarefas, ['Data Visita', 'Data_Visita', 'data_visita'])

        with st.expander("🔍 Filtros de Segmentação e Rota", expanded=True):
            col1, col2, col3, col4, col5 = st.columns(5)
            
            with col1:
                if c_cliente:
                    opt = sorted(df_todas_tarefas[c_cliente].dropna().astype(str).unique())
                    f_cliente = st.multiselect("Código Cliente", opt)
                else: f_cliente = []
            with col2:
                if c_cat:
                    opt = sorted(df_todas_tarefas[c_cat].dropna().astype(str).unique())
                    f_cat = st.multiselect("Categoria", opt)
                else: f_cat = []
            with col3:
                if c_clust:
                    opt = sorted(df_todas_tarefas[c_clust].dropna().astype(str).unique())
                    f_clust = st.multiselect("Cluster Primário", opt)
                else: f_clust = []
            with col4:
                if c_setor:
                    opt = sorted(df_todas_tarefas[c_setor].dropna().astype(str).unique())
                    f_setor = st.multiselect("Setor", opt)
                else: f_setor = []
            with col5:
                if c_data_visita:
                    opt = sorted(df_todas_tarefas[c_data_visita].dropna().astype(str).unique())
                    f_data_visita = st.multiselect("Data Visita", opt)
                else: f_data_visita = []

        df_filtrado = df_todas_tarefas.copy()
        
        if f_cliente and c_cliente: df_filtrado = df_filtrado[df_filtrado[c_cliente].astype(str).isin(f_cliente)]
        if f_cat and c_cat: df_filtrado = df_filtrado[df_filtrado[c_cat].astype(str).isin(f_cat)]
        if f_clust and c_clust: df_filtrado = df_filtrado[df_filtrado[c_clust].astype(str).isin(f_clust)]
        if f_setor and c_setor: df_filtrado = df_filtrado[df_filtrado[c_setor].astype(str).isin(f_setor)]
        if f_data_visita and c_data_visita: df_filtrado = df_filtrado[df_filtrado[c_data_visita].astype(str).isin(f_data_visita)]
        
        colunas_presentes = [c for c in colunas_ordem_tarefas if c in df_filtrado.columns]
        df_exibicao = df_filtrado[colunas_presentes]
        
        col_resumo, col_excel, col_pdf = st.columns([2, 1, 1])
        with col_resumo:
            st.markdown(f"**Total de Tarefas em Tela:** `{len(df_exibicao)} linhas`")
        
        with col_excel:
            excel_global = gerar_excel_formatado(df_exibicao)
            st.download_button(
                label="📥 Baixar Excel",
                data=excel_global,
                file_name="Planificador_Filtrado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        with col_pdf:
            pdf_global = gerar_pdf_formatado(df_exibicao)
            st.download_button(
                label="📄 Baixar PDF",
                data=pdf_global,
                file_name="Planificador_Filtrado.pdf",
                mime="application/pdf"
            )
            
        st.dataframe(df_exibicao, use_container_width=True, hide_index=True, height=600)
    else:
        st.warning("A tabela de tarefas está vazia. Execute o ficheiro `atualizar_tarefas.py` com o ficheiro Excel na pasta para subir os dados.")
