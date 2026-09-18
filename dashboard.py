import streamlit as st
import pandas as pd
import plotly.express as px
import io
import os
from datetime import datetime, timedelta, timezone 

from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

st.set_page_config(page_title="Painel ABS Distribuidora", layout="wide", page_icon="📊")

@st.cache_data
def carregar_dados_vendas():
    if os.path.exists('vendas.parquet'):
        return pd.read_parquet('vendas.parquet')
    return pd.DataFrame()

@st.cache_data
def carregar_dados_equipamentos():
    if os.path.exists('equipamentos.parquet'):
        return pd.read_parquet('equipamentos.parquet')
    return pd.DataFrame()

@st.cache_data
def carregar_dados_tarefas():
    if os.path.exists('tarefas.parquet'):
        return pd.read_parquet('tarefas.parquet')
    return pd.DataFrame()

def limpar_colunas_tarefas(df):
    if df.empty: return df
    for col in df.columns:
        col_lower = str(col).strip().lower()
        if col_lower in ['data visita', 'data_visita']:
            df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%d/%m/%Y').fillna('-')
        elif col_lower in ['setor', 'gv', 'qtd solicitada', 'qtd já comprada', 'qtd_solicitada', 'qtd_comprada']:
            df[col] = df[col].apply(lambda x: str(x)[:-2] if str(x).endswith('.0') else str(x))
            df[col] = df[col].replace(['nan', 'None', 'NaN', 'NaT'], '-')
    return df

@st.cache_data
def buscar_dados_macro():
    df_vendas = carregar_dados_vendas()
    if df_vendas.empty: return df_vendas
    
    df_vendas['Mes_Ano'] = pd.to_datetime(df_vendas['data_venda']).dt.strftime('%Y-%m')
    df_macro = df_vendas.groupby(
        ['Mes_Ano', 'nome_cliente', 'nome_produto', 'codigo_produto', 'categoria_produto', 'equipamento'], 
        as_index=False
    )[['faturamento_reais', 'volume_hl', 'volume_caixas']].sum()
    return df_macro

@st.cache_data
def buscar_dados_cliente(codigo):
    df_v = carregar_dados_vendas()
    df_e = carregar_dados_equipamentos()
    df_t = carregar_dados_tarefas()
    
    codigo_str = str(codigo).strip()
    
    df_v_cli = df_v[df_v['codigo_cliente'].astype(str) == codigo_str] if not df_v.empty else pd.DataFrame()
    df_e_cli = df_e[df_e['codigo_cliente'].astype(str) == codigo_str] if not df_e.empty else pd.DataFrame()
    
    df_t_cli = df_t[df_t['codigo_cliente'].astype(str) == codigo_str] if not df_t.empty else pd.DataFrame()
    df_t_cli = limpar_colunas_tarefas(df_t_cli)
    
    return df_v_cli, df_e_cli, df_t_cli

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
            for i in range(num_cols): worksheet.set_column(i, i, 20)
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
        if c_name in ['texto da tarefa', 'texto_da_tarefa']: col_widths.append(total_width * 0.35)
        elif c_name in ['nome fantasia', 'nome_fantasia']: col_widths.append(total_width * 0.15)
        elif c_name in ['qtd solicitada', 'qtd já comprada', 'gv', 'setor', 'operação']: col_widths.append(total_width * 0.05)
        else: col_widths.append(total_width * 0.08)
            
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

colunas_ordem_tarefas = [
    'Data Visita', 'Operação', 'codigo_cliente', 'Nome Fantasia', 'GV', 
    'Setor', 'Cluster Primário', 'Categoria', 'QTD Solicitada', 
    'QTD Já Comprada', 'Texto da Tarefa'
]

st.sidebar.image("https://cdn-icons-png.flaticon.com/512/2666/2666505.png", width=120)
st.sidebar.markdown("## Menu de Navegação")
menu = st.sidebar.radio("Selecione o nível de análise:", ["🎯 Visão Macro (Distribuidora)", "👤 Visão Micro (Por Cliente)", "📋 Planificador de Tarefas"])
st.sidebar.markdown("---")
opcao_metrica = st.sidebar.radio("Visualizar dados por:", ["Faturamento (R$)", "Volume (HL)"])

if opcao_metrica == "Faturamento (R$)":
    col_metrica = "faturamento_reais"
    prefixo_unidade, sufixo_unidade, label_kpi = "R$ ", "", "Faturamento"
else:
    col_metrica = "volume_hl"
    prefixo_unidade, sufixo_unidade, label_kpi = "", " HL", "Volume"
formato_num = ",.2f"

# ==================== PÁGINA 1 ====================
if menu == "🎯 Visão Macro (Distribuidora)":
    st.title("🎯 Visão Macro - ABS Distribuidora")
    
    with st.spinner('A carregar ficheiros locais...'):
        df_macro = buscar_dados_macro()
        
    st.markdown(f"Acompanhamento global configurado para a análise de **{opcao_metrica}**.")

    if not df_macro.empty:
        st.write("---")
        meses_disponiveis = sorted(df_macro['Mes_Ano'].dropna().unique())
        mes_inicio, mes_fim = st.select_slider("Período:", options=meses_disponiveis, value=(meses_disponiveis[0], meses_disponiveis[-1]), label_visibility="collapsed")
        df_filtrado_macro = df_macro[(df_macro['Mes_Ano'] >= mes_inicio) & (df_macro['Mes_Ano'] <= mes_fim)]

        categorias_disp = sorted(df_filtrado_macro['categoria_produto'].dropna().unique().tolist())
        cat_selecionadas = st.multiselect("Filtrar por Categoria:", options=categorias_disp, default=categorias_disp)
        if cat_selecionadas: df_filtrado_macro = df_filtrado_macro[df_filtrado_macro['categoria_produto'].isin(cat_selecionadas)]

        c1, c2 = st.columns(2)
        c1.metric(f"{label_kpi} Global", f"{prefixo_unidade}{df_filtrado_macro[col_metrica].sum():{formato_num}}".replace(",", "X").replace(".", ",").replace("X", ".") + sufixo_unidade)
        c2.metric("Clientes Positivados", f"{df_filtrado_macro['nome_cliente'].nunique()}")

        col_esq, col_dir = st.columns([2, 1])
        with col_esq:
            st.markdown(f"#### 📈 Evolução Mensal")
            df_evo = df_filtrado_macro.groupby('Mes_Ano')[col_metrica].sum().reset_index()
            if not df_evo.empty:
                df_evo['Rotulo'] = df_evo[col_metrica].apply(lambda x: f"{prefixo_unidade}{x:{formato_num}}".replace(",", "X").replace(".", ",").replace("X", ".") + sufixo_unidade)
                fig_evo = px.bar(df_evo, x='Mes_Ano', y=col_metrica, text='Rotulo')
                fig_evo.update_traces(textposition='outside', marker_color="#004A99", cliponaxis=False)
                fig_evo.update_layout(plot_bgcolor='rgba(0,0,0,0)', xaxis_title=None, yaxis_title=None)
                st.plotly_chart(fig_evo, use_container_width=True)

        with col_dir:
            st.markdown(f"#### 🍕 Mix de Categorias")
            df_mix = df_filtrado_macro.groupby('categoria_produto')[col_metrica].sum().reset_index()
            if not df_mix.empty:
                st.plotly_chart(px.pie(df_mix, values=col_metrica, names='categoria_produto', hole=0.4), use_container_width=True)

        st.subheader(f"🏆 Top 15 Clientes")
        df_top = df_filtrado_macro.groupby('nome_cliente')[col_metrica].sum().reset_index().sort_values(col_metrica, ascending=False).head(15).iloc[::-1]
        if not df_top.empty:
            df_top['Rotulo'] = df_top[col_metrica].apply(lambda x: f"{prefixo_unidade}{x:{formato_num}}".replace(",", "X").replace(".", ",").replace("X", ".") + sufixo_unidade)
            fig_cli = px.bar(df_top, x=col_metrica, y='nome_cliente', orientation='h', text='Rotulo')
            fig_cli.update_traces(textposition='outside', cliponaxis=False)
            fig_cli.update_layout(xaxis=dict(showticklabels=False), yaxis_title=None, plot_bgcolor='rgba(0,0,0,0)', height=500)
            st.plotly_chart(fig_cli, use_container_width=True, config={'displayModeBar': False})
    else:
        st.warning("Gere o ficheiro vendas.parquet usando o script local.")

# ==================== PÁGINA 2 ====================
elif menu == "👤 Visão Micro (Por Cliente)":
    st.title("👤 Portal de Autoatendimento - Cliente")
    codigo_input = st.text_input("🔎 Digite o Código do Cliente:", placeholder="Ex: 9528")

    if codigo_input:
        df_cliente, df_equip, df_tarefas = buscar_dados_cliente(codigo_input)

        if df_cliente.empty:
            st.warning(f"⚠️ Nenhum histórico encontrado para o código {codigo_input}.")
        else:
            df_cliente['data_venda'] = pd.to_datetime(df_cliente['data_venda'])
            df_cliente['Mes_Ano'] = df_cliente['data_venda'].dt.to_period('M').astype(str)
            st.subheader(f"👤 Cliente: {codigo_input} - {df_cliente['nome_cliente'].iloc[0]}")
            
            tab_resumo, tab_tarefas = st.tabs(["📊 Resumo Financeiro e Mix", "📋 Planificador do Cliente"])
            
            with tab_resumo:
                meses_disponiveis = sorted(df_cliente['Mes_Ano'].unique())
                mes_inicio, mes_fim = st.select_slider("Período:", options=meses_disponiveis, value=(meses_disponiveis[0], meses_disponiveis[-1]), label_visibility="collapsed")
                df_filtrado = df_cliente[(df_cliente['Mes_Ano'] >= mes_inicio) & (df_cliente['Mes_Ano'] <= mes_fim)]

                cat_disp = sorted(df_filtrado['categoria_produto'].dropna().unique().tolist())
                cat_sel = st.multiselect("Categoria:", options=cat_disp, default=cat_disp)
                if cat_sel: df_filtrado = df_filtrado[df_filtrado['categoria_produto'].isin(cat_sel)]

                c1, c2, c3 = st.columns(3)
                c1.metric(label_kpi, f"{prefixo_unidade}{df_filtrado[col_metrica].sum():{formato_num}}".replace(",", "X").replace(".", ",").replace("X", ".") + sufixo_unidade)
                c2.metric("Volume Físico", f"{df_filtrado['volume_caixas'].sum():,.0f} cx".replace(",", "."))
                c3.metric("Frequência de Pedidos", f"{df_filtrado['data_venda'].nunique()} dias")

                col_esq, col_dir = st.columns([2, 1])
                with col_esq:
                    resumo_grafico = df_filtrado.groupby('Mes_Ano')[col_metrica].sum().reset_index()
                    if not resumo_grafico.empty:
                        resumo_grafico['Rotulo'] = resumo_grafico[col_metrica].apply(lambda x: f"{prefixo_unidade}{x:{formato_num}}".replace(",", "X").replace(".", ",").replace("X", ".") + sufixo_unidade)
                        fig = px.bar(resumo_grafico, x='Mes_Ano', y=col_metrica, text='Rotulo')
                        fig.update_traces(textposition='outside', marker_color="#004A99", cliponaxis=False)
                        fig.update_layout(xaxis_title=None, yaxis_title=None, plot_bgcolor='rgba(0,0,0,0)')
                        st.plotly_chart(fig, use_container_width=True)

                with col_dir:
                    st.markdown("#### 🧊 Giro de Equipamentos")
                    if not df_equip.empty:
                        df_equip_agrupado = df_equip.groupby('tipo_equipamento')['quantidade'].sum().reset_index()
                        num_meses = (df_filtrado['Mes_Ano'].nunique()) or 1
                        encontrou = False
                        
                        for _, row in df_equip_agrupado.iterrows():
                            tipo_eq, qtd = row['tipo_equipamento'], int(row['quantidade'])
                            if qtd <= 0: continue
                            
                            if 'VISA' in tipo_eq: meta, fat_realizado, titulo, encontrou = 1200 * qtd * num_meses, df_filtrado[df_filtrado['equipamento'].astype(str).str.upper().str.contains('VISA', na=False)]['faturamento_reais'].sum(), f"🥤 {qtd}x VISA (NAB)", True
                            elif 'SOPI' in tipo_eq: meta, fat_realizado, titulo, encontrou = 2000 * qtd * num_meses, df_filtrado[df_filtrado['equipamento'].astype(str).str.upper().str.contains('SOPI', na=False)]['faturamento_reais'].sum(), f"🍺 {qtd}x SOPI (Cerveja)", True
                            elif 'CHOP' in tipo_eq: meta, fat_realizado, titulo, encontrou = 3870 * qtd * num_meses, df_filtrado[df_filtrado['codigo_produto'].astype(str).isin(['838', '8037'])]['faturamento_reais'].sum(), f"🍻 {qtd}x CHOPEIRA (Chopp)", True
                            else: continue
                                
                            st.write(f"**{titulo}**")
                            st.caption(f"Meta: R$ {meta:,.0f}".replace(",", "."))
                            st.progress(min(fat_realizado / meta, 1.0) if meta > 0 else 0)
                            realizado_str = f"R$ {fat_realizado:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                            if fat_realizado >= meta: st.success(f"Realizado: **{realizado_str}**")
                            else: st.error(f"Faltam R$ {(meta - fat_realizado):,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                            st.write("---")
                    else: st.info("Sem equipamentos na base.")

                st.subheader("📦 Análise por Produto")
                prod_disp = sorted(df_filtrado['nome_produto'].unique())
                prod_sel = st.multiselect("Selecione produtos:", options=prod_disp)
                if prod_sel:
                    df_mix = df_filtrado[df_filtrado['nome_produto'].isin(prod_sel)].groupby(['Mes_Ano', 'nome_produto'])['volume_caixas'].sum().reset_index()
                    if not df_mix.empty:
                        df_mix['Rotulo'] = df_mix['volume_caixas'].apply(lambda x: f"{x:,.0f} cx".replace(',', '.'))
                        fig_mix = px.bar(df_mix, x='Mes_Ano', y='volume_caixas', color='nome_produto', text='Rotulo', barmode='group')
                        fig_mix.update_traces(textposition='outside', cliponaxis=False)
                        fig_mix.update_layout(xaxis_title=None, yaxis_title=None, plot_bgcolor='rgba(0,0,0,0)')
                        st.plotly_chart(fig_mix, use_container_width=True)

            with tab_tarefas:
                if not df_tarefas.empty:
                    col_pres = [c for c in colunas_ordem_tarefas if c in df_tarefas.columns]
                    df_t_limpo = df_tarefas[col_pres]
                    c_xls, c_pdf = st.columns([1, 1])
                    c_xls.download_button("📥 Baixar Excel", gerar_excel_formatado(df_t_limpo), f"Tarefas_{codigo_input}.xlsx")
                    c_pdf.download_button("📄 Baixar PDF", gerar_pdf_formatado(df_t_limpo), f"Tarefas_{codigo_input}.pdf")
                    st.dataframe(df_t_limpo, use_container_width=True, hide_index=True)
                else: st.info("Nenhuma tarefa mapeada.")

# ==================== PÁGINA 3 ====================
elif menu == "📋 Planificador de Tarefas":
    st.title("📋 Planificador Global de Tarefas")
    
    with st.spinner("A carregar base global..."):
        df_todas_tarefas = carregar_dados_tarefas()
        df_todas_tarefas = limpar_colunas_tarefas(df_todas_tarefas)

    if not df_todas_tarefas.empty:
        def get_col(df, nomes):
            cmap = {str(c).lower().strip(): c for c in df.columns}
            return next((cmap[n] for n in nomes if n in cmap), None)
            
        c_cli, c_cat, c_clust, c_setor, c_data = get_col(df_todas_tarefas, ['codigo_cliente']), get_col(df_todas_tarefas, ['categoria']), get_col(df_todas_tarefas, ['cluster primário', 'cluster_primario']), get_col(df_todas_tarefas, ['setor']), get_col(df_todas_tarefas, ['data visita', 'data_visita'])

        with st.expander("🔍 Filtros", expanded=True):
            co1, co2, co3, co4, co5 = st.columns(5)
            f_cli = co1.multiselect("Código", sorted(df_todas_tarefas[c_cli].dropna().astype(str).unique()) if c_cli else [])
            f_cat = co2.multiselect("Categoria", sorted(df_todas_tarefas[c_cat].dropna().astype(str).unique()) if c_cat else [])
            f_clust = co3.multiselect("Cluster", sorted(df_todas_tarefas[c_clust].dropna().astype(str).unique()) if c_clust else [])
            f_setor = co4.multiselect("Setor", sorted(df_todas_tarefas[c_setor].dropna().astype(str).unique()) if c_setor else [])
            f_data = co5.multiselect("Data", sorted(df_todas_tarefas[c_data].dropna().astype(str).unique()) if c_data else [])

        df_filtro = df_todas_tarefas.copy()
        if f_cli: df_filtro = df_filtro[df_filtro[c_cli].astype(str).isin(f_cli)]
        if f_cat: df_filtro = df_filtro[df_filtro[c_cat].astype(str).isin(f_cat)]
        if f_clust: df_filtro = df_filtro[df_filtro[c_clust].astype(str).isin(f_clust)]
        if f_setor: df_filtro = df_filtro[df_filtro[c_setor].astype(str).isin(f_setor)]
        if f_data: df_filtro = df_filtro[df_filtro[c_data].astype(str).isin(f_data)]
        
        df_exibicao = df_filtro[[c for c in colunas_ordem_tarefas if c in df_filtro.columns]]
        st.markdown(f"**Total:** `{len(df_exibicao)} linhas`")
        cx, cp = st.columns(2)
        cx.download_button("📥 Excel", gerar_excel_formatado(df_exibicao), "Planificador.xlsx")
        cp.download_button("📄 PDF", gerar_pdf_formatado(df_exibicao), "Planificador.pdf")
        st.dataframe(df_exibicao, use_container_width=True, hide_index=True)
    else:
        st.warning("Gere o ficheiro tarefas.parquet usando o script local.")
