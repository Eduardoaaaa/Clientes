import streamlit as st
import pandas as pd
from sqlalchemy import create_engine
import plotly.express as px

# 1. Configuração da página
st.set_page_config(page_title="Painel de Clientes", layout="wide", page_icon="📊")

# 2. Conexão com o Banco de Dados (Mantenha o seu link Pooler aqui)
SUPABASE_DB_URL = "postgresql://postgres.zxnbyzahqmuncobkntzs:%40EduSantos157@aws-1-us-west-2.pooler.supabase.com:6543/postgres"

@st.cache_data(ttl=300)
def buscar_dados(codigo):
    try:
        engine = create_engine(SUPABASE_DB_URL)
        
        # Busca o histórico de Vendas
        query_vendas = f"SELECT * FROM vendas_consolidadas WHERE CAST(codigo_cliente AS TEXT) = '{codigo}'"
        df_vendas = pd.read_sql(query_vendas, engine)
        
        # Busca os Equipamentos do cliente específico
        query_equip = f"SELECT * FROM equipamentos_clientes WHERE TRIM(CAST(codigo_cliente AS TEXT)) = '{codigo.strip()}'"
        df_equip = pd.read_sql(query_equip, engine)
        
        return df_vendas, df_equip
    except Exception as e:
        st.error(f"Erro ao ligar à base de dados: {e}")
        return pd.DataFrame(), pd.DataFrame()

# --- TÍTULO E BUSCA ---
st.title("📊 Portal de Autoatendimento - Gerência")

# Campo de busca centralizado ou no topo
codigo_input = st.text_input("🔎 Digite o Código do Cliente para iniciar:", placeholder="Ex: 9528")

if codigo_input:
    # Agora recebemos as vendas E os equipamentos
    df_cliente, df_equip = buscar_dados(codigo_input)

    if df_cliente.empty:
        st.warning(f"⚠️ Nenhum histórico encontrado para o código {codigo_input}. Verifique se o código está correto.")
    else:
        # Prepara os dados
        df_cliente['data_venda'] = pd.to_datetime(df_cliente['data_venda'])
        df_cliente['Mes_Ano'] = df_cliente['data_venda'].dt.to_period('M').astype(str)
        nome_cliente = df_cliente['nome_cliente'].iloc[0]
        
        # --- CABEÇALHO DO CLIENTE ---
        st.subheader(f"👤 Cliente: {codigo_input} - {nome_cliente}")
        
        # --- NOVO FILTRO DE PERÍODO (PÁGINA PRINCIPAL / TOPO) ---
        st.write("---")
        st.markdown("### 📅 Selecione o Período de Análise")
        
        meses_disponiveis = sorted(df_cliente['Mes_Ano'].unique())
        
        if len(meses_disponiveis) > 1:
            # O slider agora ocupa a largura total da página
            mes_inicio, mes_fim = st.select_slider(
                "Arraste as extremidades para definir o intervalo de tempo:",
                options=meses_disponiveis,
                value=(meses_disponiveis[0], meses_disponiveis[-1]),
                label_visibility="collapsed" # Esconde o texto acima para ficar mais limpo
            )
            # Filtro lógico
            df_filtrado = df_cliente[(df_cliente['Mes_Ano'] >= mes_inicio) & (df_cliente['Mes_Ano'] <= mes_fim)]
        else:
            st.info(f"Este cliente possui apenas um mês de histórico: {meses_disponiveis[0]}")
            df_filtrado = df_cliente

        # --- SEÇÃO DE INDICADORES (KPIs) ---
        st.write("") # Espaçamento
        total_vol = df_filtrado['volume_caixas'].sum()
        total_fat = df_filtrado['faturamento_reais'].sum()
        dias_compra = df_filtrado['data_venda'].nunique()

        c1, c2, c3 = st.columns(3)
        with c1:
            st.metric("Faturamento Total", f"R$ {total_fat:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        with c2:
            st.metric("Volume Total", f"{total_vol:,.0f} cx".replace(",", "."))
        with c3:
            st.metric("Frequência de Pedidos", f"{dias_compra} dias")

        # --- SEÇÃO VISUAL: GRÁFICOS E GIRO ---
        st.write("---")
        col_esq, col_dir = st.columns([2, 1])

        with col_esq:
            st.markdown("#### 📈 Evolução Mensal")
            resumo_grafico = df_filtrado.groupby('Mes_Ano').agg({'faturamento_reais': 'sum'}).reset_index()
            
            if not resumo_grafico.empty:
                # 0. Garante a ordenação cronológica para evitar datas fora de ordem
                resumo_grafico = resumo_grafico.sort_values('Mes_Ano')

                # 1. Cria o rótulo formatado (R$ 1.234,56)
                resumo_grafico['Rotulo_Valor'] = resumo_grafico['faturamento_reais'].apply(
                    lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                )

                # 2. Monta o gráfico com Plotly
                fig = px.bar(
                    resumo_grafico,
                    x='Mes_Ano',
                    y='faturamento_reais',
                    text='Rotulo_Valor'
                )

                # 3. Ajuste do Rótulo (Aumentado para 16 e garantindo que não corte no topo)
                fig.update_traces(
                    textposition='outside', 
                    marker_color="#004A99",
                    textfont_size=16,      # Tamanho do rótulo aumentado
                    textfont_color="white",
                    cliponaxis=False       # Impede que o valor suma se a barra for alta
                )

                # 4. Ajuste do Layout e Correção do Eixo X (Datas)
                fig.update_layout(
                    xaxis_title=None,
                    yaxis_title=None,
                    plot_bgcolor='rgba(0,0,0,0)',
                    paper_bgcolor='rgba(0,0,0,0)',
                    margin=dict(t=50, b=0, l=0, r=0), # Espaço extra no topo (t=50) para o rótulo grande
                    xaxis=dict(
                        type='category', # Trata cada mês como categoria para não pular datas
                        tickangle=0      # Mantém os meses na horizontal para facilitar a leitura
                    )
                )

                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Selecione um período na barra superior para visualizar o gráfico.")

        with col_dir:
            st.markdown("#### 🧊 Giro de Equipamentos")
            
            # Se o cliente TIVER equipamento registrado
            if not df_equip.empty:
                # Agrupa caso o sistema envie o mesmo equipamento duplicado em linhas
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
                        # FILTRO EXCLUSIVO DA CHOPEIRA: Considera apenas os produtos 838 e 8037
                        codigos_chopp = ['838', '8037']
                        fat_realizado = df_filtrado[df_filtrado['codigo_produto'].astype(str).isin(codigos_chopp)]['faturamento_reais'].sum()
                        titulo = f"🍻 {qtd}x CHOPEIRA (Chopp)"
                        
                    else:
                        continue # Ignora se tiver algum equipamento estranho
                        
                    st.write(f"**{titulo}**")
                    st.caption(f"Meta período: R$ {meta:,.0f}".replace(",", "."))
                    
                    # Barra de progresso
                    pct = min(fat_realizado / meta, 1.0) if meta > 0 else 0
                    st.progress(pct)
                    
                    # Formatação do valor realizado para o padrão R$
                    realizado_str = f"R$ {fat_realizado:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                    
                    # Lógica de Meta Atingida vs GAP
                    if fat_realizado >= meta:
                        st.markdown(f"Realizado: **{realizado_str}**")
                        st.success("✅ **Meta Atingida!** Equipamento justificado.")
                    else:
                        gap = meta - fat_realizado
                        gap_str = f"R$ {gap:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                        
                        st.markdown(f"Realizado: **{realizado_str}**")
                        st.error(f"⚠️ **GAP: Faltam {gap_str}**")
                        
                    st.write("---") # Linha separadora
            else:
                # Se a tabela do cliente não trouxer equipamentos
                st.info("✅ Cliente sem equipamentos comodatados na base.")

        # --- SEÇÃO 3: ANÁLISE DE PRODUTOS ESPECÍFICOS ---
        st.divider()
        st.subheader("📦 Análise de Volume por Produto")

        # 1. Filtro de busca por produto específico
        produtos_disponiveis = sorted(df_filtrado['nome_produto'].unique())
        produtos_selecionados = st.multiselect(
            "Selecione um ou mais produtos para analisar a evolução mensal:",
            options=produtos_disponiveis,
            help="O gráfico será exibido após a seleção."
        )

        # 2. Lógica de exibição (Só mostra se algo for selecionado)
        if produtos_selecionados:
            df_mix = df_filtrado[df_filtrado['nome_produto'].isin(produtos_selecionados)]
            
            # Agrupamento mensal do volume (em caixas)
            resumo_mix = df_mix.groupby(['Mes_Ano', 'nome_produto']).agg({'volume_caixas': 'sum'}).reset_index()
            resumo_mix = resumo_mix.sort_values('Mes_Ano')

            if not resumo_mix.empty:
                resumo_mix['volume_caixas'] = resumo_mix['volume_caixas'].round(0)
                
                # Cria o rótulo com a quantidade de caixas (Ex: 15 cx)
                resumo_mix['Rotulo'] = resumo_mix['volume_caixas'].apply(lambda x: f"{x:,.0f} cx".replace(',', '.'))

                # 3. Criação do Gráfico Normal (Plotly)
                fig_mix = px.bar(
                    resumo_mix,
                    x='Mes_Ano',
                    y='volume_caixas',
                    color='nome_produto', # Se escolher 2 produtos, cria barras lado a lado
                    text='Rotulo',        # Coloca o valor da caixa na barra
                    labels={'volume_caixas': 'Volume (Cx)', 'Mes_Ano': 'Mês', 'nome_produto': 'Produto'}
                )

                # Descobre o maior valor para esticar o teto do gráfico em 20% e não cortar o rótulo
                max_vol = resumo_mix['volume_caixas'].max()

                # Ajustes visuais (rótulo para fora, gráfico não empilhado)
                fig_mix.update_traces(
                    texttemplate='%{text}',  # <-- ADICIONE ESTA LINHA AQUI!
                    textposition='outside',
                    textfont_size=16,
                    textfont_color="white",
                    cliponaxis=False
                )

                fig_mix.update_layout(
                    title="Evolução de Volume (Caixas)",
                    barmode='group', # <-- ISSO DEIXA O GRÁFICO NORMAL (LADO A LADO)
                    xaxis=dict(type='category', tickangle=0),
                    yaxis=dict(
                        range=[0, max_vol * 1.2], # Espaço extra no topo
                        showticklabels=False,     # Esconde números laterais
                        showgrid=False            # Tira as linhas de fundo
                    ),
                    xaxis_title=None,
                    yaxis_title=None,
                    plot_bgcolor='rgba(0,0,0,0)',
                    paper_bgcolor='rgba(0,0,0,0)',
                    legend=dict(
                        title="", 
                        orientation="h", 
                        yanchor="bottom", 
                        y=-0.3, 
                        xanchor="center", 
                        x=0.5
                    ),
                    margin=dict(t=40, b=80, l=0, r=0),
                    hovermode="x unified"
                )
                
                st.plotly_chart(fig_mix, use_container_width=True, config={'displayModeBar': False})
                
                # 4. Tabela de Detalhamento
                with st.expander("Ver detalhamento em tabela"):
                    tabela_mix = df_mix.groupby(['codigo_produto', 'nome_produto', 'categoria_produto']).agg({
                        'volume_caixas': 'sum'
                    }).reset_index().sort_values('volume_caixas', ascending=False)
                    
                    tabela_mix['volume_caixas'] = tabela_mix['volume_caixas'].map('{:,.0f} cx'.format).str.replace(',', '.')
                    st.dataframe(tabela_mix, use_container_width=True, hide_index=True)
        else:
            # Mensagem elegante caso a gerente ainda não tenha escolhido nada
            st.info("👆 Selecione um ou mais produtos no filtro acima para gerar o gráfico de evolução.")

else:
    # Tela inicial amigável
    st.info("👋 Bem-vindo! Utilize o campo acima para buscar as informações de um cliente pelo código.")
    st.image("https://cdn-icons-png.flaticon.com/512/2666/2666505.png", width=150)