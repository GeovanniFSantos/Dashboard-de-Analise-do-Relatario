# app.py

# Importa as bibliotecas necessárias
import pandas as pd
import streamlit as st
import plotly.express as px
import numpy as np # Importado para usar a função select na classificação de categorias

# --- Configuração Inicial ---
# O título foi ajustado para refletir o nome do relatório
st.set_page_config(layout="wide")
st.title("📊 Dashboard de Análise do Relatório Gabriel Pro")

# ==============================================================================
# 📌 PASSO 1: DEFINIÇÃO DO ARQUIVO
# Nome do seu arquivo Excel, que deve estar na mesma pasta do app.py.
# ==============================================================================
Relatorio = 'Relatorio.xlsx' 
# Definimos as colunas numéricas que precisam ser tratadas
COLUNAS_NUMERICAS = ['Valor Total', 'Pontos']
# Coluna de temporada usada para filtrar (a numérica é mais confiável)
COLUNA_NUMERO_TEMPORADA = 'Numero Temporada' 
# KPI de volume: Usaremos 'NF/Pedido' para contar o número de pedidos únicos
COLUNA_PEDIDO = 'NF/Pedido' 
# KPI de volume: Usaremos 'CPF/CNPJ' para contar pessoas únicas na aba principal
COLUNA_CNPJ_CPF = 'CPF/CNPJ' 
# Coluna de identificação do profissional
COLUNA_ESPECIFICADOR = 'Especificador/Empresa'
# NOVA CONSTANTE: Coluna de CPF na aba "Novos Cadastrados"
COLUNA_CPF_NOVO_CADASTRO = 'CPF'
# [ARQUIVO_NOVOS_CADASTRADOS removido]


# Função para carregar os dados (usa cache do streamlit para ser mais rápido)
@st.cache_data
def carregar_e_tratar_dados(caminho_arquivo):
    """Lê o arquivo Excel (2 abas), trata colunas e retorna um DataFrame do Pandas."""
    try:
        # LER A ABA PRINCIPAL (Relatório)
        df = pd.read_excel(caminho_arquivo, sheet_name=0) 
        
        # LER A ABA DE NOVOS CADASTRADOS (Assumindo a aba se chame "Novos Cadastrados")
        try:
            # Lê a aba de Novos Cadastrados do MESMO arquivo
            df_novos = pd.read_excel(caminho_arquivo, sheet_name='Novos Cadastrados')
        except ValueError:
            st.error(f"❌ Erro: A aba 'Novos Cadastrados' não foi encontrada no arquivo '{caminho_arquivo}'.")
            df_novos = pd.DataFrame()
        except FileNotFoundError:
            st.error(f"❌ Erro: O arquivo '{caminho_arquivo}' não foi encontrado.")
            df_novos = pd.DataFrame()

        # === ETAPA DE TRATAMENTO DE DADOS (DF PRINCIPAL) ===
        # 1. Tratamento de Colunas Numéricas (removendo R$ e Símbolos)
        for col in COLUNAS_NUMERICAS:
            if col in df.columns:
                # Remove espaços, vírgulas (usadas como separador de milhar) e R$
                # E converte para numérico
                df[col] = df[col].astype(str).str.replace(r'[^0-9,.]', '', regex=True)
                df[col] = df[col].str.replace(',', '.', regex=False)
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
                
        # 2. Garantir que 'Data da Venda' seja datetime e filtrar dados inválidos
        if 'Data da Venda' in df.columns:
            df['Data da Venda'] = pd.to_datetime(df['Data da Venda'], errors='coerce')
            
            # Remove linhas com 'Data da Venda' nula ou inválida
            df.dropna(subset=['Data da Venda'], inplace=True) 
            
            # 3. CRIAÇÃO DAS COLUNAS DE MÊS E ANO PARA FILTRAGEM
            df['Ano'] = df['Data da Venda'].dt.year.astype(str)
            df['Mês_num'] = df['Data da Venda'].dt.month.astype(str)
            
            # 4. CRIAÇÃO DA COLUNA DE TEMPORADA DE EXIBIÇÃO
            if COLUNA_NUMERO_TEMPORADA in df.columns:
                # Garante que a coluna de temporada é numérica e a converte para 'Temporada X'
                df[COLUNA_NUMERO_TEMPORADA] = pd.to_numeric(df[COLUNA_NUMERO_TEMPORADA], errors='coerce').fillna(0).astype(int)
                df['Temporada_Exibicao'] = 'Temporada ' + df[COLUNA_NUMERO_TEMPORADA].astype(str)
            
            # 5. Mapeamento e Formatação para o Filtro de Mês 
            nomes_meses_map = {
                '1': 'Jan (01)', '2': 'Fev (02)', '3': 'Mar (03)', '4': 'Abr (04)',
                '5': 'Mai (05)', '6': 'Jun (06)', '7': 'Jul (07)', '8': 'Ago (08)',
                '9': 'Set (09)', '10': 'Out (10)', '11': 'Nov (11)', '12': 'Dez (12)'
            }
            # Aqui, mapeamos e garantimos que apenas os meses válidos sejam exibidos.
            df['Mês_Exibicao'] = df['Mês_num'].map(nomes_meses_map)
            
            # === 6. LÓGICA DE NOVO CADASTRADO (CRÍTICO) ===
            if COLUNA_CNPJ_CPF in df.columns and COLUNA_NUMERO_TEMPORADA in df.columns:
                
                # CRÍTICO: Limpar colunas para merge
                df['CNPJ_CPF_LIMPO'] = df[COLUNA_CNPJ_CPF].astype(str).str.replace(r'[^0-9]', '', regex=True)
                
                if COLUNA_CPF_NOVO_CADASTRO in df_novos.columns:
                    df_novos['CPF_LIMPO'] = df_novos[COLUNA_CPF_NOVO_CADASTRO].astype(str).str.replace(r'[^0-9]', '', regex=True)
                
                    # 6.1. Identifica QUEM está na lista de Novos Cadastrados
                    df['Novo_Cadastro_Existe'] = df['CNPJ_CPF_LIMPO'].isin(df_novos['CPF_LIMPO'].unique())
                    
                    # 6.2. Calculamos a data da primeira compra histórica para clientes da aba principal
                    df_primeira_compra = df.groupby('CNPJ_CPF_LIMPO')['Data da Venda'].min().reset_index()
                    df_primeira_compra.columns = ['CNPJ_CPF_LIMPO', 'Data_Primeira_Compra_Historica']
                    df = pd.merge(df, df_primeira_compra, on='CNPJ_CPF_LIMPO', how='left')
                    
                    # 6.3. Marca APENAS a linha que corresponde à primeira compra histórica
                    df['Novo_Cadastrado'] = np.where(
                        (df['Novo_Cadastro_Existe'] == True) & 
                        (df['Data da Venda'] == df['Data_Primeira_Compra_Historica']), # CRÍTICO: Usa a data da venda, não a da temporada.
                        True,
                        False
                    )
                else:
                    df['Novo_Cadastrado'] = False 
        
        return df, df_novos # Retorna os dois DataFrames
    
    except FileNotFoundError:
        st.error(f"❌ Erro: Arquivo '{caminho_arquivo}' não encontrado.")
        return pd.DataFrame(), pd.DataFrame() 
    except Exception as e:
        st.error(f"Ocorreu um erro ao ler ou tratar o arquivo: {e}")
        return pd.DataFrame(), pd.DataFrame() 

# Carrega e trata os dados
df_dados_original, df_novos_cadastrados_original = carregar_e_tratar_dados(Relatorio)


# --- Aplicação Streamlit (Interface) ---

if not df_dados_original.empty:
    
    # Cria uma cópia inicial que será filtrada por Data
    df_dados_por_data = df_dados_original.copy()

    # === BARRA LATERAL (FILTROS DE DATA E TEMPORADA) ===
    st.sidebar.header("Filtros Interativos")
    
    # CRÍTICO: Inicialização de todas as variáveis de seleção no bloco principal.
    temporadas_selecionadas_exib = []
    meses_selecionados_exib = [] 
    lojas_selecionadas = []
    segmentos_selecionados = []
    
    # 1. Filtro por Temporada (Definido e Aplicado no mesmo bloco)
    if 'Temporada_Exibicao' in df_dados_original.columns:
        # Remove 'Temporada 0' e valores nulos antes de popular o filtro
        temporadas_unicas_exib = sorted(df_dados_original['Temporada_Exibicao'].loc[df_dados_original['Temporada_Exibicao'] != 'Temporada 0'].dropna().unique())
        temporadas_selecionadas_exib = st.sidebar.multiselect(
            "Selecione a Temporada:",
            options=temporadas_unicas_exib,
            default=temporadas_unicas_exib
        )
        
        # Aplicação do Filtro de Temporada
        if temporadas_selecionadas_exib:
            df_dados_por_data = df_dados_por_data[df_dados_por_data['Temporada_Exibicao'].isin(temporadas_selecionadas_exib)]

    # 2. Filtro por Mês (Estrutura Reforçada)
    # CRÍTICO: Reestrutura a lógica para eliminar o erro do Pylance
    if 'Mês_Exibicao' in df_dados_por_data.columns:
        # AQUI usamos dropna() para remover meses que não foram mapeados (os 'esquisitos')
        meses_unicos_exib = sorted(df_dados_por_data['Mês_Exibicao'].dropna().unique())
        
        # Define a variável 'meses_selecionadas_exib'
        meses_selecionados_exib = st.sidebar.multiselect(
            "Selecione o Mês:",
            options=meses_unicos_exib,
            default=meses_unicos_exib
        )
        
        # Aplicação do Filtro de Mês (AGORA DENTRO DO BLOCO IF)
        if meses_selecionados_exib:
             df_dados_por_data = df_dados_por_data[df_dados_por_data['Mês_Exibicao'].isin(meses_selecionados_exib)]


    # FIM DA FILTRAGEM DE DATA. df_dados_por_data contem apenas os dados do período selecionado.
    
    
    # === FILTROS HIERÁRQUICOS (LOJA > SEGMENTO) ===
    st.sidebar.subheader("Filtros de Entidade")

    # 3. Filtro LOJA (Primeiro Nível)
    lojas_unicas = sorted(df_dados_por_data['Loja'].unique())
    # --------------------------------------------------------------------------
    # ALTERAÇÃO CRÍTICA: FILTRO PADRÃO PARA "Bontempo"
    # --------------------------------------------------------------------------
    default_loja = ['Bontempo'] if 'Bontempo' in lojas_unicas else lojas_unicas
    
    lojas_selecionadas = st.sidebar.multiselect(
        "Selecione a Loja:",
        options=lojas_unicas,
        default=default_loja # Alterado de lojas_unicas para default_loja
    )
    # --------------------------------------------------------------------------

    # DataFrame APÓS filtro de LOJA (mas ainda dentro do período)
    df_apos_loja = df_dados_por_data[df_dados_por_data['Loja'].isin(lojas_selecionadas)]

    # 4. Filtro SEGMENTO (Segundo Nível - Hierárquico)
    # Mostra APENAS os segmentos que estão nas lojas selecionadas
    segmentos_unicos_loja = sorted(df_apos_loja['Segmento'].unique())
    segmentos_selecionados = st.sidebar.multiselect(
        "Selecione o Segmento (Aparece após a Loja):",
        options=segmentos_unicos_loja,
        default=segmentos_unicos_loja
    )

    # === CRIAÇÃO DO DATAFRAME FINAL FILTRADO ===
    # O df_filtrado contem todos os filtros aplicados (Data + Loja + Segmento)
    df_filtrado = df_apos_loja[df_apos_loja['Segmento'].isin(segmentos_selecionados)].copy()
    
    # O df_total_periodo contem apenas os filtros de data (usado para o gráfico 'Total')
    df_total_periodo = df_dados_por_data.copy()
    
    # === MÉTRICAS (KPIs) ===
    st.subheader("Métricas Chave (KPIs)")
    
    # 1. Total de Pontos
    total_pontos = df_filtrado['Pontos'].sum()
    
    # 2. Total de Pedidos Únicos
    if COLUNA_PEDIDO in df_filtrado.columns:
        total_pedidos_filtrado = df_filtrado[COLUNA_PEDIDO].nunique()
    else:
        total_pedidos_filtrado = 0

    # 3. Total de Pessoas Pontuadas
    if COLUNA_CNPJ_CPF in df_filtrado.columns:
        # A contagem de pessoas únicas usa a coluna limpa
        total_pessoas_filtrado = df_filtrado['CNPJ_CPF_LIMPO'].nunique()
    else:
        total_pessoas_filtrado = 0
    
    # 4. NOVO KPI: Valor Médio por Pedido (Pontos / Pedidos)
    if total_pedidos_filtrado > 0:
        valor_medio_por_pedido = total_pontos / total_pedidos_filtrado
    else:
        valor_medio_por_pedido = 0

    
    # NOVO LAYOUT: 4 colunas
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric(
            label="Soma Total de Pontos (Filtrado)", 
            value=f"{total_pontos:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")
        )
        
    with col2:
        st.metric(
            label="Total de Pedidos Lançados (Filtrado)", 
            value=f"{total_pedidos_filtrado:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")
        )

    with col3:
        st.metric(
            label="Total de Pessoas Pontuadas (Filtrado)", 
            value=f"{total_pessoas_filtrado:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")
        )

    with col4:
        st.metric(
            label="Valor Médio por Pedido (Pontos)", 
            value=f"{valor_medio_por_pedido:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")
        )
        

    # === GRÁFICOS E ANÁLISES (VISUALIZAÇÕES) ===
    
    # =======================================================================
    # ITEM 1: TABELA PIVÔ DE PONTUAÇÃO POR MÊS E TEMPORADA
    # =======================================================================
    st.subheader("1. Evolução da Pontuação por Mês e Temporada (Filtrado por Loja/Segmento)")
    
    if 'Mês_Exibicao' in df_filtrado.columns and 'Temporada_Exibicao' in df_filtrado.columns:
        
        # 1. Agrupamento e Soma (Pivô)
        # Usamos df_filtrado para que os filtros de Loja e Segmento sejam respeitados
        df_pivot = df_filtrado.pivot_table(
            index='Mês_Exibicao', # Linhas (Mês)
            columns='Temporada_Exibicao', # Colunas (Temporada)
            values='Pontos', # Valores a serem somados
            aggfunc='sum',
            fill_value=0 # Preenche NaNs com 0 para clareza
        ).reset_index()

        # 2. Reordenação dos Meses (Julho a Junho, seguindo o ano fiscal)
        mes_ordem_fiscal = {
            'Jul (07)': 1, 'Ago (08)': 2, 'Set (09)': 3, 'Out (10)': 4, 'Nov (11)': 5, 
            'Dez (12)': 6, 'Jan (01)': 7, 'Fev (02)': 8, 'Mar (03)': 9, 'Abr (04)': 10,
            'Mai (05)': 11, 'Jun (06)': 12
        }
        df_pivot['Ordem'] = df_pivot['Mês_Exibicao'].map(mes_ordem_fiscal)
        df_pivot.sort_values(by='Ordem', inplace=True)
        df_pivot.drop('Ordem', axis=1, inplace=True)
        
        # 3. Tratamento de Colunas
        colunas_temporada = [col for col in df_pivot.columns if col.startswith('Temporada')]
        
        # Ordenação das colunas de Temporada (T7, T8, T9, T10...)
        # CRÍTICO: Ordena as colunas de T7, T8, T9, T10...
        df_pivot = df_pivot[['Mês_Exibicao'] + sorted(colunas_temporada)]
        
        # 4. Adicionar a Linha de Total (Soma por Coluna de Temporada)
        total_row = pd.Series(df_pivot[colunas_temporada].sum(), name='Total')
        total_row['Mês_Exibicao'] = 'Total' # Define o nome da linha de total
        
        # 5. Concatena a linha de Total e Estilização
        df_pivot.set_index('Mês_Exibicao', inplace=True)
        
        # Adiciona a Linha de Total na parte inferior
        df_pivot = pd.concat([df_pivot, pd.DataFrame(total_row).T.set_index('Mês_Exibicao')])
        df_pivot.index.name = 'Mês'

        # Estilização e Exibição no Streamlit
        st.dataframe(
            df_pivot.style.format({col: lambda x: f"{x:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".") 
                                    for col in colunas_temporada + ['Total']})
                                    # CENTRALIZAÇÃO APLICADA APENAS NAS COLUNAS DE DADOS
                                    .set_properties(**{'border': '1px solid #333333', 'text-align': 'center'}, subset=pd.IndexSlice[:, colunas_temporada]),
            use_container_width=True
        )

    # 2. Gráfico de Barras (Pontos por Segmento - FILTRADO) - Item antigo 1
    st.subheader("2. Análise de Distribuição (Pontos por Segmento - Loja/Segmento Filtrado)")
    
    # Agrupa por Segmento e SOMA os Pontos no DF FILTRADO
    df_segmento_pontos = df_filtrado.groupby('Segmento')['Pontos'].sum().reset_index()
    df_segmento_pontos.columns = ['Segmento', 'Pontos_Somados']

    fig_segmento = px.bar(
        df_segmento_pontos, 
        x='Segmento', 
        y='Pontos_Somados', 
        title='Pontos Totais dos Itens Selecionados (Loja e Segmento)',
        color='Segmento'
    )
    fig_segmento.update_layout(xaxis_title="Segmento", yaxis_title="Total de Pontos (Filtrado)")
    st.plotly_chart(fig_segmento, use_container_width=True)

    # 3. Análise de Distribuição Total (Segmento - Pontos E Pedidos no Período Selecionado) - Item antigo 2
    st.subheader("3. Análise de Distribuição Total (Segmento - Pontos e Pedidos no Período Selecionado)")
    
    # === NOVO SELETOR DE MÉTRICA ===
    metrica_selecionada = st.selectbox(
        'Selecione a Métrica para Distribuição Total:',
        ('Pontos Totais', 'Pedidos Únicos')
    )
    
    # === CÁLCULO DINÂMICO DA MÉTRICA ===
    if metrica_selecionada == 'Pontos Totais':
        # Cálculo para Pontos
        df_segmento_metrica = df_total_periodo.groupby('Segmento')['Pontos'].sum().reset_index()
        df_segmento_metrica.columns = ['Segmento', 'Metrica_Somada']
        eixo_y_titulo = "Total de Pontos"
        titulo_grafico = 'Pontos Totais por Segmento (Geral do Período)'
    else:
        # Cálculo para Pedidos
        df_segmento_metrica = df_total_periodo.groupby('Segmento')[COLUNA_PEDIDO].nunique().reset_index()
        df_segmento_metrica.columns = ['Segmento', 'Metrica_Somada']
        eixo_y_titulo = "Total de Pedidos"
        titulo_grafico = 'Pedidos Únicos por Segmento (Geral do Período)'


    # === GRÁFICO ÚNICO (COMPARAÇÃO POR ALTERNÂNCIA) ===
    fig_segmento_total = px.bar(
        df_segmento_metrica, 
        x='Segmento', 
        y='Metrica_Somada', 
        title=titulo_grafico,
        color='Segmento'
    )
    fig_segmento_total.update_layout(xaxis_title="Segmento", yaxis_title=eixo_y_titulo)
    st.plotly_chart(fig_segmento_total, use_container_width=True)
    
    # FIM DA ANÁLISE 3 (Antiga 2)

    # 4. Análise de Tendência ao Longo do Tempo (Pontos Totais) - Item antigo 3
    if 'Data da Venda' in df_filtrado.columns:
        st.subheader("4. Tendência de Pontos (Pontos Totais)")
        
        # Agrupa os dados por mês/ano e soma os Pontos
        df_tendencia = df_filtrado.set_index('Data da Venda').resample('M')['Pontos'].sum().reset_index()
        df_tendencia.columns = ['Data', 'Pontos Totais']
        
        fig_tendencia = px.line(
            df_tendencia,
            x='Data',
            y='Pontos Totais',
            title='Pontos Totais por Mês/Ano',
            markers=True
        )
        fig_tendencia.update_layout(yaxis_title="Pontos Totais")
        st.plotly_chart(fig_tendencia, use_container_width=True)

    # 5. NOVO GRÁFICO: Pedidos Únicos por Mês - Item antigo 4
    if 'Mês_Exibicao' in df_filtrado.columns and COLUNA_PEDIDO in df_filtrado.columns:
        st.subheader("5. Pedidos Únicos por Mês")
        
        # Agrupa o DataFrame FILTRADO pelo Mês de Exibição e conta os pedidos únicos
        df_pedidos_por_mes = df_filtrado.groupby('Mês_Exibicao')[COLUNA_PEDIDO].nunique().reset_index()
        df_pedidos_por_mes.columns = ['Mês', 'Pedidos']
        
        # Para garantir a ordem correta dos meses (Jul, Ago, Set, etc.), precisamos ordenar pelo Mês_num original
        # Primeiro, criamos uma coluna de ordenação temporária no df_pedidos_por_mes
        mes_para_ordenacao = {
            'Jan (01)': 7, 'Fev (02)': 8, 'Mar (03)': 9, 'Abr (04)': 10,
            'Mai (05)': 11, 'Jun (06)': 12, 'Jul (07)': 1, 'Ago (08)': 2,
            'Set (09)': 3, 'Out (10)': 4, 'Nov (11)': 5, 'Dez (12)': 6
        }
        df_pedidos_por_mes['Mês_Ordem'] = df_pedidos_por_mes['Mês'].map(mes_para_ordenacao)
        df_pedidos_por_mes.sort_values(by='Mês_Ordem', inplace=True)
        
        fig_pedidos_mes = px.bar(
            df_pedidos_por_mes,
            x='Mês',
            y='Pedidos',
            title='Contagem de Pedidos Únicos por Mês',
            color='Mês'
        )
        fig_pedidos_mes.update_layout(xaxis_title="Mês", yaxis_title="Pedidos Únicos")
        st.plotly_chart(fig_pedidos_mes, use_container_width=True)


    # 6. DESEMPENHO POR PROFISSIONAL E CATEGORIA - Item antigo 5
    if COLUNA_ESPECIFICADOR in df_filtrado.columns:
        st.subheader("6. Desempenho por Profissional e Categoria")
        
        # 1. Agrupamento por Profissional (Pontuação Total)
        df_desempenho = df_filtrado.groupby(COLUNA_ESPECIFICADOR)['Pontos'].sum().reset_index()
        df_desempenho.columns = [COLUNA_ESPECIFICADOR, 'Pontuacao_Total']
        
        # 2. Definição da Lógica de Categorias (Adaptada do seu código Power BI/DAX)
        condicoes = [
            (df_desempenho['Pontuacao_Total'] >= 5000000), # Diamante
            (df_desempenho['Pontuacao_Total'] >= 2000000), # Esmeralda
            (df_desempenho['Pontuacao_Total'] >= 500000),  # Ruby
            (df_desempenho['Pontuacao_Total'] >= 150000),  # Topázio
            (df_desempenho['Pontuacao_Total'] >= 1)         # Pro
        ]
        categorias = ['Diamante', 'Esmeralda', 'Ruby', 'Topázio', 'Pro']
        
        # Aplicar a lógica usando numpy.select (equivalente ao SWITCH)
        df_desempenho['Categoria'] = np.select(
            condicoes, 
            categorias, 
            default='Sem Categoria'
        )
        
        # Ordenar por Pontuação Total (do maior para o menor)
        df_desempenho.sort_values(by='Pontuacao_Total', ascending=False, inplace=True)
        
        # 3. MATRIZ DE RESUMO (KPIs de Categoria)
        st.markdown("##### Resumo das Categorias (Contagem e Pontuação)")
        
        # Agrupar e somar os resultados por Categoria
        df_resumo_cat = df_desempenho.groupby('Categoria').agg(
            Contagem=('Categoria', 'size'),
            Pontuacao_Categoria=('Pontuacao_Total', 'sum')
        ).reset_index()
        
        # Criar as colunas para o layout da Matriz
        colunas_matriz = ['Diamante', 'Esmeralda', 'Ruby', 'Topázio', 'Pro', 'Sem Categoria']
        
        colunas_kpi_contagem = st.columns(len(colunas_matriz))
        colunas_kpi_pontuacao = st.columns(len(colunas_matriz))
        
        # Loop para exibir a contagem por categoria
        st.markdown('###### Contagem de Profissionais:')
        for i, categoria in enumerate(colunas_matriz):
            contagem = df_resumo_cat.loc[df_resumo_cat['Categoria'] == categoria, 'Contagem'].sum()
            with colunas_kpi_contagem[i]:
                 st.metric(label=categoria, value=f"{contagem:,.0f}")
                 
        # Loop para exibir a pontuação total por categoria
        st.markdown('###### Pontuação Total por Categoria:')
        for i, categoria in enumerate(colunas_matriz):
            pontuacao = df_resumo_cat.loc[df_resumo_cat['Categoria'] == categoria, 'Pontuacao_Categoria'].sum()
            with colunas_kpi_pontuacao[i]:
                 st.metric(
                    label=f"Pontos {categoria}", 
                    value=f"{pontuacao:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")
                 )

        st.markdown("---") # Divisor para separar os KPIs da Tabela
        
        # 4. TABELA DE DESEMPENHO INDIVIDUAL (Matriz)
        st.markdown("##### Tabela de Desempenho Individual")
        
        # Prepara a tabela para exibição (apenas colunas relevantes)
        df_tabela_exibicao = df_desempenho[[
            COLUNA_ESPECIFICADOR, 
            'Pontuacao_Total', 
            'Categoria'
        ]].copy()
        
        # Renomear colunas para o Português para exibição
        df_tabela_exibicao.columns = ['Especificador / Empresa', 'Pontuação', 'Categoria']
        
        # Formatar a coluna Pontuação para melhor visualização (com separador de milhar)
        df_tabela_exibicao['Pontuação'] = df_tabela_exibicao['Pontuação'].apply(
             lambda x: f"{x:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")
        )
        
        st.dataframe(df_tabela_exibicao.style.set_properties(**{'border': '1px solid #333333', 'text-align': 'center'}, subset=pd.IndexSlice[:, ['Pontuação', 'Categoria']]), use_container_width=True) # CENTRALIZAÇÃO AQUI
        
    # =======================================================================
    # ITEM 7: DESEMPENHO DE NOVOS CADASTROS (PONTOS, CONTAGEM E NOMES)
    # =======================================================================
    if COLUNA_ESPECIFICADOR in df_filtrado.columns and COLUNA_NUMERO_TEMPORADA in df_filtrado.columns:

        st.subheader("7. Análise de Novos Cadastrados (Aquisição)")
        
        # Filtra apenas os novos cadastrados no período e filtros de loja/segmento
        # ATENÇÃO: df_filtrado já contém Novo_Cadastrado marcado
        df_novos_filtrados = df_filtrado[df_filtrado['Novo_Cadastrado'] == True]
        
        # --- 7A. KPIs de Pontuação de Novos Cadastrados ---
        st.markdown("##### 7A. Pontuação e Contagem de Novos Clientes por Temporada")

        # Colunas de todas as temporadas disponíveis (T7, T8, T9, T10, etc.)
        colunas_temporada_str = [col for col in df_filtrado.columns if col.startswith('Temporada ')]
        # CRÍTICO: Ordenar colunas de T7 -> T10 para exibição dos KPIs
        colunas_temporada_str = sorted([col for col in colunas_temporada_str if col != 'Temporada_Exibicao'])
        
        # CRÍTICO: Garantir cols_kpi existe e não é zero.
        num_temporadas_kpi = max(1, len(colunas_temporada_str))
        cols_kpi_pontos = st.columns(num_temporadas_kpi)
        cols_kpi_contagem = st.columns(num_temporadas_kpi)
        
        pontos_por_temporada = {}
        contagem_por_temporada = {}
        
        st.markdown('###### Pontos por Temporada:')
        # 1. Loop para calcular e exibir Pontos por Temporada (T7, T8, T9, T10...)
        for i, t_col in enumerate(colunas_temporada_str):
            # CÁLCULO DE PONTOS: Somar Pontos onde o Novo_Cadastrado é True E a Temporada é a t_col
            pontos = df_novos_filtrados.loc[df_novos_filtrados['Temporada_Exibicao'] == t_col, 'Pontos'].sum()
            contagem_clientes = df_novos_filtrados.loc[df_novos_filtrados['Temporada_Exibicao'] == t_col, 'CNPJ_CPF_LIMPO'].nunique()
            
            pontos_por_temporada[t_col] = pontos
            contagem_por_temporada[t_col] = contagem_clientes
            
            # KPI de Pontos
            with cols_kpi_pontos[i]:
                st.metric(f"Pontos Novos {t_col.replace('Temporada ', 'T')}", 
                         f"{pontos:,.0f}".replace(",", "X").replace(".", ",").replace("X", "."))

        st.markdown('###### Clientes Novos por Temporada:')
        # 2. Loop para exibir Contagem por Temporada (T7, T8, T9, T10...)
        for i, t_col in enumerate(colunas_temporada_str):
            with cols_kpi_contagem[i]:
                st.metric(f"Clientes Novos {t_col.replace('Temporada ', 'T')}", 
                         f"{contagem_por_temporada[t_col]:,.0f}")
                
        st.markdown("---")
        
        # --- 7B. TABELA PIVÔ: Clientes Novos por Mês e Temporada ---
        st.markdown("##### 7B. Contagem de Novos Profissionais Pontuados por Mês e Temporada")

        if 'Mês_Exibicao' in df_novos_filtrados.columns:
            
            # Conta o número ÚNICO de CPF/CNPJ (Novos Clientes) por Mês e Temporada
            # USAMOS A COLUNA NOVO_CADASTRADO e AGGFUNC='sum' para contar apenas as linhas marcadas como True
            df_pivot_novos = df_novos_filtrados.pivot_table(
                index='Mês_Exibicao',
                columns='Temporada_Exibicao',
                values='Novo_Cadastrado', 
                aggfunc='sum', 
                fill_value=0
            ).reset_index()

            # Reordenação dos Meses (Julho a Junho)
            mes_ordem_fiscal = {
                'Jul (07)': 1, 'Ago (08)': 2, 'Set (09)': 3, 'Out (10)': 4, 'Nov (11)': 5, 
                'Dez (12)': 6, 'Jan (01)': 7, 'Fev (02)': 8, 'Mar (03)': 9, 'Abr (04)': 10,
                'Mai (05)': 11, 'Jun (06)': 12
            }
            df_pivot_novos['Ordem'] = df_pivot_novos['Mês_Exibicao'].map(mes_ordem_fiscal)
            df_pivot_novos.sort_values(by='Ordem', inplace=True)
            df_pivot_novos.drop('Ordem', axis=1, inplace=True)
            
            # Renomear colunas de temporada para 'Clientes T9', 'Clientes T10', etc.
            colunas_temporada_novos = [col for col in df_pivot_novos.columns if col.startswith('Temporada')]
            df_pivot_novos.columns = [
                'Mês' if col == 'Mês_Exibicao' else col.replace('Temporada ', 'Clientes T')
                for col in df_pivot_novos.columns
            ]
            
            # Ordenação das colunas de T7 -> T10
            colunas_clientes = [col for col in df_pivot_novos.columns if col.startswith('Clientes T')]
            
            # Adicionar linha de TOTAL
            total_row = pd.Series(df_pivot_novos[colunas_clientes].sum(), name='Total')
            total_row['Mês'] = 'Total'
            df_pivot_novos.set_index('Mês', inplace=True)
            df_pivot_novos = pd.concat([df_pivot_novos, pd.DataFrame(total_row).T.set_index('Mês')])

            # Estilização e Exibição
            st.dataframe(
                df_pivot_novos.style.format({col: '{:,.0f}' for col in colunas_clientes + ['Total']})
                                    .set_properties(**{'border': '1px solid #333333'}),
                use_container_width=True
            )
            
            # --- 7C. Tabela de Nomes (Detalhe dos Clientes Novos) ---
            st.markdown("##### 7C. Nomes dos Profissionais Novos (Com Compra na Temporada)")

            # Agrupa os novos clientes e mostra a primeira compra histórica para detalhe
            df_nomes_novos = df_novos_filtrados.groupby(
                [COLUNA_ESPECIFICADOR, 'CNPJ_CPF_LIMPO']
            ).agg(
                Primeira_Compra_Historica=('Data_Primeira_Compra_Historica', 'min'),
                Temporada_Cadastro=(COLUNA_NUMERO_TEMPORADA, 'first'),
                Pontos=('Pontos', 'sum')
            ).reset_index()
            
            # Renomeia e formata
            df_nomes_novos.columns = ['Nome', 'CPF/CNPJ', 'Primeira Compra', 'Temporada', 'Pontos']
            df_nomes_novos['Temporada'] = 'T' + df_nomes_novos['Temporada'].astype(str)
            df_nomes_novos['Pontos'] = df_nomes_novos['Pontos'].apply(lambda x: f"{x:,.0f}".replace(",", "X").replace(".", ",").replace("X", "."))
            df_nomes_novos['Primeira Compra'] = df_nomes_novos['Primeira Compra'].dt.strftime('%d/%m/%Y')
            
            st.dataframe(df_nomes_novos.style.set_properties(**{'border': '1px solid #333333'}), # REMOVENDO CENTRALIZAÇÃO AQUI
                         use_container_width=True)


# Mensagem se o DataFrame estiver vazio após o carregamento (não deve acontecer agora)
elif df_dados_original.empty and Relatorio == 'Relatorio.xlsx':
    st.warning("O DataFrame está vazio. Verifique se a planilha Excel tem dados.")
