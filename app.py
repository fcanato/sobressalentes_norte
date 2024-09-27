
import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO

# Definir a configuração da página antes de qualquer outra chamada do Streamlit
st.set_page_config(page_title="Relatório Completo de Sobressalentes", layout="wide")

# Função para limpar e processar os dados brutos e salvar como "planilhas.xlsx"
def processar_dados_iniciais(file):
    base = pd.read_excel(file)
    volante = pd.read_excel(file, sheet_name='volante')
    reparo = pd.read_excel(file, sheet_name='reparo')
    sob = pd.read_excel(file, sheet_name='sobressalentes')

    # Remover espaços em branco dos nomes das colunas
    base.columns = [x.strip() for x in base.columns]
    volante.columns = [x.strip() for x in volante.columns]
    reparo.columns = [x.strip() for x in reparo.columns]
    sob.columns = [x.strip() for x in sob.columns]

    # Limpeza e filtragem das planilhas
    reparo = reparo[['Cód. Produto','Desc. Produto','Complemento Remessa','RMA','MBI','Desc. Status','N° Nota Fiscal',
        'Série Nota Fiscal','Quantidade','Cód. Estoque Físico','Desc. Estoque Físico','Data Últ. Alteração',
        'Material do Fornecedor','Cód. Natureza NF','Desc. Natureza NF','Cód. Fornecedor','Fornecedor',
        'TA','Data Status Atual','Cód. RM','Data Cadastro RM','Data Empenho RM','Data Fechamento RM',
        'Cód. Doc. Entrada','Data Fech. Doc. Entrada','Complemento Doc. Entrada','Data Aguard. RMA',
        'Data Aguard. Rem. Fornec.','Data Aguard. Operacional','Data Aguard. Aprov. Contr.','Data Aguard. Escrituração',
        'Data Aguard. NF','Data Aguard. ST','Data Aguard. Coleta','Data Coletado Em Trânsito','Data Recebido Fornecedor',
        'Observação']]
    reparo = reparo[reparo['Desc. Produto'] != 'COLETADO/TRÂNSITO']

    base = base[['Cód. Produto', 'Desc. Produto', 'Qtd Estoque', 'Serial', 'Part Number', 'Code', 'Classificação', 'Id. Estoq. Físico']]

    volante = volante[['IDTEL','NOME_VOLANTE','CODIGO_PRODUTO','DESCRICAO_PRODUTO','SALDO','COMPLEMENTAR','PART_NUMBER',
         'QTDE_DIAS_ATEND_ULT_RM','DESCRICAO_CLASSIFICACAO','ITEM_CONTABIL']]

    # Ajuste de classificação
    sit = {'MATERIAL DO CLIENTE': 'DISPONÍVEL', 'RETIRADA': 'RETIRADA'}
    volante['DESCRICAO_CLASSIFICACAO'] = volante['DESCRICAO_CLASSIFICACAO'].apply(lambda x: sit.get(x, x))

    sob = sob[['Cód. Produto','Desc. Produto','Qtd Estoque','Serial','Part Number','Code','Classificação','Id. Estoq. Físico',
      'Desc. Estoque Físico']]

    # Criar um arquivo Excel com múltiplas abas limpas como "planilhas.xlsx"
    with pd.ExcelWriter('planilhas.xlsx', engine='xlsxwriter') as writer:
        base.to_excel(writer, sheet_name='Saldo', index=False)
        volante.to_excel(writer, sheet_name='Volante', index=False)
        reparo.to_excel(writer, sheet_name='Reparo', index=False)
        sob.to_excel(writer, sheet_name='Sobressalentes', index=False)

    return 'planilhas.xlsx'

# Função para exibir os dados brutos e processar o arquivo
def carregar_e_processar_dados():
    st.title("Relatorio de Sobressalentes - Norte")

    # Upload do arquivo de dados brutos
    file = st.file_uploader("Carregar arquivo Excel com os dados brutos", type=["xlsx"])

    if file:
        st.success("Arquivo carregado com sucesso. Processando os dados...")

        # Processar o arquivo e salvar como 'planilhas.xlsx'
        caminho_arquivo_processado = processar_dados_iniciais(file)

     

# Chamar a função de carregar e processar os dados
carregar_e_processar_dados()

# Função para processar "planilhas.xlsx" e gerar "planilhas_combinadas.xlsx"
def gerar_planilhas_combinadas():
    df_saldo = pd.read_excel('planilhas.xlsx', sheet_name='Saldo')
    df_volante = pd.read_excel('planilhas.xlsx', sheet_name='Volante')
    df_reparo = pd.read_excel('planilhas.xlsx', sheet_name='Reparo')
    df_sobressalentes = pd.read_excel('planilhas.xlsx', sheet_name='Sobressalentes')

    # Processamento adicional e merge das informações (exemplo do seu código anterior)
    # Aqui você pode fazer o merge com outros arquivos, adicionar colunas, ou ajustar os dados conforme a necessidade.
    
    # Criar o arquivo final "planilhas_combinadas.xlsx"
    with pd.ExcelWriter('planilhas_combinadas.xlsx', engine='xlsxwriter') as writer:
        df_saldo.to_excel(writer, sheet_name='Saldo', index=False)
        df_volante.to_excel(writer, sheet_name='Volante', index=False)
        df_reparo.to_excel(writer, sheet_name='Reparo', index=False)
        df_sobressalentes.to_excel(writer, sheet_name='Sobressalentes', index=False)

    return 'planilhas_combinadas.xlsx'




# Função para padronizar os nomes das localidades
def padronizar_localidade(df):
    df['LOCALIDADE'] = df['LOCALIDADE'].str.normalize('NFKD').str.encode('ascii', errors='ignore').str.decode('utf-8')
    return df

# Função para carregar e processar os dados para Saldo Volante
def carregar_saldo_volante():
    df_volante = pd.read_excel('planilhas_combinadas.xlsx', sheet_name='Volante')

    # Padronizar os nomes das localidades
    df_volante = padronizar_localidade(df_volante)

    # Agrupar por localidade e filtrar com base em DESCRICAO_CLASSIFICACAO
    df_volante_grouped = df_volante.groupby('LOCALIDADE').agg(
        Disponível=('SALDO', lambda x: x[df_volante['DESCRICAO_CLASSIFICACAO'] == 'DISPONÍVEL'].sum()),
        Retirada=('SALDO', lambda x: x[df_volante['DESCRICAO_CLASSIFICACAO'] == 'RETIRADA'].sum())
    ).reset_index()

    # Calcular o total geral
    df_volante_grouped['Total Geral'] = df_volante_grouped['Disponível'] + df_volante_grouped['Retirada']

    # Calcular os totais
    totais_volante = df_volante_grouped[['Disponível', 'Retirada', 'Total Geral']].sum()

    # Adicionar linha de Total Geral
    df_totais_volante = pd.DataFrame([['Total Geral', *totais_volante]], columns=df_volante_grouped.columns)
    df_volante_grouped = pd.concat([df_volante_grouped, df_totais_volante], ignore_index=True)

    return df_volante_grouped

# Função para carregar e processar os dados para Saldo de Estoque DW DM
def carregar_saldo_dw_dm():
    df_sobressalentes = pd.read_excel('planilhas_combinadas.xlsx', sheet_name='Sobressalentes')

    # Padronizar os nomes das localidades
    df_sobressalentes = padronizar_localidade(df_sobressalentes)

    # Agrupar por localidade
    df_dw_dm_grouped = df_sobressalentes.groupby('LOCALIDADE').agg(
        Disponível=('Qtd Estoque', lambda x: x[df_sobressalentes['CLASSIFICAÇÃO'] == 'DISPONÍVEL'].sum()),
        Defeito=('Qtd Estoque', lambda x: x[df_sobressalentes['CLASSIFICAÇÃO'] == 'DEFEITO'].sum())
    ).reset_index()

    # Calcular o total geral
    df_dw_dm_grouped['Total Geral'] = df_dw_dm_grouped['Disponível'] + df_dw_dm_grouped['Defeito']

    # Calcular os totais
    totais_dw_dm = df_dw_dm_grouped[['Disponível', 'Defeito', 'Total Geral']].sum()

    # Adicionar linha de Total Geral
    df_totais_dw_dm = pd.DataFrame([['Total Geral', *totais_dw_dm]], columns=df_dw_dm_grouped.columns)
    df_dw_dm_grouped = pd.concat([df_dw_dm_grouped, df_totais_dw_dm], ignore_index=True)

    return df_dw_dm_grouped

# Função para carregar e processar os dados para Saldo de Estoque
def carregar_saldo_estoque():
    df_sobressalentes = pd.read_excel('planilhas_combinadas.xlsx', sheet_name='Sobressalentes')
    df_defeito = pd.read_excel('planilhas_combinadas.xlsx', sheet_name='Defeito')  # Sheet de Defeito
    df_reparo = pd.read_excel('planilhas_combinadas.xlsx', sheet_name='Reparo')    # Sheet de Reparo
    df_disponivel = pd.read_excel('planilhas_combinadas.xlsx', sheet_name='Saldo') # Sheet Saldo

    # Padronizar os nomes das localidades em todas as sheets
    df_sobressalentes = padronizar_localidade(df_sobressalentes)
    df_defeito = padronizar_localidade(df_defeito)
    df_reparo = padronizar_localidade(df_reparo)
    df_disponivel = padronizar_localidade(df_disponivel)

    # Agrupar os dados da sheet Defeito para pegar as colunas 'Defeito' e 'Inservível'
    df_defeito_grouped = df_defeito.groupby('LOCALIDADE').agg(
        Defeito=('Qtd Estoque', lambda x: x[df_defeito['CLASSIFICAÇÃO'] == 'DEFEITO'].sum()),
        Inservível=('Qtd Estoque', lambda x: x[df_defeito['CLASSIFICAÇÃO'] == 'INSERVÍVEL'].sum())
    ).reset_index()

    # Agrupar os dados da sheet Reparo para pegar o 'Processo de Reparo'
    df_reparo_grouped = df_reparo.groupby('LOCALIDADE').agg(
        Processo_Reparo=('Quantidade', 'sum')
    ).reset_index()

    # Agrupar os dados da sheet Saldo para pegar 'Disponível'
    df_disponivel_grouped = df_disponivel.groupby('LOCALIDADE').agg(
        Disponível=('Qtd Estoque', lambda x: x[df_disponivel['CLASSIFICAÇÃO'] == 'DISPONÍVEL'].sum())
    ).reset_index()

    # Unir os dados agrupados
    df_saldo_estoque_grouped = pd.merge(df_defeito_grouped, df_disponivel_grouped, on='LOCALIDADE', how='left')
    df_saldo_estoque_grouped = pd.merge(df_saldo_estoque_grouped, df_reparo_grouped, on='LOCALIDADE', how='left')

    # Preencher valores ausentes com 0
    df_saldo_estoque_grouped.fillna(0, inplace=True)

    # Calcular o total geral
    df_saldo_estoque_grouped['Total Geral'] = df_saldo_estoque_grouped['Defeito'] + df_saldo_estoque_grouped['Disponível'] + df_saldo_estoque_grouped['Inservível'] + df_saldo_estoque_grouped['Processo_Reparo']

    # Calcular os totais para cada coluna
    totais = df_saldo_estoque_grouped[['Defeito', 'Disponível', 'Inservível', 'Processo_Reparo', 'Total Geral']].sum()

    # Adicionar linha de Total Geral
    df_totais = pd.DataFrame([['Total Geral', *totais]], columns=df_saldo_estoque_grouped.columns)
    df_saldo_estoque_grouped = pd.concat([df_saldo_estoque_grouped, df_totais], ignore_index=True)

    return df_saldo_estoque_grouped


# Aplicar o CSS personalizado
def custom_css():
    st.markdown("""
        <style>
            /* Estilizando a página */
            .reportview-container .main .block-container {
                max-width: 1200px;
                padding-top: 2rem;
                padding-right: 2rem;
                padding-left: 2rem;
                justify-content: center; /* Centraliza o conteúdo */
            }
            h1 {
                color: #8A2BE2;
                font-size: 2.5rem;
                text-align: center;
            }
            h2 {
                color: #4B0082;
                font-size: 2rem;
                text-align: center;
            }
            table {
                font-size: 0.9rem;
            }
            .css-18e3th9 {
                padding-top: 1rem;
            }
            .stDataFrame {
                margin-bottom: 2rem;
            }
                
            /* Alterar cor de fundo da página */
            .stApp {
            background-color: #F0F8FF; /* Escolha a cor desejada */
        }    
            /* Ajustar o fundo dos widgets */
            .stButton, .stTextInput, .stSelectbox, .stSidebar {
            background-color: #FFFFFF !important;
            color: black;  /* Altere conforme necessário */
        }    
       
        </style>
        """, unsafe_allow_html=True)

# Aplicar o CSS personalizado
custom_css()

# Função para gerar um resumo consolidado dos três relatórios (agrupados)
def gerar_resumo():
    df_saldo_volante = carregar_saldo_volante()
    df_saldo_dw_dm = carregar_saldo_dw_dm()
    df_saldo_estoque = carregar_saldo_estoque()

    # Organizando os dados em três colunas para ficarem lado a lado
    
    # Organizando os dados em duas colunas (Saldo Estoque e Volante) e a terceira centralizada abaixo
    col1, col2 = st.columns([1, 1])

    with col1:
        st.markdown("##### Saldo Estoque")
        st.dataframe(df_saldo_estoque.style.hide(axis='index').set_properties(**{'width': '150px'}))  # Ajusta a largura das colunas

    with col2:
        st.markdown("#####  Saldo Volante")
        st.dataframe(df_saldo_volante.style.hide(axis='index').set_properties(**{'width': '150px'}))  # Ajusta a largura das colunas

    # Terceira figura (Saldo DW DM) centralizada
    st.markdown("<div style='display: flex; justify-content: center;'>", unsafe_allow_html=True)
    st.markdown('#####  Saldo DW DM')
    st.dataframe(df_saldo_dw_dm.style.hide(axis='index').set_properties(**{'width': '300px'}))  # Ajusta a largura das colunas
    st.markdown("</div>", unsafe_allow_html=True)




# Barra lateral sempre fixa para navegação
st.sidebar.markdown("")
opcoes = st.sidebar.radio(
    "Escolha o relatório que deseja visualizar:",
    ("Página Principal", "Relatório saldo disponível", "Relatório saldo defeito", "Relatório - Saldo volante", "Controle de reparo", "Relatório - Saldo DW DM")
)

# Carregar automaticamente os dados de 'planilhas_combinadas.xlsx' após o processamento
file = 'planilhas_combinadas.xlsx'

def gerar_download_excel(df, nome_arquivo):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    writer.close() 
    dados_excel = output.getvalue()
    
    st.download_button(
        label="Baixar dados como Excel",
        data=dados_excel,
        file_name=f'{nome_arquivo}.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

# Verificação para manter o layout sempre visível
if not file:
    st.warning("Por favor, certifique-se de que o arquivo 'planilhas_combinadas.xlsx' está disponível para visualização dos relatórios.")
else:
    # Página principal com visão geral de todos os relatórios
    if opcoes == "Página Principal":
        st.markdown("", unsafe_allow_html=True)
        # Gerar e mostrar o resumo automaticamente
        gerar_resumo()

    # Relatório saldo disponível - Exibir os dados brutos da sheet 'Saldo'
    elif opcoes == "Relatório saldo disponível":
        df_saldo = pd.read_excel(file, sheet_name='Saldo')
        df_saldo = df_saldo.iloc[:,:-1]
        st.markdown("<h2 style='text-align: center;'>Relatório Saldo Disponível</h2>", unsafe_allow_html=True)
        st.dataframe(df_saldo.style.hide(axis='index'))
        gerar_download_excel(df_saldo, 'relatorio_saldo_disponivel')

    # Relatório saldo defeito - Exibir os dados brutos da sheet 'Defeito'
    elif opcoes == "Relatório saldo defeito":
        df_defeito = pd.read_excel(file, sheet_name='Defeito')
        df_defeito = df_defeito.iloc[:,:-1]
        st.markdown("<h2 style='text-align: center;'>Relatório Saldo Defeito</h2>", unsafe_allow_html=True)
        st.dataframe(df_defeito.style.hide(axis='index'))
        gerar_download_excel(df_defeito, 'relatorio_saldo_defeito')

    # Relatório Saldo Volante - Exibir os dados brutos da sheet 'Volante'
    elif opcoes == "Relatório - Saldo volante":
        df_volante = pd.read_excel(file, sheet_name='Volante')
        st.markdown("<h2 style='text-align: center;'>Relatório Saldo Volante</h2>", unsafe_allow_html=True)
        st.dataframe(df_volante.style.hide(axis='index'))
        gerar_download_excel(df_volante, 'relatorio_saldo_volante')

    # Controle de Reparo - Exibir os dados brutos da sheet 'Reparo'
    elif opcoes == "Controle de reparo":
        df_reparo = pd.read_excel(file, sheet_name='Reparo')
        # Eliminar a última coluna
        df_reparo = df_reparo.iloc[:, :-1]
        st.markdown("<h2 style='text-align: center;'>Controle de Reparo</h2>", unsafe_allow_html=True)
        st.dataframe(df_reparo.style.hide(axis='index'))
        gerar_download_excel(df_reparo, 'relatorio_controle_reparo')

    # Relatório Saldo DW DM - Exibir os dados brutos da sheet 'Sobressalentes'
    elif opcoes == "Relatório - Saldo DW DM":
        df_dw_dm = pd.read_excel(file, sheet_name='Sobressalentes')
        st.markdown("<h2 style='text-align: center;'>Relatório Saldo DW DM</h2>", unsafe_allow_html=True)
        st.dataframe(df_dw_dm.style.hide(axis='index'))
        gerar_download_excel(df_dw_dm, 'relatorio_saldo_dw_dm')

               




