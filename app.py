import streamlit as st
import pandas as pd
from io import BytesIO
import os
from datetime import datetime


import streamlit as st
import pandas as pd
import os

# Definir a configuração da página
st.set_page_config(page_title="Relatório Completo de Sobressalentes", layout="wide")

st.title("Relatório Completo de Sobressalentes")



# Upload do arquivo de dados brutos
file = st.file_uploader("Carregar arquivo Excel com os dados brutos", type=["xlsx"])

if file:
    st.success("Arquivo carregado com sucesso. Processando os dados...")

    
        # Ler os dados do arquivo carregado pelo usuário
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
    reparo = reparo[['Cód. Produto', 'Desc. Produto', 'Complemento Remessa', 'RMA', 'MBI', 'Desc. Status', 
                         'N° Nota Fiscal', 'Série Nota Fiscal', 'Quantidade', 'Cód. Estoque Físico', 
                         'Desc. Estoque Físico', 'Data Últ. Alteração', 'Material do Fornecedor', 'Cód. Natureza NF', 
                         'Desc. Natureza NF', 'Cód. Fornecedor', 'Fornecedor', 'TA', 'Data Status Atual', 'Cód. RM', 
                         'Data Cadastro RM', 'Data Empenho RM', 'Data Fechamento RM', 'Cód. Doc. Entrada', 
                         'Data Fech. Doc. Entrada', 'Complemento Doc. Entrada', 'Data Aguard. RMA', 'Data Aguard. Rem. Fornec.',
                         'Data Aguard. Operacional', 'Data Aguard. Aprov. Contr.', 'Data Aguard. Escrituração', 
                         'Data Aguard. NF', 'Data Aguard. ST', 'Data Aguard. Coleta', 'Data Coletado Em Trânsito', 
                         'Data Recebido Fornecedor', 'Observação']]

    reparo = reparo[reparo['Desc. Status'] != 'COLETADO/TRÂNSITO']

    base = base[['Cód. Produto', 'Desc. Produto', 'Qtd Estoque', 'Serial', 'Part Number', 'Code', 
                     'Classificação', 'Id. Estoq. Físico', 'Desc. Estoque Físico']]

    volante = volante[['IDTEL', 'NOME_VOLANTE', 'CODIGO_PRODUTO', 'DESCRICAO_PRODUTO', 'SALDO', 'COMPLEMENTAR', 
                           'PART_NUMBER', 'QTDE_DIAS_ATEND_ULT_RM', 'DESCRICAO_CLASSIFICACAO', 'ITEM_CONTABIL']]

        # Ajuste de classificação
    sit = {'MATERIAL DO CLIENTE': 'DISPONÍVEL', 'RETIRADA': 'RETIRADA'}
    volante['DESCRICAO_CLASSIFICACAO'] = volante['DESCRICAO_CLASSIFICACAO'].apply(lambda x: sit.get(x, x))

    sob = sob[['Cód. Produto', 'Desc. Produto', 'Qtd Estoque', 'Serial', 'Part Number', 'Classificação', 
                   'Id. Estoq. Físico', 'Desc. Estoque Físico']]

        # Sobrescrever planilhas antigas e salvar novas planilhas limpas
    with pd.ExcelWriter('planilhas.xlsx', engine='xlsxwriter') as writer:
            base.to_excel(writer, sheet_name='Saldo', index=False)
            volante.to_excel(writer, sheet_name='Volante', index=False)
            reparo.to_excel(writer, sheet_name='Reparo', index=False)
            sob.to_excel(writer, sheet_name='Sobressalentes', index=False)
    st.success("Dados processados com sucesso e planilhas criadas.")

   

        # Leitura dos arquivos gerados
    
    df = pd.read_excel('planilhas.xlsx')
    loc = pd.read_excel('localidade.xlsx')
    vol = pd.read_excel('planilhas.xlsx', sheet_name='Volante')
    id_tel = pd.read_excel('id_tel.xlsx')
    fabr = pd.read_excel('fabricante.xlsx')
    rep = pd.read_excel('planilhas.xlsx', sheet_name='Reparo')
    sob = pd.read_excel('planilhas.xlsx', sheet_name='Sobressalentes')

        # Remover espaços em branco dos nomes das colunas
    # Remover espaços em branco dos nomes das colunas
    for dataset in [df, loc, vol, rep, sob]:
        dataset.columns = [x.strip() for x in dataset.columns]



    # Mesclar "sob" com "loc" e "fabr" para adicionar localidades e fabricantes
    sob = pd.merge(sob, loc[['Id. Estoq. Físico', 'LOCALIDADE']], on='Id. Estoq. Físico', how='left')
    sob = pd.merge(sob, fabr[['Cód. Produto', 'Fabricante']], on='Cód. Produto', how='left')

    # Separar linhas com Serial vazio e remover duplicatas de Serial não vazios
    df_serial_vazio = sob[sob['Serial'].isna()]
    df_sem_duplicatas = sob.dropna(subset=['Serial']).drop_duplicates(subset=['Serial', 'Cód. Produto'], keep='first')
    sob = pd.concat([df_sem_duplicatas, df_serial_vazio], ignore_index=True)



    # Reordenar as colunas de "sob"
    colunas_inicio = ['LOCALIDADE']
    colunas_fim = ['Id. Estoq. Físico', 'Desc. Estoque Físico']
    colunas_meio = [col for col in sob.columns if col not in colunas_inicio + colunas_fim]
    sob = sob[colunas_inicio + colunas_meio + colunas_fim]

    # Mesclar "df" com "loc" e "fabr" para adicionar localidades e fabricantes
    df = pd.merge(df, loc[['Id. Estoq. Físico', 'LOCALIDADE']], on='Id. Estoq. Físico', how='left')
    df = pd.merge(df, fabr[['Cód. Produto', 'Fabricante']], on='Cód. Produto', how='left')

    # Separar linhas com Serial vazio e remover duplicatas de Serial não vazios
    df_serial_vazio = df[df['Serial'].isna()]
    df_sem_duplicatas = df.dropna(subset=['Serial']).drop_duplicates(subset=['Serial', 'Cód. Produto'], keep='first')
    df = pd.concat([df_sem_duplicatas, df_serial_vazio], ignore_index=True)

    # Criar coluna concatenada 'Serial_Conc' para "df"
    df['Serial_Conc'] = df[['Id. Estoq. Físico', 'Cód. Produto', 'Serial']].apply(lambda x: ''.join([str(i) for i in x if pd.notna(i)]), axis=1)

    # Classificação dos materiais com base no dicionário "sit"
    sit = {'MATERIAL DO CLIENTE': 'DISPONÍVEL', 'RETIRADA': 'RETIRADA', 'DEFEITO': 'DEFEITO', 'SUCATA': 'INSERVÍVEL', '0,0000': 'DISPONÍVEL'}
    df['CLASSIFICAÇÃO'] = df['Classificação'].apply(lambda x: sit.get(x, x))

    # Filtrar defeitos e materiais disponíveis
    defeito = df[(df['CLASSIFICAÇÃO'] != 'DISPONÍVEL') | (df['Id. Estoq. Físico'] == 3446)]
    df = df[(df['CLASSIFICAÇÃO'] == 'DISPONÍVEL') & (df['Id. Estoq. Físico'] != 3446)]

    # Reordenar colunas de "df"
    colunas_meio = [col for col in df.columns if col not in colunas_inicio + colunas_fim]
    
    df = df[colunas_inicio + colunas_meio + colunas_fim]

    vol.rename(columns={'CODIGO_PRODUTO':'Cód. Produto'}, inplace=True)
    vol.rename(columns={'IDTEL':'ID Tel'}, inplace=True)


    # Mesclar "vol" com "fabr" e "id_tel" para adicionar fabricantes e localidades
    vol = pd.merge(vol, fabr[['Cód. Produto', 'Fabricante']], on='Cód. Produto', how='left')
    vol = pd.merge(vol, id_tel[['ID Tel', 'LOCALIDADE', 'SUPERVISAO']], on='ID Tel', how='left')

    # Converter as colunas em uma lista para reorganização
    colunas = vol.columns.tolist()

    # Encontrar os índices das colunas 'I.C' e 'Fabricante'
    indice_ic, indice_fabricante = colunas.index('ITEM_CONTABIL'), colunas.index('Fabricante')

    # Trocar as colunas 'I.C' e 'Fabricante' de lugar
    colunas[indice_ic], colunas[indice_fabricante] = colunas[indice_fabricante], colunas[indice_ic]

    # Reorganizar o DataFrame de acordo com a nova ordem de colunas
    vol = vol[colunas]

    # Lista de códigos para manter na coluna 'I.C' da planilha 'vol'
    codigos_ic_para_manter = [11701, 11601, 10201]

    # Filtrar o DataFrame 'vol' para manter apenas os códigos na lista da coluna 'I.C'
    vol = vol[vol['ITEM_CONTABIL'].isin(codigos_ic_para_manter)]

    rep.rename(columns={'Cód. Estoque Físico':'Id. Estoq. Físico'}, inplace=True)

    reparo = pd.merge(rep, loc[['Id. Estoq. Físico', 'LOCALIDADE']], on='Id. Estoq. Físico', how='left')

    # Criar coluna concatenada 'Serial_Conc' para "reparo"
    reparo['Serial_Conc'] = reparo[['Id. Estoq. Físico', 'Cód. Produto', 'Complemento Remessa']].apply(lambda x: ''.join([str(i) for i in x if pd.notna(i)]), axis=1)

    # Reordenar colunas de "reparo" e remover colunas desnecessárias
    colunas = reparo.columns.tolist()
    colunas_inicio = ['LOCALIDADE']
    colunas_fim = ['Serial_Conc']

    colunas_meio = [col for col in reparo.columns if col not in colunas_inicio + colunas_fim]
    reparo = reparo[colunas_inicio + colunas_meio + colunas_fim]
    reparo = reparo.drop(['Material do Fornecedor', 'Cód. Natureza NF'], axis=1)

    # Calcular dias desde a 'Data Status Atual' e adicionar coluna 'Dias'
    reparo['Dias'] = (datetime.now() - reparo['Data Status Atual']).dt.days
    posicao_coluna = reparo.columns.get_loc('Data Status Atual')
    reparo.insert(posicao_coluna + 1, 'Dias', reparo.pop('Dias'))

    # Lista de códigos para manter
    codigos_para_manter = [2887, 3096, 2908, 3095, 2911, 2888, 3098, 3446, 2909, 2953, 3448,
                        2912, 2940, 2938, 3185, 3722, 3723, 3721, 3690, 3691, 3704, 3706,
                        3703, 3707, 3694, 3695, 3697, 3698, 3700, 3701, 3692, 3447, 2910, 2889]

    # Filtrar o DataFrame "reparo" para manter apenas os códigos na lista
    reparo = reparo[reparo['Id. Estoq. Físico'].isin(codigos_para_manter)]

    # Aplicar filtro de Serial_Conc para "df" e "defeito"
    df = df[~df['Serial_Conc'].isin(reparo['Serial_Conc'])]
    defeito = defeito[~defeito['Serial_Conc'].isin(reparo['Serial_Conc'])]

    # Mover 'Serial_Conc' para o final no DataFrame "df"
    colunas = df.columns.tolist()
    colunas.append(colunas.pop(colunas.index('Serial_Conc')))
    df = df[colunas]

    # Atualizar classificação dos defeitos com base no dicionário "sit"
    sit_defeito = {'RETIRADA': 'DEFEITO', 'INSERVÍVEL': 'INSERVÍVEL', 'DISPONÍVEL': 'INSERVÍVEL'}
    defeito['CLASSIFICAÇÃO'] = defeito['CLASSIFICAÇÃO'].apply(lambda x: sit_defeito.get(x, x))

    # Ajustar a classificação em "sob"
    Sob_ajuste = {'MATERIAL DO CLIENTE': 'DISPONÍVEL', 'DEFEITO': 'DEFEITO'}
    sob['CLASSIFICAÇÃO'] = sob['Classificação'].apply(lambda x: Sob_ajuste.get(x, x))

    defeito = defeito.drop('Classificação', axis=1)
    df = df.drop('Classificação', axis = 1)

    # Obter a lista de colunas e reordenar o DataFrame "defeito"
    colunas = defeito.columns.tolist()
    colunas.remove('LOCALIDADE')
    colunas.remove('Id. Estoq. Físico')
    colunas.remove('Desc. Estoque Físico')
    colunas.remove('Serial_Conc')

    colunas.insert(0, 'LOCALIDADE')
    colunas.extend(['Id. Estoq. Físico', 'Desc. Estoque Físico', 'Serial_Conc'])
    defeito = defeito[colunas]

    # Preencher valores ausentes de "Fabricante"
    for dataset in [sob, df, vol, defeito]:
        dataset['Fabricante'] = dataset['Fabricante'].fillna('NÃO LOC')

    if os.path.exists('planilhas_combinadas.xlsx'):
            os.remove('planilhas_combinadas.xlsx')    

    # Criar arquivo Excel com múltiplas abas
    with pd.ExcelWriter('planilhas_combinadas.xlsx', engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Saldo', index=False)
        vol.to_excel(writer, sheet_name='Volante', index=False)
        reparo.to_excel(writer, sheet_name='Reparo', index=False)
        defeito.to_excel(writer, sheet_name='Defeito', index=False)
        sob.to_excel(writer, sheet_name='Sobressalentes', index=False)

    print("As planilhas foram combinadas em um único arquivo com sucesso.")


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



