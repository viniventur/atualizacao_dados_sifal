import os
import pandas as pd
import wget
import ssl
from datetime import datetime, timedelta
from unidecode import unidecode
import tkinter as tk
from tkinter import messagebox
import warnings
warnings.filterwarnings('ignore')

data_atual = datetime.now()
data_ontem = data_atual - timedelta(days=1)
data_atual = data_atual.date().strftime('%d-%m-%Y')
data_ontem = data_ontem.date().strftime('%d-%m-%Y')

# mensagem de concluido
def popup_concluido():
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal
    messagebox.showinfo("Atualização das Bases", "Atualização das bases de despesa, dotação e receita concluída!")
    

# mensagem de NAO concluido
def popup_erro():
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal
    messagebox.showinfo("Atualização das Bases", "Atualização não sucedida!\nVerificar possível erro.")
    
try:

    # DESPESA
    print(' ')
    print('Atualizando DESPESA...')
    print(' ')

    try:
        # Desativar a verificação do certificado SSL
        ssl._create_default_https_context = ssl._create_unverified_context

        # URL do arquivo que você deseja baixar
        url = f'https://extrator.sefaz.al.gov.br/DESPESAS/COMPARATIVO-DESPESAS/despesa_empenhado_liquidado_pago_2024_siafe_gerado_em_{data_ontem}.xlsx'
        arquivo = os.path.join('S:/SOP/003 - GERÊNCIA DE ESTUDOS E PROJEÇÕES/DASHBOARDS POWERBI/Base python', f'despesa_empenhado_liquidado_pago_2024_siafe_gerado_em_{data_ontem}.xlsx')
        arquivo_csv = os.path.join('S:/SOP/003 - GERÊNCIA DE ESTUDOS E PROJEÇÕES/DASHBOARDS POWERBI/Base python', f'despesa_empenhado_liquidado_pago_2024_siafe_gerado_em_{data_ontem}.csv')
        base_despesas_arquivo = os.path.join('S:/SOP/003 - GERÊNCIA DE ESTUDOS E PROJEÇÕES/DASHBOARDS POWERBI/DESPESAS 2024', f'base_despesasxlsx')
        # Baixar o arquivo
        wget.download(url, arquivo)
    except: 
        # Desativar a verificação do certificado SSL
        ssl._create_default_https_context = ssl._create_unverified_context

        # URL do arquivo que você deseja baixar
        url = f'https://extrator.sefaz.al.gov.br/DESPESAS/COMPARATIVO-DESPESAS/despesa_empenhado_liquidado_pago_2024_siafe_gerado_em_{data_atual}.xlsx'
        arquivo = os.path.join('S:/SOP/003 - GERÊNCIA DE ESTUDOS E PROJEÇÕES/DASHBOARDS POWERBI/Base python', f'despesa_empenhado_liquidado_pago_2024_siafe_gerado_em_{data_atual}.xlsx')
        arquivo_csv = os.path.join('S:/SOP/003 - GERÊNCIA DE ESTUDOS E PROJEÇÕES/DASHBOARDS POWERBI/Base python', f'despesa_empenhado_liquidado_pago_2024_siafe_gerado_em_{data_atual}.csv')
        base_despesas_arquivo = os.path.join('S:/SOP/003 - GERÊNCIA DE ESTUDOS E PROJEÇÕES/DASHBOARDS POWERBI/DESPESAS 2024', f'base_despesas.xlsx')
        # Baixar o arquivo
        wget.download(url, arquivo)

    # lendo os arquivos
    df_novo = pd.read_excel(arquivo)
    df_antigo = pd.read_csv('S:/SOP/003 - GERÊNCIA DE ESTUDOS E PROJEÇÕES/5 - VINI/bases/base_despesas_csv.csv', sep=';')
    df_uo = pd.read_excel('S:/SOP/003 - GERÊNCIA DE ESTUDOS E PROJEÇÕES/5 - VINI/bases/UG_UO.xlsx', sheet_name='UO')
    df_ug = pd.read_excel('S:/SOP/003 - GERÊNCIA DE ESTUDOS E PROJEÇÕES/5 - VINI/bases/UG_UO.xlsx', sheet_name='UG')

    # manipulação e merge
    df_uo = df_uo.iloc[ : , 1:]
    df_ug = df_ug.iloc[ : , 1:]
    df_antigo = df_antigo[df_antigo.ANO != 2024]
    df = pd.concat([df_antigo, df_novo], axis=0)
    df_merge1 = pd.merge(df, df_uo, on='UO', suffixes=('', '_DROP')).filter(regex='^(?!.*_DROP)')
    df_merge = pd.merge(df_merge1, df_ug, on='UG', suffixes=('', '_DROP')).filter(regex='^(?!.*_DROP)')
    base_despesas = df_merge.copy()
    for coluna in df_merge.select_dtypes(include=[object, 'string']):
        base_despesas[coluna] = base_despesas[coluna].astype(str).replace(['\n', '\r', '  ', '   ', '    ', '             '], '', regex=True).apply(lambda x: unidecode(str(x)))
    base_despesas.to_excel('S:/SOP/003 - GERÊNCIA DE ESTUDOS E PROJEÇÕES/DASHBOARDS POWERBI/DESPESAS 2024/base_despesas.xlsx', index=False)


    # DOTACAO
    print(' ')
    print('Atualizando DOTACAO...')
    print(' ')

    try:
        # Desativar a verificação do certificado SSL
        ssl._create_default_https_context = ssl._create_unverified_context

        # URL do arquivo que você deseja baixar
        url = f'https://extrator.sefaz.al.gov.br/DESPESAS/COMPARATIVO-DOTACOES/comparativo_dotacao_despesa_2024_siafe_gerado_em_{data_ontem}.xlsx'
        arquivo = os.path.join('S:/SOP/003 - GERÊNCIA DE ESTUDOS E PROJEÇÕES/DASHBOARDS POWERBI/Base python', f'comparativo_dotacao_despesa_2024_siafe_gerado_em_{data_ontem}.xlsx')
        arquivo_csv = os.path.join('S:/SOP/003 - GERÊNCIA DE ESTUDOS E PROJEÇÕES/DASHBOARDS POWERBI/Base python', f'comparativo_dotacao_despesa_2024_siafe_gerado_em_{data_ontem}.csv')
        base_dotacao_arquivo = os.path.join('S:/SOP/003 - GERÊNCIA DE ESTUDOS E PROJEÇÕES/DASHBOARDS POWERBI/DESPESAS 2024', f'base_dotacao.csv')

        # Baixar o arquivo
        wget.download(url, arquivo)
        # lendo os arquivos
        df_novo = pd.read_excel(arquivo)
        df_novo.drop(['CODIGO_FAVORECIDO', 'NOME_FAVORECIDO'], axis=1, inplace=True)
    except: 
        # Desativar a verificação do certificado SSL
        ssl._create_default_https_context = ssl._create_unverified_context

        # URL do arquivo que você deseja baixar
        url = f'https://extrator.sefaz.al.gov.br/DESPESAS/COMPARATIVO-DOTACOES/comparativo_dotacao_despesa_2024_siafe_gerado_em_{data_atual}.xlsx'
        arquivo = os.path.join('S:/SOP/003 - GERÊNCIA DE ESTUDOS E PROJEÇÕES/DASHBOARDS POWERBI/Base python', f'comparativo_dotacao_despesa_2024_siafe_gerado_em_{data_atual}.xlsx')
        arquivo_csv = os.path.join('S:/SOP/003 - GERÊNCIA DE ESTUDOS E PROJEÇÕES/DASHBOARDS POWERBI/Base python', f'comparativo_dotacao_despesa_2024_siafe_gerado_em_{data_atual}.csv')
        base_dotacao_arquivo = os.path.join('S:/SOP/003 - GERÊNCIA DE ESTUDOS E PROJEÇÕES/DASHBOARDS POWERBI/DESPESAS 2024', f'base_dotacao.csv')
        # Baixar o arquivo
        wget.download(url, arquivo)
        # lendo os arquivos
        df_novo = pd.read_excel(arquivo)
        df_novo.drop(['CODIGO_FAVORECIDO', 'NOME_FAVORECIDO'], axis=1, inplace=True)

    # baixando os arquivos
    df_antigo = pd.read_csv('S:/SOP/003 - GERÊNCIA DE ESTUDOS E PROJEÇÕES/5 - VINI/bases/Bases_concatenadas.csv')
    df_uo = pd.read_excel('S:/SOP/003 - GERÊNCIA DE ESTUDOS E PROJEÇÕES/5 - VINI/bases/UG_UO.xlsx', sheet_name='UO')
    df_ug = pd.read_excel('S:/SOP/003 - GERÊNCIA DE ESTUDOS E PROJEÇÕES/5 - VINI/bases/UG_UO.xlsx', sheet_name='UG')
    # manipulação e merge
    df_uo = df_uo.iloc[ : , 1:]
    df_ug = df_ug.iloc[ : , 1:]
    df_antigo = df_antigo[df_antigo.ANO != 2024]
    df = pd.concat([df_antigo, df_novo], axis=0)
    df_merge1 = pd.merge(df, df_uo, on='UO', suffixes=('', '_DROP')).filter(regex='^(?!.*_DROP)')
    df_merge = pd.merge(df_merge1, df_ug, on='UG', suffixes=('', '_DROP')).filter(regex='^(?!.*_DROP)')
    base_dotacao = df_merge.copy()
    for coluna in df_merge.select_dtypes(include=[object, 'string']):
        base_dotacao[coluna] = base_dotacao[coluna].astype(str).replace(['\n', '\r', '  ', '   ', '    ', '             '], '', regex=True).apply(lambda x: unidecode(str(x)))
    base_dotacao.to_csv('S:/SOP/003 - GERÊNCIA DE ESTUDOS E PROJEÇÕES/DASHBOARDS POWERBI/DESPESAS 2024/base_dotacao.csv', index=False)

    # RECEITA
    print(' ')
    print('Atualizando RECEITA...')
    print(' ')

    try:
        # Desativar a verificação do certificado SSL
        ssl._create_default_https_context = ssl._create_unverified_context

        # URL do arquivo que você deseja baixar
        url = f'https://extrator.sefaz.al.gov.br/RECEITAS/receita_2024_siafe_gerado_em_{data_ontem}.xlsx'
        arquivo = os.path.join('S:/SOP/003 - GERÊNCIA DE ESTUDOS E PROJEÇÕES/DASHBOARDS POWERBI/Base python', f'receita_2024_siafe_gerado_em_{data_ontem}.xlsx')
        base_receita_arquivo = os.path.join('S:/SOP/003 - GERÊNCIA DE ESTUDOS E PROJEÇÕES/DASHBOARDS POWERBI/DESPESAS 2024', f'receitas_base_18a24.xlsx')

        # Baixar o arquivo
        wget.download(url, arquivo)
        # lendo os arquivos
        df_novo = pd.read_excel(arquivo)
    except: 
        # Desativar a verificação do certificado SSL
        ssl._create_default_https_context = ssl._create_unverified_context

        # URL do arquivo que você deseja baixar
        url = f'https://extrator.sefaz.al.gov.br/RECEITAS/receita_2024_siafe_gerado_em_{data_atual}.xlsx'
        arquivo = os.path.join('S:/SOP/003 - GERÊNCIA DE ESTUDOS E PROJEÇÕES/DASHBOARDS POWERBI/Base python', f'receita_2024_siafe_gerado_em_{data_atual}.xlsx')
        base_receita_arquivo = os.path.join('S:/SOP/003 - GERÊNCIA DE ESTUDOS E PROJEÇÕES/DASHBOARDS POWERBI/DESPESAS 2024', f'receitas_base_18a24.xlsx')
        # Baixar o arquivo
        wget.download(url, arquivo)
        # lendo os arquivos
        df_novo = pd.read_excel(arquivo)

    # baixando os arquivos
    df_antigo = pd.read_excel(base_receita_arquivo)
    df_ug = pd.read_excel('S:/SOP/003 - GERÊNCIA DE ESTUDOS E PROJEÇÕES/5 - VINI/bases/UG_UO.xlsx', sheet_name='UG')
    # manipulação e merge
    df_ug = df_ug.iloc[ : , 1:]
    df_antigo = df_antigo[df_antigo.ANO != 2024]
    df = pd.concat([df_antigo, df_novo], axis=0)
    #df_merge = pd.merge(df, df_ug, on='UG', suffixes=('', '_DROP')).filter(regex='^(?!.*_DROP)')
    base_receita = df.copy()
    base_receita.to_excel('S:/SOP/003 - GERÊNCIA DE ESTUDOS E PROJEÇÕES/DASHBOARDS POWERBI/DESPESAS 2024/receitas_base_18a24.xlsx', index=False)

    print(' ')
    print('ATUALIZACAO CONCLUIDA!')
    print(' ')
    popup_concluido()
except:
    print(' ')
    print('ATUALIZACAO NAO CONCLUIDA! Verificar possivel erro.')
    print(' ')
    popup_erro()
