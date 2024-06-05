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
        url = f'http://extrator.sefaz.al.gov.br/DESPESAS/COMPARATIVO-DESPESAS/CONSOLIDADO/despesa_empenhado_liquidado_pago_consolidado_2018-2024_siafe_gerado_em_{data_ontem}.csv'
        arquivo = os.path.join('S:/SOP/003 - GERÊNCIA DE ESTUDOS E PROJEÇÕES/DASHBOARDS POWERBI/Base python', f'despesa_empenhado_liquidado_pago_consolidado_2018-2024_siafe_gerado_em_{data_ontem}.xlsx')
        arquivo_csv = os.path.join('S:/SOP/003 - GERÊNCIA DE ESTUDOS E PROJEÇÕES/DASHBOARDS POWERBI/Base python', f'despesa_empenhado_liquidado_pago_consolidado_2018-2024_siafe_gerado_em_{data_ontem}.csv')
        base_despesas_arquivo = os.path.join('S:/SOP/003 - GERÊNCIA DE ESTUDOS E PROJEÇÕES/DASHBOARDS POWERBI/DESPESAS 2024', f'base_despesas.xlsx')
        base_despesas_arquivo_csv = os.path.join('S:/SOP/003 - GERÊNCIA DE ESTUDOS E PROJEÇÕES/DASHBOARDS POWERBI/DESPESAS 2024', f'base_despesas.csv')
        # Baixar o arquivo
        wget.download(url, arquivo_csv)
        wget.download(url, base_despesas_arquivo_csv)
    except: 
        url = f'http://extrator.sefaz.al.gov.br/DESPESAS/COMPARATIVO-DESPESAS/CONSOLIDADO/despesa_empenhado_liquidado_pago_consolidado_2018-2024_siafe_gerado_em_{data_atual}.csv'
        arquivo = os.path.join('S:/SOP/003 - GERÊNCIA DE ESTUDOS E PROJEÇÕES/DASHBOARDS POWERBI/Base python', f'despesa_empenhado_liquidado_pago_consolidado_2018-2024_siafe_gerado_em_{data_ontem}.xlsx')
        arquivo_csv = os.path.join('S:/SOP/003 - GERÊNCIA DE ESTUDOS E PROJEÇÕES/DASHBOARDS POWERBI/Base python', f'despesa_empenhado_liquidado_pago_consolidado_2018-2024_siafe_gerado_em_{data_ontem}.csv')
        base_despesas_arquivo = os.path.join('S:/SOP/003 - GERÊNCIA DE ESTUDOS E PROJEÇÕES/DASHBOARDS POWERBI/DESPESAS 2024', f'base_despesas.xlsx')
        base_despesas_arquivo_csv = os.path.join('S:/SOP/003 - GERÊNCIA DE ESTUDOS E PROJEÇÕES/DASHBOARDS POWERBI/DESPESAS 2024', f'base_despesas.csv')
        # Baixar o arquivo
        wget.download(url, arquivo_csv)
        wget.download(url, base_despesas_arquivo_csv)

    # DOTACAO
    print(' ')
    print('Atualizando DOTACAO...')
    print(' ')

    try:
        url = f'http://extrator.sefaz.al.gov.br/DESPESAS/COMPARATIVO-DOTACOES/CONSOLIDADO/comparativo_dotacao_despesa_consolidado_2018-2024_siafe_gerado_em_{data_ontem}.csv'
        arquivo_csv = os.path.join('S:/SOP/003 - GERÊNCIA DE ESTUDOS E PROJEÇÕES/DASHBOARDS POWERBI/Base python', f'comparativo_dotacao_despesa_consolidado_2018-2024_siafe_gerado_em_{data_ontem}.csv')
        base_dotacao_arquivo = os.path.join('S:/SOP/003 - GERÊNCIA DE ESTUDOS E PROJEÇÕES/DASHBOARDS POWERBI/DESPESAS 2024', f'base_dotacao.csv')
        # Baixar o arquivo
        wget.download(url, arquivo_csv)
        wget.download(url, base_dotacao_arquivo)
    except: 
        url = f'http://extrator.sefaz.al.gov.br/DESPESAS/COMPARATIVO-DOTACOES/CONSOLIDADO/comparativo_dotacao_despesa_consolidado_2018-2024_siafe_gerado_em_{data_atual}.csv'
        arquivo_csv = os.path.join('S:/SOP/003 - GERÊNCIA DE ESTUDOS E PROJEÇÕES/DASHBOARDS POWERBI/Base python', f'comparativo_dotacao_despesa_consolidado_2018-2024_siafe_gerado_em_{data_atual}.csv')
        base_dotacao_arquivo = os.path.join('S:/SOP/003 - GERÊNCIA DE ESTUDOS E PROJEÇÕES/DASHBOARDS POWERBI/DESPESAS 2024', f'base_dotacao.csv')
        # Baixar o arquivo
        wget.download(url, arquivo_csv)
        wget.download(url, base_dotacao_arquivo)

    # RECEITA
    print(' ')
    print('Atualizando RECEITA...')
    print(' ')
    
    try:
        url = f'http://extrator.sefaz.al.gov.br/RECEITAS/receita_consolidado_2018_2024_siafe_gerado_em_{data_ontem}.xlsx'
        arquivo = os.path.join('S:/SOP/003 - GERÊNCIA DE ESTUDOS E PROJEÇÕES/DASHBOARDS POWERBI/Base python', f'receita_consolidado_2018_2024_siafe_gerado_em_{data_ontem}.xlsx')
        base_receita_arquivo = os.path.join('S:/SOP/003 - GERÊNCIA DE ESTUDOS E PROJEÇÕES/DASHBOARDS POWERBI/DESPESAS 2024', f'receitas_base_18a24.xlsx')
        # Baixar o arquivo
        wget.download(url, arquivo)
        wget.download(url, base_receita_arquivo)
    except: 
        url = f'http://extrator.sefaz.al.gov.br/RECEITAS/receita_consolidado_2018_2024_siafe_gerado_em_{data_atual}.xlsx'
        arquivo = os.path.join('S:/SOP/003 - GERÊNCIA DE ESTUDOS E PROJEÇÕES/DASHBOARDS POWERBI/Base python', f'receita_consolidado_2018_2024_siafe_gerado_em_{data_atual}.xlsx')
        base_receita_arquivo = os.path.join('S:/SOP/003 - GERÊNCIA DE ESTUDOS E PROJEÇÕES/DASHBOARDS POWERBI/DESPESAS 2024', f'receitas_base_18a24.xlsx')
        # Baixar o arquivo
        wget.download(url, arquivo)
        wget.download(url, base_receita_arquivo)

    print(' ')
    print('ATUALIZACAO CONCLUIDA!')
    print(' ')
    popup_concluido()
except:
    print(' ')
    print('ATUALIZACAO NAO CONCLUIDA! Verificar possivel erro.')
    print(' ')
    popup_erro()