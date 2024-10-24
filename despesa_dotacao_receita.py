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


def popup_concluido():
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal
    messagebox.showinfo("Atualização das Bases", "Atualização das bases de despesa, dotação e receita concluída!")

def popup_erro():
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal
    messagebox.showinfo("Atualização das Bases", "Atualização não sucedida!\nVerificar possível erro.")

def despesa():
    try:
        print(' ')
        print('Atualizando DESPESA...')
        print(' ')

        try:
            url = f'http://extrator.sefaz.al.gov.br/DESPESAS/COMPARATIVO-DESPESAS/CONSOLIDADO/despesa_empenhado_liquidado_pago_consolidado_2018-2024_siafe_gerado_em_{data_ontem}.csv'
            #df_desp = pd.read_csv(url, sep=';', encoding='latin2')
        except: 
            url = f'http://extrator.sefaz.al.gov.br/DESPESAS/COMPARATIVO-DESPESAS/CONSOLIDADO/despesa_empenhado_liquidado_pago_consolidado_2018-2024_siafe_gerado_em_{data_atual}.csv'
            #df_desp = pd.read_csv(url, sep=';', encoding='latin2')
        
        base_despesas_arquivo_csv = os.path.join('Z:/DADOS/EXTRATOR/DESPESAS/', f'base_despesas_18_24.csv')
        
        if os.path.exists(base_despesas_arquivo_csv):
                    os.remove(base_despesas_arquivo_csv)

        wget.download(url, base_despesas_arquivo_csv)

        #df_desp.to_csv(base_despesas_arquivo_csv, index=False, encoding='latin2', decimal=',')
    except Exception as e:
        print(' ') 
        print('Erro na atualização da DESPESA:')
        print(' ')
        print(e)
        popup_erro()

def dotacao():
    try:
        print(' ')
        print('Atualizando DOTACAO...')
        print(' ')

        try:
            url = f'http://extrator.sefaz.al.gov.br/DESPESAS/COMPARATIVO-DOTACOES/CONSOLIDADO/comparativo_dotacao_despesa_consolidado_2018-2024_siafe_gerado_em_{data_ontem}.csv'
            #df_dot = pd.read_csv(url, sep=';', encoding='latin2')
        except: 
            url = f'http://extrator.sefaz.al.gov.br/DESPESAS/COMPARATIVO-DOTACOES/CONSOLIDADO/comparativo_dotacao_despesa_consolidado_2018-2024_siafe_gerado_em_{data_atual}.csv'
            #df_dot = pd.read_csv(url, sep=';', encoding='latin2')
        
        base_dotacao_arquivo = os.path.join('Z:/DADOS/EXTRATOR/DOTACOES/', f'base_dotacao_18_24.csv')
        
        if os.path.exists(base_dotacao_arquivo):
            os.remove(base_dotacao_arquivo)
        
        wget.download(url, base_dotacao_arquivo)

        #df_dot.to_csv(base_dotacao_arquivo, index=False, encoding='latin2', decimal=',')
    except Exception as e:
        print(' ') 
        print('Erro na atualização da DOTACAO:')
        print(' ')
        print(e)
        popup_erro()

def receita():
    try:
        print(' ')
        print('Atualizando RECEITA...')
        print(' ')
        

        try:
            url = f'http://extrator.sefaz.al.gov.br/RECEITAS/receita_consolidado_2018_2024_siafe_gerado_em_{data_ontem}.xlsx'

            base_receita_arquivo = os.path.join('Z:/DADOS/EXTRATOR/RECEITA_GERADA/', f'receitas_base_18a24.xlsx')

            if os.path.exists(base_receita_arquivo):
                os.remove(base_receita_arquivo)
            
            #wget.download(url, base_receita_arquivo)

            df = pd.read_excel(url)
            
        except: 
            url = f'http://extrator.sefaz.al.gov.br/RECEITAS/receita_consolidado_2018_2024_siafe_gerado_em_{data_atual}.xlsx'
            
            base_receita_arquivo = os.path.join('Z:/DADOS/EXTRATOR/RECEITA_GERADA/', f'receitas_base_18a24.xlsx')
            
            if os.path.exists(base_receita_arquivo):
                os.remove(base_receita_arquivo)
            
            # Baixar o arquivo
            #wget.download(url, base_receita_arquivo)
           
            # Baixar o arquivo
            df = pd.read_excel(url)
        
        df.to_excel('Z:/DADOS/EXTRATOR/RECEITA_GERADA/receitas_base_18a24.xlsx', index=False, engine='openpyxl')

        print(' ')
        print('ATUALIZACAO CONCLUIDA!')
        print(' ')
        popup_concluido()


    except Exception as e:
        print(' ') 
        print('Erro na atualização da RECEITA:')
        print(' ')
        print(e)
        popup_erro()


despesa()
dotacao()
receita()