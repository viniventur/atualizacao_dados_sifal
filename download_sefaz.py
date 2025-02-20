import requests
import os
import pandas as pd
from datetime import datetime
import time
import tqdm
from plyer import notification
import wget

# Datas atual e de ontem
data_atual = datetime.now().strftime("%d-%m-%Y")
data_ontem = (datetime.now() - pd.Timedelta(days=1)).strftime("%d-%m-%Y")
data_execucao = datetime.now().strftime("%d/%m/%Y - %H:%M")

# DF com o nome dos arquivos e os links
df = pd.read_excel('dados_extrator_links.xlsx')

# Diretório de salvamento
save_dir = r"O:\DADOS\EXTRATOR\CONSOLIDADAS"
os.makedirs(save_dir, exist_ok=True)

# Diretório de log
log_file = os.path.join(save_dir, "log_execucao.txt")

# Escreve log no arquivo
def escrever_log(mensagem):
    with open(log_file, "a") as log:
        log.write(f"{data_execucao}: {mensagem}\n")  

# Função para formatar o tempo de execução
def formatar_tempo(segundos):
    if segundos < 60:
        return f"{segundos:.2f} segundos"
    elif segundos < 3600:
        minutos = segundos / 60
        return f"{minutos:.2f} minutos"
    else:
        horas = segundos / 3600
        return f"{horas:.2f} horas"  

# Função para baixar e salvar arquivos usando wget
def download(arquivo, url, save_path):
    try:
        wget.download(url, save_path)
        print(f"\nArquivo {arquivo} salvo em: {save_path}")
        return True, None
    except Exception as e:
        return False, f"Erro ao baixar {arquivo}: {e}"

# Download dos arquivos indicados no DF
try:
    start_time_geral = time.time()

    # Iterando sobre o DataFrame
    for _, dado in df.iterrows():
        start_time = time.time()
        nome_arquivo = dado["dado"]
        base_url = dado["link"]

        # Construindo URLs com as datas
        url_ontem = f"{base_url}{data_ontem}.csv"
        url_atual = f"{base_url}{data_atual}.csv"

        # Caminho de salvamento
        save_path = os.path.join(save_dir, f"{nome_arquivo}.csv")

        # Verifica se o arquivo já existe e o apaga
        if os.path.exists(save_path):
            os.remove(save_path)

        # Tentativa de download
        try:
            print(f"Tentando baixar: {nome_arquivo}...")
            sucesso, erro = download(nome_arquivo, url_ontem, save_path)
            if sucesso:
                exec_time = formatar_tempo(time.time() - start_time)
                mensagem = f"Arquivo {nome_arquivo} salvo com sucesso. Tempo de execucao: {exec_time}."
                escrever_log(mensagem)
            else:
                print(f"Tentativa com a data de ontem falhou. Tentando com a data atual...")
                sucesso, erro = download(nome_arquivo, url_atual, save_path)
                if sucesso:
                    exec_time = formatar_tempo(time.time() - start_time)
                    mensagem = f"Arquivo {nome_arquivo} salvo com sucesso. Tempo de execucao: {exec_time}."
                    escrever_log(mensagem)
                else:
                    mensagem = f"Falha ao salvar o arquivo {nome_arquivo}: {erro}."
                    escrever_log(mensagem)

        except Exception as e:
            exec_time = formatar_tempo(time.time() - start_time)
            escrever_log(f"Erro inesperado ao processar {nome_arquivo}: {e}. Tempo de execucao: {exec_time}.")

    # Notificacao de sucesso 
    exec_time_geral = formatar_tempo(time.time() - start_time_geral)     
    notification.notify(
                        title="Atualização dos dados concluída!",
                        message=f'Tempo de execução: {exec_time_geral}.',
                        app_name="Extrator de Dados",
                        timeout=20
                        )
except Exception as e:
    exec_time_geral = formatar_tempo(time.time() - start_time_geral)
    # Notificacao de erro   
    notification.notify(
                        title=f'Atualização dos dados FALHOU: {e}',
                        message=f'Tempo de execução: {exec_time_geral}.',
                        app_name="Extrator de Dados",
                        timeout=20
                        )
