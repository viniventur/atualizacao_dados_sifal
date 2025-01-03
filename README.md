# Automações de Atualização de Dados de Contabilidade Pública :chart_with_upwards_trend:	
<div style="display: flex; align-items: center;">
    <img src="https://img.shields.io/badge/Python-FFD43B?style=for-the-badge&logo=python&logoColor=blue
    " alt="Python" width="120" style="margin-right: 10px;">
</div>

# Descrição do projeto :clipboard:	

Este repositório contém scripts de automação de atualização e download de arquivos de dados relacionados à contabilidade pública do Estado de Alagoas. O objetivo principal é facilitar o processo de obtenção de dados atualizados, garantindo eficiência e confiabilidade no gerenciamento de informações.

### Exclusividade dos arquivos devido a infovia :heavy_exclamation_mark:

**Os links de download só são funcionais se o usuário estiver conectado a infovia do Estado de Alagoas**, portanto, *não é possível acessar os dados publicamente*.

## Funcionalidades :hammer_and_pick:

- **Download Automatizado**: Realiza o download de arquivos a partir de URLs fornecidas em um arquivo Excel.
- **Compatibilidade com Datas**: Constrói links de download dinamicamente, considerando as datas de ontem e hoje devido a variações nas atualização dos dados. Usualmente, a data referente a atualização é a do dia anterior, porém, pode acontecer do dado ter sido atualizado no dia da execução.
- **Logs Detalhados**: Registra os resultados de cada operação de download, incluindo o tempo de execução e erros, caso ocorram.
- **Notificações de Conclusão**: Exibe um pop-up no final da execução com o resumo das últimas operações realizadas para aviso ao usuário.

## Estrutura do Projeto :scroll:

- **`dados_extrator_links.xlsx`**: Arquivo contendo os nomes e links dos arquivos a serem baixados (não contido atualmente no repositório, embora outros exemplos podem ser acessados em `antigos/`).
- **`log_execucao.txt`**: Arquivo de log que registra os detalhes de cada execução (contido na pasta final de download dos arquivos).
- **`extrator_dados_sefaz.py`**: Script principal para realizar os downloads e registrar logs.

## Pré-requisitos :gear:

1. **Python 3.7+**
2. Instalar as bibliotecas necessárias:
   ```bash
   pip install -r requirements.txt
   ```

## Como Usar :open_book:

1. Certifique-se de que o arquivo `dados_extrator_links.xlsx` está no mesmo diretório do script. Este arquivo deve conter duas colunas:
   - `dado`: Nome do arquivo a ser salvo.
   - `link`: URL base do arquivo.

2. Atualize as configurações de diretório no script, se necessário:
   - O diretório de salvamento padrão é `Z:\DADOS\EXTRATOR\CONSOLIDADAS`.

3. Execute o script:
   ```bash
   python extrator_dados_sefaz.py
   ```

4. Verifique os logs em `log_execucao.txt` para os detalhes da execução.

5. Ao final da execução, será exibida uma notificação no sistema com o resumo das últimas operações.

## Detalhes Técnicos :mag_right:

- **Geração de Links**:
  Os links de download são construídos dinamicamente a partir do arquivo Excel, adicionando as datas de ontem e hoje no final de cada URL.

- **Logs**:
  Cada operação é registrada com detalhes como:
  - Sucesso ou falha.
  - Tempo de execução (em segundos, minutos ou horas, dependendo da duração).

- **Barra de Progresso**:
  Durante o download, uma barra de progresso é exibida para indicar o andamento de cada arquivo (barra nativa do `wget`, porém pode ser utilizada a biblioteca `tqdm` para criação de barras de progresso).

## Licença :bookmark_tabs:	

Este projeto está licenciado sob a [Licença MIT](LICENSE).

