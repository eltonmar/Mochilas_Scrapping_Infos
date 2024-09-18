
import pandas as pd
import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook

# Função para extrair a informação específica de uma página
def extrair_informacao(url):
    try:
        response = requests.get(url)
        response.raise_for_status()  # Verifica se a requisição foi bem-sucedida
        soup = BeautifulSoup(response.text, 'html.parser')

        # Localiza o conteúdo da seção de especificação de materiais
        especificacao = soup.find_all('div', class_='accordion__item aud-border-b aud-overflow-hidden')

        # Verifica se há pelo menos 2 itens na lista
        if len(especificacao) >= 2:
            segunda_especificacao = especificacao[1]  # Pega o segundo elemento
            # Extraia o texto ou outra informação do segundo elemento
            informacao = segunda_especificacao.text.strip()
            return informacao
        else:
            print(f'Seção de especificação de materiais não encontrada ou insuficiente em {url}')
            return None

    except Exception as e:
        print(f'Erro ao acessar {url}: {e}')
        return None

# Carregar o arquivo Excel com a lista de sites
excel_file = "C:\\Users\\BG-PROVISORIO\\Desktop\\Vendas\\Mochilas\\Fjällräven.xlsx"
df = pd.read_excel(excel_file, sheet_name='Sheet1')  # Altere o nome da aba conforme necessário

# Cria uma lista para armazenar as informações extraídas
informacoes = []

# Itera sobre cada site na coluna
for url in df['Title_URL']:  # Substitua 'Title_URL' pelo nome da coluna que contém os sites
    informacao = extrair_informacao(url)
    informacoes.append(informacao)

# Cria um novo DataFrame com as informações extraídas
df_novo = pd.DataFrame({'Informacao': informacoes})

# Adiciona a nova aba ao arquivo Excel existente
with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a') as writer:
    df_novo.to_excel(writer, sheet_name='Nova_Aba', index=False)

print('Informações extraídas e salvas com sucesso!')
