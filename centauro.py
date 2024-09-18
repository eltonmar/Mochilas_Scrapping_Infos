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

        # Encontra a tabela de atributos
        tabela_atributos = soup.find('table', class_='AttributesTable-styled__Table-sc-5148fd8-0')

        # Inicializa a variável de material
        material = None

        # Verifica se a tabela foi encontrada
        if tabela_atributos:
            linhas = tabela_atributos.find_all('tr')

            # Itera sobre as linhas da tabela para encontrar o material
            for linha in linhas:
                th = linha.find('th')  # Cabeçalho da tabela
                td = linha.find('td')  # Conteúdo da tabela

                if th and td and th.text.strip() == 'Material':
                    material = td.text.strip()
                    break

        return material if material else 'Material não encontrado'
    except Exception as e:
        print(f'Erro ao acessar {url}: {e}')
        return None

# Carregar o arquivo Excel com a lista de sites
excel_file = "C:\\Users\\BG-PROVISORIO\\Desktop\\Vendas\\Mochilas\\Mochilas - Centauro.xlsx"
df = pd.read_excel(excel_file, sheet_name='Sheet1')  # Altere o nome da aba conforme necessário

# Cria uma lista para armazenar as informações extraídas
informacoes = []

# Itera sobre cada site na coluna
for url in df['Title_URL']:  # Substitua 'Summary_URL' pelo nome da coluna que contém os sites
    informacao = extrair_informacao(url)
    informacoes.append(informacao)

# Cria um novo DataFrame com as informações extraídas
df_novo = pd.DataFrame({'Material': informacoes})

# Adiciona a nova aba ao arquivo Excel existente
with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a') as writer:
    df_novo.to_excel(writer, sheet_name='Material_Informacoes', index=False)

print('Informações de Material extraídas e salvas com sucesso!')
