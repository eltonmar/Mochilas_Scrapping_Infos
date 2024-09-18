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

        # Encontra a seção de especificações
        especificacoes = soup.find_all('div', class_='vtex-flex-layout-0-x-flexRow vtex-flex-layout-0-x-flexRow--productSpecification')

        # Inicializa a variável de material
        material = None

        # Itera sobre as especificações para encontrar o material
        for especificacao in especificacoes:
            nome = especificacao.find('span', class_='vtex-product-specifications-1-x-specificationName')
            valor = especificacao.find('span', class_='vtex-product-specifications-1-x-specificationValue')

            if nome and valor:
                if nome.text.strip() == 'Cor':
                    material = valor.text.strip()
                    break

        return material if material else 'Cor não encontrado'
    except Exception as e:
        print(f'Erro ao acessar {url}: {e}')
        return None

# Carregar o arquivo Excel com a lista de sites
excel_file = "C:\\Users\\BG-PROVISORIO\\Desktop\\Vendas\\Mochilas\\Kipling.xlsx"
df = pd.read_excel(excel_file, sheet_name='Sheet1')  # Altere o nome da aba conforme necessário

# Cria uma lista para armazenar as informações extraídas
informacoes = []

# Itera sobre cada site na coluna
for url in df['Summary_URL']:  # Substitua 'Coluna_com_sites' pelo nome da coluna que contém os sites
    informacao = extrair_informacao(url)
    informacoes.append(informacao)

# Cria um novo DataFrame com as informações extraídas
df_novo = pd.DataFrame({'Cor': informacoes})

# Adiciona a nova aba ao arquivo Excel existente
with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a') as writer:
    df_novo.to_excel(writer, sheet_name='Material_Informacoes', index=False)

print('Informações de Material extraídas e salvas com sucesso!')
