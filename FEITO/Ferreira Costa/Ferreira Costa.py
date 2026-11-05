from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import NoSuchElementException, TimeoutException, ElementClickInterceptedException, WebDriverException
from datetime import date
import pandas as pd
import time
import sys

# Configurações do Chrome
chrome_options = Options()
chrome_options.add_argument('--headless')
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--disable-dev-shm-usage')

# Inicializar o WebDriver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

# Função para navegar até uma página específica
def navegar_para_pagina(url):
    driver.get(url)
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'body'))
        )
        print(f"Conexão bem-sucedida com {url}")
    except TimeoutException:
        print(f"Falha ao acessar {url}")
        return False
    return True

# Função para rolar até o final da página
def rolar_ate_final():
    for _ in range(15):  # Tentar rolar a página várias vezes
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(5)  # Aumentar o tempo de espera para garantir que os produtos sejam carregados

# Função para avançar para a próxima página
def avancar_pagina():
    try:
        botao_avancar = driver.find_element(By.CSS_SELECTOR, 'button[data-cy="paginator-forward-button"]')
        driver.execute_script("arguments[0].scrollIntoView(true);", botao_avancar)
        time.sleep(1)  # Pequena espera para garantir que o botão esteja visível
        botao_avancar.click()
        time.sleep(5)  # Aumentar o tempo de espera para garantir que a próxima página seja carregada
    except ElementClickInterceptedException:
        print("O clique no botão de avançar foi interceptado por outro elemento.")
        return False
    except NoSuchElementException:
        return False
    return True

# URL da página que você deseja fazer scraping
url = 'https://www.ferreiracosta.com/Categoria/pisos-e-revestimentos'

produtos_dados = []

try:
    pagina_atual = 1
    while True:
        if not navegar_para_pagina(url):
            break

        # Rolagem para carregar os produtos
        rolar_ate_final()

        # Analise a página com BeautifulSoup
        page_source = driver.page_source
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(page_source, 'html.parser')
        produtos = soup.find_all('div', {'data-cy': lambda value: value and value.startswith('card-container')})

        for produto in produtos:
            try:
                nome_empresa = "FerreiraCosta"
                data_coleta = date.today().strftime("%Y-%m-%d")
                codigo_produto = produto.find('h3', {'data-cy': 'product-card-title'}).get_text(strip=True) if produto.find('h3', {'data-cy': 'product-card-title'}) else "N/A"
                imagem_url = produto.find('img', {'data-cy': 'product-card-dynamic-image'})['src'] if produto.find('img', {'data-cy': 'product-card-dynamic-image'}) else "N/A"
                nome_produto = produto.find('h3', {'data-cy': 'product-card-title'}).get_text(strip=True) if produto.find('h3', {'data-cy': 'product-card-title'}) else "N/A"
                moeda = produto.find('span', {'data-cy': 'product-card-current-price'}).get_text(strip=True) if produto.find('span', {'data-cy': 'product-card-current-price'}) else "N/A"
                url_produto = produto.find('a', {'data-cy': 'product-card-container-as-component-link'})['href'] if produto.find('a', {'data-cy': 'product-card-container-as-component-link'}) else "N/A"

                # Montar a linha de dados do produto
                linha_produto = [nome_empresa, data_coleta, codigo_produto, imagem_url, nome_produto, moeda, url_produto]
                produtos_dados.append(linha_produto)
            except Exception as e:
                print(f"Erro ao processar produto: {e}")
                continue

        print(f"Produtos coletados na página {pagina_atual}: {len(produtos)}")

        if not avancar_pagina():
            print(f"Não foi possível avançar para a próxima página a partir da página {pagina_atual}.")
            break

        pagina_atual += 1

except WebDriverException as e:
    print(f"\033[91mO código encontrou um erro:\033[0m {e}")

finally:
    # Converter os dados coletados em um DataFrame do pandas
    df = pd.DataFrame(produtos_dados, columns=["Nome Empresa", "Data de Coleta", "Código do Produto", "URL da Imagem", "Nome do Produto", "Moeda", "URL do Produto"])

# Salvar o DataFrame em um arquivo Excel 
df.to_excel("produtos_coletados.xlsx", index=False) 
print("Os dados dos produtos foram salvos em 'produtos_coletados.xlsx'") 

driver.quit()