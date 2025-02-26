import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.options import Options
import time
import pandas as pd
from datetime import datetime

email = '***'
senha = '***'

# chrome_options = Options()
# chrome_options.add_argument("--headless")  # Executa em segundo plano (sem interface gráfica)
# chrome_options.add_argument("--disable-gpu")  # Pode ser útil em alguns ambientes
# chrome_options.add_argument("--window-size=1920,1080")  # Define o tamanho da janela virtual

# Inicializa o navegador
navegador = webdriver.Chrome()

# Acessa a página de login
navegador.get('https://app.solarz.com.br/login?logout')
campo_usuario = navegador.find_element(By.ID, 'username')
campo_senha = navegador.find_element(By.ID, "password")

campo_usuario.send_keys(email)
campo_senha.send_keys(senha)

botao_login = navegador.find_element(By.XPATH, '//*[@id="form_login"]/input[4]')
botao_login.click()
time.sleep(5)

# Muda para o iframe onde a tabela está
iframe = navegador.find_element(By.TAG_NAME, "iframe")
navegador.switch_to.frame(iframe)

# Captura o número total de páginas e converte para inteiro
campo_n_paginas = navegador.find_element(By.XPATH, '//*[@id="__next"]/div/div/div/div/div[2]/section/section/main/div[3]/div/div/div/div/ul/li[9]/a')
valor_n_paginas = int(campo_n_paginas.text)
print("Número de páginas:", valor_n_paginas)

# Lista para armazenar os DataFrames de cada página
lista_df = []

# Loop para percorrer cada página
for pagina in range(valor_n_paginas):
    print(f"Extraindo dados da página {pagina + 1} de {valor_n_paginas}...")
    
    # Aguarda 5 segundos para garantir que os dados foram carregados
    time.sleep(5)
    
    # Extrai todas as linhas da tabela
    linhas = navegador.find_elements(By.TAG_NAME, "tr")
    
    # Lista para armazenar os dados desta página
    dados = []
    for linha in linhas:
        celulas = linha.find_elements(By.TAG_NAME, "td")
        linha_dados = [celula.text for celula in celulas]
        if linha_dados:
            dados.append(linha_dados)
    
    # Se dados foram extraídos, processa o DataFrame da página
    if dados:
        df = pd.DataFrame(dados)
        # Renomeia as colunas conforme o mapeamento definido
        df = df.rename(columns={0: "indice", 1: "string_nome", 2: "Responsável", 3: "1D", 4: "15D", 5: "30D", 6: "365D", 7: "Avisos"})
        
        # Separa os dados da coluna 'string_nome' em novas colunas
        df_separado = df["string_nome"].str.split('\n', expand=True)
        
        # Junta o DataFrame separado com o original
        df_final_pagina = df_separado.join(df)
        
        # Renomeia as colunas resultantes da divisão (ajuste conforme necessário)
        df_final_pagina = df_final_pagina.rename(columns={0: "Cliente", 1: "Potência", 2: "SA", 3: "Tag"})
        
        # Remove a coluna original que foi dividida
        df_final_pagina = df_final_pagina.drop("string_nome", axis=1)
        
        # Adiciona o DataFrame desta página à lista
        lista_df.append(df_final_pagina)
    else:
        print(f"Nenhum dado encontrado na página {pagina + 1}.")
    
    # Se não for a última página, tenta clicar no botão de próxima página para avançar
    if pagina < valor_n_paginas - 1:
        xpath_options = [
            '//*[@id="__next"]/div/div/div/div/div[2]/section/section/main/div[3]/div/div/div/div/ul/li[10]/button',
            '//*[@id="__next"]/div/div/div/div/div[2]/section/section/main/div[3]/div/div/div/div/ul/li[11]/button',
            '//*[@id="__next"]/div/div/div/div/div[2]/section/section/main/div[3]/div/div/div/div/ul/li[12]/button'
        ]
        botao_prox_pagina = None
        
        for xpath in xpath_options:
            try:
                botao_prox_pagina = navegador.find_element(By.XPATH, xpath)
                if botao_prox_pagina:
                    botao_prox_pagina.click()
                    break
            except NoSuchElementException:
                continue
        
        if botao_prox_pagina is None:
            print(f"Botão de próxima página não encontrado na página {pagina + 1}. Encerrando o loop.")
            break

# Fecha o navegador após a extração de todas as páginas
navegador.quit()

# Concatena todos os DataFrames extraídos em um único DataFrame final
df_final = pd.concat(lista_df, ignore_index=True)

df_final = df_final.dropna(subset='Potência',how='all')
df_final = df_final.drop(['SA','indice','Avisos'],axis=1)

data_atual = datetime.now().strftime("%d-%m-%Y")
nome_arquivo = f"tabela-solarz-{data_atual}.xlsx"

# Salva o DataFrame no arquivo com o nome gerado
df_final.to_excel(nome_arquivo, index=False)

# Exibe o DataFrame final com os dados de todas as páginas
#display(df_final)
