import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# Caminho correto do arquivo Excel
nome = r"C:\Users\Gustavo\OneDrive\Área de Trabalho\Nova pasta (2)\Pasta1.xlsx"
Site = "https://portal.nubusnatal.com.br/TDMaxwebcommerce/"

# Leitura do arquivo Excel
df = pd.read_excel(nome)

# Inicialização do WebDriver do Chrome
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

# Acessa o site uma vez
driver.get(Site)

# Login no site
login_field = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, "//*[@id='ctl00_cpCont_login_UserName']"))
)
senha_field = driver.find_element(By.XPATH, "//*[@id='ctl00_cpCont_login_Password']")

login_field.send_keys("n02952485003830")  # Substitua pelo seu login
senha_field.send_keys("Dkja%3")  # Substitua pela sua senha

login_button = driver.find_element(By.XPATH, "//*[@id='ctl00_cpCont_login_LoginButton']")
login_button.click()

# Espera a página carregar
WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, "//*[@id='ctl00_UcMenu1_lbAssocicaoVT']"))
)

# Iteração sobre as linhas do DataFrame
for index, row in df.iloc[1:10].iterrows():
    print(f"Processando linha {index + 1}...")

    # Navega para a página de Associação de VT (se necessário)
    associacao_vt_button = driver.find_element(By.XPATH, "//*[@id='ctl00_UcMenu1_lbAssocicaoVT']")
    associacao_vt_button.click()

    # Espera o campo de cartão estar visível
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//*[@id='ctl00_cpCont_txtCartao']"))
    )

    # Preenche o campo de cartão
    numero_cartao = row['no_cartao']
    campo_cartao = driver.find_element(By.XPATH, "//*[@id='ctl00_cpCont_txtCartao']")
    campo_cartao.clear()  # Limpa o campo antes de preencher
    campo_cartao.send_keys(str(numero_cartao))

    # Clica no botão "Selecionar"
    selecionar_button = driver.find_element(By.XPATH, "//*[@id='ctl00_cpCont_btnLocalizar']")
    selecionar_button.click()

    # Seleciona o cartão
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//*[@id='ctl00_cpCont_gvCartoes_ctl02_LinkButton3']/span"))
    )
    selecionar_cartao_button = driver.find_element(By.XPATH, "//*[@id='ctl00_cpCont_gvCartoes_ctl02_LinkButton3']/span")
    selecionar_cartao_button.click()

    # Preenche o campo de usuário
    usuario = row['usuario']
    campo_usuario = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//*[@id='ctl00_cpCont_gvCartoes_ctl02_TextBox2']"))
    )
    campo_usuario.clear()  # Limpa o campo antes de preencher
    campo_usuario.send_keys(str(usuario))

    # Salva as alterações
    disquete_button = driver.find_element(By.XPATH, "//*[@id='ctl00_cpCont_gvCartoes_ctl02_ImageButton1']")
    disquete_button.click()

    print(f"Linha {index + 1} processada com sucesso!")

    # Aguarda brevemente antes de continuar para a próxima linha
    time.sleep(2)

# Fecha o navegador após completar
driver.quit()
