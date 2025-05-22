from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime, timedelta
import pandas as pd
import os
from dotenv import load_dotenv
import time

# Importa as funções do arquivo alimentacao_planilha
from subdiretorio.alimentacao_planilha import encontrar_ultimo_arquivo_swwweb, processar_arquivo_swwweb

# Caminho para a pasta de downloads desejada
download_folder = os.path.expanduser('I:\\.shortcut-targets-by-id\\1BbEijfOOPBwgJuz8LJhqn9OtOIAaEdeO\\Logdi\\Relatório e Dashboards\\01.Análise de Resultados\\download_relatorios')

# Caminho para a pasta onde a planilha será salva
planilha_folder = os.path.expanduser('I:\\.shortcut-targets-by-id\\1BbEijfOOPBwgJuz8LJhqn9OtOIAaEdeO\\Logdi\\Relatório e Dashboards\\01.Análise de Resultados')

# Caminho para a planilha de destino
planilha_destino = os.path.join(planilha_folder, "DB_Analise_Resultados.xlsx")

# Carrega as variáveis de ambiente do arquivo .env
load_dotenv("credenciais.env")

# Configurações do Edge
edge_options = Options()
edge_options.add_experimental_option('prefs', {
    "download.default_directory": download_folder,  # Define o diretório de download
    "download.prompt_for_download": False,           # Não solicitar confirmação para download
    "download.directory_upgrade": True,               # Permitir a atualização do diretório
    "safebrowsing.enabled": True                       # Habilitar navegação segura
})

def realizar_login(driver):
    # Navega até a página do formulário
    driver.get("https://sistema.ssw.inf.br/bin/ssw0422")  # Substitua pela URL do seu formulário

    # Atraso para garantir que a página carregue completamente
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "f1")))

    # Preenche os campos de login
    driver.find_element(By.NAME, "f1").send_keys(os.getenv("SSW_EMPRESA"))
    driver.find_element(By.NAME, "f2").send_keys(os.getenv("SSW_CNPJ"))
    driver.find_element(By.NAME, "f3").send_keys(os.getenv("SSW_USUARIO"))
    driver.find_element(By.NAME, "f4").send_keys(os.getenv("SSW_SENHA"))

    # Clica no botão de login diretamente
    login_button = driver.find_element(By.ID, "5")
    driver.execute_script("arguments[0].click();", login_button)
    time.sleep(5)

def preencher_formulario(driver):
    # Preenche os campos de Unidade e Opção
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "f2")))
    driver.find_element(By.NAME, "f3").send_keys("455")

    time.sleep(1)
    abas = driver.window_handles  # Lista o número de abas abertas.
    driver.switch_to.window(abas[-1])  # Muda o foco para a última aba (a nova aba)
    
    data_anterior = (datetime.now() - timedelta(days=1)).strftime('%d%m%y')

    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "f12")))
    input_element = driver.find_element(By.ID, "12")
    driver.execute_script("arguments[0].value = '';", input_element)
    time.sleep(0.3)
    driver.find_element(By.ID, "11").clear()
    time.sleep(0.3)
    driver.find_element(By.ID, "11").send_keys(data_anterior)
    time.sleep(0.3)
    driver.find_element(By.ID, "12").clear()
    time.sleep(0.3)
    driver.find_element(By.ID, "12").send_keys(data_anterior)
    time.sleep(0.3)
    driver.find_element(By.NAME, "f21").clear()
    time.sleep(0.3)
    driver.find_element(By.NAME, "f21").send_keys("t")
    time.sleep(0.3)
    driver.find_element(By.NAME, "f35").clear()
    time.sleep(0.3)
    driver.find_element(By.NAME, "f35").send_keys("e")
    time.sleep(0.3)
    input_element = driver.find_element(By.NAME, "f37")
    driver.execute_script("arguments[0].value = '';", input_element)
    driver.find_element(By.NAME, "f37").send_keys("g")
    time.sleep(0.3)
    driver.find_element(By.NAME, "f38").send_keys("h")

    login_button = driver.find_element(By.ID, "40")
    driver.execute_script("arguments[0].click();", login_button)
    time.sleep(0.8)

    # Usando ActionChains para enviar a tecla "1"
    actions = ActionChains(driver)
    actions.send_keys("1").perform()

    time.sleep(5)
    abas = driver.window_handles  # Lista o número de abas abertas.

    # Muda o foco para a última aba (a nova aba)
    driver.switch_to.window(abas[-1])
    time.sleep(2)

def capturar_seq(driver):
    # Captura o valor do seq da segunda linha da primeira coluna da tabela com o ID 'tblsr'
    tabela = driver.find_element(By.ID, "tblsr")
    linhas = tabela.find_elements(By.TAG_NAME, "tr")

    # Verifica se há pelo menos duas linhas (cabeçalho + dados)
    if len(linhas) > 1:
        seq_da_requisicao = linhas[1].find_element(By.TAG_NAME, "td").text  # Captura o valor da primeira coluna da segunda linha
        print(f"Seq da requisição: {seq_da_requisicao}")
        return seq_da_requisicao
    else:
        print("Não há linhas suficientes na tabela para capturar o seq.")
        return None

def atualizar_relatorio(driver, seq_da_requisicao):
    try:
        # Clica no botão de atualização
        update_button = driver.find_element(By.ID, "2")
        time.sleep(60)
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "2")))
        driver.execute_script("arguments[0].click();", update_button)
        time.sleep(5)
    except Exception as e:
        print("Botão de atualização não encontrado. Finalizando a automação.")
        driver.quit()
        return False  # Retorna False para indicar falha

    # Após a atualização, busca novamente a linha correspondente ao seq
    relatorios_atualizados = driver.find_elements(By.CSS_SELECTOR, "table#tblsr tr")
    relatorio_encontrado = None

    for relatorio in relatorios_atualizados[1:]:  # Ignora o cabeçalho da tabela
        seq_atual = relatorio.find_element(By.TAG_NAME, "td").text  # Captura o seq da primeira coluna
        if seq_atual == seq_da_requisicao:
            relatorio_encontrado = relatorio
            break

    # Se um relatório correspondente foi encontrado, clique no link (substituído por <u>) na mesma linha
    if relatorio_encontrado:
        link = relatorio_encontrado.find_element(By.TAG_NAME, "u")  # Supondo que o link esteja dentro de uma tag <u>
        driver.execute_script("arguments[0].click();", link)
        print("Clicou no link da requisição correspondente.")
    else:
        print("Nenhum relatório encontrado após a atualização.")
    return True  # Retorna True se tudo ocorrer bem

def main():
    # Inicializa o WebDriver com as novas configurações
    driver = webdriver.Edge(options=edge_options)

    try:
        realizar_login(driver)
        preencher_formulario(driver)

        seq_da_requisicao = capturar_seq(driver)

        if seq_da_requisicao:
            atualizar_relatorio(driver, seq_da_requisicao)

            time.sleep(10)  # Atraso para garantir que o download seja iniciado

            processar_arquivo_swwweb(download_folder, planilha_destino)

    except Exception as e:
        print(f"Ocorreu um erro: {e}")

    finally:
        # Fecha o navegador
        driver.quit()


if __name__ == "__main__":
    main()