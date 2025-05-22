import os
import pandas as pd
import pyautogui
import time
from pywinauto import Application

# Caminho para a pasta de downloads
download_folder = os.path.expanduser('I:\\.shortcut-targets-by-id\\1BbEijfOOPBwgJuz8LJhqn9OtOIAaEdeO\\Logdi\\Relatório e Dashboards\\01.Análise de Resultados\\download_relatorios')

# Caminho para a pasta onde a planilha será salva
planilha_folder = os.path.expanduser('I:\\.shortcut-targets-by-id\\1BbEijfOOPBwgJuz8LJhqn9OtOIAaEdeO\\Logdi\\Relatório e Dashboards\\01.Análise de Resultados')

# Função para encontrar o último arquivo .sswweb
def encontrar_ultimo_arquivo_swwweb(pasta):
    arquivos = [os.path.join(pasta, f) for f in os.listdir(pasta) if os.path.isfile(os.path.join(pasta, f))]
    arquivos = [f for f in arquivos if f.endswith('.sswweb')]
    return max(arquivos, key=os.path.getctime) if arquivos else None

# Função para processar o arquivo .sswweb
def processar_arquivo_swwweb(download_folder, planilha_destino):
    ultimo_arquivo = encontrar_ultimo_arquivo_swwweb(download_folder)

    if ultimo_arquivo:
        print(f"Último arquivo encontrado: {ultimo_arquivo}")

        # Lê o conteúdo do arquivo
        with open(ultimo_arquivo, 'r') as file:
            linhas = file.readlines()

        # Remove as duas primeiras linhas e a última linha
        linhas = linhas[2:-1]

        # Grava o conteúdo modificado de volta no arquivo
        with open(ultimo_arquivo, 'w') as file:
            file.writelines(linhas)

        # Abre o arquivo no Excel
        os.startfile(ultimo_arquivo)  # Força a abertura no Excel
        time.sleep(15)  # Aumenta o tempo de espera para garantir que o Excel esteja aberto

        # Foca na janela do Excel usando pywinauto
        app = Application().connect(title_re=".*Excel.*")  # Conecta à janela do Excel
        excel_window = app.top_window()  # Obtém a janela principal do Excel
        excel_window.set_focus()  # Define o foco na janela do Excel

        # Aguarda um tempo para garantir que o Excel esteja em foco
        time.sleep(1)

        # Seleciona todas as células usando pywinauto
        excel_window.type_keys('^{HOME}', with_spaces=True)  # Move para a célula A1
        time.sleep(1)
        excel_window.type_keys('^+{DOWN}')  # Seleciona até a última linha preenchida
        time.sleep(1)
        for _ in range(6):
            excel_window.type_keys('^+{RIGHT}')
        time.sleep(1)
        # Copia todos os dados restantes do Excel
        pyautogui.hotkey('ctrl', 'c')  # Copia os dados

        # Aguarda um tempo para garantir que os dados sejam copiados
        time.sleep(1)
        # Fecha a planilha do Excel
        pyautogui.hotkey('alt', 'f4')
        time.sleep(1)
        pyautogui.hotkey('enter')
        time.sleep(5)

        # Abre a planilha de destino no Excel
        if not os.path.exists(planilha_destino):
            # Cria um DataFrame vazio e salva como um novo arquivo Excel
            pd.DataFrame().to_excel(planilha_destino, index=False)
            print(f"Planilha criada: {planilha_destino}")

        os.startfile(planilha_destino)
        time.sleep(10)  # Aumenta o tempo de espera para garantir que o Excel esteja aberto

                # Foca na janela do Excel usando pywinauto
        app = Application().connect(title_re=".*Excel.*")  # Conecta à janela do Excel
        excel_window = app.top_window()  # Obtém a janela principal do Excel
        excel_window.set_focus()  # Define o foco na janela do Excel

        # Aguarda um tempo para garantir que o Excel esteja em foco
        time.sleep(1)

        # Lê a planilha de destino
        pd.read_excel(planilha_destino)

        #Seleciona a última linha preenchida
        excel_window.type_keys('^{HOME}', with_spaces=True)  # Move para a célula A1
        time.sleep(1)
        excel_window.type_keys('^{DOWN}')
        time.sleep(1)
        excel_window.type_keys('{DOWN}')
        time.sleep(0.5)

        # Cola os dados copiados na célula selecionada
        pyautogui.hotkey('ctrl', 'v')  # Cola os dados

        # Aguarda um tempo para garantir que os dados sejam colados
        time.sleep(1)

        pyautogui.hotkey('ctrl', 'b')
        time.sleep(15)
        pyautogui.hotkey('alt', 'f4')
        time.sleep(1)

        print(f"Dados copiados para a planilha: {planilha_destino}")
    else:
        print("Nenhum arquivo .sswweb encontrado na pasta de downloads.")

# Caminho para a planilha de destino
planilha_destino = os.path.join(planilha_folder, "DB_Analise_Resultados.xlsx")

# Bloco para garantir que o código só execute ao ser chamado
if __name__ == "__main__":
    # Chama a função para processar o arquivo
    processar_arquivo_swwweb(download_folder, planilha_destino)