import os
import time
from datetime import datetime
import pandas as pd
from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side
import config


# Função para configurar o driver
def configurar_driver():
    edge_options = Options()
    edge_service = EdgeService(executable_path=config.EDGE_DRIVER_PATH)
    driver = webdriver.Edge(service=edge_service, options=edge_options)
    driver.maximize_window()
    return driver

# Função para realizar o login
def realizar_login(driver, username, password):
    driver.get('https://suporte.santamarcelinacultura.org.br/planejamento/front/central.php')
    driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div/form/div/div[1]/div[2]/input').send_keys(username)
    driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div/form/div/div[1]/div[3]/input').send_keys(password)
    driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div/form/div/div[1]/div[5]/button').click()
    time.sleep(3)

# Função para navegação e ações no painel
def navegar_painel(driver, username, password):
    driver.find_element(By.XPATH, '/html/body/div[2]/header/div/div[2]/ul/li[1]/a/span').click()
    time.sleep(2)
    driver.find_element(By.XPATH, '/html/body/div[2]/header/div/div[2]/ul/li[1]/div/div/div[2]/a[5]').click()
    time.sleep(2)
    driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div/main/table/tbody/tr[2]/td/a/button').click()
    time.sleep(3)
    driver.find_element(By.XPATH, '/html/body/div/div/div[2]/form/div[1]/input').send_keys(username)
    driver.find_element(By.XPATH, '/html/body/div/div/div[2]/form/div[2]/input').send_keys(password)
    driver.find_element(By.XPATH, '/html/body/div/div/div[2]/form/button').click()
    time.sleep(3)

# Função para baixar e formatar o arquivo de indicadores
def baixar_formatar_arquivo_indicadores(username, password):
    driver = configurar_driver()
    realizar_login(driver, username, password)
    navegar_painel(driver, username, password)
    
    # Clique nos botões para baixar o arquivo
    driver.find_element(By.XPATH, '/html/body/nav/button[3]').click()
    time.sleep(5)
    driver.find_element(By.XPATH, '/html/body/div[3]/button[4]').click()
    
    # Esperar o download completar
    time.sleep(10)
    driver.quit()

    # Pasta de downloads (padrão)
    download_dir = os.path.join(os.path.expanduser('~'), 'Downloads')
    arquivos = os.listdir(download_dir)

    # Encontrar o arquivo mais recente
    arquivos = [os.path.join(download_dir, f) for f in arquivos if f.endswith('.csv')]
    if arquivos:
        arquivo_recente = max(arquivos, key=os.path.getctime)

        # Caminho para salvar o arquivo Excel
        data_atual = datetime.now().strftime("%d.%m")
        destino = f"H:/Monitoramento_e_Avaliacao/Relatórios de Metas/Mensal/Realizado/2024/Indicadores - {data_atual}.xlsx"
        
        # Carregar o CSV em um DataFrame
        df = pd.read_csv(arquivo_recente, delimiter=";")

        # Salvar como Excel
        df.to_excel(destino, index=False)
        
        # Aplicar formatação usando openpyxl
        workbook = load_workbook(destino)
        sheet = workbook.active
        
        # Formatar cabeçalho
        font_bold = Font(bold=True)
        fill_blue = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
        for cell in sheet[1]:
            cell.font = font_bold
            cell.fill = fill_blue
        
        # Adicionar bordas
        thin_border = Border(
            left=Side(border_style="thin", color="000000"),
            right=Side(border_style="thin", color="000000"),
            top=Side(border_style="thin", color="000000"),
            bottom=Side(border_style="thin", color="000000")
        )
        for row in sheet.iter_rows():
            for cell in row:
                cell.border = thin_border
        
        # Definir altura das linhas e largura das colunas
        for row in sheet.iter_rows():
            sheet.row_dimensions[row[0].row].height = 20

        # Ajustar largura de todas as colunas para 30
        for col in sheet.columns:
            column = col[0].column_letter  # Coluna atual
            sheet.column_dimensions[column].width = 30
        
        # Salvar o arquivo formatado
        workbook.save(destino)
        workbook.close()
        
        print(f"Arquivo renomeado, formatado e movido para: {destino}")
    else:
        print("Nenhum arquivo CSV encontrado na pasta de Downloads.")

if __name__ == "__main__":
    username = config.USERNAME
    password = config.PASSWORD
    baixar_formatar_arquivo_indicadores(username, password)
