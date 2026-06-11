import pandas as pd
import re
import os
import gspread
from google.oauth2.service_account import Credentials
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime, timedelta
import time

# Diretório de download
download_dir = os.path.abspath(os.path.join(os.getcwd(), 'downloads'))
if not os.path.exists(download_dir):
    os.makedirs(download_dir)

def get_codes_from_gsheet(cliente, spreadsheet_id, worksheet_name, column_name='Cód.', header_row=1):
    """
    Read codes from Google Sheets using existing authenticated client
    header_row: 0-based index of the row to use as column headers (default 1 = row 2)
    """
    # Open spreadsheet
    spreadsheet = cliente.open_by_key(spreadsheet_id)
    
    # Get worksheet by name
    worksheet = spreadsheet.worksheet(worksheet_name)
    
    # Get all data
    data = worksheet.get_all_values()
    
    if data and len(data) > header_row:
        # Use specified row as header, data starts after that
        df = pd.DataFrame(data[header_row+1:], columns=data[header_row])
        
        # Extract codes from specified column
        if column_name in df.columns:
            codigos = df[column_name].dropna().astype(int).tolist()
        else:
            # Try case-insensitive match
            for col in df.columns:
                if col.lower() == column_name.lower():
                    codigos = df[col].dropna().astype(int).tolist()
                    break
            else:
                # If column not found, use first column
                codigos = df.iloc[:, 0].dropna().astype(int).tolist()
        
        return codigos
    return []

# === INITIALIZE GOOGLE SHEETS CLIENT FIRST ===
json_credenciais = 'creds.json'
escopo = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credenciais = ServiceAccountCredentials.from_json_keyfile_name(json_credenciais, escopo)
cliente = gspread.authorize(credenciais)

# === GET CODES FROM GOOGLE SHEET ===
spreadsheet_id = "14zK_cZOpVkDc2zzlsR1L3tFIfwwUpZ6QXbAKFjGwxAE"
worksheet_name = "Página1"
codigos = get_codes_from_gsheet(cliente, spreadsheet_id, worksheet_name, "Cód.", header_row=1)
print(f"Códigos carregados: {codigos}")
print(f"Total de códigos: {len(codigos)}")

# ChromeOptions configurado
chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--window-size=1920,1080")
chrome_options.add_argument("--allow-running-insecure-content")
chrome_options.add_argument("--ignore-certificate-errors")
chrome_options.add_argument("--unsafely-treat-insecure-origin-as-secure=http://drogcidade.ddns.net:4647")
prefs = {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "directory_upgrade": True,
    "safebrowsing.enabled": True,
}
chrome_options.add_experimental_option("prefs", prefs)

navegador = webdriver.Chrome(options=chrome_options)
wait = WebDriverWait(navegador, 100)

# Acessar sistema
navegador.get('http://drogcidade.ddns.net:4647/sgfpod1/Login.pod')
wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="id_cod_usuario"]'))).send_keys('237')
wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="nom_senha"]'))).send_keys('9939')
wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="login"]'))).click()

# BOTAO SNGPC
time.sleep(2.5)
navegador.find_element(By.TAG_NAME, "body").send_keys(Keys.F11)
time.sleep(2.5)
print('Pop-up SNGPC fechado com sucesso.', flush=True)

# Caminho até o relatório
wait.until(EC.element_to_be_clickable((By.ID, "sideMenuSearch")))
navegador.find_element(By.ID, "sideMenuSearch").send_keys("Vendas por Vendedor")
navegador.find_element(By.ID, "sideMenuSearch").click()
navegador.implicitly_wait(2)
navegador.find_element(By.CSS_SELECTOR, '[title="Vendas por Vendedor"]').click()
time.sleep(2)
wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="agrup_fil_2"]'))).click()
wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="tabTabdhtmlgoodies_tabView1_1"]/a'))).click()
wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="sel_relatorio_8"]'))).click()
wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="tabTabdhtmlgoodies_tabView1_3"]/a'))).click()

# Inserir códigos (agora usando os carregados da planilha)
campo_codigo = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="cod_reduzidoEntrada"]')))
for codigo in codigos:
    campo_codigo.clear()
    campo_codigo.send_keys(codigo)
    campo_codigo.send_keys(Keys.ENTER)
    wait.until(EC.invisibility_of_element((By.XPATH, '//*[@id="divLoading"]')))

# Calcula datas de 01 até D-1
hoje = datetime.now()
datas_para_processar = [
    (hoje.replace(day=d)).strftime('%d/%m/%Y')
    for d in range(1, hoje.day + 1)
]

df_geral = pd.DataFrame()

for data_str in datas_para_processar:
    print(f"📅 Processando data: {data_str}", flush=True)

    # Limpar pasta de downloads antes de gerar novo relatório
    for f in os.listdir(download_dir):
        file_path = os.path.join(download_dir, f)
        try:
            if os.path.isfile(file_path):
                os.remove(file_path)
        except Exception as e:
            print(f"Erro ao remover {file_path}: {e}")

    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="tabTabdhtmlgoodies_tabView1_4"]/a'))).click()
    campo_ini = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="dat_inicio"]')))
    campo_ini.clear()
    campo_ini.send_keys(data_str)
    campo_fim = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="dat_fim"]')))
    campo_fim.clear()
    campo_fim.send_keys(data_str)

    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="tabTabdhtmlgoodies_tabView1_5"]/a'))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="saida_4"]'))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="runReport"]'))).click()
    wait.until(EC.invisibility_of_element((By.XPATH, '//*[@id="divLoading"]')))
    
    time.sleep(2)
    
    print(f"✅ Relatório do dia {data_str} baixado.", flush=True)

    # Verifica e seleciona o arquivo mais recente .xls/.xlsx
    files = sorted(
        [f for f in os.listdir(download_dir) if f.endswith('.xls') or f.endswith('.xlsx')],
        key=lambda x: os.path.getmtime(os.path.join(download_dir, x)),
        reverse=True
    )

    if not files:
        print(f"⚠️ Nenhum arquivo encontrado para a data {data_str}. Pulando...", flush=True)
        continue
        
    xls_file_path = os.path.join(download_dir, files[0])

    try:
        # Try openpyxl first (most common for .xlsx)
        df = pd.read_excel(xls_file_path, header=14, engine='openpyxl')
    except Exception as e1:
        try:
            # Try xlrd for older .xls files
            df = pd.read_excel(xls_file_path, header=14, engine='xlrd')
        except Exception as e2:
            print(f"Failed with openpyxl: {e1}")
            print(f"Failed with xlrd: {e2}")
            raise ValueError(f"Cannot read Excel file: {xls_file_path}")

    col_lab = df.columns.get_loc('Laboratório')
    col_codigo_prod = df.columns.get_loc('Código')
    col_vendedor = col_lab - 1
    col_nome_vendedor = col_lab
    col_nome_produto = col_codigo_prod + 1
    col_qtd_vendida = col_codigo_prod + 4
    col_valor_venda = col_codigo_prod + 7

    filial_atual = None
    codigo_vendedor = None
    nome_vendedor = None
    resultados = []

    for i in range(len(df)):
        valor_info = str(df.iat[i, col_vendedor])
        if valor_info.isnumeric():
            filial_atual = int(valor_info)  # força como número
        elif '-' in valor_info:
            match = re.match(r'^(\d+)\s*-', valor_info)
            if match:
                codigo_vendedor = match.group(1)
                nome_vendedor = str(df.iat[i, col_nome_vendedor]).strip()

        codigo_produto = df.iat[i, col_codigo_prod]
        if pd.notna(codigo_produto) and str(codigo_produto).isnumeric():
            nome_produto = df.iat[i, col_nome_produto]
            qtd_vendida = df.iat[i, col_qtd_vendida]
            valor_venda = df.iat[i, col_valor_venda]
            resultados.append([
                codigo_vendedor,
                filial_atual,
                nome_vendedor,
                int(codigo_produto),
                str(nome_produto).strip(),
                qtd_vendida,
                valor_venda,
                data_str
            ])

    df_resultado = pd.DataFrame(resultados, columns=[
        'CODIGO VENDEDOR', 'FILIAL', 'NOME VENDEDOR', 'CODIGO PRODUTO',
        'NOME PRODUTO', 'QTD VENDIDA', 'VALOR VENDA', 'DATA'
    ])

    # Reorganiza colunas: coloca DATA primeiro
    df_resultado = df_resultado[['DATA', 'CODIGO VENDEDOR', 'FILIAL', 'NOME VENDEDOR',
                                'CODIGO PRODUTO', 'NOME PRODUTO', 'QTD VENDIDA', 'VALOR VENDA']]

    # Junta ao DataFrame geral
    df_geral = pd.concat([df_geral, df_resultado], ignore_index=True)

# Encerra o navegador
navegador.quit()

# === SEND RESULTS TO GOOGLE SHEETS (using same cliente) ===
# Note: Your cliente is already authenticated from earlier
try:
    url_planilha = 'https://docs.google.com/spreadsheets/d/1hXIGivSHQLfTU9UhNsNh4cDdt1tzBZ1ssAjruJbgoRs/edit#gid=0'
    planilha = cliente.open_by_url(url_planilha)
    aba = planilha.sheet1

    # Limpa apenas o intervalo necessário e insere os dados na ordem certa
    aba.batch_clear(['A2:H'])
    if not df_geral.empty:
        aba.update('A2', df_geral.values.tolist())
        print(f"📤 {len(df_geral)} linhas enviadas com sucesso ao Google Sheets.", flush=True)
    else:
        print("⚠️ Nenhum dado para enviar ao Google Sheets.", flush=True)
except Exception as e:
    print(f"Erro ao enviar dados para o Google Sheets: {e}", flush=True)

print("✅ Processamento concluído!", flush=True)
