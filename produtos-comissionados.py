import pandas as pd
import re
import os
import gspread
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
wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="menuBar"]/li[11]/a/span[2]'))).click()
wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ul123"]/li[3]/a/span'))).click()
wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ul126"]/li[2]/a/span'))).click()
wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ul138"]/li/a/span'))).click()
wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="agrup_fil_2"]'))).click()
wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="tabTabdhtmlgoodies_tabView1_1"]/a'))).click()
wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="sel_relatorio_8"]'))).click()
wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="tabTabdhtmlgoodies_tabView1_3"]/a'))).click()

# Inserir códigos uma única vez
codigos = [
    82976, 76264, 75718, 75738, 83446, 2558, 2559, 71808, 84356, 60944,
    84029, 67967, 84336, 8001, 71807, 81362, 84920, 74943, 74944, 76039,
    78545, 9189, 9190, 9192, 9197, 9253, 9258, 69857, 75336, 69476,
    85834, 85833, 64912, 77137, 69612, 83985, 85662, 85661, 60573, 32103,
    7988, 82672, 78657, 85427, 18931, 18932, 71524, 64191, 79288, 79641,
    84954, 26932, 26936, 60401, 69858, 79638, 79637, 85399, 85400, 85401,
    64976, 64868, 64977, 65222, 61350, 59748, 51031, 54386, 81983, 69299,
    84176, 85658, 85657, 85659, 69300, 81283, 84284, 31391, 83753, 83749,
    83754, 83751, 83752, 83750, 77195, 80772, 72763, 72604, 76219, 69574,
    69339, 65731, 64238, 57856, 85831, 85832, 85835, 85836, 85837, 85839,
    85838, 85840, 85842, 85843, 85841, 65215, 85532, 18932, 82601, 75847,
    75884, 69861, 69860, 84917, 83941, 85897, 85660, 77602, 74286, 84489,
    69568, 85898, 79646, 65561, 69614, 83620, 85680, 85681, 85682, 40989,
    78270, 78271, 42727, 72579, 56381, 79553, 81830, 26898, 65886, 79648,
    72220, 48905, 69609, 76280, 76279, 84921, 83895, 85012, 76276, 76038,
    76277, 85011, 65888, 67981, 80703, 70729, 70112, 81903, 81904, 81989,
    81990, 81905, 78883, 81902, 83797, 62542, 70113
]

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
        os.remove(os.path.join(download_dir, f))

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
        df = pd.read_excel(xls_file_path, header=14)
    except Exception as e:
        print(f"Erro ao ler Excel: {e}", flush=True)
        raise

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

# Envio ao Google Sheets
json_credenciais = 'creds.json'
url_planilha = 'https://docs.google.com/spreadsheets/d/1hXIGivSHQLfTU9UhNsNh4cDdt1tzBZ1ssAjruJbgoRs/edit#gid=0'
escopo = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credenciais = ServiceAccountCredentials.from_json_keyfile_name(json_credenciais, escopo)
cliente = gspread.authorize(credenciais)
planilha = cliente.open_by_url(url_planilha)
aba = planilha.sheet1

# Limpa apenas o intervalo necessário e insere os dados na ordem certa
aba.batch_clear(['A2:H'])
aba.update('A2', df_geral.values.tolist())

print("📤 Todos os dados foram enviados com sucesso ao Google Sheets.", flush=True)















