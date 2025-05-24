import subprocess
import os
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

BASE_DIR = "/home/klab/Downloads/edu/Automaco-de-Relatorios-"
DATA_DIR = os.path.join(BASE_DIR, "data")
os.makedirs(DATA_DIR, exist_ok=True)

# 1. Executa automacao_relatorios.py para atualizar a planilha de dados
logging.info("Executando automacao_relatorios.py para atualizar a planilha de dados...")
subprocess.run(["python", os.path.join(BASE_DIR, "automacao_relatorios.py")], check=True)

# 2. Move a planilha processada para a pasta data
DADOS_FINAL = os.path.join(BASE_DIR, "Dados_Sousa_Janeiro_2023_Processado.xlsx")
DADOS_FINAL_DATA = os.path.join(DATA_DIR, "Dados_Sousa_Janeiro_2023_Processado.xlsx")
if os.path.exists(DADOS_FINAL):
    os.replace(DADOS_FINAL, DADOS_FINAL_DATA)

# 3. Converte a planilha processada para PDF
PDF_TEMP = os.path.join(DATA_DIR, "Dados_Sousa_Janeiro_2023_Processado.pdf")
if not os.path.exists(PDF_TEMP):
    logging.info("Convertendo Excel processado para PDF...")
    subprocess.run([
        "libreoffice", "--headless", "--convert-to", "pdf", "--outdir", DATA_DIR, DADOS_FINAL_DATA
    ], check=True)

# 4. Extrai as páginas desejadas do PDF para PNG
logging.info("Extraindo página 4 (Consolidado) do PDF...")
subprocess.run([
    "pdftoppm", "-png", "-f", "4", "-l", "4", PDF_TEMP, os.path.join(DATA_DIR, "consolidado_part1")
], check=True)
img_consolidado = os.path.join(DATA_DIR, "consolidado_part1-4.png")
if not os.path.exists(img_consolidado):
    img_consolidado = os.path.join(DATA_DIR, "consolidado_part1-04.png")
if os.path.exists(img_consolidado):
    os.replace(img_consolidado, os.path.join(DATA_DIR, "consolidado_part1.png"))

logging.info("Extraindo página 6 (Dashboard Custos) do PDF...")
subprocess.run([
    "pdftoppm", "-png", "-f", "6", "-l", "6", PDF_TEMP, os.path.join(DATA_DIR, "dashboard_custos")
], check=True)
img_dashboard = os.path.join(DATA_DIR, "dashboard_custos-6.png")
if not os.path.exists(img_dashboard):
    img_dashboard = os.path.join(DATA_DIR, "dashboard_custos-06.png")
if os.path.exists(img_dashboard):
    os.replace(img_dashboard, os.path.join(DATA_DIR, "dashboard_custos.png"))

# 5. Gera os PDFs finais das abas separadas
logging.info("Gerando PDFs separados para Consolidado e Dashboard Custos...")
subprocess.run(["python", os.path.join(BASE_DIR, "gerar_pdfs_abas.py")], check=True)
logging.info("Processo completo! Verifique a pasta 'data' para os arquivos gerados.")
