import os
import subprocess
from fpdf import FPDF

BASE_DIR = "/home/klab/Downloads/edu/Automaco-de-Relatorios-"
DADOS_FINAL = os.path.join(BASE_DIR, "Dados_Sousa_Janeiro_2023_Processado.xlsx")
PDF_TEMP = os.path.join(BASE_DIR, "Dados_Sousa_Janeiro_2023_Processado.pdf")
IMG_CONSOLIDADO = os.path.join(BASE_DIR, "consolidado_part1.png")
IMG_DASHBOARD = os.path.join(BASE_DIR, "dashboard_final.png")
PDF_FINAL = os.path.join(BASE_DIR, "Relatorio_Tecnico_Sousa_Janeiro_2023_GRAFICOS.pdf")

# 1. Converter Excel para PDF
if not os.path.exists(PDF_TEMP):
    print("Convertendo Excel para PDF...")
    subprocess.run([
        "libreoffice", "--headless", "--convert-to", "pdf", "--outdir", BASE_DIR, DADOS_FINAL
    ], check=True)

# 2. Extrair páginas específicas do PDF para PNG
# Consolidado: página 4, Dashboard Final: página 7
print("Extraindo página 4 (Consolidado)...")
subprocess.run([
    "pdftoppm", "-png", "-f", "4", "-l", "4", PDF_TEMP, os.path.splitext(IMG_CONSOLIDADO)[0]
], check=True)
img_consolidado_real = f"{os.path.splitext(IMG_CONSOLIDADO)[0]}-4.png"
if os.path.exists(img_consolidado_real):
    os.rename(img_consolidado_real, IMG_CONSOLIDADO)

print("Extraindo página 7 (Dashboard Final)...")
subprocess.run([
    "pdftoppm", "-png", "-f", "7", "-l", "7", PDF_TEMP, os.path.splitext(IMG_DASHBOARD)[0]
], check=True)
img_dashboard_real = f"{os.path.splitext(IMG_DASHBOARD)[0]}-7.png"
if os.path.exists(img_dashboard_real):
    os.rename(img_dashboard_real, IMG_DASHBOARD)

# 3. Montar PDF final só com os prints
print("Montando PDF final com os prints dos gráficos...")
pdf = FPDF()
for img_path in [IMG_CONSOLIDADO, IMG_DASHBOARD]:
    if os.path.exists(img_path):
        pdf.add_page()
        pdf.image(img_path, x=10, y=10, w=190)
pdf.output(PDF_FINAL)
print(f"PDF final gerado: {PDF_FINAL}")
