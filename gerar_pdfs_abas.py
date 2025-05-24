from fpdf import FPDF
import os

BASE_DIR = "/home/klab/Downloads/edu/Automaco-de-Relatorios-"
DATA_DIR = os.path.join(BASE_DIR, "data")

# Consolidado
img_consolidado = os.path.join(DATA_DIR, "consolidado_part1.png")
pdf_consolidado = os.path.join(DATA_DIR, "Consolidado.pdf")
if not os.path.exists(img_consolidado):
    # Tenta com zero à esquerda
    img_consolidado = os.path.join(DATA_DIR, "consolidado_part1-04.png")
if os.path.exists(img_consolidado):
    pdf = FPDF()
    pdf.add_page()
    pdf.image(img_consolidado, x=10, y=10, w=190)
    pdf.output(pdf_consolidado)
    print(f"PDF Consolidado gerado: {pdf_consolidado}")
else:
    print("Imagem do Consolidado não encontrada!")

# Dashboard Custos
img_dashboard = os.path.join(DATA_DIR, "dashboard_custos.png")
pdf_dashboard = os.path.join(DATA_DIR, "DashboardCustos.pdf")
if not os.path.exists(img_dashboard):
    img_dashboard = os.path.join(DATA_DIR, "dashboard_custos-06.png")
if os.path.exists(img_dashboard):
    pdf = FPDF()
    pdf.add_page()
    pdf.image(img_dashboard, x=10, y=10, w=190)
    pdf.output(pdf_dashboard)
    print(f"PDF Dashboard Custos gerado: {pdf_dashboard}")
else:
    print("Imagem do Dashboard Custos não encontrada!")
