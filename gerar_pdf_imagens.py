from fpdf import FPDF
import os

BASE_DIR = "/home/klab/Downloads/edu/Automaco-de-Relatorios-"
img_consolidado1_path = os.path.join(BASE_DIR, "consolidado_part1.png")
img_dashboard_path = os.path.join(BASE_DIR, "dashboard_final.png")
output_pdf = os.path.join(BASE_DIR, "Relatorio_Tecnico_Sousa_Janeiro_2023_Automatizado_IMAGENS.pdf")

pdf = FPDF()
for img_path in [img_consolidado1_path, img_dashboard_path]:
    if img_path and os.path.exists(img_path):
        pdf.add_page()
        pdf.image(img_path, x=10, y=10, w=190)  # Ajusta largura para caber na página A4
pdf.output(output_pdf)
print(f"PDF final gerado com imagens: {output_pdf}")
