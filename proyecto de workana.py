# Paso 0: Instalar si hace falta
# pip install pandas matplotlib seaborn reportlab openpyxl

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os

from reportlab.lib.pagesizes import A4
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle, PageBreak,
    ListFlowable, ListItem
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# -------------------------
# 1. Cargar datos
# -------------------------
os.chdir("C:/Users/NoxiePC/Desktop/proyecto de workana")
df = pd.read_excel("Copia de Base Accelerator - para proyectos.xlsx", sheet_name="Query result")
df["fecha de inscripción"] = pd.to_datetime(df["fecha de inscripción"])

# -------------------------
# 2. Generar visualizaciones
# -------------------------

sns.set(style="whitegrid")
plt.rcParams["figure.figsize"] = (10, 6)

# Para evitar FutureWarning en seaborn, se usa 'hue' y 'legend=False'

country_counts = df["country"].value_counts()
sns.barplot(
    x=country_counts.values,
    y=country_counts.index,
    hue=country_counts.index,
    palette="Blues_d",
    legend=False
)
plt.title("Distribución de talentos por país")
plt.xlabel("Cantidad de talentos")
plt.ylabel("País")
plt.tight_layout()
plt.savefig("01_talentos_por_pais.png")
plt.clf()

skill_counts = df["skill"].value_counts()
sns.barplot(
    x=skill_counts.values,
    y=skill_counts.index,
    hue=skill_counts.index,
    palette="Greens_d",
    legend=False
)
plt.title("Distribución por categoría de skills")
plt.xlabel("Cantidad de talentos")
plt.ylabel("Skill")
plt.tight_layout()
plt.savefig("02_talentos_por_skill.png")
plt.clf()

edad_por_skill = df.groupby("skill")["edad"].mean().sort_values()
sns.barplot(
    x=edad_por_skill.values,
    y=edad_por_skill.index,
    hue=edad_por_skill.index,
    palette="Purples_d",
    legend=False
)
plt.title("Edad promedio por categoría de skill")
plt.xlabel("Edad promedio")
plt.ylabel("Skill")
plt.tight_layout()
plt.savefig("03_edad_promedio_por_skill.png")
plt.clf()

inscripciones_por_mes = df["fecha de inscripción"].dt.to_period("M").value_counts().sort_index()
sns.lineplot(
    x=inscripciones_por_mes.index.astype(str),
    y=inscripciones_por_mes.values,
    marker="o",
    color="orange"
)
plt.title("Evolución de inscripciones al programa")
plt.xlabel("Mes")
plt.ylabel("Cantidad de inscripciones")
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig("04_evolucion_inscripciones.png")
plt.clf()

# -------------------------
# 3. Tabla resumen de países y skills
# -------------------------

top_paises = df["country"].value_counts().head(3).index.tolist()
resumen = []

for pais in top_paises:
    subset = df[df["country"] == pais]
    top_skill = subset["skill"].value_counts().idxmax()
    count = subset["skill"].value_counts().max()
    resumen.append([pais, top_skill, str(count)])

tabla_data = [["País", "Skill más común", "Cantidad"]] + resumen

# -------------------------
# 4. Generar PDF (limpio, sin texto 'bullet')
# -------------------------

pdf_path = r"C:\Users\NoxiePC\Desktop\informe_talentos_workana.pdf"
doc = SimpleDocTemplate(pdf_path, pagesize=A4)
styles = getSampleStyleSheet()
story = []

# Portada
story.append(Spacer(1, 100))
story.append(Paragraph("Informe de Talentos del Programa Accelerator – Workana", styles["Title"]))
story.append(Spacer(1, 20))
story.append(Paragraph("Visualización y análisis de datos a julio de 2025", styles["Heading2"]))
story.append(Spacer(1, 50))
story.append(Paragraph("Elaborado por: Luis Cortez", styles["Normal"]))
story.append(Spacer(1, 100))
story.append(Paragraph("Workana Accelerator – Talent Data Insights", styles["Italic"]))
story.append(Spacer(1, 50))
story.append(PageBreak())

# Gráficos
imagenes = [
    ("Distribución de talentos por país", "01_talentos_por_pais.png"),
    ("Distribución por categoría de skills", "02_talentos_por_skill.png"),
    ("Edad promedio por skill", "03_edad_promedio_por_skill.png"),
    ("Evolución de inscripciones al programa", "04_evolucion_inscripciones.png"),
]

for titulo, path in imagenes:
    story.append(Paragraph(titulo, styles["Heading2"]))
    story.append(Image(path, width=450, height=250))
    story.append(Spacer(1, 20))

# Tabla resumen
story.append(Paragraph("Resumen de principales países y skills", styles["Heading2"]))
tabla = Table(tabla_data, colWidths=[100, 200, 100])
tabla.setStyle(TableStyle([
    ("BACKGROUND", (0, 0), (-1, 0), colors.lightblue),
    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
    ("BOTTOMPADDING", (0, 0), (-1, 0), 12),
    ("GRID", (0, 0), (-1, -1), 1, colors.grey),
]))
story.append(tabla)
story.append(Spacer(1, 30))

# Conclusiones usando lista de viñetas sin texto extra
story.append(Paragraph("Conclusiones", styles["Heading2"]))

lista_conclusiones = [
    "La mayoría de los talentos provienen de Argentina, Colombia y Brasil.",
    "Las habilidades más comunes son Writing & Translation, Admin Support y Sales & Marketing.",
    "El promedio de edad varía según skill, indicando diversidad generacional.",
    "Se observa un crecimiento de inscripciones en los últimos meses."
]

lista_style = ParagraphStyle(
    'lista_style',
    parent=styles['BodyText'],
    leftIndent=20,
    bulletIndent=10,
    spaceAfter=5,
)

lista = ListFlowable(
    [ListItem(Paragraph(item, lista_style)) for item in lista_conclusiones],
    bulletType='bullet',
    bulletFontName='Helvetica',
    bulletFontSize=12,
    bulletDedent=5,
)

story.append(lista)
story.append(Spacer(1, 50))

# Footer
story.append(Spacer(1, 20))
story.append(Paragraph("Informe generado automáticamente con Python por Luis Cortez", styles["Italic"]))

# Exportar PDF
doc.build(story)
print(f"✅ PDF generado con éxito en: {pdf_path}")
