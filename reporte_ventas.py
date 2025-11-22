import sys
from pathlib import Path
from datetime import datetime
import pandas as pd
import requests
import os
import numpy as np
# PDF
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.platypus import Table, TableStyle
from reportlab.lib import colors
import matplotlib.pyplot as plt
from pathlib import Path
from reportlab.platypus import Image
# =========================
# CONFIGURACIÓN BÁSICA
# =========================

# Excel por defecto (ajústalo a tu ruta)
DEFAULT_EXCEL_PATH = Path(r"ventas.xlsx")

# Columnas esperadas
COL_ZONE = "Zone"
COL_APROBADO = "Valor Aprobado (Real)"
COL_GASTOS_TOTAL = "Gastos Total"
COL_UTILIDAD = "Utilidad ($)"
COL_FECHA = "Invoice Date"

# Config DeepSeek/Ollama
OLLAMA_URL = "OLLAMA_URL"
MODEL_NAME = "VERSIO_DEEPSEEK"

# Logo opcional
LOGO_PATH = None   # Ej: r"C:\logo.png"



# =========================
# FUNCIONES
# =========================

def cargar_excel(path: Path) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(f"No encuentro el archivo Excel: {path}")
    df = pd.read_excel(path)

    for col in [COL_ZONE, COL_APROBADO, COL_GASTOS_TOTAL, COL_UTILIDAD, COL_FECHA]:
        if col not in df.columns:
            raise KeyError(f"Falta la columna '{col}' en {path}")

    for col in [COL_APROBADO, COL_GASTOS_TOTAL, COL_UTILIDAD]:
        df[col] = (
            df[col]
            .astype(str)
            .str.replace(r"[$,]", "", regex=True)
            .replace("nan", 0)
            .astype(float)
        )

    df[COL_FECHA] = pd.to_datetime(df[COL_FECHA], errors="coerce")
    df = df.dropna(subset=[COL_FECHA])
    return df


def resumir_mes(df, year, month):
    mask = (df[COL_FECHA].dt.year == year) & (df[COL_FECHA].dt.month == month)
    df_mes = df.loc[mask]

    if df_mes.empty:
        return None, pd.DataFrame()

    resumen = (
        df_mes
        .groupby(COL_ZONE)
        .agg({
            COL_APROBADO: "sum",
            COL_UTILIDAD: "sum",
            COL_GASTOS_TOTAL: "sum"
        })
        .rename(columns={
            COL_APROBADO: "Ventas",
            COL_UTILIDAD: "Utilidad",
            COL_GASTOS_TOTAL: "Gastos"
        })
    )

    resumen["% Utilidad"] = (resumen["Utilidad"] / resumen["Ventas"]) * 100

    resumen = resumen.reset_index().rename(columns={COL_ZONE: "Zona"})

    total = resumen[["Ventas", "Utilidad", "Gastos"]].sum()
    total_row = {
        "Zona": "TOTAL GENERAL",
        "Ventas": total["Ventas"],
        "Utilidad": total["Utilidad"],
        "Gastos": total["Gastos"],
        "% Utilidad": (total["Utilidad"] / total["Ventas"] * 100 if total["Ventas"] else 0)
    }
    resumen.loc[len(resumen)] = total_row

    resumen_fmt = resumen.copy()
    resumen_fmt["Ventas"] = resumen_fmt["Ventas"].map(lambda x: f"{x:,.2f}")
    resumen_fmt["Utilidad"] = resumen_fmt["Utilidad"].map(lambda x: f"{x:,.2f}")
    resumen_fmt["Gastos"] = resumen_fmt["Gastos"].map(lambda x: f"{x:,.2f}")
    resumen_fmt["% Utilidad"] = resumen_fmt["% Utilidad"].map(lambda x: f"{x:.1f}%")

    return resumen, resumen_fmt

def crear_graficos(tabla_actual, tabla_anterior, carpeta_salida="."):
    # Quitamos la fila de TOTAL GENERAL
    df_act = tabla_actual[tabla_actual["Zona"] != "TOTAL GENERAL"].copy()
    df_ant = tabla_anterior[tabla_anterior["Zona"] != "TOTAL GENERAL"].copy()

    # Asegurar que son numéricos (por si quedaron como string)
    for col in ["Ventas", "Utilidad"]:
        df_act[col] = pd.to_numeric(df_act[col], errors="coerce")
        df_ant[col] = pd.to_numeric(df_ant[col], errors="coerce")

    # 1) Barras de ventas por zona (mes actual vs anterior)
    zonas = df_act["Zona"].tolist()
    x = range(len(zonas))

    ventas_act = df_act["Ventas"].values
    # alineamos anterior a las mismas zonas (si alguna no existe, pone 0)
    ventas_ant = [
        float(df_ant[df_ant["Zona"] == z]["Ventas"].values[0])
        if (df_ant["Zona"] == z).any() else 0.0
        for z in zonas
    ]

    plt.figure()
    ancho = 0.35
    plt.bar([i - ancho/2 for i in x], ventas_ant, width=ancho, label="Mes anterior")
    plt.bar([i + ancho/2 for i in x], ventas_act, width=ancho, label="Mes actual")
    plt.xticks(list(x), zonas, rotation=45, ha="right")
    plt.ylabel("Ventas (USD)")
    plt.title("Ventas por zona – Comparación mensual")
    plt.legend()
    ruta_barras = str(Path(carpeta_salida) / "grafico_ventas.png")
    plt.tight_layout()
    plt.savefig(ruta_barras, dpi=120)
    plt.close()

    # 2) Pie de utilidad por zona (mes actual) – con filtros para evitar NaN/0
    valores = pd.to_numeric(df_act["Utilidad"], errors="coerce")
    etiquetas = df_act["Zona"].astype(str)

    # Nos quedamos solo con utilidades válidas y > 0
    mascara = (valores.notna()) & (valores > 0)
    valores_filtrados = valores[mascara]
    etiquetas_filtradas = etiquetas[mascara]

    ruta_pie = None
    if len(valores_filtrados) > 0 and float(valores_filtrados.sum()) > 0:
        plt.figure()
        plt.pie(
            valores_filtrados.values,
            labels=etiquetas_filtradas.values,
            autopct="%1.1f%%"
        )
        plt.title("Distribución de la utilidad por zona – Mes actual")
        ruta_pie = str(Path(carpeta_salida) / "grafico_utilidad.png")
        plt.savefig(ruta_pie, dpi=120)
        plt.close()
    else:
        print("⚠ No se genera gráfico de pie: utilidades todas 0 o NaN.")

    return ruta_barras, ruta_pie

def construir_prompt(tabla_actual_txt: str,
                     etiqueta_actual: str,
                     tabla_anterior_txt: str,
                     etiqueta_anterior: str) -> str:
    return f"""
Eres un analista financiero senior. Vas a escribir un análisis corto y profesional
de ventas y utilidad para un reporte mensual interno.

TODOS los montos están en **dólares estadounidenses (USD)**.
No menciones soles ni ninguna otra moneda.

Datos del mes {etiqueta_anterior} (USD, formato US con coma de miles y punto decimal):
{tabla_anterior_txt}

Datos del mes {etiqueta_actual} (USD, formato US con coma de miles y punto decimal):
{tabla_actual_txt}

Reglas estrictas:
- No recalcules totales ni porcentajes.
- Si citas una cifra, cópiala EXACTAMENTE como aparece en las tablas,
  sin cambiar puntos por comas ni inventar millones.
- No inventes números nuevos ni “montos totales” que no estén en las tablas.
- Idioma: español neutro, tono ejecutivo.
- Máximo 250 palabras.

Escribe exactamente 5 párrafos numerados así: 
"1. ...", "2. ...", "3. ...", "4. ...", "5. ...".
Cada párrafo de 2–3 frases.

Contenido de cada párrafo:
1. Comparación general entre ambos meses (ventas y utilidad).
2. Zonas fuertes y qué las hace fuertes.
3. Zonas débiles y riesgos.
4. Comentario del margen de utilidad y de los gastos.
5. Cierre ejecutivo con 2–3 recomendaciones concretas.

No uses viñetas, no mezcles inglés y no pongas símbolos raros.

Formatea el análisis así:
- Usa subtítulos numerados: 1., 2., 3., 4., 5.
- Cada punto debe ser uno o dos párrafos cortos.
- Deja una línea en blanco entre cada punto para que se lea bien en un PDF.

Responde SOLO con el texto final, sin explicaciones extra.
IMPORTANTE – FORMATO DE NÚMEROS:
- Todos los montos ya están en dólares estadounidenses (USD).
- Usa SIEMPRE el formato: $1,234.56  (coma para miles, punto para decimales).
- No cambies la escala de los números: no conviertas miles en millones ni al revés.
- No inventes nuevas cantidades: repite exactamente los mismos importes que ves en las tablas.

Escribe TODO el análisis en español neutro, sin palabras en inglés.
Responde solo con el texto del análisis, sin explicar tu razonamiento interno.
""".strip()





import textwrap

MAX_PROMPT_CHARS = 5000  # para no matar a deepseek con un test

def llamar_deepseek(prompt: str) -> str:
    # Por si el prompt se nos fue de la mano
    if len(prompt) > MAX_PROMPT_CHARS:
        print(f"⚠️ Prompt muy largo ({len(prompt)} caracteres). "
              f"Lo recorto a {MAX_PROMPT_CHARS}.")
        prompt = prompt[:MAX_PROMPT_CHARS]

    payload = {
        "model": "deepseek-r1:7b",
        "prompt": prompt,
        "stream": False,
    }

    print("\nDEBUG – Prompt que se envía a DeepSeek (primeros 3000 caracteres):\n")
    print(textwrap.shorten(prompt, width=3000, placeholder="..."))
    print("\n=== FIN DEBUG ===\n")

    try:
        r = requests.post(OLLAMA_URL, json=payload, timeout=200)
    except Exception as e:
        print("❌ Error al conectar con Ollama:", e)
        return "No se pudo conectar con el servidor de modelos."

    if r.status_code != 200:
        print("❌ Ollama devolvió un error:")
        print("Código:", r.status_code)
        print("Cuerpo:", r.text)
        return "El servidor de modelos devolvió un error interno al procesar el análisis."

    try:
        data = r.json()
    except Exception as e:
        print("❌ No se pudo parsear la respuesta de Ollama:", e)
        print("Texto bruto:", r.text[:500])
        return "El servidor de modelos respondió algo que no pude interpretar."

    return data.get("response", "")


def limpiar_texto(s: str) -> str:
    """Saca caracteres que ReportLab no maneja bien."""
    reemplazos = {
        "•": "-",
        "●": "-",
        "■": "-",
        "►": "-",
        "–": "-",
        "—": "-",
        "“": '"',
        "”": '"',
        "’": "'",
        "…": "...",
    }
    for k, v in reemplazos.items():
        s = s.replace(k, v)
    return s


from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

def _formatear_monedas(df):
    df = df.copy()

    def _fmt(x):
        try:
            # por si viene como "1,065.50" en texto
            x_num = float(str(x).replace(",", ""))
            return f"${x_num:,.2f}"
        except Exception:
            return str(x)

    for col in ["Ventas", "Utilidad", "Gastos"]:
        if col in df.columns:
            df[col] = df[col].apply(_fmt)
    return df

def _dibujar_tabla(c, df, titulo, x, y_inicial):
    width, height = letter
    line_height = 14
    y = y_inicial

    c.setFont("Helvetica-Bold", 12)
    c.drawString(x, y, titulo)
    y -= line_height * 1.5

    df = _formatear_monedas(df)

    # encabezados
    c.setFont("Helvetica-Bold", 9)
    headers = list(df.columns)
    col_x = [x, x + 130, x + 230, x + 330, x + 430]

    for hx, h in zip(col_x, headers):
        c.drawString(hx, y, str(h))
    y -= line_height

    c.setFont("Helvetica", 9)
    for _, row in df.iterrows():
        # salto de página si no hay espacio
        if y < 60:
            c.showPage()
            y = height - 60
            c.setFont("Helvetica", 9)

        valores = [str(row[col]) for col in headers]
        for hx, val in zip(col_x, valores):
            c.drawString(hx, y, val)
        y -= line_height

    return y - line_height  # dejamos un espacio extra después de la tabla

def _dibujar_texto_largo(c, texto, x, y_inicial, font="Helvetica", size=10):
    from reportlab.pdfbase import pdfmetrics
    width, height = letter
    margen_derecho = 40
    max_ancho = width - x - margen_derecho
    line_height = 14

    c.setFont(font, size)
    y = y_inicial

    # limpiamos primero
    texto = limpiar_texto(texto)

    # 1) Respetar párrafos: separamos por saltos de línea
    lineas = texto.split("\n")

    for linea in lineas:
        # línea en blanco => salto de párrafo
        if not linea.strip():
            y -= line_height   # espacio extra entre párrafos
            continue

        # 2) Wrap de esa línea según ancho real
        palabras = linea.split()
        buffer = ""

        for palabra in palabras:
            prueba = (buffer + " " + palabra).strip()
            w = pdfmetrics.stringWidth(prueba, font, size)

            if w <= max_ancho:
                buffer = prueba
            else:
                # dibujar línea actual y saltar
                if y < 60:
                    c.showPage()
                    y = height - 60
                    c.setFont(font, size)

                c.drawString(x, y, buffer)
                y -= line_height
                buffer = palabra

        # última parte de la línea
        if buffer:
            if y < 60:
                c.showPage()
                y = height - 60
                c.setFont(font, size)
            c.drawString(x, y, buffer)
            y -= line_height


def dibujar_texto_con_parrafos(canvas_obj, text: str,
                               x: float, y: float,
                               max_width_chars: int = 110,
                               leading: int = 14):
    """
    Pinta texto en el PDF respetando párrafos separados por saltos de línea.
    `max_width_chars` controla cuántos caracteres por línea aprox.
    """
    for parrafo in text.split("\n"):
        parrafo = parrafo.strip()
        if not parrafo:
            # línea en blanco entre párrafos
            y -= leading
            continue

        for linea in textwrap.wrap(parrafo, width=max_width_chars):
            canvas_obj.drawString(x, y, linea)
            y -= leading

        # espacio extra entre párrafos
        y -= leading

    return y



def generar_pdf(ruta_pdf, tabla_actual, tabla_anterior, analisis):
    c = canvas.Canvas(ruta_pdf, pagesize=letter)
    width, height = letter

    # Encabezado
    c.setFont("Helvetica-Bold", 16)
    c.drawString(40, height - 40, "Reporte de Ventas")
    c.setFont("Helvetica", 9)
    c.drawString(40, height - 55, f"Generado: {datetime.now().strftime('%Y-%m-%d %H:%M')}")

    y = height - 80

    # Crear gráficos
    ruta_barras, ruta_pie = crear_graficos(tabla_actual, tabla_anterior, carpeta_salida=os.path.dirname(ruta_pdf))


    # Mes anterior
    y = _dibujar_tabla(c, tabla_anterior, "Mes Anterior", 40, y)

    # Mes actual
    if y < 120:
        c.showPage()
        c.setFont("Helvetica-Bold", 16)
        c.drawString(40, height - 40, "Reporte de Ventas (cont.)")
        y = height - 80

    y = _dibujar_tabla(c, tabla_actual, "Mes Actual", 40, y)

    # Nueva página para gráficos
    c.showPage()
    c.setFont("Helvetica-Bold", 16)
    c.drawString(40, height - 40, "Gráficas de Ventas y Utilidad")
    y = height - 80

    # Gráfico de barras
    c.drawImage(ruta_barras, 40, y - 250, width=260, height=220, preserveAspectRatio=True, mask="auto")

    # Gráfico de pie
    c.drawImage(ruta_pie, 320, y - 250, width=260, height=220, preserveAspectRatio=True, mask="auto")

    # Avanzas luego a la página del análisis
    c.showPage()


    # Análisis
    if y < 120:
        c.showPage()
        c.setFont("Helvetica-Bold", 16)
        c.drawString(40, height - 40, "Reporte de Ventas (Análisis)")
        y = height - 80

    c.setFont("Helvetica-Bold", 12)
    c.drawString(40, y, "Análisis")
    y -= 20

    c.setFont("Helvetica", 11)
    y = dibujar_texto_con_parrafos(c, analisis, x=40, y=y, max_width_chars=100, leading=14)

    c.save()


# =========================
# MAIN
# =========================

def main():
    excel_path = DEFAULT_EXCEL_PATH

    print(f"Usando archivo Excel: {excel_path}")

    df = cargar_excel(excel_path)

    hoy = datetime.today()
    cy, cm = hoy.year, hoy.month
    py, pm = (cy - 1, 12) if cm == 1 else (cy, cm - 1)

    etiqueta_actual = f"{cm:02d}/{cy}"
    etiqueta_anterior = f"{pm:02d}/{py}"

    resumen_actual, resumen_actual_fmt = resumir_mes(df, cy, cm)
    resumen_anterior, resumen_anterior_fmt = resumir_mes(df, py, pm)

    print(resumen_actual_fmt)

    tabla_actual_txt = resumen_actual.to_string(index=False)
    tabla_anterior_txt = resumen_anterior.to_string(index=False)

    prompt = construir_prompt(
        tabla_actual_txt, etiqueta_actual,
        tabla_anterior_txt, etiqueta_anterior
    )

    print("\nLlamando a DeepSeek...\n")
    analisis = llamar_deepseek(prompt)

    print("\n===== ANÁLISIS DEEPSEEK =====\n")
    print(analisis)

    # Crear PDF
    ruta_pdf = f"reporte_ventas_{cy}_{cm}.pdf"
    generar_pdf(ruta_pdf, resumen_actual_fmt, resumen_anterior_fmt, analisis)

    print(f"\nPDF generado: {ruta_pdf}")


if __name__ == "__main__":
    main()
