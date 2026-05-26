#!/usr/bin/env python3
"""
generarDashboardMes.py
======================
Genera la hoja 'Dashboard_Resumen' en el Excel del mes y una diapositiva
PowerPoint (.pptx) con el dashboard operativo.

Uso:
    python generarDashboardMes.py --anio 2026 --mes Abril

El fichero Excel debe estar en la misma carpeta que el script con el nombre:
    {anio}_{cod}_{mes}_Informe Mensual Aprovisionamiento de Datos.xlsx
    Ej: 2026_04_Abril_Informe Mensual Aprovisionamiento de Datos.xlsx
"""

# ===========================================================================
# 0. AUTO-INSTALACIÓN DE DEPENDENCIAS
# ===========================================================================
import subprocess, sys

def _instalar(paquete, importar_como=None):
    nombre = importar_como or paquete.replace("-", "_")
    try:
        __import__(nombre)
    except ImportError:
        print(f"  📦 Instalando {paquete}...")
        resultado = subprocess.run(
            [sys.executable, "-m", "pip", "install", paquete, "-q"],
            capture_output=True, text=True
        )
        if resultado.returncode != 0:
            print(f"  ❌ No se pudo instalar {paquete}:\n{resultado.stderr}")
            sys.exit(1)
        print(f"  ✅ {paquete} instalado.")

for _pkg, _imp in [
    ("pandas",      "pandas"),
    ("openpyxl",    "openpyxl"),
    ("python-pptx", "pptx"),
    ("matplotlib",  "matplotlib"),
]:
    _instalar(_pkg, _imp)

# ===========================================================================
# 1. IMPORTS
# ===========================================================================
import argparse
import io
import os

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

# ===========================================================================
# 2. CONFIGURACIÓN
# ===========================================================================
ORDEN = ["1", "3", "5", "7"]

MAPA = {
    "1": {"n": "Búsqueda de datos",      "c": "#51a5e1"},
    "3": {"n": "Bajada de datos",        "c": "#5a6b7d"},
    "5": {"n": "Otras consultas",        "c": "#23a985"},
    "7": {"n": "Cargas en Data Manager", "c": "#e75437"},
}

MESES_COD = {
    "Enero": "01", "Febrero": "02", "Marzo": "03",   "Abril":     "04",
    "Mayo":  "05", "Junio":  "06", "Julio": "07",   "Agosto":    "08",
    "Septiembre": "09", "Octubre": "10", "Noviembre": "11", "Diciembre": "12",
}

ORDEN_COLUMNAS = [
    "Búsqueda de datos", "Bajada de datos",
    "Otras consultas",   "Cargas en Data Manager",
]

# ===========================================================================
# 3. FUNCIONES AUXILIARES
# ===========================================================================
def hex_to_rgb(hex_color: str) -> tuple:
    h = hex_color.lstrip("#")
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))


def hex_to_rgb_pptx(hex_color: str) -> RGBColor:
    r, g, b = hex_to_rgb(hex_color)
    return RGBColor(r, g, b)


def tiempo_natural_sin_findes(inicio, fin) -> float:
    """Calcula horas laborables entre dos fechas (excluye fines de semana)."""
    if pd.isnull(inicio) or pd.isnull(fin) or fin < inicio:
        return 0.0
    segundos_totales = (fin - inicio).total_seconds()
    segundos_finde = 0.0
    current = inicio
    while current < fin:
        next_day = (current + pd.Timedelta(days=1)).normalize()
        if next_day > fin:
            next_day = fin
        if current.weekday() >= 5:          # sábado=5, domingo=6
            segundos_finde += (next_day - current).total_seconds()
        current = next_day
    return (segundos_totales - segundos_finde) / 3600.0


def horas_a_texto(horas: float) -> str:
    h = int(horas)
    m = int((horas % 1) * 60)
    return f"{h} horas {m} minutos"


# ===========================================================================
# 4. LECTURA Y PROCESAMIENTO DEL XLSX
# ===========================================================================
def procesar_datos(ruta_xlsx: str) -> tuple:
    """
    Lee las hojas del Excel y devuelve:
        data_final               – lista de dicts por categoría
        total_peticiones         – int
        valores_sheets_peticiones – dict {cat: count}
        valores_sheets_horas      – dict {cat: media_horas}
    """
    xls = pd.ExcelFile(ruta_xlsx, engine="openpyxl")
    nombres_hojas = xls.sheet_names

    # --- Hoja 9-: ratings ---
    ws9 = next((h for h in nombres_hojas if h.startswith("9")), None)
    ratings_dict: dict = {}
    if ws9:
        df9 = xls.parse(ws9)
        df9.columns = [str(c).strip() for c in df9.columns]
        if "Categorization Tier 3" in df9.columns:
            df9 = df9.rename(columns={"Categorization Tier 3": "Category"})
        replacements = {
            "BUSQUEDA DE DATOS":       "Búsqueda de datos",
            "OTRAS CONSULTAS":         "Otras consultas",
            "BAJADA DE DATOS":         "Bajada de datos",
            "CARGAS EN DATA MANAGER":  "Cargas en Data Manager",
        }
        if "Category" in df9.columns:
            df9["Category"] = df9["Category"].astype(str).str.strip().replace(replacements)
        if "Rating" in df9.columns and "Category" in df9.columns:
            df9["Rating"] = pd.to_numeric(df9["Rating"], errors="coerce")
            summary = df9.groupby("Category")["Rating"].agg(["mean", "count"])
            ratings_dict = summary.to_dict("index")

    # --- Hojas 1, 3, 5, 7: peticiones ---
    data_final = []
    total_peticiones = 0
    valores_sheets_peticiones: dict = {}
    valores_sheets_horas: dict = {}

    for id_p in ORDEN:
        ws_nombre = next((h for h in nombres_hojas if h.startswith(id_p)), None)
        if not ws_nombre:
            print(f"  ⚠️  No se encontró hoja que empiece por '{id_p}', se omite.")
            continue

        df = xls.parse(ws_nombre)
        df.columns = [str(c).strip() for c in df.columns]

        if "Status" not in df.columns:
            print(f"  ⚠️  Hoja '{ws_nombre}' sin columna 'Status', se omite.")
            continue

        df_flt = df[df["Status"].isin(["Terminado", "Cerrado"])].copy()
        if df_flt.empty:
            print(f"  ⚠️  Hoja '{ws_nombre}' sin filas Terminado/Cerrado.")

        for col in ["Submit Date", "Completed Date"]:
            if col in df_flt.columns:
                df_flt[col] = pd.to_datetime(
                    df_flt[col].astype(str).str.replace(".", "", regex=False),
                    errors="coerce",
                    format="mixed",
                )

        df_flt["horas_res"] = df_flt.apply(
            lambda row: tiempo_natural_sin_findes(
                row.get("Submit Date"), row.get("Completed Date")
            ),
            axis=1,
        )

        nombre_cat  = MAPA[id_p]["n"]
        media_horas = float(df_flt["horas_res"].mean() or 0.0)
        count       = len(df_flt)
        total_peticiones += count
        valores_sheets_peticiones[nombre_cat] = float(count)
        valores_sheets_horas[nombre_cat]      = round(media_horas, 2)

        stats      = ratings_dict.get(nombre_cat, {"mean": 0, "count": 0})
        rating_val = float(stats["mean"]) if stats["count"] > 0 else 0.0
        has_data   = stats["count"] > 0
        count_val  = int(stats["count"]) if has_data else 0
        rating_txt = f"{rating_val:.1f}/5 ({count_val})" if has_data else "N/A"

        data_final.append({
            "cat":     nombre_cat,
            "val":     count,
            "media":   media_horas,
            "txt":     f"{int(media_horas):02d}h {int((media_horas % 1) * 60):02d}m",
            "color":   MAPA[id_p]["c"],
            "rating":  rating_val,
            "r_txt":   rating_txt,
            "r_width": rating_val * 20 if has_data else 0,
        })

    return data_final, total_peticiones, valores_sheets_peticiones, valores_sheets_horas


# ===========================================================================
# 5. ACTUALIZAR HOJA Dashboard_Resumen EN EL XLSX
# ===========================================================================
def actualizar_dashboard_resumen(
    ruta_xlsx: str,
    data_final: list,
    total_peticiones: int,
    valores_sheets_peticiones: dict,
    valores_sheets_horas: dict,
) -> None:
    wb = load_workbook(ruta_xlsx)

    # Eliminar hoja existente si la hay
    if "Dashboard_Resumen" in wb.sheetnames:
        del wb["Dashboard_Resumen"]
    ws = wb.create_sheet("Dashboard_Resumen")

    media_global = (
        sum(valores_sheets_horas.values()) / len(valores_sheets_horas)
        if valores_sheets_horas else 0.0
    )

    # --- Datos ---
    fila_cabecera   = ["Categoría"]   + ORDEN_COLUMNAS          + ["TOTAL/MEDIA"]
    fila_tiempos    = ["Tiempo (h)"]  + [valores_sheets_horas.get(c, 0)      for c in ORDEN_COLUMNAS] + [round(media_global, 2)]
    fila_peticiones = ["Peticiones"]  + [valores_sheets_peticiones.get(c, 0) for c in ORDEN_COLUMNAS] + [total_peticiones]
    fila_formato    = ["Formato"]     + [horas_a_texto(valores_sheets_horas.get(c, 0)) for c in ORDEN_COLUMNAS] + [horas_a_texto(media_global)]

    for fila in [fila_cabecera, fila_tiempos, fila_peticiones, fila_formato]:
        ws.append(fila)

    # --- Estilos cabecera ---
    colores_cat_fill = {
        "Búsqueda de datos":      "FF51a5e1",
        "Bajada de datos":        "FF5a6b7d",
        "Otras consultas":        "FF23a985",
        "Cargas en Data Manager": "FFe75437",
    }
    # A1 negro
    ws["A1"].fill  = PatternFill("solid", fgColor="FF1e293b")
    ws["A1"].font  = Font(color="FFFFFFFF", bold=True)
    ws["A1"].alignment = Alignment(horizontal="left")

    for col_idx, cat in enumerate(ORDEN_COLUMNAS, start=2):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill      = PatternFill("solid", fgColor=colores_cat_fill[cat])
        cell.font      = Font(bold=True, color="FFFFFFFF")
        cell.alignment = Alignment(horizontal="center")

    # TOTAL/MEDIA (col 6)
    ws.cell(row=1, column=6).fill      = PatternFill("solid", fgColor="FF334155")
    ws.cell(row=1, column=6).font      = Font(bold=True, color="FFFFFFFF")
    ws.cell(row=1, column=6).alignment = Alignment(horizontal="center")

    # Negrita en etiquetas de fila
    for row_idx in range(2, 5):
        cell = ws.cell(row=row_idx, column=1)
        cell.font      = Font(bold=True)
        cell.alignment = Alignment(horizontal="left")

    # Alineación derecha en datos numéricos
    for row_idx in range(2, 4):
        for col_idx in range(2, 7):
            ws.cell(row=row_idx, column=col_idx).alignment = Alignment(horizontal="right")

    # Autoajuste de columnas
    for col in ws.columns:
        max_len = max((len(str(cell.value or "")) for cell in col), default=10)
        ws.column_dimensions[get_column_letter(col[0].column)].width = max_len + 4

    wb.save(ruta_xlsx)
    print("  ✅ Hoja 'Dashboard_Resumen' actualizada en el Excel.")


# ===========================================================================
# 6. GENERACIÓN DE GRÁFICOS CON MATPLOTLIB
# ===========================================================================
def _buf_png(fig) -> io.BytesIO:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", dpi=150, facecolor="white")
    plt.close(fig)
    buf.seek(0)
    return buf


def grafico_donut(data_final: list, total_peticiones: int) -> io.BytesIO:
    fig, ax = plt.subplots(figsize=(3.8, 3.8), facecolor="white")
    vals = [d["val"] for d in data_final]
    cols = [d["color"] for d in data_final]
    if sum(vals) == 0:
        vals = [1] * len(vals)

    ax.pie(
        vals, colors=cols, startangle=90,
        wedgeprops=dict(width=0.42, edgecolor="white", linewidth=2),
    )
    ax.text(0, 0, str(total_peticiones), ha="center", va="center",
            fontsize=26, fontweight="bold", color="#1e293b", fontfamily="Calibri")
    ax.set_title("TOTAL PETICIONES", fontsize=9, fontweight="bold",
                 color="#475569", pad=12, fontfamily="Calibri")

    # Leyenda manual debajo
    legend_handles = [
        mpatches.Patch(color=d["color"], label=f"{d['cat']}: {d['val']}")
        for d in data_final
    ]
    ax.legend(
        handles=legend_handles,
        loc="lower center",
        bbox_to_anchor=(0.5, -0.22),
        ncol=2,
        fontsize=7,
        frameon=False,
    )
    ax.axis("equal")
    return _buf_png(fig)


def grafico_tiempos(data_final: list) -> io.BytesIO:
    fig, ax = plt.subplots(figsize=(5.5, 3.2), facecolor="white")
    cats   = [d["cat"]   for d in data_final]
    medias = [d["media"] for d in data_final]
    cols   = [d["color"] for d in data_final]
    y_pos  = range(len(cats))

    bars = ax.barh(y_pos, medias, color=cols, height=0.5, edgecolor="white")
    for bar, d in zip(bars, data_final):
        offset = max(medias) * 0.02 if max(medias) > 0 else 0.1
        ax.text(
            bar.get_width() + offset,
            bar.get_y() + bar.get_height() / 2,
            d["txt"], va="center", fontsize=8.5,
            fontweight="bold", color=d["color"],
        )

    ax.set_yticks(list(y_pos))
    ax.set_yticklabels(cats, fontsize=8, fontweight="bold")
    ax.set_title("TIEMPO MEDIO DE RESOLUCIÓN (HORAS)",
                 fontsize=9, fontweight="bold", color="#475569", pad=10)
    ax.spines[["top", "right", "bottom"]].set_visible(False)
    ax.xaxis.set_visible(False)
    ax.set_facecolor("white")
    fig.tight_layout()
    return _buf_png(fig)


def grafico_ratings(data_final: list) -> io.BytesIO:
    fig, ax = plt.subplots(figsize=(6, 2.8), facecolor="#f8fafc")
    cats    = [d["cat"]    for d in data_final]
    ratings = [d["rating"] for d in data_final]
    cols    = [d["color"]  for d in data_final]
    y_pos   = range(len(cats))

    ax.barh(y_pos, [5] * len(cats), color="#e2e8f0", height=0.42, zorder=1)
    bars = ax.barh(y_pos, ratings, color=cols, height=0.42, zorder=2)
    for bar, d in zip(bars, data_final):
        ax.text(
            5.15, bar.get_y() + bar.get_height() / 2,
            d["r_txt"], va="center", fontsize=8, fontweight="bold", color=d["color"],
        )

    ax.set_yticks(list(y_pos))
    ax.set_yticklabels(cats, fontsize=8, fontweight="bold")
    ax.set_xlim(0, 7)
    ax.set_title("VALORACIÓN ENCUESTAS POR TIPO",
                 fontsize=9, fontweight="bold", color="#475569", pad=10)
    ax.spines[["top", "right", "bottom"]].set_visible(False)
    ax.xaxis.set_visible(False)
    ax.set_facecolor("#f8fafc")
    fig.patch.set_facecolor("#f8fafc")
    fig.tight_layout()
    return _buf_png(fig)


# ===========================================================================
# 7. GENERACIÓN DE LA DIAPOSITIVA PPTX
# ===========================================================================
def _rect(slide, left, top, width, height, fill=None, line_color=None, line_width_pt=1.0):
    """Añade un rectángulo a la diapositiva."""
    shape = slide.shapes.add_shape(1, left, top, width, height)  # 1 = RECTANGLE
    if fill:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill
    else:
        shape.fill.background()
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width     = Pt(line_width_pt)
    else:
        shape.line.fill.background()
    return shape


def _textbox(slide, text, left, top, width, height,
             size=11, bold=False, color=None, align=PP_ALIGN.LEFT,
             wrap=True, font="Calibri"):
    color = color or RGBColor(0, 0, 0)
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = wrap
    p   = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text           = text
    run.font.size      = Pt(size)
    run.font.bold      = bold
    run.font.color.rgb = color
    run.font.name      = font
    return tb


def generar_pptx(
    data_final: list,
    total_peticiones: int,
    mes: str,
    anio: str,
    ruta_salida: str,
) -> None:
    prs = Presentation()
    prs.slide_width  = Inches(13.33)   # Widescreen 16:9
    prs.slide_height = Inches(7.5)

    slide = prs.slides.add_slide(prs.slide_layouts[6])   # Blank
    W = prs.slide_width
    H = prs.slide_height

    # ── CABECERA ─────────────────────────────────────────────────────────────
    HDR_H = Inches(0.72)
    _rect(slide, 0, 0, W, HDR_H, fill=RGBColor(30, 41, 59))

    _textbox(slide, "DASHBOARD OPERATIVO",
             Inches(0.35), Inches(0.1), Inches(7.5), Inches(0.52),
             size=22, bold=True, color=RGBColor(255, 255, 255),
             align=PP_ALIGN.LEFT, font="Calibri")

    _textbox(slide, f"{mes.upper()} {anio}",
             Inches(9.5), Inches(0.15), Inches(3.5), Inches(0.44),
             size=13, bold=True, color=RGBColor(148, 163, 184),
             align=PP_ALIGN.RIGHT, font="Calibri")

    # ── TILES DE CATEGORÍAS ──────────────────────────────────────────────────
    TILE_TOP = Inches(0.85)
    TILE_H   = Inches(0.92)
    GAP      = Inches(0.12)
    MARGIN   = Inches(0.35)
    TILE_W   = (W - MARGIN * 2 - GAP * (len(data_final) - 1)) / len(data_final)

    for i, d in enumerate(data_final):
        col  = hex_to_rgb_pptx(d["color"])
        left = MARGIN + i * (TILE_W + GAP)
        _rect(slide, left, TILE_TOP, TILE_W, TILE_H,
              fill=RGBColor(255, 255, 255), line_color=col, line_width_pt=2.2)
        # Número grande
        _textbox(slide, str(d["val"]),
                 left, TILE_TOP + Inches(0.04), TILE_W, Inches(0.48),
                 size=18, bold=True, color=col, align=PP_ALIGN.CENTER, wrap=False)
        # Nombre categoría
        _textbox(slide, d["cat"].upper(),
                 left, TILE_TOP + Inches(0.48), TILE_W, Inches(0.42),
                 size=6.5, bold=True, color=col, align=PP_ALIGN.CENTER, wrap=True)

    # ── SECCIÓN CENTRAL ───────────────────────────────────────────────────────
    CHARTS_TOP = Inches(1.95)
    CHARTS_H   = Inches(2.85)

    # Donut
    buf_donut = grafico_donut(data_final, total_peticiones)
    slide.shapes.add_picture(buf_donut,
                             MARGIN,           CHARTS_TOP,
                             Inches(4.0),      CHARTS_H)

    # Tiempos
    buf_tiempos = grafico_tiempos(data_final)
    slide.shapes.add_picture(buf_tiempos,
                             Inches(4.6),      CHARTS_TOP,
                             Inches(8.38),     CHARTS_H)

    # ── FONDO RATINGS ─────────────────────────────────────────────────────────
    RAT_TOP = CHARTS_TOP + CHARTS_H + Inches(0.12)
    RAT_H   = Inches(2.35)
    _rect(slide, MARGIN, RAT_TOP, W - MARGIN * 2, RAT_H,
          fill=RGBColor(248, 250, 252))

    buf_ratings = grafico_ratings(data_final)
    slide.shapes.add_picture(buf_ratings,
                             Inches(1.2), RAT_TOP + Inches(0.05),
                             Inches(10.9), RAT_H - Inches(0.1))

    prs.save(ruta_salida)
    print(f"  ✅ Diapositiva PowerPoint generada: {os.path.basename(ruta_salida)}")


# ===========================================================================
# 8. MAIN
# ===========================================================================
def main():
    parser = argparse.ArgumentParser(
        description="Genera resumen mensual y diapositiva PowerPoint (.pptx)."
    )
    parser.add_argument("--anio", required=True,
                        help="Año del informe (ej: 2026)")
    parser.add_argument("--mes",  required=True,
                        help="Mes del informe (ej: Abril)")
    args = parser.parse_args()

    anio = args.anio.strip()
    mes  = args.mes.strip().capitalize()

    if mes not in MESES_COD:
        print(f"\n❌ Mes '{mes}' no reconocido.")
        print(f"   Opciones válidas: {', '.join(MESES_COD.keys())}")
        sys.exit(1)

    cod         = MESES_COD[mes]
    nombre_base = f"{anio}_{cod}_{mes}_Informe Mensual Aprovisionamiento de Datos"
    if getattr(sys, 'frozen', False):
        # Ejecutando como .exe compilado
        carpeta = os.path.dirname(sys.executable)
    else:
        # Ejecutando como script .py normal
        carpeta = os.path.dirname(os.path.abspath(__file__))
    ruta_xlsx   = os.path.join(carpeta, f"{nombre_base}.xlsx")

    if not os.path.exists(ruta_xlsx):
        print(f"\n❌ Fichero no encontrado:\n   {ruta_xlsx}")
        print(f"   Asegúrate de que el fichero Excel está en la misma carpeta que el script.")
        sys.exit(1)

    print(f"\n⏳ Procesando {mes} {anio} ...")
    print(f"   Fichero: {ruta_xlsx}\n")

    data_final, total_peticiones, valores_peticiones, valores_horas = procesar_datos(ruta_xlsx)

    if not data_final:
        print("❌ No se encontraron datos para procesar. Revisa las hojas del Excel.")
        sys.exit(1)

    print(f"   Total peticiones procesadas: {total_peticiones}")
    for d in data_final:
        print(f"   · {d['cat']}: {d['val']} peticiones — {d['txt']} resolución media")

    print()
    actualizar_dashboard_resumen(
        ruta_xlsx, data_final, total_peticiones, valores_peticiones, valores_horas
    )

    nombre_pptx = f"{anio}_{cod}_{mes}_Dashboard_KPIs.pptx"
    ruta_pptx   = os.path.join(carpeta, nombre_pptx)
    generar_pptx(data_final, total_peticiones, mes, anio, ruta_pptx)

    print(f"\n✅ Proceso completado.")
    print(f"   Excel actualizado : {os.path.basename(ruta_xlsx)}")
    print(f"   PowerPoint creado : {nombre_pptx}")


if __name__ == "__main__":
    main()
