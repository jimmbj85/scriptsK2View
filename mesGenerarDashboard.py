"""
Dashboard Mensual — versión LOCAL
===================================
Uso:
    python generarDashboardMes.py --anio 2026 --mes marzo

    o para varios meses a la vez:
    python generarDashboardMes.py --anio 2026 --mes enero febrero marzo

Estructura de carpetas (mismo directorio que el script):
    2026_01_Enero_Informe Mensual Aprovisionamiento de Datos.xlsx
    2026_02_Febrero_Informe Mensual Aprovisionamiento de Datos.xlsx
    ...

Salidas (misma carpeta):
    2026_01_Enero_Dashboard_KPIs_Final.pptx
    2026_01_Enero_Dashboard_Resumen.xlsx    <- tabla + gráfico actualizados

Dependencias:
    pip install pandas openpyxl python-pptx playwright
    # No se necesita instalar nada: usa Microsoft Edge del sistema
"""

import argparse
import asyncio
import os
import sys

import pandas as pd
import openpyxl
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from playwright.async_api import async_playwright

# =============================================================================
# CONFIGURACIÓN
# =============================================================================

MESES_NOMBRE = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
]
NOMBRE_A_COD = {n: f"{i+1:02d}" for i, n in enumerate(MESES_NOMBRE)}

ORDEN = ["1", "3", "5", "7"]
MAPA = {
    "1": {"n": "Búsqueda de datos",      "c": "#51a5e1", "rgb": (81,  165, 225), "i": "fa-magnifying-glass"},
    "3": {"n": "Bajada de datos",        "c": "#5a6b7d", "rgb": (90,  107, 125), "i": "fa-cloud-arrow-down"},
    "5": {"n": "Otras consultas",        "c": "#23a985", "rgb": (35,  169, 133), "i": "fa-comment-dots"},
    "7": {"n": "Cargas en Data Manager", "c": "#e75437", "rgb": (231,  84,  55), "i": "fa-gear"},
}
ORDEN_CATS = [MAPA[k]["n"] for k in ORDEN]

CAT_PREFIX = {
    "1-": MAPA["1"]["n"],
    "3-": MAPA["3"]["n"],
    "5-": MAPA["5"]["n"],
    "7-": MAPA["7"]["n"],
}

REPLACEMENTS_CAT = {
    "BUSQUEDA DE DATOS":      "Búsqueda de datos",
    "OTRAS CONSULTAS":        "Otras consultas",
    "BAJADA DE DATOS":        "Bajada de datos",
    "CARGAS EN DATA MANAGER": "Cargas en Data Manager",
}

def rgb_hex(r, g, b):
    return f"#{r:02X}{g:02X}{b:02X}"

# =============================================================================
# HELPERS
# =============================================================================

def tiempo_natural_sin_findes(row):
    inicio = pd.to_datetime(row["Submit Date"], errors="coerce")
    fin    = pd.to_datetime(row["Completed Date"], errors="coerce")
    if pd.isnull(inicio) or pd.isnull(fin) or fin < inicio:
        return 0.0
    segundos_totales = (fin - inicio).total_seconds()
    segundos_finde   = 0.0
    current = inicio
    while current < fin:
        next_day = (current + pd.Timedelta(days=1)).normalize()
        if next_day > fin:
            next_day = fin
        if current.weekday() >= 5:
            segundos_finde += (next_day - current).total_seconds()
        current = next_day
    return (segundos_totales - segundos_finde) / 3600.0


def normalizar_filas(filas, n_cols):
    """Recorta o rellena cada fila para que tenga exactamente n_cols columnas."""
    resultado = []
    for fila in filas:
        fila = list(fila)
        if len(fila) < n_cols:
            fila += [None] * (n_cols - len(fila))
        else:
            fila = fila[:n_cols]
        resultado.append(fila)
    return resultado


# =============================================================================
# PASO 1: PARSEO DE ARGUMENTOS
# =============================================================================

def parsear_args():
    parser = argparse.ArgumentParser(description="Genera el Dashboard Mensual en local.")
    parser.add_argument("--anio", required=True, type=int)
    parser.add_argument("--mes",  required=True, type=str, nargs="+",
                        help="Uno o varios meses (ej: enero  o  enero febrero marzo)")
    return parser.parse_args()


def resolver_meses(anio, nombres):
    meses = []
    for nombre in nombres:
        cap = nombre.strip().capitalize()
        if cap not in NOMBRE_A_COD:
            print(f"❌ Mes '{nombre}' no reconocido.")
            sys.exit(1)
        meses.append({"mes": cap, "cod": NOMBRE_A_COD[cap], "anio": str(anio)})
    return meses


# =============================================================================
# PASO 2: LECTURA DEL XLSX DEL MES
# =============================================================================

def encontrar_xlsx(carpeta, anio, cod, mes):
    """Busca el xlsx del mes por patrón {anio}_{cod} en el nombre."""
    for f in os.listdir(carpeta):
        if f.lower().endswith(".xlsx") and f"{anio}_{cod}" in f:
            # Excluir los que son salidas de este propio script
            if "Dashboard" not in f:
                return os.path.join(carpeta, f)
    return None


def leer_datos_mes(ruta_xlsx):
    """
    Lee las pestañas 1-, 3-, 5-, 7- y 9- del xlsx.
    Devuelve (data_final, ratings_dict).
    """
    wb = openpyxl.load_workbook(ruta_xlsx, read_only=True, data_only=True)

    # --- Ratings (pestaña 9-) ---
    ratings_dict = {}
    ws_9 = next((ws for ws in wb.worksheets if ws.title.startswith("9")), None)
    if ws_9:
        filas_9 = list(ws_9.iter_rows(values_only=True))
        if len(filas_9) > 1:
            cab = [str(c).strip() if c else "" for c in filas_9[0]]
            n   = len(cab)
            df9 = pd.DataFrame(normalizar_filas(filas_9[1:], n), columns=cab)
            # Renombrar columna de categoría si viene como Tier 3
            if "Categorization Tier 3" in df9.columns:
                df9 = df9.rename(columns={"Categorization Tier 3": "Category"})
            if "Category" in df9.columns:
                df9["Category"] = df9["Category"].astype(str).str.strip().replace(REPLACEMENTS_CAT)
            if "Rating" in df9.columns and "Category" in df9.columns:
                df9["Rating"] = pd.to_numeric(df9["Rating"], errors="coerce")
                summary = df9.groupby("Category")["Rating"].agg(["mean", "count"])
                ratings_dict = summary.to_dict("index")

    # --- Datos de categorías (1-, 3-, 5-, 7-) ---
    data_final         = []
    total_peticiones   = 0
    valores_horas      = {}
    valores_peticiones = {}

    for id_p in ORDEN:
        ws = next((w for w in wb.worksheets if w.title.startswith(id_p + "-")), None)
        if not ws:
            continue

        filas = list(ws.iter_rows(values_only=True))
        if len(filas) < 2:
            continue

        cabecera = [str(c).strip() if c is not None else "" for c in filas[0]]
        n_cols   = len(cabecera)
        df = pd.DataFrame(normalizar_filas(filas[1:], n_cols), columns=cabecera)

        if "Status" not in df.columns:
            continue
        df_flt = df[df["Status"].astype(str).str.strip().isin(["Terminado", "Cerrado"])].copy()
        df_flt = df_flt.dropna(subset=["Submit Date", "Completed Date"])

        df_flt["horas_res"] = df_flt.apply(tiempo_natural_sin_findes, axis=1)

        nombre_cat  = MAPA[id_p]["n"]
        count       = len(df_flt)
        media_horas = df_flt["horas_res"].mean() if count > 0 else 0.0

        total_peticiones += count
        valores_peticiones[nombre_cat] = float(count)
        valores_horas[nombre_cat]      = round(float(media_horas), 2)

        h_int = int(media_horas)
        m_int = int((media_horas % 1) * 60)

        stats       = ratings_dict.get(nombre_cat, {"mean": 0, "count": 0})
        rating_val  = stats["mean"] if stats["count"] > 0 else 0
        has_data    = stats["count"] > 0
        count_val   = int(stats["count"]) if has_data else 0
        rating_txt  = f"{rating_val:.1f}/5 ({count_val})" if has_data else "N/A"
        rating_barw = (rating_val * 20) if has_data else 0

        data_final.append({
            "cat":     nombre_cat,
            "val":     count,
            "media":   media_horas,
            "txt":     f"{h_int:02d}h {m_int:02d}m",
            "color":   MAPA[id_p]["c"],
            "rgb":     MAPA[id_p]["rgb"],
            "icon":    MAPA[id_p]["i"],
            "r_txt":   rating_txt,
            "r_width": rating_barw,
            "r_val":   round(rating_val, 2) if has_data else None,
            "r_count": count_val,
        })

    wb.close()
    return data_final, total_peticiones, valores_peticiones, valores_horas


SVG_ICONS = {
    "fa-magnifying-glass": '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512" style="width:16px;height:16px;fill:currentColor;margin-bottom:4px;"><path d="M416 208c0 45.9-14.9 88.3-40 122.7L502.6 457.4c12.5 12.5 12.5 32.8 0 45.3s-32.8 12.5-45.3 0L330.7 376c-34.4 25.2-76.8 40-122.7 40C93.1 416 0 322.9 0 208S93.1 0 208 0S416 93.1 416 208zM208 352a144 144 0 1 0 0-288 144 144 0 1 0 0 288z"/></svg>',
    "fa-cloud-arrow-down": '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 640 512" style="width:16px;height:16px;fill:currentColor;margin-bottom:4px;"><path d="M144 480C64.5 480 0 415.5 0 336c0-62.8 40.2-116.2 96.2-135.9c-.1-2.7-.2-5.4-.2-8.1c0-88.4 71.6-160 160-160c59.3 0 111 32.2 138.7 80.2C409.9 102 428.3 96 448 96c53 0 96 43 96 96c0 12.2-2.3 23.8-6.4 34.6C596 238.4 640 290.1 640 352c0 70.7-57.3 128-128 128H144zm79-167l80 80c9.4 9.4 24.6 9.4 33.9 0l80-80c9.4-9.4 9.4-24.6 0-33.9s-24.6-9.4-33.9 0l-39 39V184c0-13.3-10.7-24-24-24s-24 10.7-24 24V318.1l-39-39c-9.4-9.4-24.6-9.4-33.9 0s-9.4 24.6 0 33.9z"/></svg>',
    "fa-comment-dots":    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512" style="width:16px;height:16px;fill:currentColor;margin-bottom:4px;"><path d="M256 448c141.4 0 256-93.1 256-208S397.4 32 256 32S0 125.1 0 240c0 45.1 17.7 86.8 47.7 120.9c-1.9 24.5-11.4 46.3-21.4 62.9c-5.5 9.2-11.1 16.6-15.2 21.6c-2.1 2.5-3.7 4.4-4.9 5.7c-.6 .6-1 1.1-1.3 1.4l-.3 .3c0 0 0 0 0 0c0 0 0 0 0 0s0 0 0 0s0 0 0 0c-4.6 4.6-5.9 11.4-3.4 17.4c2.5 6 8.3 9.9 14.8 9.9c28.7 0 57.6-8.9 81.6-19.3c22.9-10 42.4-21.9 54.3-30.6c31.8 11.5 67 17.9 104.1 17.9zM128 272c-17.7 0-32-14.3-32-32s14.3-32 32-32s32 14.3 32 32s-14.3 32-32 32zm128 0c-17.7 0-32-14.3-32-32s14.3-32 32-32s32 14.3 32 32s-14.3 32-32 32zm160-32c0 17.7-14.3 32-32 32s-32-14.3-32-32s14.3-32 32-32s32 14.3 32 32z"/></svg>',
    "fa-gear":            '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512" style="width:16px;height:16px;fill:currentColor;margin-bottom:4px;"><path d="M495.9 166.6c3.2 8.7 .5 18.4-6.4 24.6l-43.3 39.4c1.1 8.3 1.7 16.8 1.7 25.4s-.6 17.1-1.7 25.4l43.3 39.4c6.9 6.2 9.6 15.9 6.4 24.6c-4.4 11.9-9.7 23.3-15.8 34.3l-4.7 8.1c-6.6 11-14 21.4-22.1 31.2c-5.9 7.2-15.7 9.6-24.5 6.8l-55.7-17.7c-13.4 10.3-28.2 18.9-44 25.4l-12.5 57.1c-2 9.1-9 16.3-18.2 17.8c-13.8 2.3-28 3.5-42.5 3.5s-28.7-1.2-42.5-3.5c-9.2-1.5-16.2-8.7-18.2-17.8l-12.5-57.1c-15.8-6.5-30.6-15.1-44-25.4L83.1 425.9c-8.8 2.8-18.6 .3-24.5-6.8c-8.1-9.8-15.5-20.2-22.1-31.2l-4.7-8.1c-6.1-11-11.4-22.4-15.8-34.3c-3.2-8.7-.5-18.4 6.4-24.6l43.3-39.4C64.6 273.1 64 264.6 64 256s.6-17.1 1.7-25.4L22.4 191.2c-6.9-6.2-9.6-15.9-6.4-24.6c4.4-11.9 9.7-23.3 15.8-34.3l4.7-8.1c6.6-11 14-21.4 22.1-31.2c5.9-7.2 15.7-9.6 24.5-6.8l55.7 17.7c13.4-10.3 28.2-18.9 44-25.4l12.5-57.1c2-9.1 9-16.3 18.2-17.8C227.3 1.2 241.5 0 256 0s28.7 1.2 42.5 3.5c9.2 1.5 16.2 8.7 18.2 17.8l12.5 57.1c15.8 6.5 30.6 15.1 44 25.4l55.7-17.7c8.8-2.8 18.6-.3 24.5 6.8c8.1 9.8 15.5 20.2 22.1 31.2l4.7 8.1c6.1 11 11.4 22.4 15.8 34.3zM256 336a80 80 0 1 0 0-160 80 80 0 1 0 0 160z"/></svg>',
}

# =============================================================================
# PASO 3: GENERACIÓN DEL HTML DEL DASHBOARD
# =============================================================================

def construir_html(data_final, total_peticiones, mes_nombre, anio):
    max_h = max((d["media"] for d in data_final), default=1) or 1

    # Gradiente para el donut
    stops, current_deg = [], 0
    for d in data_final:
        pct = (d["val"] / total_peticiones * 100) if total_peticiones > 0 else 0
        stops.append(f"{d['color']} {current_deg}% {current_deg + pct}%")
        current_deg += pct
    gradient = ",".join(stops)

    tiles_html = "".join([
        f'<div style="flex:1;height:65px;border:2px solid {d["color"]};border-radius:8px;display:flex;flex-direction:column;align-items:center;justify-content:center;background:#fff;color:{d["color"]};">'
        + SVG_ICONS.get(d["icon"], "")
        + f'<div style="font-weight:800;color:{d["color"]};font-size:9px;letter-spacing:0.5px;text-align:center;">{d["cat"].upper()}</div></div>'
        for d in data_final
    ])

    legend_grid = "".join([
        f'<div style="display:flex;align-items:center;gap:6px;">'
        f'<div style="width:10px;height:10px;background-color:{d["color"]};border-radius:2px;"></div>'
        f'<div style="font-size:11px;color:{d["color"]};font-weight:700;white-space:nowrap;">{d["cat"]}: <b>{d["val"]}</b></div></div>'
        for d in data_final
    ])

    barras_tiempo = "".join([
        f'<div style="display:flex;align-items:center;margin-bottom:12px;">'
        f'<div style="width:160px;font-size:11px;font-weight:700;color:{d["color"]};white-space:nowrap;">{d["cat"]}</div>'
        f'<div style="flex-grow:1;background:#f1f5f9;height:10px;border-radius:5px;overflow:hidden;">'
        f'<div style="width:{int(d["media"]/max_h*100) if max_h>0 else 0}%;background:{d["color"]};height:100%;"></div></div>'
        f'<div style="width:60px;text-align:right;font-size:11px;font-weight:700;color:{d["color"]};">{d["txt"]}</div></div>'
        for d in data_final
    ])

    ratings_vertical = "".join([
        f'<div style="display:flex;align-items:center;margin-bottom:12px;width:100%;">'
        f'<div style="width:160px;font-size:11px;font-weight:800;color:{d["color"]};text-transform:uppercase;text-align:right;padding-right:15px;white-space:nowrap;">{d["cat"]}</div>'
        f'<div style="flex-grow:1;background:#e2e8f0;height:12px;border-radius:6px;overflow:hidden;">'
        f'<div style="width:{d["r_width"]}%;background:{d["color"]};height:100%;"></div></div>'
        f'<div style="width:60px;text-align:left;font-size:11px;font-weight:700;color:{d["color"]};padding-left:10px;">{d["r_txt"]}</div></div>'
        for d in data_final
    ])

    return f"""<html>
<head><style>body{{margin:0;padding:0;}}</style></head>
<body style="margin:0;padding:20px;background:#fff;width:1080px;height:600px;font-family:sans-serif;box-sizing:border-box;">
  <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:15px;border-bottom:2px solid #f1f5f9;padding-bottom:10px;">
    <div style="font-weight:800;font-size:20px;color:#1e293b;">DASHBOARD OPERATIVO</div>
    <div style="font-weight:600;font-size:14px;color:#64748b;text-transform:uppercase;">{mes_nombre} {anio}</div>
  </div>
  <div style="display:flex;gap:10px;margin-bottom:15px;">{tiles_html}</div>
  <div style="display:flex;gap:20px;height:240px;margin-bottom:20px;">
    <div style="width:35%;border:1px solid #e2e8f0;border-radius:12px;padding:15px;display:flex;flex-direction:column;align-items:center;justify-content:space-between;">
      <div style="font-weight:800;color:#475569;font-size:10px;">TOTAL PETICIONES</div>
      <div style="width:110px;height:110px;border-radius:50%;background:conic-gradient({gradient});display:flex;align-items:center;justify-content:center;position:relative;">
        <div style="width:80px;height:80px;background:white;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:20px;font-weight:800;color:#1e293b;">{total_peticiones}</div>
      </div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;width:100%;">{legend_grid}</div>
    </div>
    <div style="flex:1;border:1px solid #e2e8f0;border-radius:12px;padding:20px;display:flex;flex-direction:column;justify-content:center;">
      <div style="font-weight:800;color:#475569;font-size:10px;margin-bottom:15px;">TIEMPO MEDIO RESOLUCIÓN (HORAS)</div>
      {barras_tiempo}
    </div>
  </div>
  <div style="width:65%;margin:0 auto;border:1px solid #e2e8f0;border-radius:12px;padding:20px;background:#f8fafc;display:flex;flex-direction:column;justify-content:center;">
    <div style="font-weight:800;color:#475569;font-size:11px;margin-bottom:15px;text-align:center;">VALORACIÓN ENCUESTAS POR TIPO</div>
    {ratings_vertical}
  </div>
</body></html>"""


# =============================================================================
# PASO 4: CAPTURA HTML → PNG con Playwright
# =============================================================================

async def capturar_png(html_content, ruta_png):
    async with async_playwright() as p:
        # Usar Microsoft Edge instalado en el sistema (no requiere instalar nada extra).
        browser = await p.chromium.launch(
            headless=True,
            channel="msedge",
        )
        page = await browser.new_page(viewport={"width": 1080, "height": 600})
        # Todo el contenido es inline (SVGs, estilos) — no hay red que esperar.
        await page.set_content(html_content, wait_until="domcontentloaded")
        await page.screenshot(path=ruta_png)
        await browser.close()
    print(f"  📸 PNG generado: {ruta_png}")


# =============================================================================
# PASO 5: POWERPOINT con la imagen incrustada
# =============================================================================

def generar_pptx(ruta_png, ruta_pptx, mes_nombre, anio):
    prs  = Presentation()

    # Diapositiva en blanco (layout 6 = blank en la mayoría de temas)
    blank = next((l for l in prs.slide_layouts if l.name == "Blank"), prs.slide_layouts[6])
    slide = prs.slides.add_slide(blank)

    # Fondo blanco
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    # Incrustar la imagen manteniendo el ratio original del PNG (1080×600 = 1.8:1).
    # Forzar width Y height al mismo tiempo estira el donut; aquí se ajusta por
    # ancho y se centra verticalmente el espacio sobrante.
    from PIL import Image as PilImage
    with PilImage.open(ruta_png) as _im:
        img_w_px, img_h_px = _im.size
    img_ratio = img_w_px / img_h_px  # ej. 1080/600 = 1.8

    slide_w = prs.slide_width
    slide_h = prs.slide_height
    margin  = Inches(0.1)

    avail_w = slide_w - 2 * margin
    avail_h = slide_h - 2 * margin

    # Ajustar por ancho; si la altura resultante supera la disponible, ajustar por alto
    pic_w = avail_w
    pic_h = int(pic_w / img_ratio)
    if pic_h > avail_h:
        pic_h = avail_h
        pic_w = int(pic_h * img_ratio)

    left = margin + (avail_w - pic_w) // 2   # centrado horizontal
    top  = margin + (avail_h - pic_h) // 2   # centrado vertical

    slide.shapes.add_picture(
        ruta_png,
        left=left, top=top,
        width=pic_w,
        height=pic_h,
    )

    prs.save(ruta_pptx)
    print(f"  📊 PPTX guardado: {ruta_pptx}")


# =============================================================================
# PASO 6: XLSX DE RESUMEN con tabla + gráfico (openpyxl nativo)
# =============================================================================

def generar_xlsx_resumen(ruta_xlsx_origen, data_final, total_peticiones, valores_peticiones, valores_horas, mes_nombre, anio):
    """
    Abre el xlsx del informe mensual y añade/reemplaza la hoja Dashboard_Resumen.
    Usa openpyxl (API nativa) para escribir la tabla y el gráfico de columnas.
    """
    import zipfile, shutil
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
    from openpyxl.chart import BarChart, Reference
    from openpyxl.chart.series import SeriesLabel
    from openpyxl.chart.label import DataLabelList
    from openpyxl.utils import get_column_letter
    from lxml import etree

    orden_cols = ORDEN_CATS

    # ── Helpers de estilo ────────────────────────────────────────────────────
    def solid_fill(hex6):
        return PatternFill(fill_type="solid", fgColor=hex6.lstrip("#"))

    def thin_border():
        s = Side(style="thin")
        return Border(left=s, right=s, top=s, bottom=s)

    # ── Abrir workbook destino y (re)crear la hoja ───────────────────────────
    wb = openpyxl.load_workbook(ruta_xlsx_origen)
    if "Dashboard_Resumen" in wb.sheetnames:
        del wb["Dashboard_Resumen"]
    ws = wb.create_sheet("Dashboard_Resumen")

    # ── Anchos de columna ────────────────────────────────────────────────────
    ws.column_dimensions["A"].width = 18
    for ci in range(5):
        ws.column_dimensions[get_column_letter(ci + 2)].width = 22

    # ── Fila 1: cabecera ─────────────────────────────────────────────────────
    ws["A1"].value     = "Categoría"
    ws["A1"].font      = Font(bold=True, color="FFFFFF")
    ws["A1"].fill      = solid_fill("000000")
    ws["A1"].border    = thin_border()
    ws["A1"].alignment = Alignment(horizontal="left")

    for ci, cat in enumerate(orden_cols):
        id_p = ORDEN[ci]
        r, g, b = MAPA[id_p]["rgb"]
        c = ws.cell(row=1, column=ci + 2)
        c.value     = cat
        c.font      = Font(bold=True, color="FFFFFF")
        c.fill      = solid_fill(f"{r:02X}{g:02X}{b:02X}")
        c.border    = thin_border()
        c.alignment = Alignment(horizontal="center")

    ws["F1"].value     = "TOTAL/MEDIA"
    ws["F1"].font      = Font(bold=True, color="FFFFFF")
    ws["F1"].fill      = solid_fill("334155")
    ws["F1"].border    = thin_border()
    ws["F1"].alignment = Alignment(horizontal="center")

    # ── Fila 2: Tiempo (Horas) ───────────────────────────────────────────────
    ws["A2"].value     = "Tiempo (Horas)"
    ws["A2"].font      = Font(bold=True)
    ws["A2"].fill      = solid_fill("f8fafc")
    ws["A2"].border    = thin_border()
    ws["A2"].alignment = Alignment(horizontal="left")

    for ci, cat in enumerate(orden_cols):
        c = ws.cell(row=2, column=ci + 2)
        c.value        = valores_horas.get(cat, 0)
        c.number_format = "0.00"
        c.alignment    = Alignment(horizontal="right")
        c.border       = thin_border()

    ws["F2"].value        = "=ROUND(AVERAGE(B2:E2),2)"
    ws["F2"].number_format = "0.00"
    ws["F2"].alignment    = Alignment(horizontal="right")
    ws["F2"].border       = thin_border()

    # ── Fila 3: Peticiones ───────────────────────────────────────────────────
    ws["A3"].value     = "Peticiones"
    ws["A3"].font      = Font(bold=True)
    ws["A3"].fill      = solid_fill("f8fafc")
    ws["A3"].border    = thin_border()
    ws["A3"].alignment = Alignment(horizontal="left")

    for ci, cat in enumerate(orden_cols):
        c = ws.cell(row=3, column=ci + 2)
        c.value        = int(valores_peticiones.get(cat, 0))
        c.number_format = "0"
        c.alignment    = Alignment(horizontal="right")
        c.border       = thin_border()

    ws["F3"].value        = "=SUM(B3:E3)"
    ws["F3"].number_format = "0"
    ws["F3"].alignment    = Alignment(horizontal="right")
    ws["F3"].border       = thin_border()

    # ── Fila 4: Formato legible ──────────────────────────────────────────────
    ws["A4"].value     = "Formato"
    ws["A4"].font      = Font(bold=True)
    ws["A4"].fill      = solid_fill("f8fafc")
    ws["A4"].border    = thin_border()
    ws["A4"].alignment = Alignment(horizontal="left")

    for ci in range(4):
        col_l = get_column_letter(ci + 2)
        c = ws.cell(row=4, column=ci + 2)
        c.value     = f'=INT({col_l}2)&" horas "&ROUND(({col_l}2-INT({col_l}2))*60,0)&" minutos"'
        c.alignment = Alignment(horizontal="right")
        c.border    = thin_border()

    ws["F4"].value     = '=INT(F2)&" horas "&ROUND((F2-INT(F2))*60,0)&" minutos"'
    ws["F4"].alignment = Alignment(horizontal="right")
    ws["F4"].border    = thin_border()

    # ── Fila 5: Valoración media (numérica) ──────────────────────────────────
    ws["A5"].value     = "Valoración Media"
    ws["A5"].font      = Font(bold=True)
    ws["A5"].fill      = solid_fill("f8fafc")
    ws["A5"].border    = thin_border()
    ws["A5"].alignment = Alignment(horizontal="left")

    valores_rating = {}
    for d in data_final:
        valores_rating[d["cat"]] = d.get("r_val")

    for ci, cat in enumerate(orden_cols):
        c = ws.cell(row=5, column=ci + 2)
        val = valores_rating.get(cat)
        if val is not None:
            c.value        = val
            c.number_format = "0.00"
        else:
            c.value = "N/A"
        c.alignment = Alignment(horizontal="right")
        c.border    = thin_border()

    # TOTAL/MEDIA de valoración: media ponderada real (suma(val*count) / suma(count))
    # Se calcula en Python con los datos disponibles para evitar fórmulas array complejas
    # con celdas "N/A" mezcladas con números.
    vals_num   = [(d["r_val"], d["r_count"]) for d in data_final if d["r_val"] is not None and d["r_count"] > 0]
    if vals_num:
        media_pond = round(sum(v * c for v, c in vals_num) / sum(c for _, c in vals_num), 2)
        ws["F5"].value        = media_pond
        ws["F5"].number_format = "0.00"
    else:
        ws["F5"].value = "N/A"
    ws["F5"].alignment = Alignment(horizontal="right")
    ws["F5"].border    = thin_border()

    # ── Fila 6: Valoraciones (nº total de encuestas por tipo) ────────────────
    ws["A6"].value     = "Valoraciones"
    ws["A6"].font      = Font(bold=True)
    ws["A6"].fill      = solid_fill("f8fafc")
    ws["A6"].border    = thin_border()
    ws["A6"].alignment = Alignment(horizontal="left")

    for d in data_final:
        ci = orden_cols.index(d["cat"])
        c  = ws.cell(row=6, column=ci + 2)
        c.value        = d["r_count"]
        c.number_format = "0"
        c.alignment    = Alignment(horizontal="right")
        c.border       = thin_border()

    ws["F6"].value        = "=SUM(B6:E6)"
    ws["F6"].number_format = "0"
    ws["F6"].alignment    = Alignment(horizontal="right")
    ws["F6"].border       = thin_border()

    # ── Helpers de anclaje exacto (TwoCellAnchor: fila/col 0-based) ────────────
    from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor, AnchorMarker
    from openpyxl.drawing.xdr import XDRPositiveSize2D

    def _anchor(col_from, row_from, col_to, row_to):
        """Devuelve un TwoCellAnchor que ocupa exactamente las celdas indicadas (0-based)."""
        anchor = TwoCellAnchor()
        anchor._from = AnchorMarker(col=col_from, row=row_from, colOff=0, rowOff=0)
        anchor.to    = AnchorMarker(col=col_to,   row=row_to,   colOff=0, rowOff=0)
        return anchor

    # ── Gráfico 1: Tiempo Medio (Horas) — columnas A–D, filas 8–20 ─────────
    # AnchorMarker es 0-based: col A=0, fila 8=7 (índice); fila 20=19 (límite to)
    chart = BarChart()
    chart.type         = "col"
    chart.grouping     = "clustered"
    chart.title        = "Tiempo Medio (Horas)"
    chart.y_axis.title = "Horas"
    chart.x_axis.title = "Categorías"

    cats = Reference(ws, min_col=2, max_col=5, min_row=1, max_row=1)

    for i, id_p in enumerate(ORDEN):
        r, g, b  = MAPA[id_p]["rgb"]
        hex_rgb  = f"{r:02X}{g:02X}{b:02X}"
        col      = i + 2
        cat_name = MAPA[id_p]["n"]

        values = Reference(ws, min_col=col, max_col=col, min_row=2, max_row=2)
        chart.add_data(values)

        serie = chart.series[i]
        serie.title = SeriesLabel(v=cat_name)
        serie.graphicalProperties.solidFill      = hex_rgb
        serie.graphicalProperties.line.solidFill = hex_rgb

        dLbls = DataLabelList()
        dLbls.showVal          = True
        dLbls.dLblPos          = "inEnd"
        dLbls.showLegendKey    = False
        dLbls.showCatName      = False
        dLbls.showSerName      = False
        dLbls.showPercent      = False
        dLbls.showBubbleSize   = False
        serie.dLbls = dLbls

    chart.set_categories(cats)
    chart.legend = None
    chart.anchor = _anchor(col_from=0, row_from=7, col_to=4, row_to=19)  # A8:E20
    ws.add_chart(chart)

    # ── Gráfico 2: Valoración Media — columnas E en adelante, filas 8–20 ────
    # Solo se dibuja si al menos una categoría tiene valoración numérica.
    cats_con_rating = [d for d in data_final if d.get("r_val") is not None]
    if cats_con_rating:
        chart2 = BarChart()
        chart2.type         = "col"
        chart2.grouping     = "clustered"
        chart2.title        = "Valoración Media (sobre 5)"
        chart2.y_axis.title = "Puntuación"
        chart2.x_axis.title = "Categorías"

        cats2 = Reference(ws, min_col=2, max_col=5, min_row=1, max_row=1)

        for i, id_p in enumerate(ORDEN):
            r, g, b  = MAPA[id_p]["rgb"]
            hex_rgb  = f"{r:02X}{g:02X}{b:02X}"
            col      = i + 2
            cat_name = MAPA[id_p]["n"]

            values2 = Reference(ws, min_col=col, max_col=col, min_row=5, max_row=5)
            chart2.add_data(values2)

            serie2 = chart2.series[i]
            serie2.title = SeriesLabel(v=cat_name)
            serie2.graphicalProperties.solidFill      = hex_rgb
            serie2.graphicalProperties.line.solidFill = hex_rgb

            dLbls2 = DataLabelList()
            dLbls2.showVal          = True
            dLbls2.dLblPos          = "inEnd"
            dLbls2.showLegendKey    = False
            dLbls2.showCatName      = False
            dLbls2.showSerName      = False
            dLbls2.showPercent      = False
            dLbls2.showBubbleSize   = False
            serie2.dLbls = dLbls2

        chart2.set_categories(cats2)
        chart2.legend = None
        chart2.anchor = _anchor(col_from=4, row_from=7, col_to=9, row_to=19)  # E8:J20
        ws.add_chart(chart2)

    wb.save(ruta_xlsx_origen)
    wb.close()

    # ── Post-proceso ZIP: inyectar fuente blanca negrita en <c:dLbls> ────────
    # serie._element no existe en series creadas programáticamente (solo en las
    # leídas desde XML), así que la fuente se añade directamente sobre el chart
    # XML dentro del ZIP una vez guardado.
    _patch_chart_labels_font(ruta_xlsx_origen)

    print(f"  📋 Hoja Dashboard_Resumen actualizada en: {os.path.basename(ruta_xlsx_origen)}")


def _patch_chart_labels_font(xlsx_path):
    """
    Abre el xlsx como ZIP, localiza todos los chart*.xml y añade fuente blanca
    negrita a cada nodo <c:dLbls> que no tenga ya un <c:txPr>.
    """
    import zipfile
    from lxml import etree

    _nsA = "http://schemas.openxmlformats.org/drawingml/2006/main"
    _nsC = "http://schemas.openxmlformats.org/drawingml/2006/chart"

    def _txPr_blanco():
        txPr = etree.Element(f"{{{_nsC}}}txPr")
        bodyPr = etree.SubElement(txPr, f"{{{_nsA}}}bodyPr")
        bodyPr.set("rot", "0")
        etree.SubElement(txPr, f"{{{_nsA}}}lstStyle")
        p   = etree.SubElement(txPr, f"{{{_nsA}}}p")
        pPr = etree.SubElement(p,   f"{{{_nsA}}}pPr")
        dPr = etree.SubElement(pPr, f"{{{_nsA}}}defRPr")
        dPr.set("b",    "1")
        dPr.set("sz",   "1000")   # 10 pt
        dPr.set("lang", "es-ES")
        sf   = etree.SubElement(dPr,  f"{{{_nsA}}}solidFill")
        srgb = etree.SubElement(sf,   f"{{{_nsA}}}srgbClr")
        srgb.set("val", "FFFFFF")
        return txPr

    tmp_path = xlsx_path + "._patch.tmp"
    with zipfile.ZipFile(xlsx_path, "r") as zin,          zipfile.ZipFile(tmp_path,  "w", zipfile.ZIP_DEFLATED) as zout:

        for item in zin.infolist():
            data = zin.read(item.filename)
            if (item.filename.startswith("xl/charts/chart")
                    and item.filename.endswith(".xml")):
                root = etree.fromstring(data)
                for dLbls in root.iter(f"{{{_nsC}}}dLbls"):
                    if dLbls.find(f"{{{_nsC}}}txPr") is None:
                        # Insertar txPr justo después de numFmt si existe
                        nf = dLbls.find(f"{{{_nsC}}}numFmt")
                        txPr = _txPr_blanco()
                        if nf is not None:
                            nf.addnext(txPr)
                        else:
                            dLbls.insert(0, txPr)
                data = etree.tostring(
                    root, xml_declaration=True, encoding="UTF-8", standalone=True
                )
            zout.writestr(item, data)

    os.replace(tmp_path, xlsx_path)



# =============================================================================
# MAIN
# =============================================================================

async def procesar_mes(carpeta, config):
    mes_nombre = config["mes"]
    mes_cod    = config["cod"]
    anio       = config["anio"]

    print(f"\n⏳ Procesando {mes_nombre} {anio}...")

    # Buscar xlsx de origen
    ruta_xlsx = encontrar_xlsx(carpeta, anio, mes_cod, mes_nombre)
    if not ruta_xlsx:
        print(f"  ❌ No se encontró el xlsx para {anio}_{mes_cod}. Archivos disponibles:")
        print(f"     {[f for f in os.listdir(carpeta) if f.endswith('.xlsx')]}")
        return

    print(f"  📂 Leyendo: {os.path.basename(ruta_xlsx)}")

    # Leer datos
    data_final, total_peticiones, valores_peticiones, valores_horas = leer_datos_mes(ruta_xlsx)

    if not data_final:
        print(f"  ⚠️  Sin datos procesables en {mes_nombre}.")
        return

    # Generar HTML
    html_content = construir_html(data_final, total_peticiones, mes_nombre, anio)

    # Guardar HTML (opcional, para depuración)
    ruta_html = os.path.join(carpeta, f"dash_{mes_cod}.html")
    with open(ruta_html, "w", encoding="utf-8") as f:
        f.write(html_content)

    # Capturar PNG
    ruta_png = os.path.join(carpeta, f"dash_{mes_cod}.png")
    await capturar_png(html_content, ruta_png)

    # Generar PPTX con la imagen
    ruta_pptx = os.path.join(carpeta, f"{anio}_{mes_cod}_{mes_nombre}_Dashboard_KPIs_Final.pptx")
    generar_pptx(ruta_png, ruta_pptx, mes_nombre, anio)

    # Actualizar hoja Dashboard_Resumen dentro del xlsx de origen
    generar_xlsx_resumen(ruta_xlsx, data_final, total_peticiones, valores_peticiones, valores_horas, mes_nombre, anio)

    # Limpiar temporales
    for tmp in [ruta_html, ruta_png]:
        if os.path.exists(tmp):
            os.remove(tmp)

    print(f"  ✅ {mes_nombre} {anio} completado.")


async def main():
    args   = parsear_args()
    meses  = resolver_meses(args.anio, args.mes)

    if getattr(sys, "frozen", False):
        carpeta = os.path.dirname(os.path.abspath(sys.executable))
    else:
        carpeta = os.path.dirname(os.path.abspath(__file__))

    print(f"\n🚀 Iniciando — {args.anio}, meses: {[m['mes'] for m in meses]}")
    print(f"   Carpeta: {carpeta}\n")

    for config in meses:
        await procesar_mes(carpeta, config)

    print(f"\n🎉 Todo listo.")


if __name__ == "__main__":
    asyncio.run(main())
