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
    pip install pandas openpyxl xlsxwriter python-pptx playwright
    playwright install chromium
"""

import argparse
import asyncio
import os
import sys

import pandas as pd
import openpyxl
import xlsxwriter
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
        })

    wb.close()
    return data_final, total_peticiones, valores_peticiones, valores_horas


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
        f'<div style="flex:1;height:65px;border:2px solid {d["color"]};border-radius:8px;display:flex;flex-direction:column;align-items:center;justify-content:center;background:#fff;">'
        f'<i class="fa-solid {d["icon"]}" style="font-size:16px;color:{d["color"]};margin-bottom:4px;"></i>'
        f'<div style="font-weight:800;color:{d["color"]};font-size:9px;letter-spacing:0.5px;text-align:center;">{d["cat"].upper()}</div></div>'
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
<head><link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css"></head>
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
        # channel="msedge" le dice a Playwright que use el Edge del SO en lugar
        # de un Chromium descargado aparte.
        browser = await p.chromium.launch(
            headless=True,
            channel="msedge",
        )
        page = await browser.new_page(viewport={"width": 1080, "height": 600})
        await page.set_content(html_content)
        await page.wait_for_timeout(2000)   # esperar a que cargue Font Awesome
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

    # Incrustar la imagen capturada ocupando casi toda la diapositiva
    slide_w = prs.slide_width
    slide_h = prs.slide_height
    margin  = Inches(0.1)
    slide.shapes.add_picture(
        ruta_png,
        left=margin, top=margin,
        width=slide_w - 2 * margin,
        height=slide_h - 2 * margin,
    )

    prs.save(ruta_pptx)
    print(f"  📊 PPTX guardado: {ruta_pptx}")


# =============================================================================
# PASO 6: XLSX DE RESUMEN con tabla + gráfico (xlsxwriter)
# =============================================================================

def generar_xlsx_resumen(ruta_salida, data_final, total_peticiones, valores_peticiones, valores_horas, mes_nombre, anio):
    wb = xlsxwriter.Workbook(ruta_salida)
    ws = wb.add_worksheet("Dashboard_Resumen")

    orden_cols = ORDEN_CATS

    def rgb_xlsw(r, g, b):
        return f"#{r:02X}{g:02X}{b:02X}"

    # --- Formatos ---
    fmt_hdr_negro = wb.add_format({
        "bold": True, "font_color": "white", "bg_color": "#000000",
        "align": "left", "border": 1
    })
    fmt_hdr_total = wb.add_format({
        "bold": True, "font_color": "white", "bg_color": "#334155",
        "align": "center", "border": 1
    })
    fmt_fila_lbl  = wb.add_format({"bold": True, "align": "left", "border": 1, "bg_color": "#f8fafc"})
    fmt_num       = wb.add_format({"num_format": "0.00", "align": "right", "border": 1})
    fmt_int       = wb.add_format({"num_format": "0",    "align": "right", "border": 1})
    fmt_txt       = wb.add_format({"align": "right", "border": 1})

    cat_fmts = {}
    for id_p in ORDEN:
        r, g, b = MAPA[id_p]["rgb"]
        cat_fmts[MAPA[id_p]["n"]] = wb.add_format({
            "bold": True, "font_color": "white",
            "bg_color": rgb_xlsw(r, g, b),
            "align": "center", "border": 1
        })

    # --- Cabecera ---
    ws.set_column(0, 0, 18)
    ws.set_column(1, 4, 22)
    ws.set_column(5, 5, 20)

    ws.write(0, 0, "Categoría",   fmt_hdr_negro)
    for ci, cat in enumerate(orden_cols):
        ws.write(0, ci + 1, cat, cat_fmts[cat])
    ws.write(0, 5, "TOTAL/MEDIA", fmt_hdr_total)

    # --- Fila tiempos ---
    ws.write(1, 0, "Tiempo (Horas)", fmt_fila_lbl)
    for ci, cat in enumerate(orden_cols):
        ws.write(1, ci + 1, valores_horas.get(cat, 0), fmt_num)
    ws.write_formula(1, 5, "=ROUND(AVERAGE(B2:E2),2)", fmt_num)

    # --- Fila peticiones ---
    ws.write(2, 0, "Peticiones", fmt_fila_lbl)
    for ci, cat in enumerate(orden_cols):
        ws.write(2, ci + 1, int(valores_peticiones.get(cat, 0)), fmt_int)
    ws.write_formula(2, 5, "=SUM(B3:E3)", fmt_int)

    # --- Fila formato horas/minutos ---
    ws.write(3, 0, "Formato", fmt_fila_lbl)
    for ci in range(4):
        col_letra = chr(ord("B") + ci)
        ws.write_formula(3, ci + 1,
            f'=INT({col_letra}2)&" horas "&ROUND(({col_letra}2-INT({col_letra}2))*60,0)&" minutos"',
            fmt_txt)
    ws.write_formula(3, 5,
        '=INT(F2)&" horas "&ROUND((F2-INT(F2))*60,0)&" minutos"', fmt_txt)

    # --- Gráfico de barras por categoría ---
    chart = wb.add_chart({"type": "column"})
    chart.set_title({"name": "Tiempo Medio (Horas)"})
    chart.set_x_axis({"name": "Categorías"})
    chart.set_y_axis({"name": "Horas"})
    chart.set_legend({"none": True})

    for ci, cat in enumerate(orden_cols):
        r, g, b = next(MAPA[k]["rgb"] for k in ORDEN if MAPA[k]["n"] == cat)
        chart.add_series({
            "name":       cat,
            "categories": ["Dashboard_Resumen", 0, ci + 1, 0, ci + 1],
            "values":     ["Dashboard_Resumen", 1, ci + 1, 1, ci + 1],
            "fill":       {"color": rgb_xlsw(r, g, b)},
            "data_labels": {
                "value": True,
                "num_format": "0.00",
                "font": {"bold": True, "color": "white", "size": 10},
                "position": "inside_end",
            },
        })

    ws.insert_chart("A6", chart, {"x_scale": 2.0, "y_scale": 1.5})
    wb.close()
    print(f"  📋 Excel resumen guardado: {ruta_salida}")


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

    # Generar xlsx de resumen
    ruta_resumen = os.path.join(carpeta, f"{anio}_{mes_cod}_{mes_nombre}_Dashboard_Resumen.xlsx")
    generar_xlsx_resumen(ruta_resumen, data_final, total_peticiones, valores_peticiones, valores_horas, mes_nombre, anio)

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
