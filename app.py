import os
import re
import io
import math
import tempfile
import openpyxl
import pypdf
import pytesseract
from PIL import Image, ImageSequence
import streamlit as st

st.set_page_config(
    page_title="Validador de Precalificación CUB",
    page_icon="📊",
    layout="wide"
)

def convertir_vaciado_pdf_a_workbook(ruta_pdf):
    wb_virtual = openpyxl.Workbook()
    ws_res = wb_virtual.active
    ws_res.title = "Edo. Res."
    ws_bal = wb_virtual.create_sheet("Bal. Gral.")

    reader = pypdf.PdfReader(ruta_pdf)
    texto_total = "\n".join([(p.extract_text() or "") for p in reader.pages])
    lineas = [l.strip() for l in texto_total.split("\n") if l.strip()]

    fechas = re.findall(r"(?:diciembre|junio|julio|marzo|septiembre)\s+20\d{2}|20\d{2}-(?:12-31|06-30|07-31)", texto_total, re.IGNORECASE)
    fechas_unicas = list(dict.fromkeys(fechas))[:4] or ["2025-12-31"]
    for idx_f, f_txt in enumerate(fechas_unicas):
        col_f = 6 + (idx_f * 2)
        ws_res.cell(4, col_f, f_txt)
        ws_bal.cell(4, col_f, f_txt)

    r_res, r_bal = 7, 9
    en_balance = False
    for linea in lineas:
        low = linea.lower()
        if any(k in low for k in ["balance general", "estado de situación financiera", "activos circulantes", "activo total"]):
            en_balance = True
        elif any(k in low for k in ["estado de resultados", "ventas netas", "costo de ventas"]):
            en_balance = False

        linea_sin_pct = re.sub(r"-?\d+(?:\.\d+)?\s*%", "", linea)
        nums = re.findall(r"-?\d{1,3}(?:,\d{3})*(?:\.\d+)?", linea_sin_pct)
        concepto = re.sub(r"[-$\d,.\s]+$", "", linea_sin_pct).strip()

        if concepto and len(concepto) > 2 and nums:
            vals = [float(n.replace(",", "")) for n in nums]
            if any(abs(v) > 0 for v in vals):
                ws_dest = ws_bal if en_balance else ws_res
                fila_dest = r_bal if en_balance else r_res
                ws_dest.cell(fila_dest, 1, concepto)
                for i_v, v_num in enumerate(vals[:len(fechas_unicas)]):
                    ws_dest.cell(fila_dest, 6 + (i_v * 2), v_num)
                if en_balance:
                    r_bal += 1
                else:
                    r_res += 1
    return wb_virtual


def convertir_modelo_pdf_a_workbook(ruta_pdf):
    wb_virtual = openpyxl.Workbook()
    reader = pypdf.PdfReader(ruta_pdf)
    texto_total = "\n".join([(p.extract_text() or "") for p in reader.pages])
    nombre_hoja = "Pymes(Anexo 21)" if "anexo 21" in texto_total.lower() else "Corporativos(Anexo 22)"
    ws = wb_virtual.active
    ws.title = nombre_hoja

    m_nom = re.search(r"Nombre:\s*([^\n]+)", texto_total, re.IGNORECASE)
    if m_nom:
        ws["B4"] = f"Nombre: {m_nom.group(1).strip()}"
    for sec in ["Manufactura", "Comercio", "Servicios", "Agrícola", "Explotación, Energía y Construcción"]:
        if sec.lower() in texto_total.lower():
            ws["C10"] = sec
            break
    m_ventas = re.search(r"Ventas\s+Netas[^\d]*([\d,]+(?:\.\d+)?)", texto_total, re.IGNORECASE)
    if m_ventas:
        ws["G18"] = float(m_ventas.group(1).replace(",", ""))
    m_udi = re.search(r"UDI[^\d]*(\d+\.\d+)", texto_total, re.IGNORECASE)
    if m_udi:
        ws["G19"] = float(m_udi.group(1))
    return wb_virtual


def mapear_scian_a_sector_cub(codigo):
    cod_limpio = re.sub(r"\D", "", str(codigo or ""))
    if len(cod_limpio) < 2:
        return (None, "Código inválido o vacío")
    prefijo = cod_limpio[:2]
    reglas = {
        "11": ("Agrícola", "Clave 11: Agricultura, ganadería, silvicultura, pesca y caza"),
        "21": ("Explotación, Energía y Construcción", "Clave 21: Minería"),
        "22": ("Explotación, Energía y Construcción", "Clave 22: Electricidad, agua y gas"),
        "23": ("Explotación, Energía y Construcción", "Clave 23: Construcción"),
        "31": ("Manufactura", "Clave 31: Industrias manufactureras"),
        "32": ("Manufactura", "Clave 32: Industrias manufactureras"),
        "33": ("Manufactura", "Clave 33: Industrias manufactureras"),
        "43": ("Comercio", "Clave 43: Comercio al por mayor"),
        "46": ("Comercio", "Clave 46: Comercio al por menor"),
    }
    if prefijo in reglas:
        return reglas[prefijo]
    elif prefijo in ["48", "49", "51", "52", "53", "54", "55", "56", "61", "62", "71", "72", "81", "93"]:
        return ("Servicios", f"Clave {prefijo}: Servicios")
    return (None, "Clave no catalogada en CUB")


def evaluar_intervalo(val, rango_str):
    txt = str(rango_str).strip()
    if not txt or "sin informaci" in txt.lower():
        return False
    es_pct = "%" in txt
    v = val * 100.0 if es_pct else float(val)
    nums = [float(x) for x in re.findall(r"-?\d+(?:\.\d+)?", txt)]
    if not nums:
        return False
    if len(nums) >= 2:
        a, b = nums[0], nums[1]
        incluye_izq = not txt.lstrip().startswith(("(", "（"))
        incluye_der = txt.rstrip().endswith(("]", "］"))
        ok_izq = (v >= a - 1e-6) if incluye_izq else (v > a + 1e-6)
        ok_der = (v <= b + 1e-6) if incluye_der else (v < b - 1e-6)
        return ok_izq and ok_der
    lim = nums[0]
    if any(op in txt for op in ["≤", "<="]):
        return v <= lim + 1e-6
    if "<" in txt:
        return v < lim - 1e-6
    if any(op in txt for op in ["≥", ">="]):
        return v >= lim + 1e-6
    if ">" in txt:
        return v > lim + 1e-6
    if lim in (1.0, 100.0):
        return abs(val - 1.0) < 1e-6 or abs(val - 100.0) < 1e-6
    return False


def escribir_celda_segura(ws, fila, col, valor):
    for rango in ws.merged_cells.ranges:
        if rango.min_row <= fila <= rango.max_row and rango.min_col <= col <= rango.max_col:
            ws.cell(rango.min_row, rango.min_col).value = valor
            return
    ws.cell(fila, col).value = valor


def simular_escenario_desde_comentario(wb_valores, hoja_activa, sector_eval, comentario_analista):
    txt = comentario_analista.strip()
    if not txt:
        return None
    ws_m = wb_valores[hoja_activa]
    fecha_bc = str(ws_m["C5"].value or "")[:10]
    calif_act = str(ws_m["G11"].value or "").strip()
    res_act_val = ws_m["G10"].value
    res_act_txt = f"{float(res_act_val)*100:.2f}%" if isinstance(res_act_val, (int, float)) else str(res_act_val or "")

    fila_enc = 26 if "Corporativos" in hoja_activa else 25
    col_rango = None
    for c in range(1, ws_m.max_column + 1):
        if sector_eval.lower()[:6] in str(ws_m.cell(fila_enc, c).value or "").lower():
            col_rango = c
            break
    if not col_rango:
        return None

    col_num = col_rango - 2
    col_pts = col_rango + 4
    low = txt.lower()

    def max_pts_tabla(f_ini):
        m_pts = -9999
        rt = f_ini
        while ws_m.cell(rt, col_rango).value is not None:
            op = str(ws_m.cell(rt, col_rango).value).lower()
            p = ws_m.cell(rt, col_rango + 1).value
            if isinstance(p, (int, float)) and "sin informaci" not in op and p > m_pts:
                m_pts = p
            rt += 1
        return m_pts if m_pts > -9999 else 0

    filas_1a = [r for r in range(26, 44) if ws_m.cell(r, col_num).value in [1, 2, 3, 1.0, 2.0, 3.0]]
    ganancias_1a = {}
    for r_ind in filas_1a:
        num_i = int(ws_m.cell(r_ind, col_num).value)
        pts_actual = float(ws_m.cell(r_ind, col_pts).value or 0)
        pts_max = max_pts_tabla(r_ind)
        ganancias_1a[num_i] = max(0.0, pts_max - pts_actual)

    es_quita = any(k in low for k in ["quita", "castigo", "reestructura", "quitas"])
    solo_no_bancos = any(k in low for k in ["no banc", "proveedor", "comercial", "bardahl"])
    solo_bancos = any(k in low for k in ["banc", "banco", "tarjeta", "crédito bancario"])

    if es_quita and 3 in ganancias_1a and ganancias_1a[3] > 0:
        delta_puntos = ganancias_1a[3]
        motivo_cambio = "la aclaración/eliminación de la clave de quita o reestructura en su buró"
    elif solo_no_bancos and ganancias_1a.get(2, 0) > 0:
        delta_puntos = ganancias_1a[2]
        motivo_cambio = "la actualización de su buró y con ello se elimina el atraso presentado"
    elif solo_bancos and ganancias_1a.get(1, 0) > 0:
        delta_puntos = ganancias_1a[1]
        motivo_cambio = "la actualización de su buró y con ello se elimina el atraso presentado"
    else:
        delta_puntos = sum(ganancias_1a.values())
        motivo_cambio = "la actualización de su buró y con ello se elimina el atraso presentado"

    if "Pymes" in hoja_activa:
        puntaje_act = float(ws_m["T5"].value or 0)
        peso_cuanti = 1.0
    else:
        fila_tot = 84
        for r_chk in range(75, 90):
            if "total de puntaje" in str(ws_m.cell(r_chk, col_rango + 1).value or "").lower():
                fila_tot = r_chk
                break
        puntaje_act = float(ws_m.cell(fila_tot, col_pts).value or 0)
        peso_cuanti = 0.85 if "mayores o iguales a 54" in str(ws_m["C14"].value or "").lower() else 1.0

    puntaje_nuevo = puntaje_act + (delta_puntos * peso_cuanti)
    pi_nueva = 1.0 / (1.0 + math.exp(-(500.0 - puntaje_nuevo) * (math.log(2.0) / 40.0)))
    sp_val = float(ws_m["G6"].value or 0.55)
    res_nueva_dec = pi_nueva * sp_val
    res_nueva_pct = res_nueva_dec * 100.0

    calif_nueva = calif_act
    if "Grado de Riesgo" in wb_valores.sheetnames:
        ws_gr = wb_valores["Grado de Riesgo"]
        for rg in range(4, 13):
            grado_txt = str(ws_gr.cell(rg, 1).value or "").strip()
            lim_sup = ws_gr.cell(rg, 6).value
            if grado_txt and isinstance(lim_sup, (int, float)):
                if res_nueva_dec <= float(lim_sup) + 1e-7:
                    calif_nueva = grado_txt
                    break

    return {
        "fecha_bc": fecha_bc,
        "calif_act": calif_act,
        "res_act": res_act_txt,
        "calif_nueva": calif_nueva,
        "res_nueva": f"{res_nueva_pct:.2f}%",
        "puntaje_act": round(puntaje_act, 1),
        "puntaje_nuevo": round(puntaje_nuevo, 1),
        "delta": round(delta_puntos, 1),
        "motivo_cambio": motivo_cambio,
    }


def redactar_correo_alison(cliente_nombre, errores, ws_valores, esc_buro=None):
    empresa_corta = re.sub(
        r",?\s*(S\.?A\.?\s*de\s*C\.?V\.?|S\.?\s*de\s*R\.?L\.?\s*de\s*C\.?V\.?|S\.?A\.?P\.?I\.?\s*de\s*C\.?V\.?)",
        "", cliente_nombre, flags=re.IGNORECASE
    ).strip()

    lineas_correo = [
        f"Asunto: Validación precalificación - {empresa_corta}",
        "",
        "Hola [Nombre], buen día",
        "",
    ]

    if not errores:
        lineas_correo.append("La precalificación se llenó correctamente")
        if esc_buro:
            f_tag = f" consultada el {esc_buro['fecha_bc']}" if esc_buro["fecha_bc"] else ""
            lineas_correo.append(
                f"Sobre el tema que mencionas, considerando tal cual la información de buró{f_tag} que es la que me compartes, "
                f"el cliente tendría una calificación de {esc_buro['calif_act']} con porcentaje de reserva del {esc_buro['res_act']}. "
                f"Sin embargo, si se considera lo que comenta el ejecutivo sobre {esc_buro['motivo_cambio']}, "
                f"el cliente tendría una calificación de {esc_buro['calif_nueva']} con porcentaje de reserva del {esc_buro['res_nueva']}."
            )
        else:
            calif = str(ws_valores["G11"].value or "").strip()
            res_val = ws_valores["G10"].value
            res_txt = f"{float(res_val)*100:.2f}%" if isinstance(res_val, (int, float)) else str(res_val or "")
            if calif and res_txt:
                lineas_correo.append(f"La calificación es {calif} con un % reserva del {res_txt}.")
        lineas_correo.extend(["", "Saludos."])
        return "\n".join(lineas_correo)

    lineas_correo.append("Adjunto precalificación junto con sus comentarios y/o modificaciones:")
    lineas_correo.append("")

    grupos = {"Datos Generales": [], "1A": [], "1B": [], "1C": [], "Cualitativos": [], "Otros": []}

    def agregar_a_grupo(categoria, texto_item):
        texto_limpio = texto_item.strip().rstrip(".")
        if texto_limpio and texto_limpio not in grupos[categoria]:
            grupos[categoria].append(texto_limpio)

    for e in errores:
        mot = e["motivo"]
        fila_e = e.get("fila")
        if "Sector" in mot and "C10" in mot:
            agregar_a_grupo("Datos Generales", f"Sector ({e['nuevo_val']})")
        elif "Celda G18" in mot:
            agregar_a_grupo("Datos Generales", f"Ventas Netas (${float(e['nuevo_val']):,.2f})")
        elif "Anexo Incorrecto" in mot:
            agregar_a_grupo("Datos Generales", f"Pestaña normativa ({e['nuevo_val'] or 'Anexo correspondiente según UDIs'})")
        elif "Factor 1-A" in mot:
            if "No Bancos" in mot:
                agregar_a_grupo("1A", "No Bancos")
            elif "Bancos" in mot:
                agregar_a_grupo("1A", "Bancos")
            else:
                agregar_a_grupo("1A", "Indicador de Buró")
        elif "Factor 1-B" in mot:
            m_1b = re.search(r"Factor 1-B \('([^']+)'\)", mot)
            agregar_a_grupo("1B", m_1b.group(1) if m_1b else "Indicador de experiencia interna")
        elif "1-C" in mot or (fila_e and 55 <= fila_e <= 85):
            if "Insumo '" not in mot:
                m_ind = re.search(r"'([^']+)'", mot)
                agregar_a_grupo("1C", m_ind.group(1) if m_ind else "Indicador financiero")
        elif fila_e and 94 <= fila_e <= 120:
            if "F108" in mot or "auditad" in mot.lower():
                agregar_a_grupo("Cualitativos", "Estados financieros auditados")
            elif "F112" in mot or "gobierno" in mot.lower() or "organizacional" in mot.lower():
                agregar_a_grupo("Cualitativos", "Estructura organizacional")
            elif "F117" in mot or "accionaria" in mot.lower():
                agregar_a_grupo("Cualitativos", "Composición accionaria")
            else:
                agregar_a_grupo("Cualitativos", mot)
        else:
            agregar_a_grupo("Otros", mot)

    if grupos["Datos Generales"]:
        lineas_correo.append(f"  • Datos Generales: {', '.join(grupos['Datos Generales'])}")
    if grupos["1A"]:
        if len(grupos["1A"]) >= 2:
            lineas_correo.append("  • 1A: Ambos indicadores, se actualiza conforme a Buró de Crédito")
        else:
            lineas_correo.append(f"  • 1A ({grupos['1A'][0]}): se ajusta conforme a Buró de Crédito")
    if grupos["1B"]:
        if len(grupos["1B"]) >= 2:
            lineas_correo.append("  • 1B: Ambos indicadores, se pone el mejor puntaje")
        else:
            lineas_correo.append(f"  • 1B ({grupos['1B'][0]}): se pone el mejor puntaje")
    if grupos["1C"]:
        lineas_correo.append(f"  • 1C: {', '.join(grupos['1C'])}")
    if grupos["Cualitativos"]:
        lineas_correo.append(f"  • Cualitativos: {', '.join(grupos['Cualitativos'])}")
    for item_otro in grupos["Otros"]:
        lineas_correo.append(f"  • {item_otro}")

    if esc_buro:
        f_tag = f" consultada el {esc_buro['fecha_bc']}" if esc_buro["fecha_bc"] else ""
        lineas_correo.extend([
            "",
            f"Sobre el tema que mencionas, considerando tal cual la información de buró{f_tag}, el cliente tendría una calificación de "
            f"{esc_buro['calif_act']} con porcentaje de reserva del {esc_buro['res_act']}. Sin embargo, si se considera lo que comenta "
            f"el ejecutivo sobre {esc_buro['motivo_cambio']}, el cliente tendría una calificación de {esc_buro['calif_nueva']} con "
            f"porcentaje de reserva del {esc_buro['res_nueva']}."
        ])

    lineas_correo.extend(["", "Saludos."])
    return "\n".join(lineas_correo)


def ejecutar_auditoria_completa(path_vac, path_buro, path_pre, path_tif, comentario_analista):
    logs_pasos = []
    errores_llenado = []
    alertas_sin_impacto = []

    def registrar_error(hoja, fila, col, nuevo_val, motivo):
        for e in errores_llenado:
            if e["hoja"] == hoja and e["fila"] == fila and e["col"] == col and e["motivo"] == motivo:
                return
        errores_llenado.append({"hoja": hoja, "fila": fila, "col": col, "nuevo_val": nuevo_val, "motivo": motivo})

    if path_vac.lower().endswith(".pdf"):
        wb_vac = convertir_vaciado_pdf_a_workbook(path_vac)
    else:
        wb_vac = openpyxl.load_workbook(path_vac, data_only=True)

    reader_bc = pypdf.PdfReader(path_buro)
    texto_buro_completo = "\n".join([(p.extract_text() or "") for p in reader_bc.pages])

    texto_tif_paginas = []
    if path_tif and os.path.exists(path_tif):
        with Image.open(path_tif) as img:
            for frame in ImageSequence.Iterator(img):
                texto_tif_paginas.append(pytesseract.image_to_string(frame.convert("L"), lang="spa"))
    texto_tif_completo = "\n".join(texto_tif_paginas)

    es_pre_pdf = path_pre.lower().endswith(".pdf")
    if es_pre_pdf:
        wb_pre_val = convertir_modelo_pdf_a_workbook(path_pre)
        wb_pre_form = wb_pre_val
    else:
        wb_pre_val = openpyxl.load_workbook(path_pre, data_only=True)
        wb_pre_form = openpyxl.load_workbook(path_pre, data_only=False)

    hoja_modelo = None
    for candidata in ["Pymes(Anexo 21)", "Corporativos(Anexo 22)", "Entidades Financieras(Anexo 20)"]:
        if candidata in wb_pre_val.sheetnames and wb_pre_val[candidata]["C6"].value not in [None, 0, "0", ""]:
            hoja_modelo = candidata
            break
    if not hoja_modelo:
        hoja_modelo = next((c for c in ["Corporativos(Anexo 22)", "Pymes(Anexo 21)"] if c in wb_pre_val.sheetnames), wb_pre_val.sheetnames[0])

    ws_val = wb_pre_val[hoja_modelo]
    ws_form = wb_pre_form[hoja_modelo]
    nombre_cliente = str(ws_val["B4"].value or "Cliente").replace("Nombre:", "").strip()
    logs_pasos.append(f"**[Paso 1]** Empresa detectada: `{nombre_cliente}` | Pestaña activa: `{hoja_modelo}`")

    tabla_scian = re.search(
        r"C[oó]digo\s+SCIAN\s+Descripci[oó]n(.*?)(?:Previsor|Mensajes|INFORMACI[OÓ]N CREDITICIA|Cr[eé]ditos Activos)",
        texto_buro_completo, re.DOTALL | re.IGNORECASE
    )
    sectores_validos_buro = []
    codigos_buro = []
    if tabla_scian:
        pares_scian = re.findall(r"\b(\d{5,6})\b", tabla_scian.group(1))
        for cod in pares_scian:
            sec_cub, _ = mapear_scian_a_sector_cub(cod)
            if sec_cub:
                sectores_validos_buro.append(sec_cub)
                codigos_buro.append(cod)

    sector_excel = str(ws_val["C10"].value or "").strip()
    fila_cod_excel = 14 if "Pymes" in hoja_modelo else 15
    codigo_excel_num = re.sub(r"\D", "", str(ws_val[f"C{fila_cod_excel}"].value or ""))
    sector_desde_codigo_excel, _ = mapear_scian_a_sector_cub(codigo_excel_num)
    sector_activo = sector_excel

    if not sector_excel:
        msg = f"[Hoja: {hoja_modelo} | Celda C10] El campo 'Sector' está vacío."
        registrar_error(hoja_modelo, 10, 3, sector_desde_codigo_excel or (sectores_validos_buro[0] if sectores_validos_buro else "Comercio"), msg)
    elif sector_desde_codigo_excel and sector_excel.lower() != sector_desde_codigo_excel.lower():
        msg = f"[Hoja: {hoja_modelo} | Celda C10] Inconsistencia interna: C{fila_cod_excel} tiene '{codigo_excel_num}' ({sector_desde_codigo_excel}), pero C10 dice '{sector_excel}'."
        registrar_error(hoja_modelo, 10, 3, sector_desde_codigo_excel, msg)
    elif sectores_validos_buro and sector_excel not in sectores_validos_buro:
        logs_pasos.append(f"**[Paso 2 ⚠️]** Excel dice `{sector_excel}`, Buró reporta `{', '.join(codigos_buro)}` ({sectores_validos_buro[0]}). Se respeta Excel.")
    else:
        logs_pasos.append(f"**[Paso 2 ✅]** Sector `{sector_activo}` validado correctamente.")

    hoja_edo_res = next((wb_vac[h] for h in wb_vac.sheetnames if h.strip().lower() in ["edo. res.", "edo res", "estado de resultados", "resultados"]), wb_vac[wb_vac.sheetnames[0]])
    fila_ventas = next((r for r in range(1, 20) if str(hoja_edo_res.cell(r, 1).value or "").strip().lower() in ["ventas netas", "ventas"]), None)
    ventas_vac, col_periodo_res, fecha_periodo_res = None, None, ""

    if fila_ventas:
        cols_anuales, cols_interinas = [], []
        for c in range(2, 11):
            val_v = hoja_edo_res.cell(fila_ventas, c).value
            enc_fecha = hoja_edo_res.cell(4, c).value
            if isinstance(val_v, (int, float)) and abs(float(val_v)) > 10 and enc_fecha is not None:
                if "12-31" in str(enc_fecha).lower() or "diciembre" in str(enc_fecha).lower() or getattr(enc_fecha, "month", None) == 12:
                    cols_anuales.append(c)
                else:
                    cols_interinas.append(c)
        col_periodo_res = cols_anuales[-1] if cols_anuales else (cols_interinas[-1] if cols_interinas else None)
        if col_periodo_res:
            ventas_vac = abs(float(hoja_edo_res.cell(fila_ventas, col_periodo_res).value))
            fecha_periodo_res = str(hoja_edo_res.cell(4, col_periodo_res).value)[:10]

    ventas_modelo = float(ws_val["G18"].value or 0)
    udi_modelo = float(ws_val["G19"].value or 0)

    if ventas_vac is not None:
        if abs(ventas_vac - ventas_modelo) > 1.5:
            msg = f"[Hoja: {hoja_modelo} | Celda G18] Dice '{ventas_modelo:,.2f}', pero en Edo. Res. ({fecha_periodo_res}) reporta '{ventas_vac:,.2f}'."
            registrar_error(hoja_modelo, 18, 7, round(ventas_vac, 2), msg)
        if udi_modelo > 0:
            ventas_udis = ventas_vac / udi_modelo
            if ventas_udis < 14000 and "Corporativos" in hoja_modelo:
                registrar_error(hoja_modelo, None, None, None, f"[Anexo Incorrecto] Equivale a {ventas_udis:,.2f} miles de UDIs (<14M). Debe usarse 'Pymes(Anexo 21)'.")
            elif ventas_udis >= 14000 and "Pymes" in hoja_modelo:
                registrar_error(hoja_modelo, None, None, None, f"[Anexo Incorrecto] Equivale a {ventas_udis:,.2f} miles de UDIs (>=14M). Debe usarse 'Corporativos(Anexo 22)'.")
        logs_pasos.append(f"**[Paso 3]** Ventas Vaciado: `${ventas_vac:,.2f}` vs Modelo: `${ventas_modelo:,.2f}`.")

    def extraer_metrica_buro(nombre_campo, texto):
        m = re.search(rf"{nombre_campo}\s+(\d+(?:\.\d+)?)", texto)
        return float(m.group(1)) if m else None

    bk_sat = extraer_metrica_buro("BK12_PCT_SAT", texto_buro_completo)
    nbk_sat = extraer_metrica_buro("NBK12_PCT_SAT", texto_buro_completo)
    total_cred_bancarios = (extraer_metrica_buro("BK12_NUM_CRED", texto_buro_completo) or 0) + (extraer_metrica_buro("BK12_NUM_TC_ACT", texto_buro_completo) or 0)
    nbk_num = extraer_metrica_buro("NBK12_NUM_CRED", texto_buro_completo) or 0

    fila_encabezado = 26 if "Corporativos" in hoja_modelo else 25
    col_sector_inicio = next((c for c in range(1, ws_val.max_column + 1) if sector_activo.lower()[:6] in str(ws_val.cell(fila_encabezado, c).value or "").lower()), None)

    def obtener_mejor_opcion_tabla(fila_inicio, col_rango_idx):
        mejor_op, mejor_pts, r_t = None, -9999, fila_inicio
        while ws_val.cell(r_t, col_rango_idx).value is not None:
            op_txt = str(ws_val.cell(r_t, col_rango_idx).value).strip()
            pts_val = ws_val.cell(r_t, col_rango_idx + 1).value
            if isinstance(pts_val, (int, float)) and "sin informaci" not in op_txt.lower() and pts_val > mejor_pts:
                mejor_pts, mejor_op = pts_val, op_txt
            r_t += 1
        return mejor_op, mejor_pts

    if col_sector_inicio:
        col_rango, col_sel, col_pts = col_sector_inicio, col_sector_inicio + 3, col_sector_inicio + 4
        letra_sel = openpyxl.utils.get_column_letter(col_sel)
        fila_1a_bk = 28 if "Corporativos" in hoja_modelo else 27
        fila_1a_nbk = 35 if "Corporativos" in hoja_modelo else 34

        sel_bk = str(ws_val.cell(fila_1a_bk, col_sel).value or "").strip()
        sel_nbk = str(ws_val.cell(fila_1a_nbk, col_sel).value or "").strip()
        mejor_bk, pts_bk = obtener_mejor_opcion_tabla(fila_1a_bk, col_rango)
        mejor_nbk, pts_nbk = obtener_mejor_opcion_tabla(fila_1a_nbk, col_rango)

        if total_cred_bancarios > 0 and bk_sat == 1.0 and mejor_bk and sel_bk != mejor_bk:
            registrar_error(hoja_modelo, fila_1a_bk, col_sel, mejor_bk, f"[Hoja: {hoja_modelo} | Celda {letra_sel}{fila_1a_bk}] Factor 1-A Bancos dice '{sel_bk}', debe ser '{mejor_bk}' ({pts_bk} pts).")
        if nbk_num > 0 and nbk_sat == 1.0 and mejor_nbk and sel_nbk != mejor_nbk:
            registrar_error(hoja_modelo, fila_1a_nbk, col_sel, mejor_nbk, f"[Hoja: {hoja_modelo} | Celda {letra_sel}{fila_1a_nbk}] Factor 1-A No Bancos dice '{sel_nbk}', debe ser '{mejor_nbk}' ({pts_nbk} pts).")

        filas_1b = [44, 50] if "Corporativos" in hoja_modelo else [48, 53, 54]
        for r_1b in filas_1b:
            num_ind = ws_val.cell(r_1b, col_rango - 2).value
            nom_ind = str(ws_val.cell(r_1b, col_rango - 1).value or "").strip()
            if num_ind in [4, 5, 4.0, 5.0] and nom_ind:
                sel_1b = str(ws_val.cell(r_1b, col_sel).value or "").strip()
                mejor_op_1b, mejor_pts_1b = obtener_mejor_opcion_tabla(r_1b, col_rango)
                pts_act_1b = ws_val.cell(r_1b, col_pts).value
                if mejor_op_1b and sel_1b != mejor_op_1b and pts_act_1b != mejor_pts_1b:
                    registrar_error(hoja_modelo, r_1b, col_sel, mejor_op_1b, f"[Hoja: {hoja_modelo} | Celda {letra_sel}{r_1b}] Factor 1-B ('{nom_ind}'): Tiene '{sel_1b}', debe ser '{mejor_op_1b}'.")

    if "Corporativos" in hoja_modelo and col_sector_inicio:
        hoja_bal_gral = next((wb_vac[h] for h in wb_vac.sheetnames if h.strip().lower() in ["bal. gral.", "bal gral", "balance general"]), None)
        if hoja_bal_gral and col_periodo_res:
            def indexar(ws_o):
                m = {}
                for r in range(1, ws_o.max_row + 1):
                    ta = str(ws_o.cell(r, 1).value or "").strip()
                    tb = str(ws_o.cell(r, 2).value or "").strip()
                    nom = ta if ta and not ta.replace(".", "").isdigit() else tb
                    if nom:
                        vals = {c: float(ws_o.cell(r, c).value) for c in range(2, ws_o.max_column + 1) if isinstance(ws_o.cell(r, c).value, (int, float))}
                        m.setdefault(re.sub(r"\s+", " ", nom.lower()).strip(), []).append((r, vals))
                return m

            mapa_res, mapa_bal = indexar(hoja_edo_res), indexar(hoja_bal_gral)
            col_rango = col_sector_inicio
            col_num, col_nombre = col_rango - 2, col_rango - 1
            col_sel, col_calc, col_insumo = col_rango + 3, col_rango + 6, col_rango + 7
            col_etiqueta = col_insumo + 1

            filas_ind_1c = [r for r in range(55, 85) if isinstance(ws_val.cell(r, col_num).value, int) and ws_val.cell(r, col_nombre).value]
            reglas_ins = [
                (["ingresos anuales", "ventas netas"], mapa_res, ["ventas netas", "ventas"]),
                (["costo de ventas"], mapa_res, ["costo de ventas"]),
                (["gasto por interes"], mapa_res, ["interés pagado", "interes pagado", "gastos financieros"]),
                (["utilidad bruta"], mapa_res, ["utilidad bruta"]),
                (["activo total"], mapa_bal, ["activo total", "total activo"]),
                (["activo circulante"], mapa_bal, ["total activos circulantes", "total activo circulante"]),
                (["pasivo circulante"], mapa_bal, ["total pasivos circulantes", "total pasivo circulante"]),
                (["efectivo"], mapa_bal, ["efectivo", "efectivo y equivalentes"]),
                (["inversiones a corto plazo", "inversiones temporales"], mapa_bal, ["inversiones temporales", "inversiones a corto plazo"]),
                (["clientes", "cuentas y documentos por cobrar"], mapa_bal, ["clientes"]),
                (["proveedores", "acreedores"], mapa_bal, ["proveedores"]),
            ]

            insumos_corr, discrepancias = {}, []
            for r_ins in range(55, 85):
                et_raw = str(ws_val.cell(r_ins, col_etiqueta).value or "").strip()
                v_ex = ws_val.cell(r_ins, col_insumo).value
                if et_raw and isinstance(v_ex, (int, float)):
                    coord_i = f"{openpyxl.utils.get_column_letter(col_insumo)}{r_ins}"
                    for claves, mp, noms in reglas_ins:
                        if any(k in et_raw.lower() for k in claves):
                            val_v = next((abs(cd[col_periodo_res]) for nm in noms if nm in mp for fr, cd in mp[nm] if fr < 88 and col_periodo_res in cd), None)
                            if val_v is not None and abs(float(v_ex) - val_v) > 1.5:
                                insumos_corr[coord_i] = val_v
                                r_padre = max([ri for ri in filas_ind_1c if ri <= r_ins], default=None)
                                discrepancias.append({"fila": r_ins, "col": col_insumo, "coord": coord_i, "etiqueta": et_raw, "val_excel": float(v_ex), "val_v": val_v, "r_padre": r_padre})
                            break

            ind_err = set()
            for idx, r_ind in enumerate(filas_ind_1c):
                nom_ind = ws_val.cell(r_ind, col_nombre).value
                r_sig = filas_ind_1c[idx + 1] if (idx + 1) < len(filas_ind_1c) else 85
                fila_sel = r_ind if ws_val.cell(r_ind, col_sel).value is not None else (r_ind + 1)
                sel_act = str(ws_val.cell(fila_sel, col_sel).value or "").strip()
                f_str = next((str(ws_form.cell(fc, col_calc).value or "").strip() for fc in range(r_ind, r_sig) if str(ws_form.cell(fc, col_calc).value or "").startswith("=") and "VLOOKUP" not in str(ws_form.cell(fc, col_calc).value or "").upper()), "")
                val_calc = None
                if f_str:
                    try:
                        expr = f_str.lstrip("=")
                        expr_num = re.sub(r"\b[A-Z]+\d+\b", lambda m: str(insumos_corr[m.group(0).upper()]) if m.group(0).upper() in insumos_corr else (str(float(ws_val[m.group(0).upper()].value)) if isinstance(ws_val[m.group(0).upper()].value, (int, float)) else "0.0"), expr, flags=re.IGNORECASE)
                        val_calc = float(eval(expr_num, {"__builtins__": None}, {}))
                    except Exception:
                        pass
                if val_calc is not None:
                    rango_esp = next((str(ws_val.cell(rt, col_rango).value).strip() for rt in range(r_ind, r_sig) if ws_val.cell(rt, col_rango).value and evaluar_intervalo(val_calc, str(ws_val.cell(rt, col_rango).value))), None)
                    if rango_esp and sel_act != rango_esp:
                        ind_err.add(r_ind)
                        coord_s = f"{openpyxl.utils.get_column_letter(col_sel)}{fila_sel}"
                        registrar_error(hoja_modelo, fila_sel, col_sel, rango_esp, f"[Hoja: {hoja_modelo} | Celda {coord_s}] '{nom_ind}': Tiene '{sel_act}', debe ser '{rango_esp}'.")

            for d in discrepancias:
                if d["r_padre"] not in ind_err:
                    alertas_sin_impacto.append(f"Celda {d['coord']} ({d['etiqueta']}): Excel={d['val_excel']:,.2f} vs Vaciado={d['val_v']:,.2f} (sin impacto en rango CUB).")
                else:
                    registrar_error(hoja_modelo, d["fila"], d["col"], round(d["val_v"], 2), f"[Hoja: {hoja_modelo} | Celda {d['coord']}] Insumo '{d['etiqueta']}': Excel={d['val_excel']:,.2f} vs Vaciado={d['val_v']:,.2f}.")

    escenario = simular_escenario_desde_comentario(wb_pre_val, hoja_modelo, sector_activo, comentario_analista)
    correo_final = redactar_correo_alison(nombre_cliente, errores_llenado, ws_val, escenario)

    excel_corregido_bytes = None
    if errores_llenado and not es_pre_pdf:
        wb_salida = openpyxl.load_workbook(path_pre, data_only=False)
        for e in errores_llenado:
            if e["fila"] and e["col"] and e["nuevo_val"] is not None:
                escribir_celda_segura(wb_salida[e["hoja"]], e["fila"], e["col"], e["nuevo_val"])
        buffer = io.BytesIO()
        wb_salida.save(buffer)
        excel_corregido_bytes = buffer.getvalue()

    calif_final = str(ws_val["G11"].value or "N/D").strip()
    res_val = ws_val["G10"].value
    res_final = f"{float(res_val)*100:.2f}%" if isinstance(res_val, (int, float)) else str(res_val or "N/D")

    return {
        "nombre_cliente": nombre_cliente,
        "hoja_modelo": hoja_modelo,
        "calif": calif_final,
        "reserva": res_final,
        "errores": errores_llenado,
        "alertas": alertas_sin_impacto,
        "logs": logs_pasos,
        "correo": correo_final,
        "excel_bytes": excel_corregido_bytes,
        "escenario": escenario,
    }


st.title("📊 Validador Automático de Precalificación (CUB)")
st.markdown("Sube los archivos del expediente para validar automáticamente el llenado del modelo y generar el dictamen.")

col1, col2, col3, col4 = st.columns(4)
with col1:
    up_vaciado = st.file_uploader("1️⃣ Vaciado / EEFF", type=["xlsx", "xlsm", "pdf"])
with col2:
    up_buro = st.file_uploader("2️⃣ Buró de Crédito (PDF)", type=["pdf"])
with col3:
    up_precalif = st.file_uploader("3️⃣ Modelo Precalificación", type=["xlsx", "xlsm", "pdf"])
with col4:
    up_tif = st.file_uploader("4️⃣ Gobierno Corp. (Opcional)", type=["tif", "tiff"])

comentario_opcional = st.text_area(
    "💬 Comentario del ejecutivo / aclaración de Buró (Opcional — solo para simular doble escenario)",
    placeholder="Pega aquí el párrafo si hay alguna aclaración de buró o quita...",
    height=80
)

if st.button("🔍 Auditar Precalificación", type="primary", use_container_width=True):
    if not (up_vaciado and up_buro and up_precalif):
        st.error("⚠️ Por favor sube los 3 archivos obligatorios: Vaciado/EEFF, Buró de Crédito y Modelo de Precalificación.")
    else:
        with st.spinner("Auditando expediente contra normativa CUB..."):
            with tempfile.TemporaryDirectory() as tmpdir:
                p_vac = os.path.join(tmpdir, up_vaciado.name)
                p_bur = os.path.join(tmpdir, up_buro.name)
                p_pre = os.path.join(tmpdir, up_precalif.name)
                with open(p_vac, "wb") as f: f.write(up_vaciado.getbuffer())
                with open(p_bur, "wb") as f: f.write(up_buro.getbuffer())
                with open(p_pre, "wb") as f: f.write(up_precalif.getbuffer())

                p_tif = None
                if up_tif:
                    p_tif = os.path.join(tmpdir, up_tif.name)
                    with open(p_tif, "wb") as f: f.write(up_tif.getbuffer())

                res = ejecutar_auditoria_completa(p_vac, p_bur, p_pre, p_tif, comentario_opcional)

        st.divider()
        st.subheader(f"🏢 Expediente: {res['nombre_cliente']}")

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Anexo Aplicable", res["hoja_modelo"])
        m2.metric("Calificación Actual", res["calif"])
        m3.metric("% Reserva", res["reserva"])
        m4.metric("Errores Detectados", len(res["errores"]))

        if res["escenario"]:
            esc = res["escenario"]
            st.info(
                f"📊 **Simulación Doble Escenario:** Actual = **{esc['calif_act']}** ({esc['res_act']}, {esc['puntaje_act']} pts) "
                f"➡️ Ajustado = **{esc['calif_nueva']}** ({esc['res_nueva']}, {esc['puntaje_nuevo']} pts | Delta: +{esc['delta']} pts)"
            )

        if not res["errores"]:
            st.success("✅ La precalificación se llenó correctamente. No se requieren modificaciones.")
        else:
            st.warning(f"⚠️ Se detectaron **{len(res['errores'])}** error(es) de llenado:")
            for idx, err in enumerate(res["errores"], 1):
                st.write(f"**{idx}.** {err['motivo']}")

        if res["excel_bytes"]:
            nombre_limpio = re.sub(r'[\\/*?:"<>|]', "", res["nombre_cliente"]).strip()
            st.download_button(
                label="⬇️ Descargar Excel Corregido",
                data=res["excel_bytes"],
                file_name=f"Modelo Precalificación - {nombre_limpio} (Corregido).xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

        st.subheader("✉️ Borrador de Correo Listo para Copiar")
        st.code(res["correo"], language=None)

        with st.expander("🔎 Ver trazabilidad técnica paso a paso"):
            for log in res["logs"]:
                st.markdown(f"- {log}")
            if res["alertas"]:
                st.markdown("**Alertas informativas sin impacto en puntaje:**")
                for al in res["alertas"]:
                    st.markdown(f"- ⚠️ {al}")
                    