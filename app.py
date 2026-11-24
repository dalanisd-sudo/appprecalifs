import os
import re
import io
import tempfile
import openpyxl
import pypdf
import pytesseract
from PIL import Image, ImageSequence
import streamlit as st

st.set_page_config(
    page_title="Autocontrol de Precalificación CUB",
    page_icon="🛡️",
    layout="wide"
)

# Control de llave para reiniciar los cargadores de archivos al limpiar sesión
if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0
if "resultado_auditoria" not in st.session_state:
    st.session_state["resultado_auditoria"] = None


def limpiar_memoria_y_reiniciar():
    st.session_state["resultado_auditoria"] = None
    st.session_state["uploader_key"] += 1


# -------------------------------------------------------------------------
# MOTOR DE LECTURA Y AUDITORÍA NORMATIVA (CUB ANEXOS 20, 21 Y 22)
# -------------------------------------------------------------------------
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


def ejecutar_auditoria_analista(path_vac, path_buro, path_pre, path_tif):
    logs_pasos = []
    errores_llenado = []
    alertas_informativas = []

    def registrar_error(bloque, celda_str, hoja, fila, col, val_actual, nuevo_val, motivo):
        for e in errores_llenado:
            if e["hoja"] == hoja and e["fila"] == fila and e["col"] == col and e["motivo"] == motivo:
                return
        errores_llenado.append({
            "bloque": bloque,
            "celda": celda_str,
            "hoja": hoja,
            "fila": fila,
            "col": col,
            "val_actual": str(val_actual),
            "nuevo_val": nuevo_val,
            "motivo": motivo
        })

    # 1. Lectura de archivos
    wb_vac = convertir_vaciado_pdf_a_workbook(path_vac) if path_vac.lower().endswith(".pdf") else openpyxl.load_workbook(path_vac, data_only=True)
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

    # 2. Detección de pestaña activa
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
    logs_pasos.append(f"**[Paso 1 — Pestaña]** Empresa: `{nombre_cliente}` | Pestaña evaluada: `{hoja_modelo}`")

    # 3. Validación SCIAN y Sector
    tabla_scian = re.search(
        r"C[oó]digo\s+SCIAN\s+Descripci[oó]n(.*?)(?:Previsor|Mensajes|INFORMACI[OÓ]N CREDITICIA|Cr[eé]ditos Activos)",
        texto_buro_completo, re.DOTALL | re.IGNORECASE
    )
    sectores_validos_buro, codigos_buro = [], []
    if tabla_scian:
        for cod in re.findall(r"\b(\d{5,6})\b", tabla_scian.group(1)):
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
        sugerido = sector_desde_codigo_excel or (sectores_validos_buro[0] if sectores_validos_buro else "Comercio")
        registrar_error("Datos Generales", "C10", hoja_modelo, 10, 3, "Vacío", sugerido, "El campo 'Sector' (C10) está vacío.")
    elif sector_desde_codigo_excel and sector_excel.lower() != sector_desde_codigo_excel.lower():
        registrar_error(
            "Datos Generales", "C10", hoja_modelo, 10, 3, sector_excel, sector_desde_codigo_excel,
            f"Inconsistencia interna: en C{fila_cod_excel} capturaste el código '{codigo_excel_num}' ({sector_desde_codigo_excel}), pero en C10 elegiste '{sector_excel}'."
        )
    elif sectores_validos_buro and sector_excel not in sectores_validos_buro:
        alertas_informativas.append(f"Diferencia de Sector vs Buró: en Excel elegiste '{sector_excel}', mientras que el Buró reporta claves {', '.join(codigos_buro)} ({sectores_validos_buro[0]}). Verifica que el giro real de la empresa corresponda al sector capturado.")
    else:
        logs_pasos.append(f"**[Paso 2 — Sector]** Sector `{sector_activo}` consistente con claves SCIAN.")

    # 4. Validación de Ventas Netas y UDIs
    hoja_edo_res = next((wb_vac[h] for h in wb_vac.sheetnames if h.strip().lower() in ["edo. res.", "edo res", "estado de resultados", "resultados"]), wb_vac[wb_vac.sheetnames[0]])
    fila_ventas = next((r for r in range(1, 20) if str(hoja_edo_res.cell(r, 1).value or "").strip().lower() in ["ventas netas", "ventas"]), None)
    ventas_vac, col_periodo_res, fecha_periodo_res, es_dictaminado = None, None, "", False

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
            es_dictaminado = "dictaminad" in str(hoja_edo_res.cell(5, col_periodo_res).value or "").lower()

    ventas_modelo = float(ws_val["G18"].value or 0)
    udi_modelo = float(ws_val["G19"].value or 0)

    if ventas_vac is not None:
        if abs(ventas_vac - ventas_modelo) > 1.5:
            registrar_error(
                "Datos Generales", "G18", hoja_modelo, 18, 7, f"{ventas_modelo:,.2f}", round(ventas_vac, 2),
                f"Ventas Netas en G18 dice '{ventas_modelo:,.2f}', pero en Edo. Res. ({fecha_periodo_res}) reporta '{ventas_vac:,.2f}'."
            )
        if udi_modelo > 0:
            ventas_udis = ventas_vac / udi_modelo
            if ventas_udis < 14000 and "Corporativos" in hoja_modelo:
                registrar_error(
                    "Pestaña Normativa", "Hoja", hoja_modelo, None, None, hoja_modelo, "Pymes(Anexo 21)",
                    f"Las ventas equivalen a {ventas_udis:,.2f} miles de UDIs (<14M UDIs). Debe llenarse la pestaña 'Pymes(Anexo 21)'."
                )
            elif ventas_udis >= 14000 and "Pymes" in hoja_modelo:
                registrar_error(
                    "Pestaña Normativa", "Hoja", hoja_modelo, None, None, hoja_modelo, "Corporativos(Anexo 22)",
                    f"Las ventas equivalen a {ventas_udis:,.2f} miles de UDIs (>=14M UDIs). Debe llenarse la pestaña 'Corporativos(Anexo 22)'."
                )
        logs_pasos.append(f"**[Paso 3 — Ventas y UDIs]** Edo. Res. ({fecha_periodo_res}): `${ventas_vac:,.2f}` vs Modelo G18: `${ventas_modelo:,.2f}`.")

    # 5. Factor 1-A (Buró), 1-B (Experiencia) y Cualitativos (.tif)
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
            registrar_error(
                "Factor 1-A (Buró)", f"{letra_sel}{fila_1a_bk}", hoja_modelo, fila_1a_bk, col_sel, sel_bk, mejor_bk,
                f"Créditos Bancarios tiene '{sel_bk}', pero BK12_PCT_SAT = 1 (100% puntual). Debe ser '{mejor_bk}' ({pts_bk} pts)."
            )
        if nbk_num > 0 and nbk_sat == 1.0 and mejor_nbk and sel_nbk != mejor_nbk:
            registrar_error(
                "Factor 1-A (Buró)", f"{letra_sel}{fila_1a_nbk}", hoja_modelo, fila_1a_nbk, col_sel, sel_nbk, mejor_nbk,
                f"Créditos No Bancarios tiene '{sel_nbk}', pero NBK12_PCT_SAT = 1 (100% puntual). Debe ser '{mejor_nbk}' ({pts_nbk} pts)."
            )

        filas_1b = [44, 50] if "Corporativos" in hoja_modelo else [48, 53, 54]
        for r_1b in filas_1b:
            num_ind = ws_val.cell(r_1b, col_rango - 2).value
            nom_ind = str(ws_val.cell(r_1b, col_rango - 1).value or "").strip()
            if num_ind in [4, 5, 4.0, 5.0] and nom_ind:
                sel_1b = str(ws_val.cell(r_1b, col_sel).value or "").strip()
                mejor_op_1b, mejor_pts_1b = obtener_mejor_opcion_tabla(r_1b, col_rango)
                pts_act_1b = ws_val.cell(r_1b, col_pts).value
                if mejor_op_1b and sel_1b != mejor_op_1b and pts_act_1b != mejor_pts_1b:
                    registrar_error(
                        "Factor 1-B (Experiencia)", f"{letra_sel}{r_1b}", hoja_modelo, r_1b, col_sel, sel_1b, mejor_op_1b,
                        f"Indicador '{nom_ind}' tiene '{sel_1b}' ({pts_act_1b} pts). Conforme a política debe llevar '{mejor_op_1b}' ({mejor_pts_1b} pts)."
                    )
        logs_pasos.append(f"**[Paso 4 — Buró y Exp. Interna]** BK12_PCT_SAT={bk_sat} | NBK12_PCT_SAT={nbk_sat}.")

    # Revisión de Gobierno Corporativo (.tif) en Corporativos
    if "Corporativos" in hoja_modelo:
        tiene_consejo = bool(re.search(r"consejo de administraci[oó]n|administrador [uú]nico|comisario|asamblea", texto_tif_completo, re.IGNORECASE))
        accionista_33 = bool(re.search(r"([3-9][4-9]|[4-9]\d|100)\s*%", texto_tif_completo))
        sel_aud = str(ws_val["F108"].value or "").strip()
        sel_gob = str(ws_val["F112"].value or "").strip()
        sel_acc = str(ws_val["F117"].value or "").strip()

        if not es_dictaminado and "dos últimos años" in sel_aud.lower():
            registrar_error("Cualitativos", "F108", hoja_modelo, None, None, sel_aud, "Revisar", "F108 dice que los EEFF son auditados los dos últimos años, pero el Vaciado no es dictaminado.")
        if path_tif and tiene_consejo and ("no cuenta" in sel_gob.lower() or "débil" in sel_gob.lower()):
            registrar_error("Cualitativos", "F112", hoja_modelo, None, None, sel_gob, "Revisar", "F112 indica estructura débil/sin gobierno, pero el .tif acredita órganos de administración.")
        if path_tif and accionista_33 and "más del 33%" not in sel_acc.lower():
            registrar_error("Cualitativos", "F117", hoja_modelo, None, None, sel_acc, "Revisar", "F117 no refleja concentración accionaria >33%, pero el .tif muestra participación superior al 33%.")

    # 6. Factor 1-C (Razones Financieras en Corporativos)
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
                        es_pct = "%" in rango_esp
                        val_fmt = f"{val_calc*100:.2f}%" if es_pct else f"{val_calc:.2f}"
                        registrar_error(
                            "Factor 1-C (Financiero)", coord_s, hoja_modelo, fila_sel, col_sel, sel_act, rango_esp,
                            f"Indicador '{nom_ind}' (cálculo real = {val_fmt}): tiene '{sel_act}', debe ser '{rango_esp}'."
                        )

            for d in discrepancias:
                if d["r_padre"] not in ind_err:
                    alertas_informativas.append(f"Celda {d['coord']} ({d['etiqueta']}): capturaste {d['val_excel']:,.2f} vs Vaciado {d['val_v']:,.2f}. No altera puntaje ni calificación porque el ratio permanece en el mismo rango.")
                else:
                    registrar_error(
                        "Factor 1-C (Insumo)", d["coord"], hoja_modelo, d["fila"], d["col"], f"{d['val_excel']:,.2f}", round(d["val_v"], 2),
                        f"Insumo '{d['etiqueta']}': capturaste {d['val_excel']:,.2f}, pero en el Vaciado reporta {d['val_v']:,.2f}."
                    )
            logs_pasos.append(f"**[Paso 5 — Factor 1-C]** {len(filas_ind_1c)} indicadores financieros auditados.")

    # Detector de #REF! y vínculos externos rotos
    for r in range(1, min(ws_val.max_row + 1, 130)):
        for c in range(1, min(ws_val.max_column + 1, 70)):
            coord_err = f"{openpyxl.utils.get_column_letter(c)}{r}"
            val_c = str(ws_val.cell(r, c).value or "").strip()
            form_c = str(ws_form.cell(r, c).value or "").strip()
            if any(err in val_c.upper() for err in ["#REF!", "#VALUE!", "#VALOR!", "#NAME?", "#NOMBRE?", "#DIV/0!", "#N/A"]) or "#REF!" in form_c.upper():
                alertas_informativas.append(f"Error de fórmula en celda {coord_err} ({val_c or '#REF!'}). Pega los valores numéricos directos sin vínculos externos.")
            elif form_c.startswith("=") and re.search(r"\[\d+\]|\.xls", form_c, re.IGNORECASE):
                alertas_informativas.append(f"Vínculo externo detectado en {coord_err} ({form_c}). Conviene pegarlo como valor fijo.")

    # Generar Excel corregido
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
        "sector": sector_activo,
        "calif": calif_final,
        "reserva": res_final,
        "errores": errores_llenado,
        "alertas": list(dict.fromkeys(alertas_informativas)),
        "logs": logs_pasos,
        "excel_bytes": excel_corregido_bytes,
    }


# -------------------------------------------------------------------------
# INTERFAZ VISUAL PARA EL ANALISTA
# -------------------------------------------------------------------------
st.title("🛡️ Autocontrol de Precalificación (CUB)")
st.caption("🔒 **Privacidad garantizada:** Los archivos se procesan exclusivamente en memoria temporal y se eliminan automáticamente al finalizar. Ningún expediente se guarda en base de datos.")

k = st.session_state["uploader_key"]
col1, col2, col3, col4 = st.columns(4)
with col1:
    up_vaciado = st.file_uploader("1️⃣ Vaciado / EEFF", type=["xlsx", "xlsm", "pdf"], key=f"vac_{k}", help="Excel de vaciado o estados financieros en PDF")
with col2:
    up_buro = st.file_uploader("2️⃣ Buró de Crédito (PDF)", type=["pdf"], key=f"bur_{k}", help="Reporte de Buró de Crédito Persona Moral")
with col3:
    up_precalif = st.file_uploader("3️⃣ Modelo Precalificación", type=["xlsx", "xlsm", "pdf"], key=f"pre_{k}", help="Plantilla de Precalificación IFRS9 / CUB")
with col4:
    up_tif = st.file_uploader("4️⃣ Gobierno Corp. (Opcional)", type=["tif", "tiff"], key=f"tif_{k}", help="Documento escaneado .tif de estructura accionaria/gobierno")

st.markdown("")
if st.button("🔍 Validar mi Precalificación", type="primary", use_container_width=True):
    if not (up_vaciado and up_buro and up_precalif):
        st.error("⚠️ Sube los 3 archivos obligatorios (Vaciado/EEFF, Buró de Crédito y Modelo de Precalificación) para iniciar la validación.")
    else:
        with st.spinner("Verificando datos contables, claves SCIAN, Buró y rangos CUB..."):
            # Carpeta temporal que se destruye automáticamente al terminar este bloque 'with'
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

                st.session_state["resultado_auditoria"] = ejecutar_auditoria_analista(p_vac, p_bur, p_pre, p_tif)

# Renderizar resultados si existen en la sesión activa
res = st.session_state["resultado_auditoria"]
if res:
    st.divider()
    st.subheader(f"🏢 {res['nombre_cliente']}")

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Anexo Evaluado", res["hoja_modelo"])
    m2.metric("Sector Activo", res["sector"] or "N/D")
    m3.metric("Calificación (Archivo)", f"{res['calif']} ({res['reserva']})")
    m4.metric("Correcciones Requeridas", len(res["errores"]))

    if not res["errores"]:
        st.success("✅ **¡Precalificación aprobada!** Todos los campos y rangos coinciden con la documentación soporte. Ya puedes enviarla.")
    else:
        st.error(f"❌ **Se detectaron {len(res['errores'])} punto(s) por corregir antes de enviar:**")
        tabla_datos = [
            {
                "Bloque": e["bloque"],
                "Celda": e["celda"],
                "Valor Capturado": e["val_actual"],
                "Valor Correcto (CUB)": str(e["nuevo_val"]),
                "Detalle": e["motivo"]
            }
            for e in res["errores"]
        ]
        st.dataframe(tabla_datos, use_container_width=True, hide_index=True)

        if res["excel_bytes"]:
            nombre_limpio = re.sub(r'[\\/*?:"<>|]', "", res["nombre_cliente"]).strip()
            st.download_button(
                label="⬇️ Descargar Excel con Correcciones Aplicadas",
                data=res["excel_bytes"],
                file_name=f"Modelo Precalificación - {nombre_limpio} (Corregido).xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
            st.caption("💡 *Nota: Abre el archivo descargado en Excel y guárdalo para que se actualicen las fórmulas de puntaje y reserva.*")

    if res["alertas"]:
        with st.expander(f"⚠️ Ver {len(res['alertas'])} observación(es) preventiva(s) (no bloquean envío)"):
            for al in res["alertas"]:
                st.warning(al)

    with st.expander("🔎 Ver resumen de validación paso a paso"):
        for log in res["logs"]:
            st.markdown(f"- {log}")

    st.divider()
    st.button(
        "🗑️ Borrar expediente de memoria y limpiar pantalla",
        on_click=limpiar_memoria_y_reiniciar,
        use_container_width=True,
        help="Elimina inmediatamente todos los archivos cargados y resultados de esta sesión."
    )
    