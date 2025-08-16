from __future__ import annotations
from pathlib import Path
import io
import math
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
import xlsxwriter

# -------------- Utilidades --------------

def compute_bsa_mosteller(peso_kg: float, talla_cm: float) -> float:
    """BSA (m²) = sqrt( (peso * talla) / 3600 )"""
    return ((peso_kg * talla_cm) / 3600.0) ** 0.5

def _read_excel_bytes(file_or_path) -> bytes:
    if hasattr(file_or_path, "read"):
        return file_or_path.read()
    p = Path(file_or_path)
    return p.read_bytes()

# -------------- Catálogo desde "Listas" --------------

def extract_catalog_from_excel(xlsx_file) -> pd.DataFrame:
    """
    Lee la hoja 'Listas' y devuelve un DataFrame de catálogo:
    columnas: bucket, Medicamento, Dosis, Solución, VS, Tiempo, Via
    """
    raw = pd.read_excel(io.BytesIO(_read_excel_bytes(xlsx_file)), sheet_name="Listas", header=None)
    cat_frames = []

    # Estrategia robusta: busca filas que contengan "Medicamento" en alguna columna
    # y asume que las 5-6 columnas siguientes contienen la tabla.
    for r in range(min(200, raw.shape[0])):
        for c in range(min(40, raw.shape[1]-1)):
            if str(raw.iat[r, c]).strip().lower() == "medicamento":
                # intenta leer 6 columnas: Medicamento, Dosis, Solución, VS, Tiempo, Via
                header = raw.iloc[r, c:c+6].tolist()
                if "Medicamento" not in header[0]:
                    continue
                # consume hacia abajo hasta que toda la fila sean NaN
                rows = []
                rr = r + 1
                while rr < raw.shape[0]:
                    row = raw.iloc[rr, c:c+6]
                    if row.isna().all():
                        break
                    rows.append(row.tolist())
                    rr += 1
                df = pd.DataFrame(rows, columns=["Medicamento","Dosis","Solución","VS","Tiempo","Via"])
                # intenta deducir bucket mirando títulos cercanos (en columnas previas)
                bucket = _infer_bucket(raw, r, c)
                df["bucket"] = bucket
                # limpia filas vacías o separadores "-"
                df = df[~df["Medicamento"].astype(str).str.strip().isin(["-", "nan", "None"])]
                cat_frames.append(df)

    if not cat_frames:
        return pd.DataFrame(columns=["bucket","Medicamento","Dosis","Solución","VS","Tiempo","Via"])

    cat = pd.concat(cat_frames, ignore_index=True)
    # normaliza bucket
    cat["bucket"] = cat["bucket"].fillna("Desconocido")
    # quita duplicados conservando la primera entrada (suele haber listas repetidas)
    cat = cat.drop_duplicates(subset=["bucket","Medicamento"], keep="first")
    return cat[["bucket","Medicamento","Dosis","Solución","VS","Tiempo","Via"]]

def _infer_bucket(raw: pd.DataFrame, r: int, c: int) -> str:
    # explora hacia arriba en mismas columnas buscando títulos típicos
    for up in range(max(0, r-8), r)[::-1]:
        row_vals = raw.iloc[up, max(0,c-2):c+1].astype(str).str.lower().tolist()
        txt = " ".join(row_vals)
        if "premedic" in txt:
            return "Premedicación"
        if "anticuerp" in txt:
            return "Anticuerpos"
        if "quimiot" in txt:
            return "Quimioterapia"
        if "otros" in txt:
            return "Otros"
    return "Catálogo"

def human_bucket_choices(df_cat: pd.DataFrame):
    def choices(bucket):
        return sorted(df_cat[df_cat["bucket"]==bucket]["Medicamento"].astype(str).tolist())
    prem = choices("Premedicación")
    acs  = choices("Anticuerpos")
    qx   = choices("Quimioterapia")
    otros= choices("Otros")
    return prem, acs, qx, otros

# -------------- Generación de XLSX imprimible --------------

HEADER_MAP = [
    ("Nombre:",               "nombre"),
    ("Sexo:",                 "sexo"),
    ("Diagnóstico:",          "diagnostico"),
    ("Objetivo:",             "objetivo"),
    ("Ciclo:",                "ciclo"),
    ("Peso (Kg):",            "peso"),
    ("Talla (cm):",           "talla"),
    ("Cr:",                   "cr"),
    ("SC (m²):",              "bsa"),
    ("Alergias:",             "alergias"),
    ("Fecha aplicación:",     "fecha_aplicacion"),
]

TABLE_ORDER = [
    ("P R E M E D I C A C I Ó N", "Premedicación"),
    ("A N T I C U E R P O S  M O N O C L O N A L E S", "Anticuerpos"),
    ("Q U I M I O T E R A P I A", "Quimioterapia"),
    ("O T R O S", "Otros"),
]

TABLE_COLS = ["Medicamento","Dosis (mg o mg/m²)","Solución","Volumen","Tiempo","Vía"]

def generate_indication_xlsx(
    plantilla_path: Path,
    output_path: Path,
    patient: dict,
    df_catalog: pd.DataFrame,
    selections: dict,
):
    """
    Crea un XLSX con formato limpio, inspirado en tu plantilla,
    listo para impresión/convertir a PDF.
    """
    wb = xlsxwriter.Workbook(str(output_path))
    ws = wb.add_worksheet("Indicaciones")

    fmt_title = wb.add_format({"bold": True, "font_size": 14})
    fmt_h = wb.add_format({"bold": True, "bg_color":"#EFEFEF", "border":1})
    fmt_k = wb.add_format({"bold": True})
    fmt_v = wb.add_format({})
    fmt_tbl = wb.add_format({"border":1})
    fmt_tbl_h = wb.add_format({"border":1, "bold": True, "bg_color":"#F5F5F5"})

    ws.write("A1", "INDICACIONES MÉDICAS", fmt_title)

    row = 3
    for label, key in HEADER_MAP:
        ws.write(row, 0, label, fmt_k)
        ws.write(row, 1, "" if patient.get(key) is None else patient.get(key), fmt_v)
        row += 1

    row += 1

    def write_table(title, bucket, start_row):
        ws.write(start_row, 0, title, fmt_k)
        start_row += 1
        # encabezados
        for j, col in enumerate(TABLE_COLS):
            ws.write(start_row, j, col, fmt_tbl_h)
        start_row += 1

        meds = selections.get(bucket.lower(), []) if isinstance(selections, dict) else []
        # si selections es dict con claves 'premedicacion','anticuerpos','quimioterapia','otros':
        if bucket == "Premedicación":
            meds = selections.get("premedicacion", [])
        elif bucket == "Anticuerpos":
            meds = selections.get("anticuerpos", [])
        elif bucket == "Quimioterapia":
            meds = selections.get("quimioterapia", [])
        elif bucket == "Otros":
            meds = selections.get("otros", [])

        cat = df_catalog[df_catalog["bucket"]==bucket]
        for med in meds:
            row_df = cat[cat["Medicamento"].astype(str)==str(med)].head(1)
            if row_df.empty:
                continue
            r = row_df.iloc[0]
            ws.write(start_row, 0, med, fmt_tbl)
            ws.write(start_row, 1, r.get("Dosis",""), fmt_tbl)
            ws.write(start_row, 2, r.get("Solución",""), fmt_tbl)
            ws.write(start_row, 3, r.get("VS",""), fmt_tbl)
            ws.write(start_row, 4, r.get("Tiempo",""), fmt_tbl)
            ws.write(start_row, 5, r.get("Via",""), fmt_tbl)
            start_row += 1
        return start_row + 1

    # Ajustes de ancho
    ws.set_column(0, 0, 38)
    ws.set_column(1, 1, 22)
    ws.set_column(2, 5, 16)

    for title, bucket in TABLE_ORDER:
        row = write_table(title, bucket, row)

    wb.close()

