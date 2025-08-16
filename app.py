import gradio as gr
from pathlib import Path
import tempfile, subprocess, io, json
import pandas as pd
from datetime import date

from generate_prescription import (
    extract_catalog_from_excel,
    generate_indication_xlsx,
    human_bucket_choices,
    compute_bsa_mosteller,
)

TITLE = "Generador de Indicaciones Médicas (Quimioterapia)"
DESC = """
Sube tu plantilla Excel (con hojas **Indicaciones Médicas** y **Listas**).
Captura los datos del paciente, selecciona medicamentos desde los catálogos de **Listas** y genera un PDF imprimible.
"""

def load_catalog(file_obj):
    if file_obj is None:
        return gr.update(choices=[]), gr.update(choices=[]), gr.update(choices=[]), gr.update(choices=[])
    df_cat = extract_catalog_from_excel(file_obj)
    # choices agrupados
    prem_choices, ac_choices, qx_choices, otros_choices = human_bucket_choices(df_cat)
    return (
        gr.update(choices=prem_choices, value=[]),
        gr.update(choices=ac_choices,   value=[]),
        gr.update(choices=qx_choices,   value=[]),
        gr.update(choices=otros_choices,value=[]),
    )

def on_generate(
    plantilla, nombre, sexo, dx, objetivo, ciclo,
    peso, talla, cr, alergias, fecha_aplicacion,
    prem, acs, qx, otros
):
    if plantilla is None:
        return None, "Sube una plantilla Excel.", "{}"

    with tempfile.TemporaryDirectory() as tmpd:
        xlsx_in = Path(tmpd) / "plantilla.xlsx"
        xlsx_in.write_bytes(plantilla.read())

        # catálogo
        df_cat = extract_catalog_from_excel(xlsx_in.open("rb"))
        # superficie corporal
        try:
            bsa = compute_bsa_mosteller(float(peso), float(talla))
        except Exception:
            bsa = None

        patient = {
            "nombre": nombre.strip(),
            "sexo": sexo,
            "diagnostico": dx.strip(),
            "objetivo": objetivo,
            "ciclo": int(ciclo) if ciclo is not None else None,
            "peso": float(peso) if peso not in (None, "") else None,
            "talla": float(talla) if talla not in (None, "") else None,
            "cr": float(cr) if cr not in (None, "") else None,
            "alergias": alergias.strip() if alergias else "",
            "fecha_aplicacion": fecha_aplicacion.strip() if fecha_aplicacion else "",
            "bsa": bsa,
        }

        selections = {
            "premedicacion": prem or [],
            "anticuerpos": acs or [],
            "quimioterapia": qx or [],
            "otros": otros or [],
        }

        # crea XLSX con layout imprimible
        xlsx_out = Path(tmpd) / f"Indicacion_{date.today()}.xlsx"
        generate_indication_xlsx(
            plantilla_path=xlsx_in,
            output_path=xlsx_out,
            patient=patient,
            df_catalog=df_cat,
            selections=selections,
        )

        # convierte a PDF con LibreOffice
        pdf_out = Path(tmpd) / f"Indicacion_{date.today()}.pdf"
        subprocess.run(
            ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", str(Path(tmpd)), str(xlsx_out)],
            check=True
        )

        debug_json = json.dumps(
            {"patient": patient, "selections": selections}, ensure_ascii=False, indent=2
        )

        return pdf_out.read_bytes(), "PDF generado ✅", debug_json

with gr.Blocks(title=TITLE) as demo:
    gr.Markdown(f"# {TITLE}\n\n{DESC}")

    plantilla = gr.File(label="Plantilla Excel (.xlsx)")

    with gr.Row():
        nombre = gr.Textbox(label="Nombre completo")
        sexo = gr.Dropdown(choices=["FEM","MAS"], label="Sexo")
        objetivo = gr.Dropdown(
            choices=["Adyuvante","Neoadyuvante","Inducción","Mantenimiento","Paliativo","Concomitante","No aplica"],
            label="Objetivo"
        )
        ciclo = gr.Number(label="Ciclo", value=1, precision=0)

    dx = gr.Textbox(label="Diagnóstico")
    alergias = gr.Textbox(label="Alergias")

    with gr.Row():
        peso = gr.Number(label="Peso (kg)")
        talla = gr.Number(label="Talla (cm)")
        cr = gr.Number(label="Creatinina sérica (mg/dL)")
        fecha_aplicacion = gr.Textbox(label="Fecha de aplicación (dd.mm.aaaa)")

    gr.Markdown("### Selección de medicamentos (desde **Listas**)")
    with gr.Row():
        prem = gr.CheckboxGroup(choices=[], label="Premedicación")
        acs  = gr.CheckboxGroup(choices=[], label="Anticuerpos monoclonales")
    with gr.Row():
        qx   = gr.CheckboxGroup(choices=[], label="Quimioterapia")
        otros= gr.CheckboxGroup(choices=[], label="Otros")

    # Al subir plantilla, carga catálogos
    plantilla.change(load_catalog, [plantilla], [prem, acs, qx, otros])

    btn = gr.Button("Generar PDF")
    archivo = gr.File(label="Descarga el PDF", file_types=[".pdf"])
    status  = gr.Markdown()
    debug   = gr.Code(label="(Opcional) Datos enviados", language="json")

    btn.click(
        on_generate,
        [plantilla, nombre, sexo, dx, objetivo, ciclo, peso, talla, cr, alergias, fecha_aplicacion, prem, acs, qx, otros],
        [archivo, status, debug]
    )

if __name__ == "__main__":
    demo.launch()
