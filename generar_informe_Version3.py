#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generador de informes veterinarios a partir del Excel Datos_informe_veterinario.xlsx
Lee los datos, inserta textos y fotografías y guarda el Word en Informes_generados/.
"""

import os
import sys
import traceback
import re
import datetime
from io import BytesIO

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
from PIL import Image, ImageOps, UnidentifiedImageError

# Configuración
PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILENAME = os.path.join(PROJECT_DIR, "Datos_informe_veterinario.xlsx")
TEMPLATE_FILENAME = os.path.join(PROJECT_DIR, "Plantilla_informe_veterinario.docx")
FOTOGRAFIAS_DIR = os.path.join(PROJECT_DIR, "Fotografias")
OUTPUT_DIR = os.path.join(PROJECT_DIR, "Informes_generados")
REGISTRO_FILE = os.path.join(PROJECT_DIR, "registro_generacion.txt")

ALLOWED_EXT = {'.jpg', '.jpeg', '.png', '.tif', '.tiff', '.bmp', '.webp'}

# Campos esperados (hoja "Datos": columna A nombre del campo, columna B valor)
EXPECTED_FIELDS = {
    "Número de caso": None,
    "Número de historia": None,
    "Especie": None,
    "Edad": None,
    "Sexo": None,
    "Identificación del animal": None,
    "Paraje": None,
    "Término municipal": None,
    "Provincia": None,
    "UTV": None,
    "Coordenadas": None,
    "Número de precinto o acta": None,
    "Fecha de ingreso en el CRFS AMUS": None,
    "Tipo de ingreso": None,
    "Fecha de necropsia": None,
    "Conservación del cadáver": None,
    "Examen radiográfico realizado": None,
    "Análisis toxicológico realizado": None,
    "Examen externo": None,
    "Examen interno": None,
    "Resultado del examen radiográfico": None,
    "Laboratorio de toxicología": None,
    "Fecha del análisis toxicológico": None,
    "Muestras enviadas": None,
    "Resultados toxicológicos": None,
    "Diagnóstico y conclusiones": None,
    "Fecha de emisión": None,
    "Nombre del veterinario firmante": None,
    "Número de colegiado": None,
    "Provincia del colegio profesional": None,
    "Centro de trabajo": None
}

# Secciones que deben empezar en página nueva
SECTIONS_NEW_PAGE = [
    "EXAMEN EXTERNO",
    "EXAMEN INTERNO",
    "EXAMEN RADIOGRÁFICO",
    "ANÁLISIS TOXICOLÓGICO",
    "DIAGNÓSTICO Y CONCLUSIONES"
]

def ensure_dirs():
    os.makedirs(FOTOGRAFIAS_DIR, exist_ok=True)
    for sub in ("Examen_externo", "Examen_interno", "Examen_radiografico"):
        os.makedirs(os.path.join(FOTOGRAFIAS_DIR, sub), exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)

def natural_sort_key(s):
    # produce key for natural sort
    _nsre = re.compile('([0-9]+)')
    return [int(text) if text.isdigit() else text.lower() for text in _nsre.split(s)]

def list_images(folder):
    files = []
    if not os.path.isdir(folder):
        return files
    for fn in os.listdir(folder):
        fp = os.path.join(folder, fn)
        if os.path.isfile(fp):
            ext = os.path.splitext(fn)[1].lower()
            if ext in ALLOWED_EXT:
                files.append(fn)
    files.sort(key=natural_sort_key)
    return files

def sanitize_filename_component(s):
    if s is None:
        return ""
    # Remove invalid Windows filename chars: <>:"/\|?* and control chars. Trim trailing spaces and dots.
    invalid = r'<>:"/\\|?*\0'
    res = "".join("_" if c in invalid else c for c in s)
    res = res.strip().rstrip(".")
    # Replace newlines and tabs
    res = res.replace("\n", " ").replace("\r", " ").replace("\t", " ")
    # Avoid empty string
    return res if res else "_"

def build_output_filename(num_historia, especie):
    if not num_historia or not especie:
        raise ValueError("Número de historia y Especie son obligatorios para construir el nombre del archivo.")
    a = sanitize_filename_component(str(num_historia))
    b = sanitize_filename_component(str(especie))
    base = f"Informe veterinario {a} {b}.docx"
    path = os.path.join(OUTPUT_DIR, base)
    if not os.path.exists(path):
        return path
    # create alternative
    i = 2
    name_wo_ext = base[:-5]  # remove .docx
    while True:
        alt = f"{name_wo_ext} ({i}).docx"
        alt_path = os.path.join(OUTPUT_DIR, alt)
        if not os.path.exists(alt_path):
            return alt_path
        i += 1

def read_excel_fields(excel_path):
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"No se encuentra el Excel: {excel_path}")
    try:
        wb = load_workbook(excel_path, data_only=True)
    except PermissionError as e:
        raise PermissionError("No se puede abrir el Excel. ¿Está abierto en Excel? Por favor, guarde y cierre el archivo antes de ejecutar.") from e
    if "Datos" not in wb.sheetnames:
        raise ValueError("El libro debe contener una hoja llamada 'Datos'.")
    ws = wb["Datos"]
    # read A->B pairs until empty A
    data = {}
    for row in ws.iter_rows(min_row=1, max_col=2):
        key_cell = row[0].value
        val_cell = row[1].value if len(row) > 1 else None
        if key_cell is None:
            continue
        key = str(key_cell).strip()
        # store only if key in expected fields or allow additional fields
        data[key] = val_cell
    # Map expected keys, fill None if missing
    mapped = {}
    for k in EXPECTED_FIELDS.keys():
        mapped[k] = data.get(k, None)
    # Also return leftover (additional) fields
    extras = {k: v for k, v in data.items() if k not in EXPECTED_FIELDS}
    return mapped, extras

def image_to_stream(image_path, max_width_px=1600, max_height_px=1200):
    """
    Load image with Pillow, apply EXIF orientation and scale so it's not huge.
    Return BytesIO (PNG) suitable for python-docx add_picture.
    """
    img = Image.open(image_path)
    img = ImageOps.exif_transpose(img)
    # Limit size to avoid too large insertion
    img.thumbnail((max_width_px, max_height_px), Image.LANCZOS)
    bio = BytesIO()
    # Save as PNG to avoid lossy conversion issues
    if img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in img.info):
        img.save(bio, format="PNG")
    else:
        img = img.convert("RGB")
        img.save(bio, format="JPEG", quality=90)
    bio.seek(0)
    return bio

def safe_insert_image(cell, image_stream, width_inches=3.2):
    # Insert image into a table cell paragraph
    paragraph = cell.paragraphs[0]
    run = paragraph.add_run()
    try:
        run.add_picture(image_stream, width=Inches(width_inches))
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        return True
    except Exception as e:
        return False

def add_heading(doc, text, level=1):
    p = doc.add_heading(text, level=level)
    return p

def add_plain_paragraph(doc, text, bold=False):
    p = doc.add_paragraph()
    if text is None:
        text = ""
    # preserve newlines: split and add runs with line breaks
    parts = str(text).splitlines()
    for i, line in enumerate(parts):
        run = p.add_run(line)
        if bold:
            run.bold = True
        if i < len(parts) - 1:
            p.add_run("\n")
    return p

def ensure_heading_styles(doc):
    """
    Asegura que existan los estilos 'Heading 1' y 'Heading 2' en el documento.
    Si no existen, los crea con formato básico para evitar errores al usar doc.add_heading().
    """
    styles = doc.styles
    # Heading 1
    try:
        _ = styles['Heading 1']
    except KeyError:
        try:
            h1 = styles.add_style('Heading 1', WD_STYLE_TYPE.PARAGRAPH)
            h1.font.name = 'Times New Roman'
            h1.font.size = Pt(16)
            h1.font.bold = True
        except Exception:
            # si no se puede crear el estilo, ignorar y dejar que python-docx use estilos por defecto
            pass
    # Heading 2
    try:
        _ = styles['Heading 2']
    except KeyError:
        try:
            h2 = styles.add_style('Heading 2', WD_STYLE_TYPE.PARAGRAPH)
            h2.font.name = 'Times New Roman'
            h2.font.size = Pt(12)
            h2.font.bold = True
        except Exception:
            pass

def safe_set_table_style(table, style_name='Table Grid'):
    """
    Intentar asignar un estilo a la tabla; si no existe en la plantilla, ignorar el fallo.
    """
    try:
        table.style = style_name
    except Exception:
        # No interrumpir si el estilo no existe en la plantilla
        pass

def generate_report():
    ensure_dirs()
    registro = {
        "timestamp": datetime.datetime.now().isoformat(sep=" ", timespec="seconds"),
        "numero_historia": None,
        "especie": None,
        "archivo_generado": None,
        "imagenes_encontradas": {},
        "imagenes_insertadas": 0,
        "omitidos": [],
        "errores": []
    }
    try:
        fields, extras = read_excel_fields(EXCEL_FILENAME)
    except Exception as e:
        msg = f"Error leyendo Excel: {e}"
        registro["errores"].append(msg)
        write_log(registro)
        print(msg)
        return False

    numero_historia = fields.get("Número de historia")
    especie = fields.get("Especie")
    registro["numero_historia"] = numero_historia
    registro["especie"] = especie

    if not numero_historia or not especie:
        msg = "Número de historia y/o Especie vacíos. Complete esos campos en el Excel antes de generar el informe."
        registro["errores"].append(msg)
        write_log(registro)
        print(msg)
        return False

    # Prepare output filename
    try:
        out_path = build_output_filename(numero_historia, especie)
    except Exception as e:
        registro["errores"].append(str(e))
        write_log(registro)
        print("No se puede construir el nombre del archivo:", e)
        return False

    registro["archivo_generado"] = os.path.basename(out_path)

    # Load template or start a new document
    if os.path.exists(TEMPLATE_FILENAME):
        doc = Document(TEMPLATE_FILENAME)
    else:
        # create minimal doc
        doc = Document()
        # basic styling
        try:
            style = doc.styles['Normal']
            style.font.name = 'Calibri'
            style.font.size = Pt(11)
        except Exception:
            pass

    # Ensure heading styles exist to avoid KeyError on doc.add_heading
    try:
        ensure_heading_styles(doc)
    except Exception:
        # no stop, continue
        pass

    # Title
    h = doc.add_heading("CASO / INFORME VETERINARIO", level=1)

    # DATOS PREVIOS
    doc.add_heading("DATOS PREVIOS", level=2)
    # add table for data previa
    t = doc.add_table(rows=0, cols=2)
    safe_set_table_style(t, 'Table Grid')

    def add_kv(k, v):
        row = t.add_row()
        row.cells[0].text = str(k)
        row.cells[1].text = "" if v is None else str(v)

    # Selected fields for data previa
    for k in ["Número de caso", "Número de historia", "Especie", "Edad", "Sexo",
              "Identificación del animal", "Paraje", "Término municipal", "Provincia",
              "UTV", "Coordenadas", "Número de precinto o acta",
              "Fecha de ingreso en el CRFS AMUS", "Tipo de ingreso", "Fecha de necropsia",
              "Conservación del cadáver", "Examen radiográfico realizado", "Análisis toxicológico realizado"]:
        add_kv(k, fields.get(k))

    # NECROPSIA section
    doc.add_heading("NECROPSIA", level=2)
    add_plain_paragraph(doc, "Resumen de necropsia:")

    # Sections and their content
    section_map = [
        ("EXAMEN EXTERNO", "Examen externo"),
        ("EXAMEN INTERNO", "Examen interno"),
        ("EXAMEN RADIOGRÁFICO", "Resultado del examen radiográfico"),
        ("ANÁLISIS TOXICOLÓGICO", "Resultados toxicológicos"),
        ("DIAGNÓSTICO Y CONCLUSIONES", "Diagnóstico y conclusiones")
    ]

    # Add the sections, respecting page breaks
    for sec_title, field_name in section_map:
        if sec_title in SECTIONS_NEW_PAGE:
            doc.add_page_break()
        doc.add_heading(sec_title, level=2)
        add_plain_paragraph(doc, fields.get(field_name))

    # FIRMA
    doc.add_heading("FIRMA DEL VETERINARIO", level=2)
    firma_tbl = doc.add_table(rows=4, cols=2)
    safe_set_table_style(firma_tbl, 'Table Grid')
    firma_tbl.cell(0,0).text = "Fecha de emisión"
    firma_tbl.cell(0,1).text = "" if fields.get("Fecha de emisión") is None else str(fields.get("Fecha de emisión"))
    firma_tbl.cell(1,0).text = "Nombre del veterinario firmante"
    firma_tbl.cell(1,1).text = "" if fields.get("Nombre del veterinario firmante") is None else str(fields.get("Nombre del veterinario firmante"))
    firma_tbl.cell(2,0).text = "Número de colegiado"
    firma_tbl.cell(2,1).text = "" if fields.get("Número de colegiado") is None else str(fields.get("Número de colegiado"))
    firma_tbl.cell(3,0).text = "Centro de trabajo / Provincia del colegio"
    firma_tbl.cell(3,1).text = "{} / {}".format(fields.get("Centro de trabajo") or "", fields.get("Provincia del colegio profesional") or "")

    # ANEXO FOTOGRÁFICO (groups)
    annex_groups = [
        ("ANEXO FOTOGRÁFICO: EXAMEN EXTERNO", os.path.join(FOTOGRAFIAS_DIR, "Examen_externo")),
        ("ANEXO FOTOGRÁFICO: EXAMEN INTERNO", os.path.join(FOTOGRAFIAS_DIR, "Examen_interno")),
        ("ANEXO FOTOGRÁFICO: ANEXO FOTOGRÁFICO: EXAMEN RADIOGRÁFICO", os.path.join(FOTOGRAFIAS_DIR, "Examen_radiografico"))
    ]

    total_inserted = 0
    for group_title, folder in annex_groups:
        imgs = list_images(folder)
        registro["imagenes_encontradas"][os.path.basename(folder)] = len(imgs)
        if len(imgs) == 0:
            # register empty but continue
            doc.add_paragraph()  # nothing to do
            continue
        # start group on new page
        doc.add_page_break()
        doc.add_heading(group_title, level=2)
        # For each 4 images make a page with 2x2 table
        img_counter = 0
        page_img_index = 0
        for i, imgname in enumerate(imgs):
            img_path = os.path.join(folder, imgname)
            # if starting a new page table
            if page_img_index == 0:
                table = doc.add_table(rows=2, cols=2)
                safe_set_table_style(table, 'Table Grid')
            # Determine cell row/col
            r = page_img_index // 2
            c = page_img_index % 2
            cell = table.cell(r, c)
            # Try to open and check image
            try:
                stream = image_to_stream(img_path)
            except UnidentifiedImageError:
                motivo = f"Imagen dañada o formato no soportado: {imgname} en {folder}"
                registro["omitidos"].append((imgname, motivo))
                print("Omitida:", motivo)
                page_img_index += 1
                if page_img_index >= 4:
                    page_img_index = 0
                continue
            except Exception as e:
                motivo = f"Error procesando imagen {imgname}: {e}"
                registro["omitidos"].append((imgname, motivo))
                print("Omitida:", motivo)
                page_img_index += 1
                if page_img_index >= 4:
                    page_img_index = 0
                continue

            # Insert image centered in cell
            success = safe_insert_image(cell, stream, width_inches=3.0)
            if success:
                total_inserted += 1
                # Add caption under the image in the same cell
                cap_para = cell.add_paragraph()
                cap_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                cap_run = cap_para.add_run(f"Imagen {total_inserted}. ________________________________")
            else:
                registro["omitidos"].append((imgname, "Fallo al insertar en el documento"))
                print("Fallo al insertar imagen:", imgname)
            page_img_index += 1
            if page_img_index >= 4:
                page_img_index = 0
        # end group
    registro["imagenes_insertadas"] = total_inserted

    # Save document
    try:
        doc.save(out_path)
    except PermissionError as e:
        msg = f"No se pudo guardar el archivo. ¿Existe un documento abierto con ese nombre? {e}"
        registro["errores"].append(msg)
        write_log(registro)
        print(msg)
        return False
    except Exception as e:
        registro["errores"].append(f"Error guardando Word: {e}")
        write_log(registro)
        print("Error guardando documento:", e)
        return False

    # Final registro
    write_log(registro)
    # Console summary
    print(f"Informe generado: {out_path}")
    print("Imágenes encontradas por carpeta:")
    for k, v in registro["imagenes_encontradas"].items():
        print(f" - {k}: {v}")
    print(f"Imágenes insertadas correctamente: {registro['imagenes_insertadas']}")
    if registro["omitidos"]:
        print("Imágenes omitidas:")
        for fn, motivo in registro["omitidos"]:
            print(f" - {fn}: {motivo}")
    if registro["errores"]:
        print("Errores:")
        for e in registro["errores"]:
            print(" -", e)
    return True

def write_log(reg):
    # Append summary to registro_generacion.txt
    try:
        with open(REGISTRO_FILE, "a", encoding="utf-8") as f:
            f.write("=== Ejecución: {} ===\n".format(reg.get("timestamp", "")))
            f.write(f"Número de historia: {reg.get('numero_historia')}\n")
            f.write(f"Especie: {reg.get('especie')}\n")
            f.write(f"Archivo generado: {reg.get('archivo_generado')}\n")
            f.write("Imágenes encontradas por carpeta:\n")
            for k, v in reg.get("imagenes_encontradas", {}).items():
                f.write(f"  - {k}: {v}\n")
            f.write(f"Imágenes insertadas: {reg.get('imagenes_insertadas')}\n")
            if reg.get("omitidos"):
                f.write("Omitidos:\n")
                for fn, motivo in reg.get("omitidos"):
                    f.write(f"  - {fn}: {motivo}\n")
            if reg.get("errores"):
                f.write("Errores:\n")
                for e in reg.get("errores"):
                    f.write(f"  - {e}\n")
            f.write("\n")
    except Exception as e:
        print("No se pudo actualizar el registro:", e)

if __name__ == "__main__":
    try:
        ok = generate_report()
        if ok:
            print("Proceso finalizado correctamente.")
        else:
            print("Proceso finalizado con errores (ver mensaje anterior).")
    except Exception as exc:
        print("Excepción no controlada:")
        traceback.print_exc()
        sys.exit(1)