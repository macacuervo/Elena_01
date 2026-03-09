#!/usr/bin/env python3
"""Extrae datos de un PDF de accidente y rellena una plantilla .docm.

Versión reforzada para principiantes:
- valida entradas (ficheros y formato),
- informa qué campos se encontraron y cuáles faltan,
- muestra mensajes de error claros y sencillos.
"""
from __future__ import annotations

import argparse
import re
import sys
import zipfile
import zlib
from pathlib import Path

FIELD_SPECS: list[dict[str, str | bool]] = [
    {"key": "trabnombre", "label": "Nombre", "required": True},
    {"key": "trabapellido1", "label": "Apellido 1º", "required": True},
    {"key": "trabapellido2", "label": "Apellido 2º", "required": True},
    {"key": "fechaacci", "label": "Fecha de accidente", "required": True},
    {"key": "fechanac", "label": "Fecha nacimiento", "required": False},
    {"key": "empresa", "label": "Empresa", "required": True},
    {"key": "direccion1", "label": "Dirección trabajador", "required": False},
    {"key": "localidad1", "label": "Municipio trabajador", "required": False},
    {"key": "provincia1", "label": "Provincia trabajador", "required": False},
    {"key": "ocupacion", "label": "Ocupación", "required": False},
    {"key": "fechaingreso", "label": "Fecha ingreso", "required": False},
    {"key": "empresacalle", "label": "Calle empresa/centro", "required": False},
    {"key": "empresamunicipio", "label": "Municipio empresa/centro", "required": False},
    {"key": "empresaprovincia", "label": "Provincia empresa/centro", "required": False},
]


def extract_pdf_text(pdf_path: Path) -> str:
    raw = pdf_path.read_bytes()
    pieces: list[str] = []

    for m in re.finditer(rb"stream\r?\n(.*?)\r?\nendstream", raw, re.S):
        block = m.group(1)
        try:
            decoded = zlib.decompress(block).decode("latin1", "ignore")
        except Exception:
            continue

        for txt in re.findall(r"\((.*?)\)\s*Tj", decoded, re.S):
            pieces.append(txt)

        for arr in re.findall(r"\[(.*?)\]\s*TJ", decoded, re.S):
            pieces.extend(re.findall(r"\((.*?)\)", arr, re.S))

    text = "\n".join(pieces)
    text = text.replace(r"\(", "(").replace(r"\)", ")").replace(r"\\", "\\")
    text = text.replace("\xa0", " ")
    text = re.sub(r"\r\n?", "\n", text)
    text = re.sub(r"[ \t]+", " ", text)
    return text


def pick_after_label(text: str, label_regex: str, fallback: str = "") -> str:
    m = re.search(label_regex, text, re.I | re.S)
    return m.group(1).strip() if m else fallback


def extract_fields(text: str) -> dict[str, str]:
    data: dict[str, str] = {}

    data["trabnombre"] = pick_after_label(text, r"Nombre:\s*\n([^\n]+)")
    data["trabapellido1"] = pick_after_label(text, r"Apellido\s*1º:\s*\n([^\n]+)")
    data["trabapellido2"] = pick_after_label(text, r"Apellido\s*2º:\s*\n([^\n]+)")
    data["fechaacci"] = pick_after_label(text, r"Fecha de accidente:\s*\n([^\n]+)")
    data["fechanac"] = pick_after_label(text, r"Fecha nacimiento:\s*\n([^\n]+)")
    data["empresa"] = pick_after_label(text, r"Nombre o Razón Social:\s*\n([^\n]+)")

    domicilio = pick_after_label(text, r"Domicilio:\s*\n([^\n]+)")
    piso = pick_after_label(text, r"Domicilio:\s*\n[^\n]+\n([^\n]+)")
    data["direccion1"] = (domicilio + (" " + piso if piso else "")).strip()

    data["localidad1"] = pick_after_label(text, r"Municipio:\s*\n([^\n]+)")
    data["provincia1"] = pick_after_label(text, r"Provincia:\s*\n([^\n]+)")
    data["ocupacion"] = pick_after_label(text, r"Ocupación:\s*\n([^\n]+)")
    data["fechaingreso"] = pick_after_label(text, r"Fecha de ingreso en la empresa:\s*\n([^\n]+)")

    data["empresacalle"] = pick_after_label(
        text, r"Nombre o Razón Social:\s*\n[^\n]+\n\s*\nDomicilio:\s*\n([^\n]+)"
    )
    data["empresamunicipio"] = pick_after_label(
        text, r"Código Postal:\s*\n[^\n]+\n\s*\nMunicipio:\s*\n([^\n]+)"
    )
    data["empresaprovincia"] = pick_after_label(text, r"Municipio:\s*\n[^\n]+\n\s*\nProvincia:\s*\n([^\n]+)")

    return {k: v for k, v in data.items() if v}


def build_report(fields: dict[str, str]) -> tuple[list[str], list[str], list[str]]:
    found: list[str] = []
    missing_required: list[str] = []
    missing_optional: list[str] = []

    for spec in FIELD_SPECS:
        key = str(spec["key"])
        label = str(spec["label"])
        required = bool(spec["required"])

        if fields.get(key):
            found.append(f"{key} ({label})")
        elif required:
            missing_required.append(f"{key} ({label})")
        else:
            missing_optional.append(f"{key} ({label})")

    return found, missing_required, missing_optional


def find_placeholders(xml: str, keys: list[str]) -> tuple[list[str], list[str]]:
    present, absent = [], []
    for key in keys:
        if re.search(rf">{re.escape(key)}<", xml):
            present.append(key)
        else:
            absent.append(key)
    return present, absent


def replace_placeholders_in_docm(template_docm: Path, output_docm: Path, values: dict[str, str]) -> tuple[int, list[str]]:
    with zipfile.ZipFile(template_docm, "r") as zin:
        xml = zin.read("word/document.xml").decode("utf-8", "ignore")
        keys = sorted(values.keys())
        present, _absent = find_placeholders(xml, keys)

        replacements = 0
        for placeholder in present:
            value = values[placeholder]
            xml, count = re.subn(rf">{re.escape(placeholder)}<", f">{value}<", xml)
            replacements += count

        with zipfile.ZipFile(output_docm, "w") as zout:
            for item in zin.infolist():
                data = xml.encode("utf-8") if item.filename == "word/document.xml" else zin.read(item.filename)
                zout.writestr(item, data)

    return replacements, present


def validate_inputs(pdf: Path, template: Path) -> None:
    if not pdf.exists():
        raise FileNotFoundError(f"No existe el PDF: {pdf}")
    if not template.exists():
        raise FileNotFoundError(f"No existe la plantilla Word: {template}")
    if pdf.suffix.lower() != ".pdf":
        raise ValueError(f"El archivo PDF no tiene extensión .pdf: {pdf}")
    if template.suffix.lower() not in {".docm", ".docx"}:
        raise ValueError(f"La plantilla debe ser .docm o .docx: {template}")


def print_user_report(fields: dict[str, str], replaced: int, matched_placeholders: list[str], output: Path) -> None:
    found, missing_required, missing_optional = build_report(fields)

    print("\n=== INFORME DE EXTRACCIÓN ===")
    print(f"Campos encontrados: {len(found)}")
    for item in found:
        key = item.split(" ")[0]
        print(f"  ✅ {item}: {fields[key]}")

    print(f"\nCampos obligatorios no encontrados: {len(missing_required)}")
    for item in missing_required:
        print(f"  ❌ {item}")

    print(f"\nCampos opcionales no encontrados: {len(missing_optional)}")
    for item in missing_optional:
        print(f"  ⚠️ {item}")

    print("\n=== INFORME DE RELLENO DE PLANTILLA ===")
    print(f"Placeholders encontrados en Word y reemplazados: {len(matched_placeholders)}")
    print(f"Número total de reemplazos realizados: {replaced}")
    print(f"Documento generado: {output}")


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Extrae datos de un PDF y rellena una plantilla Word (.docm/.docx)."
    )
    parser.add_argument("--pdf", required=True, type=Path, help="Ruta del PDF de entrada")
    parser.add_argument("--template", required=True, type=Path, help="Ruta de la plantilla Word (.docm)")
    parser.add_argument("--output", required=True, type=Path, help="Ruta del Word de salida")
    args = parser.parse_args()

    try:
        validate_inputs(args.pdf, args.template)
        text = extract_pdf_text(args.pdf)
        if not text.strip():
            raise ValueError("No se pudo extraer texto del PDF. Puede ser un PDF escaneado o no compatible.")

        fields = extract_fields(text)
        replaced, matched_placeholders = replace_placeholders_in_docm(args.template, args.output, fields)
        print_user_report(fields, replaced, matched_placeholders, args.output)

        _found, missing_required, _missing_optional = build_report(fields)
        if missing_required:
            print(
                "\nConsejo: faltan campos obligatorios. Revisa el PDF o ajusta las reglas del script para ese formato."
            )
        return 0

    except FileNotFoundError as exc:
        print(f"\n❌ Error: {exc}")
        print("Consejo: revisa que las rutas sean correctas y vuelve a ejecutar el comando.")
        return 1
    except (ValueError, zipfile.BadZipFile) as exc:
        print(f"\n❌ Error: {exc}")
        print("Consejo: usa archivos válidos (.pdf y .docm/.docx) y vuelve a intentarlo.")
        return 1
    except Exception as exc:
        print(f"\n❌ Error inesperado: {exc}")
        print("Consejo: vuelve a ejecutar y, si persiste, comparte este mensaje para revisarlo.")
        return 1


if __name__ == "__main__":
    sys.exit(main())
