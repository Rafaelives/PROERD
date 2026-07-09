from collections import Counter, defaultdict
import csv
from datetime import date, timedelta
from functools import lru_cache
import html
import json
from math import sqrt
import os
from pathlib import Path
import re
import socket
import xml.etree.ElementTree as ET
from urllib.request import urlopen
import unicodedata
from zipfile import ZipFile

from flask import Flask, jsonify, render_template, send_file

app = Flask(__name__)

XLSX_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
EXCEL_EPOCH = date(1899, 12, 30)
MAP_CRIME_CATEGORIES = {
    "CVLI": {"label": "CVLI", "color": "#d62828"},
    "FURTO": {"label": "Furto", "color": "#2fb344"},
    "CVP": {"label": "CVP", "color": "#116b35"},
    "DISPARO DE ARMA": {"label": "Disparo de arma de fogo", "color": "#6b7280"},
    "MARIA DA PENHA": {"label": "Maria da Penha", "color": "#7c3aed"},
    "PERTURBAÇÃO AO SOSSEGO ALHEIO": {
        "label": "Perturbação do sossego alheio",
        "color": "#1664d9",
    },
    "LESÃO CORPORAL": {"label": "Lesão corporal", "color": "#b95622"},
    "ACHADO DE CADÁVER": {"label": "Achado de cadáver", "color": "#111111"},
    "TENTATIVA DE HOMICÍDIO": {"label": "Tentativa de homicídio", "color": "#12c8c6"},
}
MAPPING_SHEET_NAME = "Mapeamento"
TEACHING_SHEET_NAME = "monit_instr"
CEARA_MALHA_URL = (
    "https://servicodados.ibge.gov.br/api/v3/malhas/estados/23"
    "?formato=application/vnd.geo+json&qualidade=minima&intrarregiao=municipio"
)
CEARA_MUNICIPIOS_URL = "https://servicodados.ibge.gov.br/api/v1/localidades/estados/23/municipios"


def _format_number(value, decimals=0):
    if decimals:
        text = f"{value:,.{decimals}f}"
    else:
        text = f"{value:,.0f}"
    return text.replace(",", "X").replace(".", ",").replace("X", ".")


def _to_float(value):
    try:
        return float(str(value).replace(",", "."))
    except (TypeError, ValueError):
        return None


def _column_index(cell_ref):
    letters = re.match(r"([A-Z]+)", cell_ref).group(1)
    index = 0
    for letter in letters:
        index = index * 26 + ord(letter) - 64
    return index - 1


def _shared_strings(workbook):
    if "xl/sharedStrings.xml" not in workbook.namelist():
        return []

    root = ET.fromstring(workbook.read("xl/sharedStrings.xml"))
    return [
        "".join(node.text or "" for node in item.iter(f"{XLSX_NS}t"))
        for item in root.findall(f"{XLSX_NS}si")
    ]


def _sheet_path_by_name(workbook, sheet_name):
    workbook_xml = ET.fromstring(workbook.read("xl/workbook.xml"))
    rels_xml = ET.fromstring(workbook.read("xl/_rels/workbook.xml.rels"))
    rels = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels_xml}

    for sheet in workbook_xml.find(f"{XLSX_NS}sheets"):
        if sheet.attrib["name"] == sheet_name:
            return f"xl/{rels[sheet.attrib[f'{REL_NS}id']]}"

    available = [sheet.attrib["name"] for sheet in workbook_xml.find(f"{XLSX_NS}sheets")]
    raise ValueError(f"Aba {sheet_name!r} não encontrada. Abas disponíveis: {available}")


def _read_xlsx_sheet(path, sheet_name):
    with ZipFile(path) as workbook:
        shared_strings = _shared_strings(workbook)
        sheet_path = _sheet_path_by_name(workbook, sheet_name)
        root = ET.fromstring(workbook.read(sheet_path))

        rows = []
        for row in root.find(f"{XLSX_NS}sheetData").findall(f"{XLSX_NS}row"):
            values = []
            for cell in row.findall(f"{XLSX_NS}c"):
                index = _column_index(cell.attrib["r"])
                while len(values) <= index:
                    values.append("")

                cell_type = cell.attrib.get("t")
                raw_value = cell.find(f"{XLSX_NS}v")

                if cell_type == "s" and raw_value is not None:
                    value = shared_strings[int(raw_value.text)]
                elif cell_type == "inlineStr":
                    value = "".join(node.text or "" for node in cell.iter(f"{XLSX_NS}t"))
                else:
                    value = raw_value.text if raw_value is not None else ""

                values[index] = value
            rows.append(values)

    headers = rows[0]
    records = []
    for row in rows[1:]:
        records.append({header: row[index] if index < len(row) else "" for index, header in enumerate(headers)})
    return records


def _workbook_path():
    project_dir = Path(__file__).resolve().parent
    matches = list(project_dir.glob("Ana*lise.xlsx"))
    if not matches:
        raise FileNotFoundError("Planilha Análise.xlsx não encontrada na pasta do projeto.")
    return matches[0]


def _project_dir():
    return Path(__file__).resolve().parent


def _ceara_asset_paths():
    data_dir = _project_dir() / "static" / "data"
    return {
        "data_dir": data_dir,
        "geojson": data_dir / "ceara_municipios.geojson",
        "kml": data_dir / "ceara_municipios.kml",
        "csv": data_dir / "ceara_municipios.csv",
        "pdf": data_dir / "ceara_municipios.pdf",
    }


def _fetch_json(url):
    with urlopen(url, timeout=40) as response:
        return json.loads(response.read().decode("utf-8"))


def _ceara_geojson():
    geojson = _fetch_json(CEARA_MALHA_URL)
    names = {
        str(item["id"]): item["nome"]
        for item in _fetch_json(CEARA_MUNICIPIOS_URL)
    }

    for feature in geojson.get("features", []):
        properties = feature.setdefault("properties", {})
        code = str(properties.get("codarea", ""))
        properties["nome"] = names.get(code, code)

    geojson["features"].sort(key=lambda item: item.get("properties", {}).get("nome", ""))
    return geojson


def _kml_coordinates(ring):
    return " ".join(f"{lon},{lat},0" for lon, lat in ring)


def _kml_polygon(coordinates):
    rings = []
    if not coordinates:
        return ""

    rings.append(
        "<outerBoundaryIs><LinearRing><coordinates>"
        f"{_kml_coordinates(coordinates[0])}"
        "</coordinates></LinearRing></outerBoundaryIs>"
    )
    for inner_ring in coordinates[1:]:
        rings.append(
            "<innerBoundaryIs><LinearRing><coordinates>"
            f"{_kml_coordinates(inner_ring)}"
            "</coordinates></LinearRing></innerBoundaryIs>"
        )

    return f"<Polygon>{''.join(rings)}</Polygon>"


def _feature_to_kml_placemark(feature):
    properties = feature.get("properties", {})
    name = html.escape(properties.get("nome") or properties.get("codarea") or "Município")
    code = html.escape(str(properties.get("codarea", "")))
    geometry = feature.get("geometry") or {}
    geometry_type = geometry.get("type")
    coordinates = geometry.get("coordinates") or []

    if geometry_type == "Polygon":
        geometry_markup = _kml_polygon(coordinates)
    elif geometry_type == "MultiPolygon":
        geometry_markup = "<MultiGeometry>" + "".join(_kml_polygon(polygon) for polygon in coordinates) + "</MultiGeometry>"
    else:
        geometry_markup = ""

    return (
        "<Placemark>"
        f"<name>{name}</name>"
        "<styleUrl>#municipio</styleUrl>"
        "<ExtendedData>"
        f'<Data name="codarea"><value>{code}</value></Data>'
        f'<Data name="nome"><value>{name}</value></Data>'
        "</ExtendedData>"
        f"{geometry_markup}"
        "</Placemark>"
    )


def _geojson_to_kml(geojson):
    placemarks = "\n".join(_feature_to_kml_placemark(feature) for feature in geojson.get("features", []))
    return (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<kml xmlns="http://www.opengis.net/kml/2.2">\n'
        "<Document>\n"
        "<name>Divisas municipais do Ceará - IBGE</name>\n"
        "<Style id=\"municipio\">"
        "<LineStyle><color>ff0b4f6c</color><width>1.8</width></LineStyle>"
        "<PolyStyle><color>261ccad8</color></PolyStyle>"
        "</Style>\n"
        f"{placemarks}\n"
        "</Document>\n"
        "</kml>\n"
    )


def _kml_style_id(properties):
    dados = properties.get("dados") or {}
    schools = dados.get("escolas", 0)
    if not dados.get("has_data"):
        return "sem_dados"
    if schools >= 10:
        return "alta"
    if schools >= 5:
        return "media"
    return "baixa"


def _kml_data(name, value):
    return f'<Data name="{html.escape(name)}"><value>{html.escape(str(value))}</value></Data>'


def _feature_to_school_kml_placemark(feature):
    properties = feature.get("properties", {})
    dados = properties.get("dados") or {}
    name = html.escape(properties.get("nome") or properties.get("codarea") or "Município")
    code = properties.get("codarea", "")
    schools = int(dados.get("escolas") or 0)
    students = int(dados.get("alunos") or 0)
    average = round(students / schools) if schools else 0
    geometry = feature.get("geometry") or {}
    geometry_type = geometry.get("type")
    coordinates = geometry.get("coordinates") or []

    if geometry_type == "Polygon":
        geometry_markup = _kml_polygon(coordinates)
    elif geometry_type == "MultiPolygon":
        geometry_markup = "<MultiGeometry>" + "".join(_kml_polygon(polygon) for polygon in coordinates) + "</MultiGeometry>"
    else:
        geometry_markup = ""

    description = html.escape(
        f"Escolas: {schools}\n"
        f"Alunos: {students}\n"
        f"Alunos por escola: {average}"
    )
    extended_data = "".join(
        [
            _kml_data("codigo_ibge", code),
            _kml_data("municipio", properties.get("nome", "")),
            _kml_data("municipio_tabela", dados.get("municipio_tabela", "")),
            _kml_data("crpm", dados.get("crpm", "")),
            _kml_data("ais", dados.get("ais", "")),
            _kml_data("escolas", schools),
            _kml_data("alunos", students),
            _kml_data("alunos_por_escola", average),
            _kml_data("faixa_escolas", dados.get("faixa_escolas", "")),
            _kml_data("faixa_alunos", dados.get("faixa_alunos", "")),
            _kml_data("tem_dados", "sim" if dados.get("has_data") else "nao"),
        ]
    )

    return (
        "<Placemark>"
        f"<name>{name}</name>"
        f"<description>{description}</description>"
        f"<styleUrl>#{_kml_style_id(properties)}</styleUrl>"
        f"<ExtendedData>{extended_data}</ExtendedData>"
        f"{geometry_markup}"
        "</Placemark>"
    )


def _school_geojson_to_kml(geojson):
    placemarks = "\n".join(_feature_to_school_kml_placemark(feature) for feature in geojson.get("features", []))
    return (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<kml xmlns="http://www.opengis.net/kml/2.2">\n'
        "<Document>\n"
        "<name>Escolas por município do Ceará</name>\n"
        "<description>Mapa exportado da aba DADOS com quantidade de escolas e alunos por município.</description>\n"
        '<Style id="sem_dados"><LineStyle><color>ff78858c</color><width>1.1</width></LineStyle>'
        '<PolyStyle><color>55e5e0d8</color></PolyStyle></Style>\n'
        '<Style id="baixa"><LineStyle><color>ff34535d</color><width>1.2</width></LineStyle>'
        '<PolyStyle><color>9940b4f0</color></PolyStyle></Style>\n'
        '<Style id="media"><LineStyle><color>ff34535d</color><width>1.4</width></LineStyle>'
        '<PolyStyle><color>99cb7a1f</color></PolyStyle></Style>\n'
        '<Style id="alta"><LineStyle><color>ff34535d</color><width>1.7</width></LineStyle>'
        '<PolyStyle><color>997a7f08</color></PolyStyle></Style>\n'
        f"{placemarks}\n"
        "</Document>\n"
        "</kml>\n"
    )


def export_school_kml():
    paths = _ceara_asset_paths()
    paths["data_dir"].mkdir(parents=True, exist_ok=True)
    data = get_territorial_data()
    paths["kml"].write_text(_school_geojson_to_kml(data["geojson"]), encoding="utf-8")
    return paths["kml"]


def _school_export_rows(data):
    rows = []
    for record in data["mapped_rows"]:
        schools = int(record.get("escolas") or 0)
        students = int(record.get("alunos") or 0)
        average = round(students / schools) if schools else 0
        rows.append(
            {
                "municipio": record.get("municipio_tabela", ""),
                "municipio_mapeado": record.get("municipio_mapeado", ""),
                "crpm": record.get("crpm", ""),
                "ais": record.get("ais", ""),
                "escolas": schools,
                "alunos": students,
                "alunos_por_escola": average,
                "faixa_escolas": record.get("faixa_escolas", ""),
                "faixa_alunos": record.get("faixa_alunos", ""),
            }
        )
    return rows


def export_school_csv():
    paths = _ceara_asset_paths()
    paths["data_dir"].mkdir(parents=True, exist_ok=True)
    data = get_territorial_data()
    rows = _school_export_rows(data)
    fieldnames = [
        "municipio",
        "municipio_mapeado",
        "crpm",
        "ais",
        "escolas",
        "alunos",
        "alunos_por_escola",
        "faixa_escolas",
        "faixa_alunos",
    ]

    with paths["csv"].open("w", encoding="utf-8-sig", newline="") as output:
        writer = csv.DictWriter(output, fieldnames=fieldnames, delimiter=";", lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)

    return paths["csv"]


def _pdf_text(value):
    text = str(value)
    replacements = {
        "\\": "\\\\",
        "(": "\\(",
        ")": "\\)",
        "\r": " ",
        "\n": " ",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    return text


def _pdf_cell(value, width):
    text = str(value)
    return text if len(text) <= width else f"{text[: max(0, width - 3)]}..."


def _simple_pdf(title, lines):
    page_width = 595
    page_height = 842
    margin_x = 42
    line_height = 14
    lines_per_page = 50
    pages = [lines[index:index + lines_per_page] for index in range(0, len(lines), lines_per_page)] or [[]]
    objects = []

    def add_object(content):
        objects.append(content)
        return len(objects)

    font_id = add_object("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>")
    page_ids = []

    for page_number, page_lines in enumerate(pages, start=1):
        commands = [
            "BT",
            "/F1 16 Tf",
            f"{margin_x} {page_height - 48} Td",
            f"({_pdf_text(title)}) Tj",
            "/F1 9 Tf",
            f"0 -{line_height * 2} Td",
        ]

        for line in page_lines:
            commands.append(f"({_pdf_text(line)}) Tj")
            commands.append(f"0 -{line_height} Td")

        commands.extend(
            [
                "/F1 8 Tf",
                f"0 -{line_height} Td",
                f"(Pagina {page_number} de {len(pages)}) Tj",
                "ET",
            ]
        )
        stream = "\n".join(commands)
        content_id = add_object(f"<< /Length {len(stream.encode('cp1252', errors='replace'))} >>\nstream\n{stream}\nendstream")
        page_id = add_object(
            f"<< /Type /Page /Parent 0 0 R /MediaBox [0 0 {page_width} {page_height}] "
            f"/Resources << /Font << /F1 {font_id} 0 R >> >> /Contents {content_id} 0 R >>"
        )
        page_ids.append(page_id)

    pages_id = len(objects) + 1
    for page_id in page_ids:
        objects[page_id - 1] = objects[page_id - 1].replace("/Parent 0 0 R", f"/Parent {pages_id} 0 R")
    kids = " ".join(f"{page_id} 0 R" for page_id in page_ids)
    add_object(f"<< /Type /Pages /Kids [{kids}] /Count {len(page_ids)} >>")
    catalog_id = add_object(f"<< /Type /Catalog /Pages {pages_id} 0 R >>")

    output = bytearray(b"%PDF-1.4\n")
    offsets = [0]
    for index, content in enumerate(objects, start=1):
        offsets.append(len(output))
        output.extend(f"{index} 0 obj\n".encode("ascii"))
        output.extend(content.encode("cp1252", errors="replace"))
        output.extend(b"\nendobj\n")

    xref_offset = len(output)
    output.extend(f"xref\n0 {len(objects) + 1}\n".encode("ascii"))
    output.extend(b"0000000000 65535 f \n")
    for offset in offsets[1:]:
        output.extend(f"{offset:010d} 00000 n \n".encode("ascii"))
    output.extend(
        (
            "trailer\n"
            f"<< /Size {len(objects) + 1} /Root {catalog_id} 0 R >>\n"
            "startxref\n"
            f"{xref_offset}\n"
            "%%EOF\n"
        ).encode("ascii")
    )
    return bytes(output)


def export_school_pdf():
    paths = _ceara_asset_paths()
    paths["data_dir"].mkdir(parents=True, exist_ok=True)
    data = get_territorial_data()
    summary = data["summary"]
    rows = _school_export_rows(data)
    lines = [
        f"Fonte: {data['source']}",
        f"Municipios com dados: {summary['municipalities_with_data']} de {summary['municipalities_total']}",
        f"Total de escolas: {_format_number(summary['schools_total'])}",
        f"Total de alunos: {_format_number(summary['students_total'])}",
        "",
        f"{'Municipio':<28} {'Escolas':>8} {'Alunos':>10} {'Al./esc.':>8}",
        "-" * 60,
    ]

    for row in rows:
        lines.append(
            f"{_pdf_cell(row['municipio'], 28):<28} "
            f"{_format_number(row['escolas']):>8} "
            f"{_format_number(row['alunos']):>10} "
            f"{_format_number(row['alunos_por_escola']):>8}"
        )

    paths["pdf"].write_bytes(_simple_pdf("Escolas por municipio do Ceara", lines))
    return paths["pdf"]


def _coordinate_bounds(geojson):
    lats = []
    lons = []

    def visit(value):
        if isinstance(value, list) and value and isinstance(value[0], (int, float)):
            lons.append(value[0])
            lats.append(value[1])
            return
        if isinstance(value, list):
            for item in value:
                visit(item)

    for feature in geojson.get("features", []):
        visit((feature.get("geometry") or {}).get("coordinates", []))

    if not lats or not lons:
        return [[-7.9, -41.5], [-2.7, -37.1]]
    return [[min(lats), min(lons)], [max(lats), max(lons)]]


def ensure_ceara_boundary_assets():
    paths = _ceara_asset_paths()
    paths["data_dir"].mkdir(parents=True, exist_ok=True)

    if paths["geojson"].exists() and paths["kml"].exists():
        return paths

    geojson = _ceara_geojson()
    paths["geojson"].write_text(json.dumps(geojson, ensure_ascii=False), encoding="utf-8")
    paths["kml"].write_text(_geojson_to_kml(geojson), encoding="utf-8")
    return paths


def _normalize_key(value):
    text = unicodedata.normalize("NFKD", str(value or "").strip())
    text = "".join(char for char in text if not unicodedata.combining(char))
    return re.sub(r"[^a-z0-9]+", "", text.lower())


def _clean_dados_record(record):
    return {
        "ais": (record.get("AIS") or "").strip(),
        "crpm": (record.get("CRPM") or "").strip(),
        "batalhao": (record.get("BATALHÃO") or "").strip(),
        "cidade": (record.get("CIDADE") or "").strip(),
        "bairro": (record.get("BAIRRO") or "").strip(),
    }


def _load_ceara_geojson():
    paths = ensure_ceara_boundary_assets()
    return json.loads(paths["geojson"].read_text(encoding="utf-8"))


def _count_options(records, key):
    counts = Counter()
    for record in records:
        value = record.get(key)
        if isinstance(value, list):
            counts.update(item for item in value if item)
        elif value:
            counts[value] += 1
    return [
        {"label": label, "count": count}
        for label, count in sorted(counts.items(), key=lambda item: (-item[1], item[0]))
    ]


def _unique_labels(values):
    return sorted({value for value in values if value and value != "Não informado"})


def _join_labels(values):
    labels = _unique_labels(values)
    return ", ".join(labels) if labels else "Não informado"


MUNICIPALITY_ALIASES = {
    "itapage": "itapaje",
    "limoeiro": "limoeirodonorte",
}


def _municipality_key(value):
    key = _normalize_key(value)
    return MUNICIPALITY_ALIASES.get(key, key)


def _to_int(value):
    number = _to_float(value)
    return int(number) if number is not None else 0


def _school_range(value):
    if value >= 15:
        return "15+ escolas"
    if value >= 10:
        return "10 a 14 escolas"
    if value >= 5:
        return "5 a 9 escolas"
    if value >= 1:
        return "1 a 4 escolas"
    return "Sem escolas"


def _student_range(value):
    if value >= 1500:
        return "1.500+ alunos"
    if value >= 1000:
        return "1.000 a 1.499 alunos"
    if value >= 500:
        return "500 a 999 alunos"
    if value >= 1:
        return "1 a 499 alunos"
    return "Sem alunos"


def _clean_mapping_label(record, key):
    return (record.get(key) or "Não informado").strip() or "Não informado"


def _first_mapping_label(record, *keys):
    for key in keys:
        value = (record.get(key) or "").strip()
        if value:
            return value
    return "Não informado"


def _mapping_cities(value):
    cities = [
        city.strip()
        for city in re.split(r"[/;]", str(value or ""))
        if city.strip()
    ]
    return cities or ["Não informado"]


def _mapping_role_group(value):
    normalized = _normalize_key(value)
    if "instrutor" in normalized:
        return "Instrutor"
    if "monitor" in normalized:
        return "Monitor"
    if "coord" in normalized:
        return "Coordenador"
    return "Outra função"


def _mapping_option_counts(records, key):
    counter = Counter()
    for record in records:
        value = record.get(key)
        if isinstance(value, list):
            counter.update(item for item in value if item)
        elif value:
            counter[value] += 1
    return [
        {"label": label, "count": count}
        for label, count in sorted(counter.items(), key=lambda item: (-item[1], item[0]))
    ]


def _mapping_teaching_details(records):
    details_by_registration = {}

    for record in records:
        registration_key = _normalize_key(record.get("MATRICULA"))
        if not registration_key:
            continue

        details = details_by_registration.setdefault(
            registration_key,
            {
                "school_names": set(),
                "classes": set(),
                "students_total": 0,
            },
        )
        school = (record.get("ESCOLA") or "").strip()
        class_name = (record.get("TURMAS") or "").strip()
        if school:
            details["school_names"].add(school)
        if class_name:
            details["classes"].add((school, class_name))
        details["students_total"] += _to_int(record.get("ALUNOS"))

    return {
        registration_key: {
            "school_names": sorted(details["school_names"], key=_normalize_key),
            "schools_total": len(details["school_names"]),
            "classes_total": len(details["classes"]),
            "students_total": details["students_total"],
        }
        for registration_key, details in details_by_registration.items()
    }


def _mapping_city_rows(records, geojson):
    municipality_by_key = {
        _municipality_key(feature.get("properties", {}).get("nome")): feature
        for feature in geojson.get("features", [])
    }
    rows_by_municipality = {}
    unmatched_rows = []

    for record in records:
        matched_any = False
        for city in record["cities"]:
            feature = municipality_by_key.get(_municipality_key(city))
            if not feature:
                continue

            matched_any = True
            municipality = feature["properties"]["nome"]
            key = _municipality_key(municipality)
            row = rows_by_municipality.setdefault(
                key,
                {
                    "municipio_mapeado": municipality,
                    "registros": 0,
                    "regions": [],
                    "battalions": [],
                    "functions": [],
                    "role_groups": [],
                    "post_grads": [],
                    "post_roles": [],
                    "people": [],
                    "rows": [],
                },
            )
            row["registros"] += 1
            row["rows"].append(record)
            row["people"].append(record["name"])
            row["regions"] = _unique_labels([*row["regions"], record["region"]])
            row["battalions"] = _unique_labels([*row["battalions"], record["battalion"]])
            row["functions"] = _unique_labels([*row["functions"], record["function"]])
            row["role_groups"] = _unique_labels([*row["role_groups"], record["role_group"]])
            row["post_grads"] = _unique_labels([*row["post_grads"], record["post_grad"]])
            row["post_roles"] = _unique_labels([*row["post_roles"], record["post_role"]])

        if not matched_any:
            unmatched_rows.append(record)

    for feature in geojson.get("features", []):
        properties = feature.setdefault("properties", {})
        row = rows_by_municipality.get(_municipality_key(properties.get("nome")))
        properties["mapeamento"] = {
            "has_data": bool(row),
            "registros": row["registros"] if row else 0,
            "regions": row["regions"] if row else [],
            "battalions": row["battalions"] if row else [],
            "functions": row["functions"] if row else [],
            "role_groups": row["role_groups"] if row else [],
            "post_grads": row["post_grads"] if row else [],
            "post_roles": row["post_roles"] if row else [],
            "people": row["people"] if row else [],
        }

    return list(rows_by_municipality.values()), unmatched_rows


def get_territorial_data():
    geojson = _load_ceara_geojson()
    raw_records = _read_xlsx_sheet(_workbook_path(), "DADOS")
    dados_records = []

    for index, record in enumerate(raw_records, start=1):
        municipality = (record.get("Quantidade de Escolas por Município") or "").strip()
        schools = _to_int(record.get("ESCOLAS"))
        students = _to_int(record.get("qtd_alunos"))
        if not municipality:
            continue

        dados_records.append(
            {
                "id": index,
                "municipio_tabela": municipality,
                "municipio_mapeado": "",
                "crpm": (record.get("CRPM") or "Não informado").strip() or "Não informado",
                "ais": (record.get("AIS") or "Não informado").strip() or "Não informado",
                "escolas": schools,
                "alunos": students,
                "faixa_escolas": _school_range(schools),
                "faixa_alunos": _student_range(students),
            }
        )

    municipality_by_key = {
        _municipality_key(feature.get("properties", {}).get("nome")): feature
        for feature in geojson.get("features", [])
    }
    rows_by_municipality = {}
    unmatched_rows = []

    for row in dados_records:
        municipality_key = _municipality_key(row["municipio_tabela"])
        feature = municipality_by_key.get(municipality_key)
        if feature:
            row["municipio_mapeado"] = feature["properties"]["nome"]
            grouped_row = rows_by_municipality.setdefault(
                municipality_key,
                {
                    "id": row["id"],
                    "municipio_tabela": row["municipio_tabela"],
                    "municipio_mapeado": row["municipio_mapeado"],
                    "crpm": "",
                    "ais": "",
                    "crpm_values": [],
                    "ais_values": [],
                    "escolas": 0,
                    "alunos": 0,
                    "faixa_escolas": "",
                    "faixa_alunos": "",
                    "rows": [],
                },
            )
            grouped_row["rows"].append(row)
            grouped_row["escolas"] += row["escolas"]
            grouped_row["alunos"] += row["alunos"]
            grouped_row["crpm_values"] = _unique_labels([*grouped_row["crpm_values"], row["crpm"]])
            grouped_row["ais_values"] = _unique_labels([*grouped_row["ais_values"], row["ais"]])
            grouped_row["crpm"] = _join_labels(grouped_row["crpm_values"])
            grouped_row["ais"] = _join_labels(grouped_row["ais_values"])
            grouped_row["faixa_escolas"] = _school_range(grouped_row["escolas"])
            grouped_row["faixa_alunos"] = _student_range(grouped_row["alunos"])
        else:
            unmatched_rows.append(row)

    for feature in geojson.get("features", []):
        properties = feature.setdefault("properties", {})
        key = _municipality_key(properties.get("nome"))
        row = rows_by_municipality.get(key)

        properties["dados"] = {
            "has_data": bool(row),
            "municipio_tabela": row["municipio_tabela"] if row else "",
            "escolas": row["escolas"] if row else 0,
            "alunos": row["alunos"] if row else 0,
            "crpm": row["crpm"] if row else "",
            "ais": row["ais"] if row else "",
            "crpm_values": row["crpm_values"] if row else [],
            "ais_values": row["ais_values"] if row else [],
            "faixa_escolas": row["faixa_escolas"] if row else "Sem escolas",
            "faixa_alunos": row["faixa_alunos"] if row else "Sem alunos",
            "rows": row["rows"] if row else [],
        }

    mapped_rows = list(rows_by_municipality.values())
    schools_total = sum(record["escolas"] for record in dados_records)
    students_total = sum(record["alunos"] for record in dados_records)
    max_schools = max((record["escolas"] for record in mapped_rows), default=0)
    max_students = max((record["alunos"] for record in mapped_rows), default=0)

    return {
        "source": "Municípios - Escolas/CRPM",
        "geojson": geojson,
        "records": dados_records,
        "mapped_rows": mapped_rows,
        "unmatched_rows": unmatched_rows,
        "filters": {
            "faixa_escolas": _count_options(dados_records, "faixa_escolas"),
            "faixa_alunos": _count_options(dados_records, "faixa_alunos"),
            "crpm": _count_options(mapped_rows, "crpm_values"),
            "ais": _count_options(mapped_rows, "ais_values"),
            "cidade": _count_options(mapped_rows, "municipio_mapeado"),
        },
        "summary": {
            "total_rows": len(dados_records),
            "mapped_rows": len(mapped_rows),
            "unmatched_rows": len(unmatched_rows),
            "municipalities_total": len(geojson.get("features", [])),
            "municipalities_with_data": len(rows_by_municipality),
            "schools_total": schools_total,
            "students_total": students_total,
            "max_schools": max_schools,
            "max_students": max_students,
        },
    }


def _available_port(preferred_port):
    for port in range(preferred_port, preferred_port + 50):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            try:
                sock.bind(("127.0.0.1", port))
            except OSError:
                continue
            return port
    raise RuntimeError(f"Nenhuma porta livre encontrada a partir de {preferred_port}.")


def _excel_date_range(records):
    serials = []
    for record in records:
        serial = _to_float(record.get("DATA"))
        if serial is not None:
            serials.append(int(serial))

    if not serials:
        return "Período não identificado"

    start = EXCEL_EPOCH + timedelta(days=min(serials))
    end = EXCEL_EPOCH + timedelta(days=max(serials))
    return f"{start.strftime('%d/%m/%Y')} a {end.strftime('%d/%m/%Y')}"


def _excel_dates(records):
    dates = []
    for record in records:
        serial = _to_float(record.get("DATA"))
        if serial is not None:
            dates.append(EXCEL_EPOCH + timedelta(days=int(serial)))
    return dates


def _top_counter(records, column, limit=5):
    return Counter(record.get(column) or "Não informado" for record in records).most_common(limit)


def _pearson_correlation(points):
    if len(points) < 2:
        return None

    xs = [point["x"] for point in points]
    ys = [point["y"] for point in points]
    mean_x = sum(xs) / len(xs)
    mean_y = sum(ys) / len(ys)
    covariance = sum((x - mean_x) * (y - mean_y) for x, y in zip(xs, ys))
    variance_x = sum((x - mean_x) ** 2 for x in xs)
    variance_y = sum((y - mean_y) ** 2 for y in ys)
    denominator = sqrt(variance_x * variance_y)
    return covariance / denominator if denominator else None


def _map_crime_key(crime):
    normalized = (crime or "").strip().upper()
    if normalized.startswith("DISPARO DE ARMA"):
        return "DISPARO DE ARMA"
    return normalized if normalized in MAP_CRIME_CATEGORIES else None


def _crime_color(crime):
    key = _map_crime_key(crime)
    if key:
        return MAP_CRIME_CATEGORIES[key]["color"]
    return "#1f2937"


def _city_metrics(records):
    cities = {}
    crimes_by_city = defaultdict(Counter)

    for record in records:
        city = record.get("CICADE") or record.get("CIDADE") or "Não informado"
        crime = record.get("CRIME") or "Não informado"
        population = _to_float(record.get("POPULAÇÃO"))
        lat = _to_float(record.get("LAT"))
        lon = _to_float(record.get("LONG"))

        if city not in cities:
            cities[city] = {"count": 0, "population": None, "lat": [], "lon": []}

        cities[city]["count"] += 1
        crimes_by_city[city][crime] += 1

        if population:
            cities[city]["population"] = population
        if lat is not None and lon is not None:
            cities[city]["lat"].append(lat)
            cities[city]["lon"].append(lon)

    metrics = []
    for city, data in cities.items():
        rate = None
        if data["population"]:
            rate = data["count"] / data["population"] * 10000

        metrics.append(
            {
                "city": city,
                "count": data["count"],
                "rate": rate,
                "top_crime": crimes_by_city[city].most_common(1)[0][0],
                "lat": sum(data["lat"]) / len(data["lat"]) if data["lat"] else None,
                "lon": sum(data["lon"]) / len(data["lon"]) if data["lon"] else None,
            }
        )

    return metrics


def _unique_sorted(records, column, limit=120):
    values = sorted({record.get(column) for record in records if record.get(column)})
    return values[:limit]


def _map_layer_points(records, location_type):
    locations = {}

    for record in records:
        city = record.get("CICADE") or record.get("CIDADE") or ""
        if location_type == "fortaleza" and city.strip().lower() != "fortaleza":
            continue

        crime_key = _map_crime_key(record.get("CRIME"))
        if crime_key is None:
            continue

        if location_type == "fortaleza":
            location = record.get("BAIRRO") or "Não informado"
            layer_label = "Bairro"
        else:
            location = city or "Não informado"
            layer_label = "Município"

        lat = _to_float(record.get("LAT"))
        lon = _to_float(record.get("LONG"))
        if lat is None or lon is None:
            continue

        if location not in locations:
            locations[location] = {"lat": [], "lon": [], "crimes": Counter(), "layer_label": layer_label}

        locations[location]["crimes"][crime_key] += 1
        locations[location]["lat"].append(lat)
        locations[location]["lon"].append(lon)

    points = []
    for location, data in locations.items():
        lat = sum(data["lat"]) / len(data["lat"])
        lon = sum(data["lon"]) / len(data["lon"])
        total = sum(data["crimes"].values())
        for crime_key, count in data["crimes"].most_common():
            category = MAP_CRIME_CATEGORIES[crime_key]
            points.append(
                {
                    "name": location,
                    "type": data["layer_label"],
                    "crime": crime_key,
                    "crime_label": category["label"],
                    "color": category["color"],
                    "count": count,
                    "total": total,
                    "lat": round(lat, 6),
                    "lon": round(lon, 6),
                }
            )
    return sorted(points, key=lambda point: (point["name"], -point["count"], point["crime_label"]))


def _fortaleza_neighborhood_table(records):
    rows = {}
    ais_by_neighborhood = defaultdict(Counter)
    crime_by_neighborhood = defaultdict(Counter)

    for record in records:
        city = record.get("CICADE") or record.get("CIDADE") or ""
        if city.strip().lower() != "fortaleza":
            continue

        neighborhood = record.get("BAIRRO") or "Não informado"
        ais = record.get("AIS") or "-"
        crime = record.get("CRIME") or "Não informado"

        rows.setdefault(neighborhood, {"bairro": neighborhood, "total": 0, "cvp": 0, "maria": 0})
        rows[neighborhood]["total"] += 1
        ais_by_neighborhood[neighborhood][ais] += 1
        crime_by_neighborhood[neighborhood][crime] += 1

        if crime == "CVP":
            rows[neighborhood]["cvp"] += 1
        if crime == "MARIA DA PENHA":
            rows[neighborhood]["maria"] += 1

    table = []
    for index, row in enumerate(sorted(rows.values(), key=lambda item: item["total"], reverse=True), start=1):
        row["ranking"] = f"{index}º"
        row["ais"] = ais_by_neighborhood[row["bairro"]].most_common(1)[0][0]
        row["top_crime"] = crime_by_neighborhood[row["bairro"]].most_common(1)[0][0]
        table.append(row)
    return table[:14]


def _crime_tree(records):
    root_total = 0
    crime_counts = Counter()
    neighborhoods_by_crime = defaultdict(Counter)

    for record in records:
        city = record.get("CICADE") or record.get("CIDADE") or ""
        if city.strip().lower() != "fortaleza":
            continue

        crime = record.get("CRIME") or "Não informado"
        neighborhood = record.get("BAIRRO") or "Não informado"
        root_total += 1
        crime_counts[crime] += 1
        neighborhoods_by_crime[crime][neighborhood] += 1

    children = []
    max_crime = max(crime_counts.values()) if crime_counts else 1
    for crime, count in crime_counts.most_common(8):
        neighborhoods = [
            {
                "name": name,
                "count": value,
                "width": round(value / count * 100, 1) if count else 0,
            }
            for name, value in neighborhoods_by_crime[crime].most_common(10)
        ]
        children.append(
            {
                "name": crime,
                "count": count,
                "color": _crime_color(crime),
                "width": round(count / max_crime * 100, 1),
                "neighborhoods": neighborhoods,
            }
        )

    return {"root_label": "Bairros e suas maiores ocorrências", "root_total": root_total, "children": children}


@lru_cache(maxsize=1)
def get_dashboard_data():
    geojson = _load_ceara_geojson()
    raw_records = _read_xlsx_sheet(_workbook_path(), MAPPING_SHEET_NAME)
    teaching_records = _read_xlsx_sheet(_workbook_path(), TEACHING_SHEET_NAME)
    teaching_details = _mapping_teaching_details(teaching_records)
    records = []

    for index, record in enumerate(raw_records, start=1):
        raw_city = _clean_mapping_label(record, "CIDADE")
        post_grad = _clean_mapping_label(record, "POST/GRAD")
        function = _clean_mapping_label(record, "FUNÇÃO")
        role_group = _mapping_role_group(function)
        seniority = _first_mapping_label(record, "ANTIGUIDADE", "ORD ANT.")
        registration = _clean_mapping_label(record, "MATRICULA")
        teaching = teaching_details.get(
            _normalize_key(registration),
            {
                "school_names": [],
                "schools_total": 0,
                "classes_total": 0,
                "students_total": 0,
            },
        )
        records.append(
            {
                "id": index,
                "seniority": seniority,
                "seniority_sort": _to_int(seniority),
                "region": _clean_mapping_label(record, "REGIÃO_ATIVIDADE"),
                "battalion": _clean_mapping_label(record, "BATALHÃO"),
                "lotacao": _clean_mapping_label(record, "LOTAÇÃO"),
                "raw_city": raw_city,
                "cities": _mapping_cities(raw_city),
                "name_upper": _clean_mapping_label(record, "NOME_MAIUSCULO"),
                "numeral": _clean_mapping_label(record, "NUMERAL"),
                "post_grad": post_grad,
                "matricula": registration,
                "name": _clean_mapping_label(record, "NOME"),
                "function": function,
                "role_group": role_group,
                "post_role": f"{post_grad} · {role_group}",
                **teaching,
            }
        )

    mapped_rows, unmatched_rows = _mapping_city_rows(records, geojson)
    total = len(records)
    mapped_assignments = sum(row["registros"] for row in mapped_rows)
    municipalities_total = len(geojson.get("features", []))
    city_rows = sorted(mapped_rows, key=lambda row: (-row["registros"], row["municipio_mapeado"]))
    region_counts = _mapping_option_counts(records, "region")
    battalion_counts = _mapping_option_counts(records, "battalion")
    function_counts = _mapping_option_counts(records, "function")
    role_group_counts = _mapping_option_counts(records, "role_group")
    post_grad_counts = _mapping_option_counts(records, "post_grad")
    post_role_counts = _mapping_option_counts(records, "post_role")
    top_city = city_rows[0] if city_rows else {"municipio_mapeado": "-", "registros": 0}
    top_region = region_counts[0] if region_counts else {"label": "-", "count": 0}
    top_post_grad = post_grad_counts[0] if post_grad_counts else {"label": "-", "count": 0}
    instructor_total = sum(1 for record in records if record["role_group"] == "Instrutor")
    monitor_total = sum(1 for record in records if record["role_group"] == "Monitor")
    coordinator_total = sum(1 for record in records if record["role_group"] == "Coordenador")
    instructors_with_classes = [
        record
        for record in records
        if record["role_group"] == "Instrutor" and record["classes_total"] > 0
    ]
    top_class_instructor = max(
        instructors_with_classes,
        key=lambda record: (record["classes_total"], record["schools_total"], -record["seniority_sort"]),
        default=None,
    )
    top_5_total = sum(row["registros"] for row in city_rows[:5])
    hhi_regional = sum((item["count"] / total) ** 2 for item in region_counts) if total else 0
    period = "Mapeamento por cidade, região de atividade, função e POST/GRAD"

    return {
        "source": f"Análise.xlsx / aba {MAPPING_SHEET_NAME}",
        "period": period,
        "geojson": geojson,
        "records": records,
        "mapped_rows": city_rows,
        "unmatched_rows": unmatched_rows,
        "filters": {
            "regions": region_counts,
            "battalions": battalion_counts,
            "functions": function_counts,
            "role_groups": role_group_counts,
            "post_grads": post_grad_counts,
            "post_roles": post_role_counts,
            "cities": [
                {"label": row["municipio_mapeado"], "count": row["registros"]}
                for row in city_rows
            ],
        },
        "summary": {
            "total_records": total,
            "mapped_assignments": mapped_assignments,
            "unmatched_rows": len(unmatched_rows),
            "municipalities_total": municipalities_total,
            "municipalities_with_data": len(city_rows),
            "regions_total": len(region_counts),
            "battalions_total": len(battalion_counts),
            "functions_total": len(function_counts),
            "post_grads_total": len(post_grad_counts),
            "instructors_total": instructor_total,
            "monitors_total": monitor_total,
            "coordinators_total": coordinator_total,
            "schools_with_instructor_data": len(
                {
                    school
                    for record in records
                    for school in record["school_names"]
                }
            ),
            "max_records": max((row["registros"] for row in city_rows), default=0),
        },
        "top_class_instructor": {
            "name": top_class_instructor["name"] if top_class_instructor else "-",
            "classes_total": top_class_instructor["classes_total"] if top_class_instructor else 0,
            "schools_total": top_class_instructor["schools_total"] if top_class_instructor else 0,
        },
        "kpis": [
            {
                "label": "Região por cidade",
                "value": _format_number(len(city_rows)),
                "trend": f"{_format_number(len(region_counts))} região(ões) de atividade",
            },
            {
                "label": "Instrutores",
                "value": _format_number(instructor_total),
                "trend": "Filtrado por cidade, região e POST/GRAD",
            },
            {
                "label": "Monitores",
                "value": _format_number(monitor_total),
                "trend": f"{_format_number(coordinator_total)} coordenador(es)",
            },
            {
                "label": "POST/GRAD predominante",
                "value": top_post_grad["label"],
                "trend": f"{_format_number(top_post_grad['count'])} registro(s)",
            },
        ],
        "scientific_metrics": [
            {
                "label": "Média por município",
                "value": _format_number(mapped_assignments / len(city_rows), 1) if city_rows else "0",
                "note": "Registros territorializados por município mapeado",
            },
            {
                "label": "Cobertura territorial",
                "value": f"{len(city_rows) / municipalities_total * 100:.1f}%".replace(".", ",") if municipalities_total else "0%",
                "note": f"{len(city_rows)} de {municipalities_total} municípios da malha oficial",
            },
            {
                "label": "Concentração territorial",
                "value": f"{top_5_total / mapped_assignments * 100:.1f}%".replace(".", ",") if mapped_assignments else "0%",
                "note": "Participação dos 5 municípios com mais registros",
            },
            {
                "label": "HHI regional",
                "value": f"{hhi_regional:.3f}".replace(".", ","),
                "note": "Índice de concentração por REGIÃO_ATIVIDADE",
            },
        ],
        "charts": {
            "regions": region_counts[:12],
            "battalions": battalion_counts[:12],
            "functions": function_counts[:12],
            "role_groups": role_group_counts[:12],
            "post_grads": post_grad_counts[:12],
            "post_roles": post_role_counts[:12],
            "cities": [
                {"label": row["municipio_mapeado"], "count": row["registros"]}
                for row in city_rows[:12]
            ],
        },
        "top_region": top_region,
    }


@app.route("/")
def home():
    map_page = get_territorial_data()
    scientific_page = get_dashboard_data()
    return render_template("index.html", map_page=map_page, scientific_page=scientific_page)


@app.route("/dados/ceara-municipios.kml")
def ceara_municipalities_kml():
    kml_path = export_school_kml()
    return send_file(kml_path, mimetype="application/vnd.google-earth.kml+xml", as_attachment=True)


@app.route("/dados/ceara-municipios.csv")
def ceara_municipalities_csv():
    csv_path = export_school_csv()
    return send_file(csv_path, mimetype="text/csv", as_attachment=True)


@app.route("/dados/ceara-municipios.pdf")
def ceara_municipalities_pdf():
    pdf_path = export_school_pdf()
    return send_file(pdf_path, mimetype="application/pdf", as_attachment=True)


@app.route("/api/dashboard")
def dashboard_api():
    return jsonify(get_territorial_data())


@app.route("/api/analise-cientifica")
def scientific_analysis_api():
    return jsonify(get_dashboard_data())


if __name__ == "__main__":
    preferred_port = int(os.environ.get("PORT", 5010))
    port = _available_port(preferred_port)
    print(f"Dashboard disponível em http://127.0.0.1:{port}")
    app.run(debug=True, port=port, use_reloader=False)
