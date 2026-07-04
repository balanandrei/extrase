#!/usr/bin/env python3
"""Extrage XML-ul atasat din PDF-urile de extras de cont Trezorerie si produce,
pentru fiecare folder de input, un XLSX + un DBF + un XML concatenat.

Rulare:
    python3 extract_extrase.py [config.toml]

Daca nu se da niciun argument, se foloseste `config.toml` din folderul scriptului.
Nu are dependinte externe (isi scrie singur XLSX-ul si DBF-ul).
"""
from __future__ import annotations

import argparse
import re
import sys
import tomllib
import zipfile
import zlib
from datetime import date
from decimal import Decimal, InvalidOperation
from collections import OrderedDict
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape as xml_text_escape


DEFAULT_GLOB = "TREZ321_ExtrasEP_PDFCLI_*_XML_SIGNED_*.pdf"
DEFAULT_DBF_ENCODING = "cp1250"
CUI_RE = re.compile(r"_PDFCLI_(\d+)_XML_")


# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

class Config:
    def __init__(self, input_folders: list[Path], output_folder: Path,
                 glob_pattern: str, dbf_encoding: str, cui: str) -> None:
        self.input_folders = input_folders
        self.output_folder = output_folder
        self.glob_pattern = glob_pattern
        self.dbf_encoding = dbf_encoding
        self.cui = cui  # override; "" => auto-detect din numele fisierului


def load_config(path: Path) -> Config:
    if not path.is_file():
        raise SystemExit(f"Nu am gasit fisierul de config: {path}")

    with path.open("rb") as fh:
        data = tomllib.load(fh)

    raw_inputs = data.get("input_folders")
    if not raw_inputs or not isinstance(raw_inputs, list):
        raise SystemExit("Config invalid: 'input_folders' trebuie sa fie o lista nevida.")

    raw_output = data.get("output_folder")
    if not raw_output or not isinstance(raw_output, str):
        raise SystemExit("Config invalid: 'output_folder' trebuie setat (string).")

    input_folders = [Path(p).expanduser() for p in raw_inputs]
    output_folder = Path(raw_output).expanduser()

    return Config(
        input_folders=input_folders,
        output_folder=output_folder,
        glob_pattern=data.get("glob_pattern", DEFAULT_GLOB),
        dbf_encoding=data.get("dbf_encoding", DEFAULT_DBF_ENCODING),
        cui=str(data.get("cui", "")).strip(),
    )


def detect_cui(filename: str) -> str:
    match = CUI_RE.search(filename)
    return match.group(1) if match else ""


def sanitize_name(name: str) -> str:
    cleaned = re.sub(r"[^\w.-]+", "_", name.strip())
    return cleaned.strip("_") or "output"


# ---------------------------------------------------------------------------
# Parsare XML
# ---------------------------------------------------------------------------

def extract_embedded_xml(pdf_path: Path) -> bytes:
    data = pdf_path.read_bytes()
    pattern = re.compile(
        rb"(?P<num>\d+)\s+0\s+obj\s*(?P<dict><<.*?/Type\s*/EmbeddedFile.*?>>)\s*stream\r?\n(?P<body>.*?)\r?\nendstream",
        re.S,
    )
    matches = pattern.findall(data)
    if not matches:
        raise ValueError("nu am gasit /Type /EmbeddedFile")
    if len(matches) > 1:
        print(f"    ATENTIE: {pdf_path.name} contine {len(matches)} atasamente; "
              f"il folosesc doar pe primul.", file=sys.stderr)

    _num, dictionary, body = matches[0]
    if b"/FlateDecode" in dictionary:
        body = zlib.decompress(body)

    if not body.lstrip().startswith(b"<"):
        raise ValueError("atasamentul extras nu pare XML")
    return body


def text_of(parent: ET.Element, path: str) -> str:
    value = parent.findtext(path)
    return (value or "").strip()


def attrs_of(parent: ET.Element, path: str) -> dict[str, str]:
    child = parent.find(path)
    return dict(child.attrib) if child is not None else {}


def money_value(value: str) -> str:
    """Valorile din XML sunt sumele reale in lei (ex. "17226" = 17226.00,
    "530641.5" = 530641.50). Intoarce mereu un string cu 2 zecimale
    (".00" unde nu exista fractie) sau "" daca inputul e gol/invalid."""
    text = (value or "").strip()
    if not text:
        return ""
    normalized = text.replace(",", ".")
    try:
        amount = Decimal(normalized)
    except InvalidOperation:
        return ""
    return f"{amount:.2f}"


def budget_parts(account_code: str, fiscal_code: str) -> OrderedDict[str, str]:
    parts: OrderedDict[str, str] = OrderedDict()
    code = (account_code or "").strip()
    budget_code = code
    if fiscal_code and code.endswith(fiscal_code):
        budget_code = code[: -len(fiscal_code)]

    parts["cont_cod_complet"] = code
    parts["cont_cod_bugetar"] = budget_code
    parts["cont_cui"] = fiscal_code if fiscal_code and code.endswith(fiscal_code) else ""

    match = re.match(r"^(?P<prefix>\d{2}A)(?P<body>\d+)$", budget_code)
    if not match:
        parts["cont_prefix"] = ""
        parts["indicator_functional"] = ""
        parts["articol_economic"] = ""
        return parts

    prefix = match.group("prefix")
    body = match.group("body")
    parts["cont_prefix"] = prefix

    if len(body) >= 6:
        functional = body[:6]
        parts["indicator_functional"] = f"{functional[:2]}.{functional[2:4]}.{functional[4:6]}"
    else:
        parts["indicator_functional"] = body

    if len(body) > 6:
        economic = body[6:]
        if len(economic) >= 6:
            parts["articol_economic"] = f"{economic[:2]}.{economic[2:4]}.{economic[4:6]}"
        else:
            parts["articol_economic"] = economic
    else:
        parts["articol_economic"] = ""

    return parts


def operation_rows(root: ET.Element, source_pdf: str, fiscal_code: str) -> list[OrderedDict[str, str]]:
    rows: list[OrderedDict[str, str]] = []

    for cont_ext_index, cont_ext in enumerate(root.findall("cont_ext"), start=1):
        cont_value = text_of(cont_ext, "cont")
        cont_attrs = attrs_of(cont_ext, "cont")
        common: OrderedDict[str, str] = OrderedDict()
        common["source_pdf"] = source_pdf
        common["Data_extras"] = root.attrib.get("Data_extras", "")
        common["Denumire_EP"] = root.attrib.get("Denumire_EP", "")
        common["Nr_ext"] = root.attrib.get("Nr_ext", "")
        common["cont_ext_index"] = str(cont_ext_index)
        common.update(budget_parts(cont_value, fiscal_code))
        common["cont_iban"] = cont_attrs.get("Cod_IBAN", "")
        common["cont_nr_op"] = cont_attrs.get("Nr_op", "")
        common["Sold_precedent_debit"] = money_value(text_of(cont_ext, "Sold_precedent/Sumad"))
        common["Sold_precedent_credit"] = money_value(text_of(cont_ext, "Sold_precedent/Sumac"))
        common["Rulaj_zi_debit"] = money_value(text_of(cont_ext, "Rulaj_zi/Sumad"))
        common["Rulaj_zi_credit"] = money_value(text_of(cont_ext, "Rulaj_zi/Sumac"))
        common["Total_sume_debit"] = money_value(text_of(cont_ext, "Total_sume/Sumad"))
        common["Total_sume_credit"] = money_value(text_of(cont_ext, "Total_sume/Sumac"))
        common["Sold_final_debit"] = money_value(text_of(cont_ext, "Sold_final/Sumad"))
        common["Sold_final_credit"] = money_value(text_of(cont_ext, "Sold_final/Sumac"))

        misc_items = cont_ext.findall("cont_misc")
        if not misc_items:
            rows.append(common)
            continue

        for misc_index, misc in enumerate(misc_items, start=1):
            row = OrderedDict(common)
            row["cont_misc_index"] = str(misc_index)
            row["nrdoc"] = text_of(misc, "nrdoc")
            row["datadoc"] = text_of(misc, "datadoc")
            row["databan"] = text_of(misc, "databan")
            row["nrrefdest"] = text_of(misc, "nrrefdest")
            row["ibanbfpl"] = text_of(misc, "ibanbfpl")
            row["platitor"] = text_of(misc, "platitor")
            row["numepb"] = text_of(misc, "numepb")
            row["suma_debit"] = money_value(text_of(misc, "sumad"))
            row["suma_credit"] = money_value(text_of(misc, "sumac"))
            row["codcontract"] = text_of(misc, "codcontract")
            row["randcontract"] = text_of(misc, "randcontract")
            row["codprogram"] = text_of(misc, "codprogram")
            row["explicatii"] = text_of(misc, "explicatii")
            rows.append(row)

    return rows


# ---------------------------------------------------------------------------
# Scriere XLSX (numeric pentru coloanele de sume/indexuri)
# ---------------------------------------------------------------------------

# Coloanele de sume: numerice cu format "0.00" (afiseaza mereu 2 zecimale).
MONEY_COLUMNS = {
    "Sold_precedent_debit",
    "Sold_precedent_credit",
    "Rulaj_zi_debit",
    "Rulaj_zi_credit",
    "Total_sume_debit",
    "Total_sume_credit",
    "Sold_final_debit",
    "Sold_final_credit",
    "suma_debit",
    "suma_credit",
}

# Indexuri: numerice intregi.
INDEX_COLUMNS = {"cont_ext_index", "cont_misc_index"}

NUMERIC_COLUMNS = MONEY_COLUMNS | INDEX_COLUMNS


def xml_escape(value: object) -> str:
    return xml_text_escape("" if value is None else str(value))


def column_name(index: int) -> str:
    name = ""
    while index:
        index, rem = divmod(index - 1, 26)
        name = chr(65 + rem) + name
    return name


def _is_number(text: str) -> bool:
    if not text:
        return False
    try:
        float(text)
        return True
    except ValueError:
        return False


def write_xlsx(path: Path, headers: list[str], rows: list[dict[str, str]], sheet_name: str) -> None:
    numeric_indexes = {i for i, h in enumerate(headers, start=1) if h in NUMERIC_COLUMNS}
    money_indexes = {i for i, h in enumerate(headers, start=1) if h in MONEY_COLUMNS}
    sheet_rows: list[str] = []

    def make_cell(idx: int, value: str, row_number: int) -> str:
        ref = f"{column_name(idx)}{row_number}"
        if idx in numeric_indexes and _is_number(value):
            # style s="1" = format "0.00" (mereu 2 zecimale) pentru sume
            style = ' s="1"' if idx in money_indexes else ""
            return f'<c r="{ref}"{style}><v>{value}</v></c>'
        if value == "":
            return f'<c r="{ref}"/>'
        return f'<c r="{ref}" t="inlineStr"><is><t xml:space="preserve">{xml_escape(value)}</t></is></c>'

    header_cells = "".join(
        f'<c r="{column_name(idx)}1" t="inlineStr"><is><t>{xml_escape(h)}</t></is></c>'
        for idx, h in enumerate(headers, start=1)
    )
    sheet_rows.append(f'<row r="1">{header_cells}</row>')

    for row_number, row in enumerate(rows, start=2):
        cells = "".join(
            make_cell(idx, row.get(header, ""), row_number)
            for idx, header in enumerate(headers, start=1)
        )
        sheet_rows.append(f'<row r="{row_number}">{cells}</row>')

    sheet_ref = f"A1:{column_name(max(len(headers), 1))}{max(len(rows) + 1, 1)}"
    sheet_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="{sheet_ref}"/>
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData>{"".join(sheet_rows)}</sheetData>
</worksheet>'''

    safe_sheet = xml_escape(sheet_name[:31]) or "Extrase"
    workbook_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="{safe_sheet}" sheetId="1" r:id="rId1"/></sheets>
</workbook>'''

    rels_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>'''

    workbook_rels_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''

    # cellXfs[0] = format general (implicit); cellXfs[1] = numFmtId 2 = "0.00".
    styles_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''

    content_types_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>'''

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as xlsx:
        xlsx.writestr("[Content_Types].xml", content_types_xml)
        xlsx.writestr("_rels/.rels", rels_xml)
        xlsx.writestr("xl/workbook.xml", workbook_xml)
        xlsx.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml)
        xlsx.writestr("xl/styles.xml", styles_xml)
        xlsx.writestr("xl/worksheets/sheet1.xml", sheet_xml)


# ---------------------------------------------------------------------------
# Scriere DBF (sumele = campuri numerice N cu 2 zecimale)
# ---------------------------------------------------------------------------

DBF_FIELDS = [
    ("source_pdf", "SRC_PDF", "C"),
    ("Data_extras", "DATA_EXT", "C"),
    ("Denumire_EP", "DEN_EP", "C"),
    ("Nr_ext", "NR_EXT", "C"),
    ("cont_ext_index", "CONT_IDX", "N"),
    ("cont_cod_complet", "CONT_FULL", "C"),
    ("cont_cod_bugetar", "CONT_BUG", "C"),
    ("cont_cui", "CONT_CUI", "C"),
    ("cont_prefix", "PREFIX", "C"),
    ("indicator_functional", "IND_FUNC", "C"),
    ("articol_economic", "ART_ECON", "C"),
    ("cont_iban", "IBAN", "C"),
    ("cont_nr_op", "NR_OP", "C"),
    ("Sold_precedent_debit", "SOLDPD", "N"),
    ("Sold_precedent_credit", "SOLDPC", "N"),
    ("Rulaj_zi_debit", "RUL_D", "N"),
    ("Rulaj_zi_credit", "RUL_C", "N"),
    ("Total_sume_debit", "TOTAL_D", "N"),
    ("Total_sume_credit", "TOTAL_C", "N"),
    ("Sold_final_debit", "SOLDF_D", "N"),
    ("Sold_final_credit", "SOLDF_C", "N"),
    ("cont_misc_index", "MISC_IDX", "N"),
    ("nrdoc", "NRDOC", "C"),
    ("datadoc", "DATADOC", "D"),
    ("databan", "DATABAN", "D"),
    ("nrrefdest", "NRREFDEST", "C"),
    ("ibanbfpl", "IBANBFPL", "C"),
    ("platitor", "PLATITOR", "C"),
    ("numepb", "NUMEPB", "C"),
    ("suma_debit", "SUMA_D", "N"),
    ("suma_credit", "SUMA_C", "N"),
    ("codcontract", "CODCONTR", "C"),
    ("randcontract", "RANDCONTR", "C"),
    ("codprogram", "CODPROG", "C"),
    ("explicatii", "EXPLIC", "C"),
]

# Campurile de bani: numerice cu 2 zecimale.
DBF_MONEY_FIELDS = {
    "Sold_precedent_debit",
    "Sold_precedent_credit",
    "Rulaj_zi_debit",
    "Rulaj_zi_credit",
    "Total_sume_debit",
    "Total_sume_credit",
    "Sold_final_debit",
    "Sold_final_credit",
    "suma_debit",
    "suma_credit",
}


def dbf_text(value: object, length: int, encoding: str) -> bytes:
    encoded = str(value or "").encode(encoding, errors="replace")[:length]
    return encoded.ljust(length, b" ")


def dbf_number(value: object, length: int, decimals: int, field_name: str) -> bytes:
    text = str(value or "").strip()
    if not text:
        return b" " * length
    try:
        number = Decimal(text.replace(",", "."))
        rendered = f"{number:.{decimals}f}" if decimals else str(int(number))
    except (InvalidOperation, ValueError):
        print(f"    ATENTIE: valoare numerica invalida in {field_name!r}: {text!r}", file=sys.stderr)
        return b" " * length
    if len(rendered) > length:
        # Nu trunchiem tacut un numar; l-am corupe. Semnalam si lasam gol.
        print(f"    ATENTIE: valoarea {rendered!r} nu incape in {field_name!r} "
              f"(max {length}); las camp gol.", file=sys.stderr)
        return b" " * length
    return rendered.rjust(length, " ").encode("ascii", errors="replace")


def dbf_date(value: object) -> bytes:
    text = str(value or "").strip()
    if re.fullmatch(r"\d{8}", text):
        return text.encode("ascii")
    return b" " * 8


def write_dbf(path: Path, rows: list[dict[str, str]], encoding: str) -> None:
    field_specs = []
    for source_key, dbf_name, field_type in DBF_FIELDS:
        if field_type == "D":
            length, decimals = 8, 0
        elif field_type == "N":
            if source_key in DBF_MONEY_FIELDS:
                length, decimals = 18, 2
            else:
                length, decimals = 10, 0
        else:
            max_len = max(
                [len(str(row.get(source_key, "")).encode(encoding, errors="replace")) for row in rows]
                + [1]
            )
            length = min(max(max_len, len(dbf_name)), 254)
            decimals = 0
        field_specs.append((source_key, dbf_name, field_type, length, decimals))

    today = date.today()
    record_count = len(rows)
    header_length = 32 + 32 * len(field_specs) + 1
    record_length = 1 + sum(field[3] for field in field_specs)

    header = bytearray()
    header.append(0x03)
    header.extend(bytes([today.year - 1900, today.month, today.day]))
    header.extend(record_count.to_bytes(4, "little"))
    header.extend(header_length.to_bytes(2, "little"))
    header.extend(record_length.to_bytes(2, "little"))
    header.extend(b"\x00" * 17)
    header.append(0xC8)
    header.extend(b"\x00" * 2)

    for _, dbf_name, field_type, length, decimals in field_specs:
        name = dbf_name.encode("ascii")[:10]
        header.extend(name + b"\x00" * (11 - len(name)))
        header.extend(field_type.encode("ascii"))
        header.extend(b"\x00" * 4)
        header.append(length)
        header.append(decimals)
        header.extend(b"\x00" * 14)
    header.append(0x0D)

    with path.open("wb") as dbf:
        dbf.write(header)
        for row in rows:
            record = bytearray(b" ")
            for source_key, dbf_name, field_type, length, decimals in field_specs:
                value = row.get(source_key, "")
                if field_type == "N":
                    record.extend(dbf_number(value, length, decimals, dbf_name))
                elif field_type == "D":
                    record.extend(dbf_date(value))
                else:
                    record.extend(dbf_text(value, length, encoding))
            dbf.write(record)
        dbf.write(b"\x1A")


# ---------------------------------------------------------------------------
# Procesare per folder
# ---------------------------------------------------------------------------

def process_folder(input_folder: Path, output_base: Path, config: Config) -> bool:
    print(f"\n=== Folder: {input_folder} ===")
    if not input_folder.is_dir():
        print(f"  EROARE: folderul nu exista, il sar.", file=sys.stderr)
        return False

    pdfs = sorted(input_folder.glob(config.glob_pattern))
    if not pdfs:
        print(f"  Nu am gasit PDF-uri ({config.glob_pattern}), il sar.", file=sys.stderr)
        return False

    out_dir = output_base / sanitize_name(input_folder.name)
    xml_dir = out_dir / "xml_extrase"
    xml_dir.mkdir(parents=True, exist_ok=True)

    merged_root = ET.Element("extrase_concat")
    rows: list[dict[str, str]] = []
    headers: OrderedDict[str, None] = OrderedDict()
    failed: list[tuple[str, str]] = []

    for pdf_path in pdfs:
        try:
            xml_bytes = extract_embedded_xml(pdf_path)
            root = ET.fromstring(xml_bytes)
        except Exception as exc:  # noqa: BLE001 - vrem sa continuam batch-ul
            failed.append((pdf_path.name, str(exc)))
            print(f"  ESUAT: {pdf_path.name}: {exc}", file=sys.stderr)
            continue

        xml_path = xml_dir / f"{pdf_path.stem}.xml"
        xml_path.write_bytes(xml_bytes)

        fiscal_code = config.cui or detect_cui(pdf_path.name)
        wrapper = ET.SubElement(merged_root, "document", source_pdf=pdf_path.name)
        wrapper.append(root)

        for row in operation_rows(root, pdf_path.name, fiscal_code):
            rows.append(row)
            for header in row:
                headers.setdefault(header, None)

    if not rows:
        print("  Niciun rand extras (toate fisierele au esuat?).", file=sys.stderr)

    base = sanitize_name(input_folder.name)
    merged_tree = ET.ElementTree(merged_root)
    ET.indent(merged_tree, space="  ")
    merged_xml_path = out_dir / f"{base}_concat.xml"
    merged_tree.write(merged_xml_path, encoding="utf-8", xml_declaration=True)

    xlsx_path = out_dir / f"{base}.xlsx"
    write_xlsx(xlsx_path, list(headers.keys()), rows, sheet_name=base)

    dbf_path = out_dir / f"{base}.dbf"
    write_dbf(dbf_path, rows, config.dbf_encoding)

    print(f"  PDF-uri gasite: {len(pdfs)} | procesate: {len(pdfs) - len(failed)} | esuate: {len(failed)}")
    print(f"  Randuri: {len(rows)}")
    print(f"  Excel:   {xlsx_path}")
    print(f"  DBF:     {dbf_path}")
    print(f"  XML:     {merged_xml_path}")
    if failed:
        print(f"  Fisiere esuate:")
        for name, reason in failed:
            print(f"    - {name}: {reason}")

    return len(failed) == 0


def main() -> int:
    parser = argparse.ArgumentParser(description="Extrage extrase de cont din PDF -> XLSX + DBF + XML.")
    parser.add_argument(
        "config",
        nargs="?",
        default=str(Path(__file__).resolve().parent / "config.toml"),
        help="Calea catre fisierul de config TOML (default: config.toml langa script).",
    )
    args = parser.parse_args()

    config = load_config(Path(args.config).expanduser())
    config.output_folder.mkdir(parents=True, exist_ok=True)

    all_ok = True
    for input_folder in config.input_folders:
        ok = process_folder(input_folder, config.output_folder, config)
        all_ok = all_ok and ok

    print(f"\nGata. Output in: {config.output_folder}")
    return 0 if all_ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
