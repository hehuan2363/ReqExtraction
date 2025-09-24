#!/usr/bin/env python3
"""Extract numbered clauses from a standards PDF into JSON and Excel outputs."""

from __future__ import annotations

import argparse
import json
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import BinaryIO, Dict, Iterable, List, Optional, Tuple, Union
from xml.sax.saxutils import escape
import zipfile

from pdfminer.high_level import extract_pages
from pdfminer.layout import LAParams, LTChar, LTTextContainer, LTTextLine
from pdfminer.pdfdocument import PDFTextExtractionNotAllowed
from pdfminer.pdfparser import PDFSyntaxError


@dataclass
class TextChunk:
    page: int
    top: float
    left: float
    width: float
    text: str
    font_size: float
    is_bold: bool


@dataclass
class Line:
    page: int
    top: float
    chunks: List[TextChunk] = field(default_factory=list)

    def add_chunk(self, chunk: TextChunk) -> None:
        self.chunks.append(chunk)

    def sort_chunks(self) -> None:
        self.chunks.sort(key=lambda c: c.left)

    def text(self) -> str:
        self.sort_chunks()
        parts: List[str] = []
        last_right: Optional[float] = None
        for chunk in self.chunks:
            if not chunk.text:
                continue
            if last_right is not None:
                gap = chunk.left - last_right
                if gap > 1.5:
                    parts.append(" ")
            parts.append(chunk.text)
            last_right = chunk.left + chunk.width
        return "".join(parts)

    def cleaned_text(self) -> str:
        raw = self.text()
        return " ".join(raw.split())

    def max_font_size(self) -> float:
        return max((chunk.font_size for chunk in self.chunks), default=0.0)

    def bold_ratio(self) -> float:
        total = sum(len(chunk.text.strip()) for chunk in self.chunks if chunk.text.strip())
        if total == 0:
            return 0.0
        bold = sum(len(chunk.text.strip()) for chunk in self.chunks if chunk.is_bold and chunk.text.strip())
        return bold / total


@dataclass
class Heading:
    identifier: str
    title: str
    line_index: int
    line_count: int


@dataclass
class Clause:
    identifier: str
    title: str
    body_lines: List[str] = field(default_factory=list)
    children: List["Clause"] = field(default_factory=list)

    def add_line(self, line: str) -> None:
        self.body_lines.append(line)

    def text(self) -> str:
        paragraphs: List[str] = []
        buffer: List[str] = []
        for line in self.body_lines:
            if not line:
                if buffer:
                    paragraphs.append(" ".join(buffer).strip())
                    buffer = []
                continue
            if buffer and buffer[-1].endswith('-') and line and line[0].islower():
                buffer[-1] = buffer[-1][:-1] + line
            elif buffer:
                buffer.append(line)
            else:
                buffer.append(line)
        if buffer:
            paragraphs.append(" ".join(buffer).strip())
        return "\n\n".join(p for p in paragraphs if p)

    def to_dict(self) -> Dict[str, object]:
        data: Dict[str, object] = {
            "clause": self.identifier,
            "title": self.title,
            "text": self.text(),
        }
        if self.children:
            data["subclauses"] = [child.to_dict() for child in self.children]
        return data


HEADING_RE = re.compile(r"^(\d+(?:\.\d+)*)(?:\s+(.*\S))?$")

SKIP_PATTERNS = [
    re.compile(r"^copyright british standards institution", re.IGNORECASE),
    re.compile(r"^provided by accuris", re.IGNORECASE),
    re.compile(r"^licensee=", re.IGNORECASE),
    re.compile(r"^not for resale", re.IGNORECASE),
    re.compile(r"^no reproduction or networking permitted", re.IGNORECASE),
    re.compile(r"^bs en ", re.IGNORECASE),
    re.compile(r"^iec 61513", re.IGNORECASE),
    re.compile(r"^61513", re.IGNORECASE),
    re.compile(r"^raising standards worldwide", re.IGNORECASE),
    re.compile(r"^–\s*\d+\s*–"),
    re.compile(r"^--[`',.-]{5,}"),
]


def parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Split a standards PDF into JSON and Excel clause files."
    )
    parser.add_argument(
        "pdf",
        nargs="?",
        default="Standards/Test Stanard.pdf",
        help="Path to the standards PDF (default: Standards/Test Stanard.pdf)",
    )
    parser.add_argument(
        "--output-dir",
        default="-output",
        help="Directory where outputs will be written (default: -output)",
    )
    return parser.parse_args()


def _iter_text_lines(container: LTTextContainer) -> Iterable[LTTextLine]:
    for element in container:
        if isinstance(element, LTTextLine):
            yield element
        elif isinstance(element, LTTextContainer):
            yield from _iter_text_lines(element)


def _is_bold_font(fontname: Optional[str]) -> bool:
    if not fontname:
        return False
    lowered = fontname.lower()
    return "bold" in lowered or "black" in lowered or "heavy" in lowered


def _text_line_to_line(
    text_line: LTTextLine,
    page_number: int,
    page_height: float,
) -> Optional[Line]:
    raw_text = text_line.get_text()
    if not raw_text:
        return None
    cleaned = raw_text.replace("\r", " ").replace("\n", " ").replace("\xa0", " ")
    if not cleaned.strip():
        return None
    lower = cleaned.strip().lower()
    if lower.startswith("link to page"):
        return None

    x0, y0, x1, y1 = text_line.bbox
    top = max(page_height - y1, 0.0)
    width = max(x1 - x0, 0.0)

    chars = [obj for obj in text_line if isinstance(obj, LTChar)]
    font_size = max((char.size for char in chars), default=getattr(text_line, "height", 0.0))
    total_weight = 0
    bold_weight = 0
    for char in chars:
        text = char.get_text()
        weight = len(text.strip()) or len(text)
        total_weight += weight
        if _is_bold_font(getattr(char, "fontname", "")):
            bold_weight += weight
    is_bold = total_weight > 0 and (bold_weight / total_weight) >= 0.5

    chunk = TextChunk(
        page=page_number,
        top=top,
        left=float(x0),
        width=width,
        text=cleaned.strip("\x00"),
        font_size=float(font_size),
        is_bold=is_bold,
    )
    return Line(page=page_number, top=top, chunks=[chunk])


def extract_lines_from_pdf(pdf_path: Path) -> List[Line]:
    laparams = LAParams(char_margin=1.0, line_margin=0.2, word_margin=0.1)
    lines: List[Line] = []
    for page_number, page_layout in enumerate(extract_pages(str(pdf_path), laparams=laparams), start=1):
        page_height = float(getattr(page_layout, "height", 0.0))
        for text_line in _iter_text_lines(page_layout):
            line = _text_line_to_line(text_line, page_number, page_height)
            if line is None:
                continue
            lines.append(line)
    lines.sort(key=lambda line: (line.page, line.top, line.chunks[0].left if line.chunks else 0.0))
    return lines


def should_skip_line(text: str) -> bool:
    stripped = text.strip()
    if not stripped:
        return False
    if "..." in stripped and stripped.split()[-1].isdigit():
        return True
    if "--```" in stripped:
        return True
    for pattern in SKIP_PATTERNS:
        if pattern.search(stripped):
            return True
    return False


def looks_like_fragment(line: Line, text: str) -> bool:
    if not text:
        return False
    if line.bold_ratio() > 0.0:
        return False
    if text.startswith(("•", "–", "-", "(", ")")):
        return False
    if any(ch in text for ch in ".,;:!?"):
        return False
    words = text.split()
    if len(words) <= 1:
        return False
    return len(words) <= 6


def find_headings(lines: List[Line]) -> List[Heading]:
    headings: List[Heading] = []
    i = 0
    total = len(lines)
    while i < total:
        line = lines[i]
        text = line.cleaned_text()
        if not text:
            i += 1
            continue
        match = HEADING_RE.match(text)
        if not match:
            i += 1
            continue
        if line.max_font_size() < 14 or line.bold_ratio() < 0.5:
            i += 1
            continue
        identifier = match.group(1)
        remainder = (match.group(2) or "").strip()
        title = remainder
        consumed = 1
        if not title:
            j = i + 1
            title_parts: List[str] = []
            while j < total:
                candidate = lines[j]
                candidate_text = candidate.cleaned_text()
                if not candidate_text:
                    consumed += 1
                    j += 1
                    continue
                if candidate.max_font_size() >= 14 and candidate.bold_ratio() >= 0.5:
                    if HEADING_RE.match(candidate_text):
                        break
                    title_parts.append(candidate_text)
                    consumed += 1
                    j += 1
                else:
                    break
            title = " ".join(title_parts).strip()
        if not title and "." not in identifier:
            i += consumed
            continue
        headings.append(Heading(identifier=identifier, title=title, line_index=i, line_count=consumed))
        i += consumed
    return headings


def build_clauses(lines: List[Line]) -> List[Clause]:
    headings = find_headings(lines)
    clauses: List[Clause] = []
    clause_by_id: Dict[str, Clause] = {}
    for idx, heading in enumerate(headings):
        if heading.identifier in clause_by_id:
            continue
        clause = Clause(identifier=heading.identifier, title=heading.title or "")
        clause_by_id[heading.identifier] = clause
        if "." in heading.identifier:
            parent_id = ".".join(heading.identifier.split(".")[:-1])
            parent = clause_by_id.get(parent_id)
            if parent:
                parent.children.append(clause)
            else:
                clauses.append(clause)
        else:
            clauses.append(clause)
        start = heading.line_index + heading.line_count
        end = headings[idx + 1].line_index if idx + 1 < len(headings) else len(lines)
        prev_top: Optional[float] = None
        prev_page: Optional[int] = None
        for line_index in range(start, end):
            line = lines[line_index]
            text = line.cleaned_text()
            if should_skip_line(text):
                continue
            if HEADING_RE.match(text):
                continue
            if looks_like_fragment(line, text):
                continue
            if prev_page is not None:
                if line.page != prev_page or (line.top - (prev_top or line.top)) > 18:
                    clause.add_line("")
            if text:
                clause.add_line(text)
            else:
                clause.add_line("")
            prev_top = line.top
            prev_page = line.page
    ordered_clauses = [clause for clause in clauses if "." not in clause.identifier or clause.identifier.count(".") == len(clause.identifier.split(".")) - 1]
    ordered_clauses.sort(key=lambda c: [int(part) for part in c.identifier.split(".")])
    return ordered_clauses


def flatten_clauses(clause: Clause, parent: Optional[str] = None, level: int = 1) -> List[List[str]]:
    rows: List[List[str]] = [[clause.identifier, clause.title, parent or "", str(level), clause.text()]]
    for child in clause.children:
        rows.extend(flatten_clauses(child, clause.identifier, level + 1))
    return rows


def clauses_to_rows(clauses: List[Clause]) -> List[List[str]]:
    rows: List[List[str]] = [["Clause", "Title", "Parent", "Level", "Text"]]
    for clause in clauses:
        rows.extend(flatten_clauses(clause))
    return rows


def extract_pdf_clauses(pdf_path: Union[str, Path]) -> List[Clause]:
    resolved = Path(pdf_path).expanduser().resolve()
    if not resolved.exists():
        raise FileNotFoundError(resolved)
    try:
        lines = extract_lines_from_pdf(resolved)
    except PDFTextExtractionNotAllowed as exc:
        raise PermissionError("Text extraction is not permitted for this PDF.") from exc
    except PDFSyntaxError as exc:
        raise ValueError(f"Failed to parse PDF structure: {exc}") from exc
    if not lines:
        raise ValueError("No text extracted from PDF.")
    clauses = build_clauses(lines)
    if not clauses:
        raise ValueError("No clauses were detected in the document.")
    return clauses


def extract_pdf_data(pdf_path: Union[str, Path]) -> Tuple[List[Clause], List[List[str]]]:
    clauses = extract_pdf_clauses(pdf_path)
    rows = clauses_to_rows(clauses)
    return clauses, rows


def column_letter(index: int) -> str:
    result = ""
    while True:
        index, remainder = divmod(index, 26)
        result = chr(65 + remainder) + result
        if index == 0:
            break
        index -= 1
    return result


def build_sheet_xml(rows: List[List[str]]) -> str:
    xml_lines = [
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>",
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">",
        "  <sheetData>",
    ]
    for r_index, row in enumerate(rows, start=1):
        xml_lines.append(f"    <row r=\"{r_index}\">")
        for c_index, value in enumerate(row):
            cell_ref = f"{column_letter(c_index)}{r_index}"
            if not value:
                xml_lines.append(f"      <c r=\"{cell_ref}\"/>")
                continue
            text_value = escape(value).replace("\n", "&#10;")
            xml_lines.append(
                f"      <c r=\"{cell_ref}\" t=\"inlineStr\"><is><t xml:space=\"preserve\">{text_value}</t></is></c>"
            )
        xml_lines.append("    </row>")
    xml_lines.extend(["  </sheetData>", "</worksheet>"])
    return "\n".join(xml_lines)


def write_xlsx(rows: List[List[str]], output: Union[Path, BinaryIO]) -> None:
    sheet_xml = build_sheet_xml(rows)
    workbook_xml = """<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">
  <sheets>
    <sheet name=\"Clauses\" sheetId=\"1\" r:id=\"rId1\"/>
  </sheets>
</workbook>
""".strip()
    workbook_rels = """<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>
  <Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>
</Relationships>
""".strip()
    root_rels = """<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>
</Relationships>
""".strip()
    styles_xml = """<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">
  <fonts count=\"1\"><font><name val=\"Calibri\"/><family val=\"2\"/><sz val=\"11\"/></font></fonts>
  <fills count=\"1\"><fill><patternFill patternType=\"none\"/></fill></fills>
  <borders count=\"1\"><border/></borders>
  <cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>
  <cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\"><alignment wrapText=\"1\"/></xf></cellXfs>
  <cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>
</styleSheet>
""".strip()
    content_types = """<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">
  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>
  <Default Extension=\"xml\" ContentType=\"application/xml\"/>
  <Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>
  <Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>
  <Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>
</Types>
""".strip()

    if isinstance(output, Path):
        zip_target = output
    else:
        output.seek(0)
        output.truncate(0)
        zip_target = output

    with zipfile.ZipFile(zip_target, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", root_rels)
        zf.writestr("xl/workbook.xml", workbook_xml)
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        zf.writestr("xl/styles.xml", styles_xml)


def main() -> int:
    args = parse_arguments()
    pdf_path = Path(args.pdf).expanduser().resolve()
    output_dir = Path(args.output_dir).expanduser().resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    try:
        clauses, rows = extract_pdf_data(pdf_path)
    except FileNotFoundError:
        print(f"PDF not found: {pdf_path}", file=sys.stderr)
        return 1
    except PermissionError as exc:
        print(str(exc), file=sys.stderr)
        return 1
    except ValueError as exc:
        print(str(exc), file=sys.stderr)
        return 1
    except PDFSyntaxError as exc:
        print(f"Malformed PDF: {exc}", file=sys.stderr)
        return 1

    json_path = output_dir / "clauses.json"
    with json_path.open("w", encoding="utf-8") as handle:
        json.dump([clause.to_dict() for clause in clauses], handle, indent=2)

    xlsx_path = output_dir / "clauses.xlsx"
    write_xlsx(rows, xlsx_path)

    print(f"Wrote JSON: {json_path}")
    print(f"Wrote Excel: {xlsx_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
