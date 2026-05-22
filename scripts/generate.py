from __future__ import annotations

import argparse
import copy
import re
import textwrap
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_INPUT = ROOT / "作品情報フォーム.xlsx"
DEFAULT_TEMPLATE = ROOT / "冬樟展パンフ.docx"
DEFAULT_DOCX_OUTPUT = ROOT / "output" / "パンフレット_テンプレ流し込み.docx"
DEFAULT_LIST_OUTPUT = ROOT / "output" / "作品一覧.txt"

SPREADSHEET_NS = {
    "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WORD_NS = {"w": W_NS}
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"

ET.register_namespace("w", W_NS)
ET.register_namespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
ET.register_namespace("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing")
ET.register_namespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")
ET.register_namespace("mc", MC_NS)
ET.register_namespace("w14", "http://schemas.microsoft.com/office/word/2010/wordml")
ET.register_namespace("w15", "http://schemas.microsoft.com/office/word/2012/wordml")
ET.register_namespace("w16", "http://schemas.microsoft.com/office/word/2018/wordml")
ET.register_namespace("w16cex", "http://schemas.microsoft.com/office/word/2018/wordml/cex")
ET.register_namespace("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid")
ET.register_namespace("w16du", "http://schemas.microsoft.com/office/word/2023/wordml/word16du")
ET.register_namespace("w16sdtdh", "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash")
ET.register_namespace("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex")
ET.register_namespace("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing")
ET.register_namespace("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape")
ET.register_namespace("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup")
ET.register_namespace("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture")
ET.register_namespace("a14", "http://schemas.microsoft.com/office/drawing/2010/main")
ET.register_namespace("v", "urn:schemas-microsoft-com:vml")
ET.register_namespace("o", "urn:schemas-microsoft-com:office:office")


@dataclass(frozen=True)
class Work:
    number: int
    section: str
    name: str
    kana: str
    grade: str
    kind: str
    style: str
    title: str
    orientation: str
    size: str
    location: str
    mounting: str
    text: str
    comment: str
    artport: str


def normalize_space(value: str) -> str:
    return re.sub(r"[ \t\u3000]+", " ", value.strip())


def compact_for_sort(value: str) -> str:
    return re.sub(r"[\s\u3000]+", "", value.strip())


def find_column(headers: dict[str, str], *needles: str) -> str | None:
    for col, header in headers.items():
        normalized = header.replace("\n", " ").strip()
        if all(needle in normalized for needle in needles):
            return col
    return None


def cell_value(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    value = cell.find("a:v", SPREADSHEET_NS)
    inline = cell.find("a:is", SPREADSHEET_NS)

    if cell_type == "s" and value is not None:
        return shared_strings[int(value.text or "0")]
    if cell_type == "inlineStr" and inline is not None:
        return "".join(text.text or "" for text in inline.findall(".//a:t", SPREADSHEET_NS))
    if value is not None:
        return value.text or ""
    return ""


def read_xlsx(path: Path) -> dict[str, list[dict[str, str]]]:
    with zipfile.ZipFile(path) as book:
        shared_strings: list[str] = []
        if "xl/sharedStrings.xml" in book.namelist():
            shared_xml = ET.fromstring(book.read("xl/sharedStrings.xml"))
            for item in shared_xml.findall("a:si", SPREADSHEET_NS):
                shared_strings.append("".join(text.text or "" for text in item.findall(".//a:t", SPREADSHEET_NS)))

        workbook = ET.fromstring(book.read("xl/workbook.xml"))
        rels = ET.fromstring(book.read("xl/_rels/workbook.xml.rels"))
        rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}

        sheets: dict[str, list[dict[str, str]]] = {}
        for sheet in workbook.findall(".//a:sheet", SPREADSHEET_NS):
            sheet_name = sheet.attrib["name"]
            rel_id = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
            target = "xl/" + rel_map[rel_id].lstrip("/")
            worksheet = ET.fromstring(book.read(target))

            raw_rows: list[dict[str, str]] = []
            for row in worksheet.findall(".//a:sheetData/a:row", SPREADSHEET_NS):
                values: dict[str, str] = {}
                for cell in row.findall("a:c", SPREADSHEET_NS):
                    match = re.match(r"([A-Z]+)", cell.attrib.get("r", ""))
                    if match:
                        values[match.group(1)] = cell_value(cell, shared_strings).strip()
                raw_rows.append(values)

            if not raw_rows:
                sheets[sheet_name] = []
                continue

            headers = raw_rows[0]
            records: list[dict[str, str]] = []
            for row in raw_rows[1:]:
                record = {header: row.get(col, "").strip() for col, header in headers.items() if header}
                if any(record.values()):
                    records.append(record)
            sheets[sheet_name] = records

        return sheets


def grade_sort_key(grade: str) -> tuple[int, int, str]:
    text = compact_for_sort(grade)
    if match := re.fullmatch(r"B(\d+)", text, re.IGNORECASE):
        return (0, int(match.group(1)), text)
    if match := re.fullmatch(r"M(\d+)", text, re.IGNORECASE):
        return (1, int(match.group(1)), text)
    if match := re.fullmatch(r"D(\d+)", text, re.IGNORECASE):
        return (2, int(match.group(1)), text)
    if match := re.search(r"修士(\d+|一|二|三)回生", text):
        return (1, kanji_number(match.group(1)), text)
    if match := re.search(r"博士(\d+|一|二|三)回生", text):
        return (2, kanji_number(match.group(1)), text)
    if match := re.search(r"(\d+|一|二|三|四|五|六)回生", text):
        return (0, kanji_number(match.group(1)), text)
    return (9, 99, text)


def kanji_number(value: str) -> int:
    table = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6}
    return table.get(value, int(value) if value.isdigit() else 99)


def build_works(sheets: dict[str, list[dict[str, str]]]) -> list[Work]:
    personal = sheets.get("個人", [])
    collaboration = sheets.get("合作", [])

    personal_works: list[dict[str, str]] = []
    for row in personal:
        personal_works.append(
            {
                "section": "個人作品",
                "name": value_by_header(row, "氏名"),
                "kana": value_by_header(row, "ふりがな"),
                "grade": value_by_header(row, "学年"),
                "kind": value_by_header(row, "臨書 or 創作"),
                "style": value_by_header(row, "書体"),
                "title": value_by_header(row, "作品名"),
                "orientation": value_by_header(row, "作品の向き"),
                "size": value_by_header(row, "作品サイズ"),
                "location": value_by_header(row, "展示場所"),
                "mounting": value_by_header(row, "表装形式"),
                "text": value_by_header(row, "釈文"),
                "comment": value_by_header(row, "作品コメント"),
                "artport": value_by_header(row, "作品のオンライン公開"),
            }
        )

    personal_works.sort(key=lambda item: (grade_sort_key(item["grade"]), compact_for_sort(item["kana"]), compact_for_sort(item["name"])))

    collaboration_works: list[dict[str, str]] = []
    for row in collaboration:
        collaboration_works.append(
            {
                "section": "合作",
                "name": value_by_header(row, "合作参加者全員分"),
                "kana": "",
                "grade": "合作",
                "kind": value_by_header(row, "臨書 or 創作"),
                "style": value_by_header(row, "書体"),
                "title": value_by_header(row, "作品名"),
                "orientation": value_by_header(row, "作品の向き"),
                "size": value_by_header(row, "作品サイズ"),
                "location": value_by_header(row, "展示場所"),
                "mounting": value_by_header(row, "表装形式"),
                "text": value_by_header(row, "釈文"),
                "comment": value_by_header(row, "作品コメント"),
                "artport": value_by_header(row, "作品のオンライン公開"),
            }
        )

    collaboration_works.sort(key=lambda item: compact_for_sort(item["name"]))

    works: list[Work] = []
    for number, item in enumerate([*personal_works, *collaboration_works], start=1):
        works.append(Work(number=number, **item))
    return works


def value_by_header(row: dict[str, str], starts_with: str) -> str:
    for header, value in row.items():
        if header.replace("\n", " ").strip().startswith(starts_with):
            return value.strip()
    return ""


def group_personal_by_grade(works: Iterable[Work]) -> dict[str, list[Work]]:
    groups: dict[str, list[Work]] = {}
    for work in works:
        if work.section != "個人作品":
            continue
        groups.setdefault(work.grade, []).append(work)
    return dict(sorted(groups.items(), key=lambda item: grade_sort_key(item[0])))


def write_list_text(works: list[Work], path: Path) -> None:
    lines: list[str] = ["作品一覧", "", "個人作品"]
    for grade, grade_works in group_personal_by_grade(works).items():
        lines.extend(["", grade])
        for work in grade_works:
            lines.append(f"{work.number:02d}. {normalize_space(work.name)}　{work.kind}・{work.title}")

    collaboration = [work for work in works if work.section == "合作"]
    if collaboration:
        lines.extend(["", "合作"])
        for work in collaboration:
            lines.append(f"{work.number:02d}. {one_line(work.name)}　{work.kind}・{work.title}")

    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def one_line(value: str) -> str:
    return normalize_space(value.replace("\r\n", " ").replace("\n", " "))


def write_docx(works: list[Work], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    document_xml = build_document_xml(works)

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as docx:
        docx.writestr("[Content_Types].xml", CONTENT_TYPES_XML)
        docx.writestr("_rels/.rels", ROOT_RELS_XML)
        docx.writestr("word/_rels/document.xml.rels", DOCUMENT_RELS_XML)
        docx.writestr("word/styles.xml", STYLES_XML)
        docx.writestr("word/document.xml", document_xml)


def write_template_docx(works: list[Work], template_path: Path, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(template_path) as template:
        document_root = ET.fromstring(template.read("word/document.xml"))
        body = document_root.find("w:body", WORD_NS)
        if body is None:
            raise SystemExit("テンプレートの word/document.xml に本文が見つかりません。")
        remove_ignorable_namespace_hint(document_root)

        children = list(body)
        list_start = find_child_index_by_text(children, "作品一覧")
        detail_start = find_vertical_detail_start(children, list_start)
        roster_start = find_child_index_by_text(children, "【顧問・部員紹介】", start=detail_start)
        detail_end = find_previous_section_break(children, roster_start)

        list_title_template = children[list_start]
        list_section_template = children[find_child_index_by_text(children, "個人作品", start=list_start)]
        list_grade_template = children[find_child_index_by_text(children, "一回生", start=list_start)]
        list_item_template = children[find_child_index_by_text_contains(children, "創作・", start=list_start)]
        vertical_transition = [copy.deepcopy(child) for child in children[detail_start - 3 : detail_start]]
        force_vertical_two_columns(vertical_transition)
        detail_section_end_template = find_vertical_section_end_template(children, detail_start, roster_start)
        detail_templates = children[218:225]
        if len(detail_templates) < 7:
            raise SystemExit("作品詳細ブロックのテンプレート取得に失敗しました。")

        new_children: list[ET.Element] = []
        new_children.extend(copy.deepcopy(child) for child in children[:list_start])
        new_children.extend(build_template_list_elements(works, list_title_template, list_section_template, list_grade_template, list_item_template))
        new_children.extend(vertical_transition)
        new_children.extend(build_template_detail_elements(works, detail_templates))
        new_children.append(clone_section_break(detail_section_end_template))
        new_children.extend(copy.deepcopy(child) for child in children[detail_end:])

        for child in list(body):
            body.remove(child)
        for child in new_children:
            body.append(child)

        document_xml = ET.tostring(document_root, encoding="utf-8", xml_declaration=True, short_empty_elements=True)

        with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as output:
            for item in template.infolist():
                if item.filename == "word/document.xml":
                    output.writestr(item, document_xml)
                else:
                    output.writestr(item, template.read(item.filename))


def remove_ignorable_namespace_hint(document_root: ET.Element) -> None:
    # ElementTree may rewrite prefixes (for example w14 -> ns2). If mc:Ignorable
    # still refers to the original prefixes, Word treats the file as repairable.
    document_root.attrib.pop(f"{{{MC_NS}}}Ignorable", None)


def force_vertical_two_columns(elements: list[ET.Element]) -> None:
    for element in elements:
        section = element.find("w:pPr/w:sectPr", WORD_NS)
        if section is None:
            continue
        cols = section.find("w:cols", WORD_NS)
        if cols is None:
            cols = ET.SubElement(section, f"{{{W_NS}}}cols")
        cols.attrib[f"{{{W_NS}}}num"] = "2"
        cols.attrib.setdefault(f"{{{W_NS}}}space", "720")

        text_direction = section.find("w:textDirection", WORD_NS)
        if text_direction is None:
            text_direction = ET.SubElement(section, f"{{{W_NS}}}textDirection")
        text_direction.attrib[f"{{{W_NS}}}val"] = "tbRl"
        return


def find_vertical_section_end_template(children: list[ET.Element], start: int, end: int) -> ET.Element:
    candidate: ET.Element | None = None
    for child in children[start:end]:
        section = child.find("w:pPr/w:sectPr", WORD_NS)
        if section is None:
            continue
        text_direction = section.find("w:textDirection", WORD_NS)
        cols = section.find("w:cols", WORD_NS)
        if text_direction is not None and text_direction.attrib.get(f"{{{W_NS}}}val") == "tbRl":
            if cols is not None and cols.attrib.get(f"{{{W_NS}}}num") == "2":
                candidate = child
    if candidate is None:
        raise SystemExit("作品詳細末尾の縦書きセクション設定が見つかりません。")
    return candidate


def clone_section_break(template: ET.Element) -> ET.Element:
    paragraph = ET.Element(f"{{{W_NS}}}p")
    ppr = template.find("w:pPr", WORD_NS)
    if ppr is not None:
        paragraph.append(copy.deepcopy(ppr))
    return paragraph


def build_template_list_elements(
    works: list[Work],
    title_template: ET.Element,
    section_template: ET.Element,
    grade_template: ET.Element,
    item_template: ET.Element,
) -> list[ET.Element]:
    elements: list[ET.Element] = [clone_paragraph_with_text(title_template, "作品一覧")]
    elements.append(clone_paragraph_with_text(section_template, "個人作品"))

    for grade, grade_works in group_personal_by_grade(works).items():
        elements.append(clone_paragraph_with_text(grade_template, grade))
        for work in grade_works:
            elements.append(clone_paragraph_with_text(item_template, list_line_for_template(work)))

    collaboration = [work for work in works if work.section == "合作"]
    if collaboration:
        elements.append(clone_paragraph_with_text(section_template, "合作"))
        for work in collaboration:
            elements.append(clone_paragraph_with_text(item_template, list_line_for_template(work)))
    return elements


def build_template_detail_elements(works: list[Work], templates: list[ET.Element]) -> list[ET.Element]:
    elements: list[ET.Element] = []
    for work in works:
        elements.append(clone_paragraph_with_text(templates[0], display_name(work)))
        elements.append(clone_paragraph_with_text(templates[1], f"{work.kind}　{work.title}　{work.size}"))
        elements.append(clone_paragraph_with_text(templates[2], ""))
        elements.append(clone_paragraph_with_text(templates[3], quoted_text(work.text)))
        elements.append(clone_paragraph_with_text(templates[4], work.comment))
        elements.append(clone_paragraph_with_text(templates[5], "【作品画像：ここに手入力で配置】"))
        elements.append(clone_paragraph_with_text(templates[6], ""))
    return elements


def list_line_for_template(work: Work) -> str:
    page_hint = 11 + work.number - 1
    name = one_line(work.name) if work.section == "合作" else compact_for_sort(work.name)
    return f"{name}\t{work.kind}・{work.title}\t{page_hint}"


def quoted_text(text: str) -> str:
    if not text:
        return ""
    stripped = text.strip()
    if stripped.startswith("「") and stripped.endswith("」"):
        return stripped
    return f"「{stripped}」"


def clone_paragraph_with_text(template: ET.Element, text: str) -> ET.Element:
    paragraph_element = f"{{{W_NS}}}p"
    run_element = f"{{{W_NS}}}r"
    text_element = f"{{{W_NS}}}t"
    break_element = f"{{{W_NS}}}br"
    tab_element = f"{{{W_NS}}}tab"

    # Do not copy paragraph IDs such as w14:paraId. Duplicating them across
    # generated paragraphs can make Word treat the document as repairable.
    paragraph = ET.Element(paragraph_element)
    ppr = template.find("w:pPr", WORD_NS)
    if ppr is not None:
        paragraph.append(copy.deepcopy(ppr))

    run_template = template.find(".//w:r", WORD_NS)
    run_properties = run_template.find("w:rPr", WORD_NS) if run_template is not None else None

    lines = text.splitlines() or [""]
    for index, line in enumerate(lines):
        if index:
            run = ET.SubElement(paragraph, run_element)
            if run_properties is not None:
                run.append(copy.deepcopy(run_properties))
            ET.SubElement(run, break_element)

        parts = line.split("\t")
        for part_index, part in enumerate(parts):
            if part_index:
                run = ET.SubElement(paragraph, run_element)
                if run_properties is not None:
                    run.append(copy.deepcopy(run_properties))
                ET.SubElement(run, tab_element)
            if part:
                run = ET.SubElement(paragraph, run_element)
                if run_properties is not None:
                    run.append(copy.deepcopy(run_properties))
                text_node = ET.SubElement(run, text_element)
                if part.startswith(" ") or part.endswith(" "):
                    text_node.attrib["{http://www.w3.org/XML/1998/namespace}space"] = "preserve"
                text_node.text = part
    return paragraph


def paragraph_text(element: ET.Element) -> str:
    return "".join(text.text or "" for text in element.findall(".//w:t", WORD_NS)).strip()


def find_child_index_by_text(children: list[ET.Element], text: str, start: int = 0) -> int:
    for index, child in enumerate(children[start:], start=start):
        if paragraph_text(child) == text:
            return index
    raise SystemExit(f"テンプレート内に {text!r} が見つかりません。")


def find_child_index_by_text_contains(children: list[ET.Element], text: str, start: int = 0) -> int:
    for index, child in enumerate(children[start:], start=start):
        if text in paragraph_text(child):
            return index
    raise SystemExit(f"テンプレート内に {text!r} を含む段落が見つかりません。")


def find_vertical_detail_start(children: list[ET.Element], list_start: int) -> int:
    for index in range(list_start + 1, len(children)):
        if paragraph_text(children[index]) == "賛助作品" and any(has_section_break(child) for child in children[list_start:index]):
            return index
    raise SystemExit("作品詳細セクションの開始位置が見つかりません。")


def find_previous_section_break(children: list[ET.Element], before: int) -> int:
    for index in range(before - 1, -1, -1):
        if has_section_break(children[index]):
            return index
    return before


def has_section_break(element: ET.Element) -> bool:
    return element.find("w:pPr/w:sectPr", WORD_NS) is not None


def build_document_xml(works: list[Work]) -> str:
    body: list[str] = []
    body.append(paragraph("作品一覧", style="Title"))
    body.append(paragraph("個人作品", style="Heading1"))

    for grade, grade_works in group_personal_by_grade(works).items():
        body.append(paragraph(grade, style="Heading2"))
        for work in grade_works:
            body.append(paragraph(f"{work.number:02d}. {normalize_space(work.name)}　{work.kind}・{work.title}", style="ListLine"))

    collaboration = [work for work in works if work.section == "合作"]
    if collaboration:
        body.append(paragraph("合作", style="Heading1"))
        for work in collaboration:
            body.append(paragraph(f"{work.number:02d}. {one_line(work.name)}　{work.kind}・{work.title}", style="ListLine"))

    for work in works:
        body.append(page_break())
        body.append(paragraph(f"No. {work.number:02d}", style="Heading1"))
        body.append(paragraph("作品画像：ここに手入力で配置", style="ImagePlaceholder"))
        body.append(paragraph(display_name(work), style="Heading2"))
        body.append(paragraph(f"{work.kind}　{work.title}"))
        body.append(paragraph(f"書体：{work.style}　サイズ：{work.size}　向き：{work.orientation}"))
        body.append(paragraph(f"展示場所：{work.location}　表装形式：{work.mounting}"))
        if work.text:
            body.append(paragraph("釈文", style="Heading3"))
            body.append(paragraph(work.text))
        if work.comment:
            body.append(paragraph("作品コメント", style="Heading3"))
            body.append(paragraph(work.comment))

    body.append(SECTION_PROPERTIES_XML)
    return DOCUMENT_XML_TEMPLATE.format(body="".join(body))


def display_name(work: Work) -> str:
    if work.section == "合作":
        return one_line(work.name)
    return f"{normalize_space(work.name)}　{work.grade}"


def paragraph(text: str, style: str | None = None) -> str:
    style_xml = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
    runs = "".join(run(part) if index == 0 else '<w:r><w:br/></w:r>' + run(part) for index, part in enumerate(text.splitlines()))
    return f"<w:p>{style_xml}{runs}</w:p>"


def run(text: str) -> str:
    preserve = ' xml:space="preserve"' if text.startswith(" ") or text.endswith(" ") else ""
    return f"<w:r><w:t{preserve}>{escape(text)}</w:t></w:r>"


def page_break() -> str:
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'


def validate_works(works: list[Work]) -> list[str]:
    warnings: list[str] = []
    for work in works:
        missing = []
        for label, value in [
            ("作者", work.name),
            ("臨書 or 創作", work.kind),
            ("作品名", work.title),
            ("作品サイズ", work.size),
            ("釈文", work.text),
            ("作品コメント", work.comment),
        ]:
            if not value:
                missing.append(label)
        if missing:
            warnings.append(f"No.{work.number:02d} {display_name(work)}: {', '.join(missing)} が空です")
    return warnings


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="作品情報フォーム.xlsx からパンフレットを生成します。",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=textwrap.dedent(
            """\
            例:
              python scripts/generate.py
              python scripts/generate.py --input 作品情報フォーム.xlsx --docx output/パンフレット.docx
              python scripts/generate.py --plain-docx output/簡易版.docx
            """
        ),
    )
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT, help="入力する xlsx ファイル")
    parser.add_argument("--docx", type=Path, default=DEFAULT_DOCX_OUTPUT, help="出力する docx ファイル")
    parser.add_argument("--list", type=Path, default=DEFAULT_LIST_OUTPUT, help="確認用の作品一覧テキスト")
    parser.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE, help="書式をコピーする前回パンフレット docx")
    parser.add_argument("--plain-docx", type=Path, help="テンプレートを使わない簡易版 docx も出力する場合の出力先")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    sheets = read_xlsx(args.input)
    works = build_works(sheets)
    if not works:
        raise SystemExit("作品データが見つかりませんでした。")

    write_list_text(works, args.list)
    write_template_docx(works, args.template, args.docx)
    if args.plain_docx:
        write_docx(works, args.plain_docx)

    print(f"作品数: {len(works)}")
    print(f"docx: {args.docx}")
    print(f"list: {args.list}")
    if args.plain_docx:
        print(f"plain docx: {args.plain_docx}")

    warnings = validate_works(works)
    if warnings:
        print("\n確認事項:")
        for warning in warnings:
            print(f"- {warning}")


CONTENT_TYPES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>
"""

ROOT_RELS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""

DOCUMENT_RELS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"""

DOCUMENT_XML_TEMPLATE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>{body}</w:body>
</w:document>
"""

SECTION_PROPERTIES_XML = """
<w:sectPr>
  <w:pgSz w:w="11906" w:h="16838"/>
  <w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="708" w:footer="708" w:gutter="0"/>
</w:sectPr>
"""

STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:rPr>
      <w:rFonts w:ascii="Yu Mincho" w:hAnsi="Yu Mincho" w:eastAsia="Yu Mincho"/>
      <w:sz w:val="21"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Title">
    <w:name w:val="Title"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:after="240"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="36"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="Heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:before="240" w:after="120"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="28"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="Heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:before="160" w:after="80"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="24"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="Heading 3"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:before="120" w:after="40"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="21"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ListLine">
    <w:name w:val="List Line"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:after="40"/></w:pPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ImagePlaceholder">
    <w:name w:val="Image Placeholder"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr>
      <w:spacing w:before="160" w:after="160"/>
      <w:jc w:val="center"/>
    </w:pPr>
    <w:rPr><w:i/><w:color w:val="808080"/><w:sz w:val="24"/></w:rPr>
  </w:style>
</w:styles>
"""


if __name__ == "__main__":
    main()
