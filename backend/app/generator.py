from __future__ import annotations

import copy
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET


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


@dataclass(frozen=True)
class SheetData:
    headers: list[str]
    rows: list[dict[str, str]]


class UserFacingError(Exception):
    """An error message intended to be shown directly to non-technical users."""


REQUIRED_COLUMNS = {
    "個人": [
        "氏名",
        "ふりがな",
        "学年",
        "臨書 or 創作",
        "書体",
        "作品名",
        "作品の向き",
        "作品サイズ",
        "展示場所",
        "表装形式",
        "釈文",
        "作品コメント",
    ],
    "合作": [
        "合作参加者全員分",
        "臨書 or 創作",
        "書体",
        "作品名",
        "作品の向き",
        "作品サイズ",
        "展示場所",
        "表装形式",
        "釈文",
        "作品コメント",
    ],
}


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


def read_xlsx(path: Path) -> dict[str, SheetData]:
    try:
        book = zipfile.ZipFile(path)
    except FileNotFoundError as exc:
        raise UserFacingError(
            f"入力Excelファイルが見つかりません。\n"
            f"確認する場所: {path}\n"
            "直し方: input フォルダに 作品情報フォーム.xlsx を置いてから、もう一度実行してください。"
        ) from exc
    except zipfile.BadZipFile as exc:
        raise UserFacingError(
            f"入力Excelファイルを開けませんでした。\n"
            f"ファイル: {path}\n"
            "直し方: Excelで開ける .xlsx ファイルか確認してください。古い .xls 形式や、壊れたファイルは使えません。"
        ) from exc

    with book:
        shared_strings: list[str] = []
        if "xl/sharedStrings.xml" in book.namelist():
            shared_xml = ET.fromstring(book.read("xl/sharedStrings.xml"))
            for item in shared_xml.findall("a:si", SPREADSHEET_NS):
                shared_strings.append("".join(text.text or "" for text in item.findall(".//a:t", SPREADSHEET_NS)))

        try:
            workbook = ET.fromstring(book.read("xl/workbook.xml"))
            rels = ET.fromstring(book.read("xl/_rels/workbook.xml.rels"))
        except KeyError as exc:
            raise UserFacingError(
                f"入力Excelファイルの中身を読み取れませんでした。\n"
                f"ファイル: {path}\n"
                "直し方: Excelで開いて、別名で .xlsx として保存し直してから、もう一度実行してください。"
            ) from exc
        rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}

        sheets: dict[str, SheetData] = {}
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
                sheets[sheet_name] = SheetData(headers=[], rows=[])
                continue

            headers = raw_rows[0]
            header_names = [header for header in headers.values() if header]
            records: list[dict[str, str]] = []
            for row in raw_rows[1:]:
                record = {header: row.get(col, "").strip() for col, header in headers.items() if header}
                if any(record.values()):
                    records.append(record)
            sheets[sheet_name] = SheetData(headers=header_names, rows=records)

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
    table = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9, "十": 10}
    return table.get(value, int(value) if value.isdigit() else 99)


def number_to_kanji(value: str) -> str:
    table = {
        "1": "一",
        "2": "二",
        "3": "三",
        "4": "四",
        "5": "五",
        "6": "六",
        "7": "七",
        "8": "八",
        "9": "九",
        "10": "十",
    }
    return table.get(value, value)


def display_grade(grade: str) -> str:
    text = compact_for_sort(grade)
    if match := re.fullmatch(r"B(\d+)", text, re.IGNORECASE):
        return f"{number_to_kanji(match.group(1))}回生"
    if match := re.fullmatch(r"M(\d+)", text, re.IGNORECASE):
        return f"修士{number_to_kanji(match.group(1))}回生"
    if match := re.fullmatch(r"D(\d+)", text, re.IGNORECASE):
        return f"博士{number_to_kanji(match.group(1))}回生"
    text = re.sub(r"修士(\d+)回生", lambda m: f"修士{number_to_kanji(m.group(1))}回生", text)
    text = re.sub(r"博士(\d+)回生", lambda m: f"博士{number_to_kanji(m.group(1))}回生", text)
    text = re.sub(r"(?<!士)(\d+)回生", lambda m: f"{number_to_kanji(m.group(1))}回生", text)
    return text or grade


def display_participant_grades(value: str) -> str:
    text = value.strip()
    text = re.sub(r"（B(\d+)）", lambda m: f"（{number_to_kanji(m.group(1))}回生）", text, flags=re.IGNORECASE)
    text = re.sub(r"（M(\d+)）", lambda m: f"（修士{number_to_kanji(m.group(1))}回生）", text, flags=re.IGNORECASE)
    text = re.sub(r"（D(\d+)）", lambda m: f"（博士{number_to_kanji(m.group(1))}回生）", text, flags=re.IGNORECASE)
    text = re.sub(r"（(\d+)回生）", lambda m: f"（{number_to_kanji(m.group(1))}回生）", text)
    return text


def build_works(sheets: dict[str, SheetData]) -> list[Work]:
    validate_input_sheets(sheets)
    personal = sheets["個人"].rows
    collaboration = sheets["合作"].rows

    personal_works: list[dict[str, str]] = []
    for row in personal:
        personal_works.append(
            {
                "section": "個人作品",
                "name": value_by_header(row, "氏名"),
                "kana": value_by_header(row, "ふりがな"),
                "grade": display_grade(value_by_header(row, "学年")),
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
                "name": display_participant_grades(value_by_header(row, "合作参加者全員分")),
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


def validate_input_sheets(sheets: dict[str, SheetData]) -> None:
    missing_sheets = [sheet_name for sheet_name in REQUIRED_COLUMNS if sheet_name not in sheets]
    if missing_sheets:
        found_sheets = "、".join(sheets) if sheets else "なし"
        raise UserFacingError(
            "入力Excelファイルのシート名を確認してください。\n"
            f"足りないシート: {'、'.join(missing_sheets)}\n"
            f"見つかったシート: {found_sheets}\n"
            "直し方: Excel下部のシート名を「個人」と「合作」にしてください。余分な文字やスペースも入れないでください。"
        )

    missing_by_sheet: list[str] = []
    for sheet_name, required_columns in REQUIRED_COLUMNS.items():
        headers = sheets[sheet_name].headers
        missing_columns = [column for column in required_columns if not has_matching_header(headers, column)]
        if missing_columns:
            missing_by_sheet.append(f"{sheet_name}シート: {'、'.join(missing_columns)}")

    if missing_by_sheet:
        raise UserFacingError(
            "入力Excelファイルの列名を確認してください。\n"
            + "\n".join(missing_by_sheet)
            + "\n直し方: 1行目の列名をフォームの元の名前に戻してください。説明文が後ろに続くのは問題ありません。"
        )

    if all(not sheets[sheet_name].rows for sheet_name in REQUIRED_COLUMNS):
        raise UserFacingError(
            "入力Excelファイルに作品データがありません。\n"
            "直し方: 1行目は列名のままにして、2行目以降に作品情報を入力してください。"
        )


def has_matching_header(headers: list[str], starts_with: str) -> bool:
    return any(header.replace("\n", " ").strip().startswith(starts_with) for header in headers)


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


def write_template_docx(works: list[Work], template_path: Path, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        template = zipfile.ZipFile(template_path)
    except FileNotFoundError as exc:
        raise UserFacingError(
            f"テンプレートのWordファイルが見つかりません。\n"
            f"確認する場所: {template_path}\n"
            "直し方: サーバーに同梱されているパンフ鋳型を確認してください。"
        ) from exc
    except zipfile.BadZipFile as exc:
        raise UserFacingError(
            f"テンプレートのWordファイルを開けませんでした。\n"
            f"ファイル: {template_path}\n"
            "直し方: Wordで開ける .docx ファイルか確認してください。古い .doc 形式や、壊れたファイルは使えません。"
        ) from exc

    with template:
        try:
            document_root = ET.fromstring(template.read("word/document.xml"))
        except KeyError as exc:
            raise UserFacingError(
                f"テンプレートのWordファイルの中身を読み取れませんでした。\n"
                f"ファイル: {template_path}\n"
                "直し方: Wordで開いて、別名で .docx として保存し直してから、もう一度実行してください。"
            ) from exc
        body = document_root.find("w:body", WORD_NS)
        if body is None:
            raise UserFacingError(
                "テンプレートのWordファイルに本文が見つかりません。\n"
                "直し方: パンフ鋳型が、前回パンフレットのWordファイルか確認してください。"
            )
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
            raise UserFacingError(
                "テンプレートから作品詳細ページの書式を読み取れませんでした。\n"
                "直し方: パンフ鋳型を、前回パンフレットの元ファイルに戻してください。"
            )

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
        raise UserFacingError(
            "テンプレートの作品詳細ページの区切りを見つけられませんでした。\n"
            "直し方: パンフ鋳型を、前回パンフレットの元ファイルに戻してください。"
        )
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
            elements.append(clone_paragraph_with_text(item_template, list_line_for_template(work), vertical_last_number=True))

    collaboration = [work for work in works if work.section == "合作"]
    if collaboration:
        elements.append(clone_paragraph_with_text(section_template, "合作"))
        for work in collaboration:
            elements.append(clone_paragraph_with_text(item_template, list_line_for_template(work), vertical_last_number=True))
    return elements


def build_template_detail_elements(works: list[Work], templates: list[ET.Element]) -> list[ET.Element]:
    elements: list[ET.Element] = []
    for work in works:
        elements.append(clone_paragraph_with_text(templates[0], display_name(work)))
        elements.append(clone_paragraph_with_text(templates[1], f"{work.kind}　{work.title}　{work.size}"))
        elements.append(clone_paragraph_with_text(templates[2], ""))
        elements.append(clone_indexed_text_paragraph(templates[3], work.number, quoted_text(work.text)))
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


def clone_paragraph_with_text(template: ET.Element, text: str, vertical_last_number: bool = False) -> ET.Element:
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
                if vertical_last_number and part_index == len(parts) - 1 and part.isdigit():
                    rpr = run.find("w:rPr", WORD_NS)
                    if rpr is None:
                        rpr = ET.Element(f"{{{W_NS}}}rPr")
                        run.insert(0, rpr)
                    east_asian_layout = ET.SubElement(rpr, f"{{{W_NS}}}eastAsianLayout")
                    east_asian_layout.attrib[f"{{{W_NS}}}vert"] = "1"
                    east_asian_layout.attrib[f"{{{W_NS}}}vertCompress"] = "1"
                text_node = ET.SubElement(run, text_element)
                if part.startswith(" ") or part.endswith(" "):
                    text_node.attrib["{http://www.w3.org/XML/1998/namespace}space"] = "preserve"
                text_node.text = part
    return paragraph


def clone_indexed_text_paragraph(template: ET.Element, number: int, text: str) -> ET.Element:
    paragraph = copy.deepcopy(template)
    strip_generated_id_attrs(paragraph)
    update_drawing_ids(paragraph, number)

    text_nodes = paragraph.findall(".//w:t", WORD_NS)
    if not text_nodes:
        return clone_paragraph_with_text(template, text)

    text_nodes[0].text = str(number)
    for node in text_nodes[1:]:
        node.text = ""
    text_nodes[-1].text = text

    return paragraph


def strip_generated_id_attrs(element: ET.Element) -> None:
    for node in element.iter():
        for attr in list(node.attrib):
            local_name = attr.rsplit("}", 1)[-1]
            if local_name in {"paraId", "textId", "rsidR", "rsidRPr", "rsidRDefault", "rsidP"}:
                node.attrib.pop(attr, None)


def update_drawing_ids(element: ET.Element, number: int) -> None:
    base_id = 900000 + number
    for node in element.iter():
        local_name = node.tag.rsplit("}", 1)[-1]
        if local_name in {"docPr", "cNvPr"} and "id" in node.attrib:
            node.attrib["id"] = str(base_id)
        if local_name == "docPr" and "name" in node.attrib:
            node.attrib["name"] = f"作品番号 {number}"
        for attr in list(node.attrib):
            if attr.rsplit("}", 1)[-1] in {"anchorId", "editId"}:
                node.attrib[attr] = f"{base_id:08X}"[-8:]
            if attr == "id" and str(node.attrib[attr]).startswith("_x0000_s"):
                node.attrib[attr] = f"_x0000_s{base_id}"


def paragraph_text(element: ET.Element) -> str:
    return "".join(text.text or "" for text in element.findall(".//w:t", WORD_NS)).strip()


def find_child_index_by_text(children: list[ET.Element], text: str, start: int = 0) -> int:
    for index, child in enumerate(children[start:], start=start):
        if paragraph_text(child) == text:
            return index
    raise UserFacingError(
        "テンプレートに必要な見出しが見つかりません。\n"
        f"見つからない文字: {text}\n"
        "直し方: パンフ鋳型が、前回パンフレットのWordファイルか確認してください。"
    )


def find_child_index_by_text_contains(children: list[ET.Element], text: str, start: int = 0) -> int:
    for index, child in enumerate(children[start:], start=start):
        if text in paragraph_text(child):
            return index
    raise UserFacingError(
        "テンプレートに必要な行が見つかりません。\n"
        f"見つからない文字: {text}\n"
        "直し方: パンフ鋳型が、前回パンフレットのWordファイルか確認してください。"
    )


def find_vertical_detail_start(children: list[ET.Element], list_start: int) -> int:
    for index in range(list_start + 1, len(children)):
        if paragraph_text(children[index]) == "賛助作品" and any(has_section_break(child) for child in children[list_start:index]):
            return index
    raise UserFacingError(
        "テンプレートの作品詳細ページを見つけられませんでした。\n"
        "直し方: パンフ鋳型を、前回パンフレットの元ファイルに戻してください。"
    )


def find_previous_section_break(children: list[ET.Element], before: int) -> int:
    for index in range(before - 1, -1, -1):
        if has_section_break(children[index]):
            return index
    return before


def has_section_break(element: ET.Element) -> bool:
    return element.find("w:pPr/w:sectPr", WORD_NS) is not None


def display_name(work: Work) -> str:
    if work.section == "合作":
        return one_line(work.name)
    return f"{normalize_space(work.name)}　{work.grade}"


