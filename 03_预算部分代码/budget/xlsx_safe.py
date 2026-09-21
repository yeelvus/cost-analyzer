# -*- coding: utf-8 -*-
"""把 openpyxl 写出的 xlsx 改成 Mac Excel 能直接打开的格式。

openpyxl 3.1 会把文字写成 inlineStr、写入空 workbookProtection、
并给 AutoFilter 加 _xlnm._FilterDatabase。Mac Excel 每次打开都会弹修复。
"""

from __future__ import annotations

import math
import os
import re
import tempfile
import zipfile
from html import unescape
from pathlib import Path
from xml.sax.saxutils import escape

ILLEGAL_XML_CHARS = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")
INLINE_CELL_RE = re.compile(r"<c\b([^>]*)>(.*?)</c>", re.DOTALL)
INLINE_T_RE = re.compile(r"<t(?:\s[^>]*)?>(.*?)</t>", re.DOTALL)
RID_RE = re.compile(r'Id="rId(\d+)"')


def clean_xml_text(value):
    """去掉 Excel/XML 非法控制字符。"""
    if value is None:
        return ""
    return ILLEGAL_XML_CHARS.sub("", str(value))


def excel_cell_value(value):
    """写入单元格前清洗：NaN/Inf 不写，字符串去掉非法字符。"""
    if value is None:
        return None
    if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
        return None
    if isinstance(value, str):
        text = clean_xml_text(value)
        return text if text else None
    return value


def _intern(text: str, table: list[str], index: dict[str, int]) -> int:
    if text not in index:
        index[text] = len(table)
        table.append(text)
    return index[text]


def _convert_sheet_inline_strings(xml: str, table: list[str], index: dict[str, int]) -> str:
    def repl(match: re.Match) -> str:
        attrs, inner = match.group(1), match.group(2)
        if 't="inlineStr"' not in attrs:
            return match.group(0)
        attrs = (
            attrs.replace(' t="inlineStr"', "")
            .replace('t="inlineStr" ', "")
            .replace('t="inlineStr"', "")
            .strip()
        )
        t_match = INLINE_T_RE.search(inner)
        raw = t_match.group(1) if t_match else ""
        text = clean_xml_text(unescape(raw))
        idx = _intern(text, table, index)
        attr_prefix = f" {attrs}" if attrs else ""
        return f'<c{attr_prefix} t="s"><v>{idx}</v></c>'

    xml = INLINE_CELL_RE.sub(repl, xml)
    xml = re.sub(r"(</f>)<v></v>", r"\1", xml)
    xml = re.sub(r"(</f>)<v/>", r"\1", xml)
    return xml


def _fix_workbook_xml(xml: str) -> str:
    xml = re.sub(r"<workbookProtection\b[^>]*/>", "", xml)
    xml = re.sub(
        r"<workbookProtection\b[^>]*>.*?</workbookProtection>",
        "",
        xml,
        flags=re.DOTALL,
    )
    xml = re.sub(
        r'<definedName name="_xlnm\._FilterDatabase"[^>]*>.*?</definedName>',
        "",
        xml,
        flags=re.DOTALL,
    )
    xml = re.sub(r"<definedNames>\s*</definedNames>", "", xml)
    xml = re.sub(r'\sfullCalcOnLoad="1"', "", xml)
    return xml


def _next_rid(rels_xml: str) -> str:
    ids = [int(n) for n in RID_RE.findall(rels_xml)]
    return f"rId{(max(ids) + 1) if ids else 1}"


def _ensure_shared_strings_rels(rels_xml: str) -> str:
    if "sharedStrings" in rels_xml:
        return rels_xml
    rid = _next_rid(rels_xml)
    rel = (
        f'<Relationship Id="{rid}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" '
        'Target="sharedStrings.xml"/>'
    )
    if "</Relationships>" in rels_xml:
        return rels_xml.replace("</Relationships>", rel + "</Relationships>")
    return rels_xml


def _ensure_shared_strings_content_type(types_xml: str) -> str:
    if "sharedStrings.xml" in types_xml:
        return types_xml
    override = (
        '<Override PartName="/xl/sharedStrings.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
    )
    if "</Types>" in types_xml:
        return types_xml.replace("</Types>", override + "</Types>")
    return types_xml


def _build_sst(table: list[str]) -> str:
    items = []
    for text in table:
        extra = ' xml:space="preserve"' if text != text.strip() else ""
        items.append(f"<si><t{extra}>{escape(text)}</t></si>")
    body = "".join(items)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        f'count="{len(table)}" uniqueCount="{len(table)}">{body}</sst>'
    )


def sanitize_xlsx(path) -> None:
    """就地改写 xlsx，去掉 Mac Excel 会报修的内容。"""
    src = Path(path)
    if not src.exists():
        return

    with zipfile.ZipFile(src, "r") as zin:
        names = zin.namelist()
        files = {name: zin.read(name) for name in names}

    changed = False
    table: list[str] = []
    index: dict[str, int] = {}

    if "xl/sharedStrings.xml" in files:
        # 已有 sst 时仍转换新增 inlineStr，追加到表尾（本流程新建簿通常没有 sst）
        pass

    for name in list(files):
        if not name.startswith("xl/worksheets/sheet") or not name.endswith(".xml"):
            continue
        xml = files[name].decode("utf-8")
        if 't="inlineStr"' not in xml and "<v></v>" not in xml:
            continue
        new_xml = _convert_sheet_inline_strings(xml, table, index)
        if new_xml != xml:
            files[name] = new_xml.encode("utf-8")
            changed = True

    if "xl/workbook.xml" in files:
        xml = files["xl/workbook.xml"].decode("utf-8")
        new_xml = _fix_workbook_xml(xml)
        if new_xml != xml:
            files["xl/workbook.xml"] = new_xml.encode("utf-8")
            changed = True

    if table:
        files["xl/sharedStrings.xml"] = _build_sst(table).encode("utf-8")
        if "xl/_rels/workbook.xml.rels" in files:
            rels = files["xl/_rels/workbook.xml.rels"].decode("utf-8")
            files["xl/_rels/workbook.xml.rels"] = _ensure_shared_strings_rels(rels).encode("utf-8")
        if "[Content_Types].xml" in files:
            types = files["[Content_Types].xml"].decode("utf-8")
            files["[Content_Types].xml"] = _ensure_shared_strings_content_type(types).encode("utf-8")
        changed = True

    if not changed:
        return

    fd, tmp_name = tempfile.mkstemp(suffix=".xlsx", dir=str(src.parent))
    os.close(fd)
    tmp_path = Path(tmp_name)
    try:
        with zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for name, data in files.items():
                zout.writestr(name, data)
        os.replace(tmp_path, src)
    except Exception:
        if tmp_path.exists():
            tmp_path.unlink()
        raise


def safe_save_workbook(wb, path) -> None:
    """保存工作簿并做 Mac Excel 兼容处理；先写临时文件再替换。"""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp_name = tempfile.mkstemp(suffix=".xlsx", dir=str(path.parent))
    os.close(fd)
    tmp_path = Path(tmp_name)
    try:
        if getattr(wb, "security", None) is not None:
            wb.security = None
        calc = getattr(wb, "calculation", None)
        if calc is not None and hasattr(calc, "fullCalcOnLoad"):
            calc.fullCalcOnLoad = False
        wb.save(tmp_path)
        sanitize_xlsx(tmp_path)
        os.replace(tmp_path, path)
    except Exception:
        if tmp_path.exists():
            tmp_path.unlink()
        raise
