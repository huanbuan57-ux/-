"""DOCX 技术通知模板渲染与校验。

模板使用普通 Word 占位符。包含 ``{{row.*}}`` 的表格行会按
``notification_data['rows']`` 重复；其余占位符在正文、表格、页眉和页脚中替换。
"""
from __future__ import annotations

from copy import deepcopy
from pathlib import Path
import re
from typing import Any, Iterable

from docx import Document
from docx.table import _Cell, _Row, Table
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn


# 这些是模板契约中必须存在的槽位。其它业务字段（country、delivery_place 等）
# 可按项目需要加入模板，不会被校验器误判为非法。
REQUIRED_GLOBAL_PLACEHOLDERS = (
    "{{number}}",
    "{{date}}",
    "{{contract_no}}",
    "{{model}}",
    "{{general_requirements}}",
    "{{label_url}}",
)
REQUIRED_ROW_FIELDS = (
    "code",
    "cn",
    "en",
    "quantity",
    "modulus",
    "units",
    "cartons",
    "station",
    "size",
    "weight",
    "logo_type", "country", "contract_no", "model",
)
_TOKEN_RE = re.compile(r"\{\{\s*([^{}]+?)\s*\}\}")
_ROW_TOKEN_RE = re.compile(r"\{\{\s*row\.([^{}]+?)\s*\}\}")
GLOBAL_ALIASES = {"number": "通知编号", "date": "实施日期", "contract_no": "合同号",
                  "model": "车型", "general_requirements": "通用技术要求", "label_url": "标签网址",
                  "country": "国家", "logo_type": "LOGO类型", "delivery_place": "交货地点",
                  "delivery_time": "交货时间", "finish_date": "完成日期"}
ROW_ALIASES = {"code": "图号", "cn": "中文名称", "en": "英文名称", "quantity": "数量",
               "modulus": "包装模数", "units": "单箱数量", "cartons": "纸箱总数",
               "station": "工位号", "size": "纸箱尺寸", "weight": "重量说明",
               "logo_type": "LOGO类型", "country": "国家", "contract_no": "合同号", "model": "车型"}


def _iter_table_cells(table: Table) -> Iterable[_Cell]:
    for row in table.rows:
        for cell in row.cells:
            yield cell
            # python-docx exposes nested tables through cell.tables.
            for nested in cell.tables:
                yield from _iter_table_cells(nested)


def _iter_paragraphs(document: Document) -> Iterable[Paragraph]:
    # XML traversal includes nested tables and text boxes, which .paragraphs
    # alone deliberately excludes.
    for element in document.element.body.iter(qn("w:p")):
        yield Paragraph(element, document)
    seen = set()
    for section in document.sections:
        for part in (section.header, section.first_page_header, section.even_page_header,
                     section.footer, section.first_page_footer, section.even_page_footer):
            if not part._has_definition:
                continue
            if id(part._element) in seen:
                continue
            seen.add(id(part._element))
            for element in part._element.iter(qn("w:p")):
                yield Paragraph(element, part)


def _paragraph_text(paragraph: Paragraph) -> str:
    return "".join(run.text or "" for run in paragraph.runs)


def _replace_paragraph(paragraph: Paragraph, replacements: dict[str, Any]) -> None:
    """跨 run 替换占位符，保留未修改文字的 run 格式。"""
    if not paragraph.runs:
        return
    original = _paragraph_text(paragraph)
    if not original:
        return

    runs = paragraph.runs
    spans, offset = [], 0
    for run in runs:
        spans.append((offset, offset + len(run.text)))
        offset += len(run.text)
    for match in reversed(list(_TOKEN_RE.finditer(original))):
        key = match.group(1).strip()
        if key not in replacements:
            continue
        value = "" if replacements[key] is None else str(replacements[key])
        start = next(i for i, (a, b) in enumerate(spans) if a <= match.start() < b)
        end = next(i for i, (a, b) in enumerate(spans) if a < match.end() <= b)
        left = match.start() - spans[start][0]
        right = match.end() - spans[end][0]
        if start == end:
            runs[start].text = runs[start].text[:left] + value + runs[start].text[right:]
        else:
            runs[start].text = runs[start].text[:left] + value
            for i in range(start + 1, end):
                runs[i].text = ""
            runs[end].text = runs[end].text[right:]


def _replace_container(container: Any, replacements: dict[str, Any]) -> None:
    for paragraph in container.paragraphs:
        _replace_paragraph(paragraph, replacements)
    for table in container.tables:
        for cell in _iter_table_cells(table):
            for paragraph in cell.paragraphs:
                _replace_paragraph(paragraph, replacements)


def _row_text(row: _Row) -> str:
    return "\n".join(cell.text for cell in row.cells)


def _row_has_tokens(row: _Row) -> bool:
    return bool(_ROW_TOKEN_RE.search(_row_text(row))) or "{{图号}}" in _row_text(row)


def _replace_row(row: _Row, row_data: dict[str, Any], globals_: dict[str, Any]) -> None:
    replacements = dict(globals_)
    normalized = dict(row_data)
    for key, alias in ROW_ALIASES.items():
        if key not in normalized and alias in row_data:
            normalized[key] = row_data[alias]
        if key not in normalized and key in globals_:
            normalized[key] = globals_[key]
    replacements.update({f"row.{key}": value for key, value in normalized.items()})
    replacements.update({alias: normalized[key] for key, alias in ROW_ALIASES.items() if key in normalized})
    for cell in row.cells:
        for element in cell._element.iter(qn("w:p")):
            _replace_paragraph(Paragraph(element, cell), replacements)


def _expand_table_rows(table: Table, rows: list[dict[str, Any]], globals_: dict[str, Any]) -> None:
    # Snapshot indices first: inserting/removing XML rows while iterating the
    # live python-docx collection otherwise skips adjacent template rows.
    prototypes = [(idx, row) for idx, row in enumerate(table.rows) if _row_has_tokens(row)]
    for original_idx, prototype in reversed(prototypes):
        tr = prototype._tr
        parent = tr.getparent()
        insert_at = parent.index(tr)
        for row_data in rows:
            clone = deepcopy(tr)
            parent.insert(insert_at, clone)
            new_row = _Row(clone, table)
            _replace_row(new_row, row_data, globals_)
            insert_at += 1
        parent.remove(tr)


def _expand_all_tables(document: Document, rows: list[dict[str, Any]], globals_: dict[str, Any]) -> None:
    def expand_tables(tables: Iterable[Table]) -> None:
        for table in tables:
            _expand_table_rows(table, rows, globals_)
            for row in table.rows:
                for cell in row.cells:
                    expand_tables(cell.tables)
    expand_tables(document.tables)
    for section in document.sections:
        for part in (section.header, section.first_page_header, section.even_page_header,
                     section.footer, section.first_page_footer, section.even_page_footer):
            if part._has_definition:
                expand_tables(part.tables)


def _all_text(document: Document) -> str:
    return "\n".join(_paragraph_text(p) for p in _iter_paragraphs(document))


def validate_template(path: str | Path) -> list[str]:
    """返回模板缺失的占位符列表。文件不可读也视为模板错误。"""
    path = Path(path)
    if not path.exists():
        raise ValueError(f"模板不存在：{path}")
    try:
        document = Document(path)
    except Exception as exc:  # python-docx 对旧版 .doc 会走到这里
        raise ValueError(f"无法读取 DOCX 模板：{path}") from exc
    text = _all_text(document)
    tokens = set(_TOKEN_RE.findall(text))
    missing = [token for token in REQUIRED_GLOBAL_PLACEHOLDERS
               if token[2:-2] not in tokens and GLOBAL_ALIASES[token[2:-2]] not in tokens]
    # Required product slots must be in the repeating row, not scattered in
    # document metadata where they cannot produce one record per product.
    dynamic_text = "\n".join(_row_text(row) for table in document.tables for row in table.rows if _row_has_tokens(row))
    dynamic_tokens = set(_TOKEN_RE.findall(dynamic_text))
    missing.extend(f"{{{{row.{field}}}}}" for field in REQUIRED_ROW_FIELDS
                   if f"row.{field}" not in dynamic_tokens and ROW_ALIASES[field] not in dynamic_tokens)
    return missing


def _validate_data(notification_data: dict[str, Any]) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    if not isinstance(notification_data, dict):
        raise ValueError("notification_data 必须是 dict")
    rows = notification_data.get("rows", [])
    if rows is None:
        rows = []
    if not isinstance(rows, list) or any(not isinstance(row, dict) for row in rows):
        raise ValueError("notification_data['rows'] 必须是 dict 列表")
    if not isinstance(notification_data.get("globals", {}), dict):
        raise ValueError("notification_data['globals'] 必须是 dict")
    globals_ = dict(notification_data.get("globals", {}))
    globals_.update({key: value for key, value in notification_data.items() if key not in ("rows", "globals")})
    for key, alias in GLOBAL_ALIASES.items():
        if key not in globals_ and alias in globals_:
            globals_[key] = globals_[alias]
        if alias not in globals_ and key in globals_:
            globals_[alias] = globals_[key]
    return globals_, rows


def render_notification(template, notification_data: dict[str, Any], output_path) -> Path:
    """按模板渲染技术通知并返回输出路径。

    ``template`` 和 ``output_path`` 接受字符串或 :class:`pathlib.Path`。模板
    错误、输入结构错误、或渲染后仍有占位符时抛出 ``ValueError``。
    """
    template = Path(template)
    output_path = Path(output_path)
    if template.resolve() == output_path.resolve():
        raise ValueError("输出文件不能覆盖模板")
    missing = validate_template(template)
    if missing:
        raise ValueError("模板缺少占位符：" + "、".join(missing))
    globals_, rows = _validate_data(notification_data)
    try:
        document = Document(template)
    except Exception as exc:
        raise ValueError(f"无法读取 DOCX 模板：{template}") from exc

    _expand_all_tables(document, rows, globals_)
    for paragraph in _iter_paragraphs(document):
        _replace_paragraph(paragraph, globals_)

    unresolved = sorted(set(_TOKEN_RE.findall(_all_text(document))))
    if unresolved:
        raise ValueError("渲染后仍存在未解析占位符：" + "、".join("{{" + t + "}}" for t in unresolved))
    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_path)
    return output_path
