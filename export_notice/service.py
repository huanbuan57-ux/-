from __future__ import annotations

from dataclasses import dataclass, field, asdict
from datetime import date, datetime
from pathlib import Path
from typing import Any

from .core import filter_owned_products, validate_rows, resolve_field
from .io import load_batch, load_product_master
from .requirements import extract_requirements


@dataclass
class Analysis:
    rows: list
    all_rows: list
    excluded: list
    requirements: Any
    master: Any
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    requirement_sources: list[dict] = field(default_factory=list)


def analyze_rows(rows, master, texts) -> Analysis:
    owned, excluded = filter_owned_products(rows, master)
    req = extract_requirements(texts)
    for row in owned:
        if not row.fields.get('包装模数') and req.modulus != '/':
            row.fields['包装模数'] = req.modulus
        # 工位号必须来自本批文件；资料表中即使有值也不回退。
    validated = validate_rows(owned, master)
    errors = list(dict.fromkeys(req.errors + validated.errors))
    warnings = validated.warnings[:]
    finish_dates=set(str(row.fields['完成日期']) for row in owned if row.fields.get('完成日期'))
    if len(finish_dates)>1:
        errors.append('完成日期存在冲突：'+'、'.join(sorted(finish_dates)))
    if not owned:
        errors.append('本批次没有本人负责且启用的产品')
    unique = {}
    for row in owned:
        if row.standard_code in unique:
            warnings.append(f'{row.location}：重复图号不重复生成；首次来源：{unique[row.standard_code].location}')
        else:
            unique[row.standard_code] = row
    return Analysis(list(unique.values()), rows, excluded, req, master, errors, warnings)


def analyze_batch(folder, master_path, aliases_path=None):
    batch = load_batch(folder, aliases_path=aliases_path)
    result = analyze_rows(batch.rows, load_product_master(master_path), batch.requirement_texts)
    result.errors.extend(batch.errors)
    result.requirement_sources=batch.requirement_sources
    return result


def notification_data(analysis, number, effective_date):
    req = analysis.requirements
    products = []
    for row in analysis.rows:
        v = row.resolved
        products.append({
            'code': row.standard_code, 'cn': v['中文名称'], 'en': v['英文名称'],
            'quantity': row.quantity, 'modulus': v['包装模数'], 'units': v['单箱数量'],
            'cartons': row.carton_count, 'station': v['工位号'], 'size': v['纸箱尺寸'],
            'weight': v['重量说明'], 'country': req.country, 'logo_type': req.logo_type,
            'contract_no': req.contract_no, 'model': req.model,
        })
    return dict(number=number, date=effective_date, contract_no=req.contract_no,
                model=req.model, country=req.country, logo_type=req.logo_type,
                delivery_place=req.delivery_place, delivery_time=req.delivery_time,
                label_url=req.label_url, general_requirements=req.general_requirements,
                finish_date=next((row.fields['完成日期'] for row in analysis.rows if row.fields.get('完成日期')), '/'), rows=products)


def generate(analysis, template, output_root, batch_name, number='', effective_date=''):
    from .document import render_notification, validate_template
    from .report import write_report
    try:
        template_errors=validate_template(template)
    except ValueError as exc:
        template_errors=[str(exc)]
    errors = list(dict.fromkeys(analysis.errors + template_errors))
    safe_name = ''.join(c if c not in '<>:"/\\|?*' else '_' for c in batch_name).strip() or '批次'
    safe_name=safe_name.rstrip('. ')[:100] or '批次'
    stamp = datetime.now().strftime('%Y%m%d_%H%M%S_%f')
    output = Path(output_root) / f'{safe_name}_{stamp}'
    output.mkdir(parents=True, exist_ok=False)
    if errors:
        write_report(analysis, output/'校验报告.xlsx', extra_errors=errors)
        return None, output/'校验报告.xlsx', errors
    try:
        doc = render_notification(template, notification_data(analysis, number or '/', effective_date or date.today().isoformat()),
                                  output/'出口技术通知.docx')
    except Exception as exc:
        errors.append(f'Word 生成失败：{exc}')
        write_report(analysis, output/'校验报告.xlsx', extra_errors=errors)
        return None, output/'校验报告.xlsx', errors
    write_report(analysis, output/'校验报告.xlsx')
    return doc, output/'校验报告.xlsx', []
