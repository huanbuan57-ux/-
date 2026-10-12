from __future__ import annotations
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
import re
from typing import Any

def normalize_part_no(raw_code: Any) -> str:
    return re.sub(r'(?:[-+]CK)+$', '', str(raw_code or '').strip().upper()).strip()

def resolve_field(excel_value, master_value, default=None):
    for value in (excel_value, master_value, default):
        if value is not None and str(value).strip():
            return str(value).strip()
    return None

@dataclass
class ProductRecord:
    standard_code: str
    responsible: bool = True
    enabled: bool = True
    chinese_name: str = ''
    english_name: str = ''
    packaging_modulus: int | None = None
    units_per_carton: int | None = None
    carton_size: str = ''
    station_no: str = ''
    weight_note: str = ''

class ProductMaster(dict[str, ProductRecord]):
    pass

@dataclass
class BatchRow:
    raw_code: str
    chinese_name: str
    quantity: Any
    fields: dict[str, Any]
    source_file: str
    source_sheet: str
    source_row: int
    standard_code: str = ''
    carton_count: int | None = None
    resolved: dict[str, Any] = field(default_factory=dict)
    def __post_init__(self):
        self.standard_code = normalize_part_no(self.raw_code)
    @property
    def location(self):
        return f'{self.source_file} / {self.source_sheet} / 第{self.source_row}行 / {self.standard_code}'

@dataclass
class ValidationResult:
    rows: list[BatchRow]
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

def positive_integer(value):
    try:
        number = Decimal(str(value).replace(',', '').strip())
        if number.is_finite() and number > 0 and number == number.to_integral_value():
            return int(number)
    except (InvalidOperation, ValueError, TypeError):
        pass
    return None

def filter_owned_products(rows, master):
    owned, excluded = [], []
    for row in rows:
        record = master.get(row.standard_code)
        if record and record.responsible and record.enabled:
            owned.append(row)
        else:
            excluded.append({'raw_code': row.raw_code, 'standard_code': row.standard_code, 'source_file': row.source_file, 'source_sheet': row.source_sheet, 'source_row': row.source_row, 'reason': '不在本人负责产品资料表'})
    return owned, excluded

def validate_rows(rows, master):
    result = ValidationResult(rows)
    seen = {}
    for row in rows:
        record = master.get(row.standard_code)
        if not record or not record.responsible or not record.enabled:
            result.errors.append(f'{row.location}：不在本人负责产品资料表')
            continue
        for label, attr in [('英文名称', 'english_name'), ('包装模数', 'packaging_modulus'), ('单箱数量', 'units_per_carton'), ('纸箱尺寸', 'carton_size'), ('重量说明', 'weight_note')]:
            raw = row.fields.get(label)
            baseline = getattr(record, attr)
            row.resolved[label] = resolve_field(raw, baseline)
            if raw not in (None, '') and baseline not in (None, '') and str(raw).strip() != str(baseline).strip():
                result.warnings.append(f'{row.location}：{label}使用 Excel 值 {raw}，资料表值 {baseline}')
        row.resolved['工位号'] = resolve_field(row.fields.get('工位号'), None, '/')
        row.resolved['中文名称'] = resolve_field(row.chinese_name, record.chinese_name)
        for label in ['中文名称', '英文名称', '重量说明']:
            if not row.resolved[label]:
                result.errors.append(f'{row.location}：{label}缺失')
        missing = [label for label in ['包装模数', '单箱数量', '纸箱尺寸'] if not row.resolved[label]]
        if missing:
            result.errors.append(f'{row.location}：包装参数缺失（{"、".join(missing)}）')
        qty = positive_integer(row.quantity)
        units = positive_integer(row.resolved['单箱数量'])
        modulus = positive_integer(row.resolved['包装模数'])
        for label, value, number in [('数量', row.quantity, qty), ('单箱数量', row.resolved['单箱数量'], units), ('包装模数', row.resolved['包装模数'], modulus)]:
            if value not in (None, '') and number is None:
                result.errors.append(f'{row.location}：{label}须为正整数（当前值：{value}）')
        if qty is None:
            result.errors.append(f'{row.location}：数量无效')
        row.carton_count = None
        if qty and units:
            if qty % units:
                result.errors.append(f'{row.location}：数量 {qty} 不能被单箱数量 {units} 整除')
            else:
                row.carton_count = qty // units
        fingerprint = (qty, row.resolved, row.fields)
        if row.standard_code in seen and seen[row.standard_code][1] != fingerprint:
            result.errors.append(f'{row.location}：同一图号存在数量或要求冲突；另一来源：{seen[row.standard_code][0]}')
        seen[row.standard_code] = (row.location, fingerprint)
    return result
