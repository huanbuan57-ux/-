import unittest

from export_notice.core import (
    ProductMaster,
    ProductRecord,
    BatchRow,
    normalize_part_no,
    filter_owned_products,
    resolve_field,
    validate_rows,
)


class CoreTests(unittest.TestCase):
    def test_normalize_part_no_removes_repeated_ck_suffixes(self):
        self.assertEqual(normalize_part_no(" s111900000182-ck-ck "), "S111900000182")
        self.assertEqual(normalize_part_no("s111900000294+ck"), "S111900000294")

    def test_filter_owned_products_uses_master_not_ck_suffix(self):
        master = ProductMaster({
            "S111900000182": ProductRecord(standard_code="S111900000182", responsible=True, enabled=True),
        })
        rows = [
            BatchRow("S111900000182-ck", "A", 10, {}, "f.xlsx", "计划", 2),
            BatchRow("L012500003529", "B", 10, {}, "f.xlsx", "计划", 3),
        ]
        owned, excluded = filter_owned_products(rows, master)
        self.assertEqual([r.standard_code for r in owned], ["S111900000182"])
        self.assertEqual(excluded[0]["reason"], "不在本人负责产品资料表")

    def test_excel_value_has_priority_and_missing_station_is_slash(self):
        self.assertEqual(resolve_field(" English from Excel ", "Master English"), "English from Excel")
        self.assertEqual(resolve_field("", "Master English"), "Master English")

    def test_validation_computes_cartons_and_blocks_missing_packaging(self):
        good = BatchRow(
            "S111900000182-ck", "空滤器总成", 120,
            {"英文名称": "Air filter", "包装模数": "24", "单箱数量": "2", "纸箱尺寸": "1120*250*240"},
            "f.xlsx", "计划", 2,
        )
        bad = BatchRow("S111900000294-ck", "谐振腔总成", 120, {}, "f.xlsx", "计划", 3)
        master = ProductMaster({
            "S111900000182": ProductRecord("S111900000182", True, True, english_name="Master", packaging_modulus=24, units_per_carton=2, carton_size="x"),
            "S111900000294": ProductRecord("S111900000294", True, True),
        })
        result = validate_rows([good, bad], master)
        self.assertEqual(result.rows[0].carton_count, 60)
        self.assertTrue(any("包装参数缺失" in e for e in result.errors))


if __name__ == "__main__":
    unittest.main()
