import unittest
from export_notice.core import BatchRow, ProductMaster, ProductRecord
from export_notice.service import analyze_rows

class ServiceTests(unittest.TestCase):
    def master(self):
        return ProductMaster({'A':ProductRecord('A',chinese_name='产品',english_name='Product',
                  packaging_modulus=12,units_per_carton=2,carton_size='100*200*300',weight_note='实际称重')})

    def test_batch_modulus_overrides_master_station_remains_slash(self):
        row=BatchRow('A-ck','产品',120,{},'a.xlsx','计划',2)
        batch=analyze_rows([row],self.master(),
             ['合同号：ABC 车型：C1 国家：/ LOGO类型：/ 包装模数：24 标签网址：https://example.com'])
        self.assertEqual(batch.rows[0].resolved['包装模数'],'24')
        self.assertEqual(batch.rows[0].resolved['工位号'],'/')
        self.assertEqual(batch.rows[0].carton_count,60)
        self.assertEqual(batch.errors,[])

    def test_identical_rows_deduplicated_with_trace(self):
        rows=[BatchRow('A-ck','产品',120,{},'a.xlsx','计划',2),
              BatchRow('A+ck','产品',120,{},'b.xlsx','计划',3)]
        batch=analyze_rows(rows,self.master(),['合同号：ABC 车型：C1 国家：/ LOGO类型：/ 标签网址：https://example.com'])
        self.assertEqual(len(batch.rows),1)
        self.assertEqual(len(batch.all_rows),2)
        self.assertTrue(batch.warnings)

    def test_nondivisible_and_duplicate_conflict_block(self):
        rows=[BatchRow('A','产品',121,{},'a.xlsx','计划',2),
              BatchRow('A','产品',120,{},'b.xlsx','计划',3)]
        batch=analyze_rows(rows,self.master(),['合同号：ABC 车型：C1 国家：/ LOGO类型：/ 标签网址：https://example.com'])
        self.assertTrue(any('整除' in x for x in batch.errors))
        self.assertTrue(any('冲突' in x for x in batch.errors))
