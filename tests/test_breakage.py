import os, shutil, sys, tempfile, unittest
from datetime import date
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from core import Store
import breakage, reports

TEMPLATE = os.path.join(os.path.dirname(__file__), "..", "template.xlsx")


class BreakageTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        p = os.path.join(self.tmp, "x.xlsx")
        shutil.copy(TEMPLATE, p)
        self.s = Store(p)
        self.s.add_item("A1", "كاسة", "أدوات", "قطعة", 100, 1000, 5, carton=12)
        self.s.add_item("B1", "صحن", "أدوات", "قطعة", 50, 2500, 5)

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def bal(self):
        return {i["name"]: i["balance"] for i in self.s.all()["items"]}

    def test_saves_lines_but_does_not_touch_stock(self):
        n = breakage.add(self.s, date.today(), "المطبخ", "غسيل", [{"item": "كاسة", "qty": 2, "unit": "كارتون"}, {"item": "صحن", "qty": 3}])
        self.assertEqual(n, 2)
        d = self.s.all()
        self.assertEqual(self.bal(), {"كاسة": 100, "صحن": 50})            # الرصيد ما تغيّر
        self.assertEqual(d["moves"], [])                                   # وما تسجل حركة
        self.assertEqual([(b["item"], b["qty"], b["dept"], b["price"], b["value"]) for b in d["breakage"]],
                         [("كاسة", 24, "المطبخ", 1000, 24000), ("صحن", 3, "المطبخ", 2500, 7500)])

    def test_price_is_kept_at_record_time(self):
        breakage.add(self.s, date.today(), "البار", "", [{"item": "صحن", "qty": 2}])
        self.s.update_item("B1", "أدوات", "قطعة", 9000, 5)
        self.assertEqual(self.s.all()["breakage"][0]["price"], 2500)

    def test_bad_input_saves_nothing(self):
        for dept, lines in (("المطبخ", []), ("المطبخ", [{"item": "صحن", "qty": 0}]),
                            ("المطبخ", [{"item": "صحن", "qty": 1}, {"item": "ما موجود", "qty": 1}])):
            with self.assertRaises(ValueError):
                breakage.add(self.s, date.today(), dept, "", lines)
        self.assertEqual(self.s.all()["breakage"], [])

    def test_delete_keeps_the_rest(self):
        breakage.add(self.s, date.today(), "المطبخ", "", [{"item": "كاسة", "qty": 1}, {"item": "صحن", "qty": 2}, {"item": "كاسة", "qty": 3}])
        rows = [b["row"] for b in self.s.all()["breakage"]]
        breakage.delete(self.s, rows[1])
        left = self.s.all()["breakage"]
        self.assertEqual([(b["item"], b["qty"], b["value"]) for b in left], [("كاسة", 1, 1000), ("كاسة", 3, 3000)])
        with self.assertRaises(ValueError):
            breakage.delete(self.s, 99)

    def test_report_has_prices_and_totals(self):
        breakage.add(self.s, date.today(), "", "x", [{"item": "كاسة", "qty": 4}, {"item": "صحن", "qty": 2}])
        breakage.add(self.s, date.today(), "", "", [{"item": "كاسة", "qty": 1}])
        r = reports.build("breakage", self.s.all())
        cards = {c["label"]: c["value"] for c in r["cards"]}
        self.assertEqual(cards["إجمالي القطع المتكسرة"], 7)
        self.assertEqual(cards["قيمة التكسير (دينار)"], 5 * 1000 + 2 * 2500)
        by_item = r["tables"][0]["rows"]
        self.assertEqual((by_item[0][1], by_item[0][3], by_item[0][4]), ("كاسة", 1000, 5000))
        self.assertEqual(len(r["tables"]), 2)
        self.assertEqual(len(r["tables"][1]["rows"]), 3)
        self.assertEqual(reports.build("breakage", self.s.all(), "2000-01-01", "2000-01-31")["cards"][1]["value"], 0)
        self.assertTrue(reports.build_xlsx(r, "مطعم")[:2] == b"PK")

    def test_old_file_without_the_sheet_is_upgraded(self):
        import openpyxl
        wb = openpyxl.load_workbook(self.s.path)
        del wb["تكسير_الأقسام"]
        wb.save(self.s.path)
        self.assertEqual(self.s.all()["breakage"], [])


if __name__ == "__main__":
    unittest.main()
