import os, shutil, sys, tempfile, unittest
from datetime import date
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from core import Store
import reports

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

    def test_batch_saves_every_line_and_lowers_stock(self):
        n = self.s.add_breakage(date.today(), [{"item": "كاسة", "qty": 2, "unit": "كارتون"}, {"item": "صحن", "qty": 3}], "", "غسيل")
        self.assertEqual(n, 2)
        d = self.s.all()
        bal = {i["name"]: i["balance"] for i in d["items"]}
        self.assertEqual(bal["كاسة"], 100 - 24)
        self.assertEqual(bal["صحن"], 47)
        self.assertEqual([m["type"] for m in d["moves"]], ["كسر وضياع", "كسر وضياع"])

    def test_bad_line_saves_nothing(self):
        with self.assertRaises(ValueError):
            self.s.add_breakage(date.today(), [{"item": "صحن", "qty": 1}, {"item": "ما موجود", "qty": 1}])
        with self.assertRaises(ValueError):
            self.s.add_breakage(date.today(), [])
        with self.assertRaises(ValueError):
            self.s.add_breakage(date.today(), [{"item": "صحن", "qty": 0}])
        self.assertEqual(self.s.all()["moves"], [])

    def test_report_has_prices_and_totals(self):
        self.s.add_breakage(date.today(), [{"item": "كاسة", "qty": 4}, {"item": "صحن", "qty": 2}], "المطبخ", "x")
        r = reports.build("breakage", self.s.all())
        self.assertEqual(r["title"], "تقرير التكسير")
        cards = {c["label"]: c["value"] for c in r["cards"]}
        self.assertEqual(cards["إجمالي القطع المتكسرة"], 6)
        self.assertEqual(cards["قيمة التكسير (دينار)"], 4 * 1000 + 2 * 2500)
        by_item = r["tables"][0]["rows"]
        self.assertEqual(by_item[0][1], "صحن")                       # الأعلى قيمة أول
        self.assertEqual(by_item[0][3], 2500)
        self.assertEqual(by_item[0][4], 5000)
        self.assertEqual(r["tables"][1]["rows"][0][0], "المطبخ")
        self.assertEqual(len(r["tables"][2]["rows"]), 2)
        # الفترة
        self.assertEqual(reports.build("breakage", self.s.all(), "2000-01-01", "2000-01-31")["cards"][1]["value"], 0)
        # الإكسل يشتغل
        self.assertTrue(reports.build_xlsx(r, "مطعم")[:2] == b"PK")


if __name__ == "__main__":
    unittest.main()
