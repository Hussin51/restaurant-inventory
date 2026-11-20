import io, os, shutil, sys, tempfile, unittest
from datetime import date
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from openpyxl import Workbook
from core import Store
import importer, reports, storage

ROOT = os.path.join(os.path.dirname(__file__), "..")


def book(sheets):
    wb = Workbook()
    wb.remove(wb.active)
    for title, rows in sheets:
        ws = wb.create_sheet(title)
        for r in rows:
            ws.append(r)
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


class BreakageTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.data = storage.Data(self.tmp)
        self.data.ensure(os.path.join(ROOT, "template.xlsx"))
        self.s = Store(self.data.xlsx)
        self.s.add_supplier("مورد", "1")
        self.s.add_item("PL-001", "صحن", "صحون وكاسات", "قطعة", 100, 1500, 10)
        self.s.add_item("CP-001", "كأس", "صحون وكاسات", "قطعة", 50, 2, 10, currency="دولار")
        self.s.add_invoice(dict(type="إعارة", to="المطبخ", due="2099-01-01", lines=[dict(code="PL-001", unit="قطعة", qty=20), dict(code="CP-001", unit="قطعة", qty=10)]))
        self.s.add_invoice(dict(type="إعارة", to="البار", due="2099-01-01", lines=[dict(code="CP-001", unit="قطعة", qty=10)]))
        items = self.s.items()
        self.by_code = {i["code"].casefold(): i for i in items}
        self.by_name = {i["name"].casefold(): i for i in items}
        self.depts = ["المطبخ", "البار", "الصالة"]

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_sheet_per_department_and_dept_column(self):
        raw = book([("المطبخ", [["الكود", "الكمية", "ملاحظة"], ["PL-001", 3, "غسيل"], ["CP-001", 0, None]]),
                    ("كل", [["القسم", "المادة", "الكمية"], ["البار", "كأس", 2], ["غير معروف", "كأس", 1], ["البار", "مادة وهمية", 1]])])
        r = importer.parse_breakage(raw, self.by_code, self.by_name, self.depts)
        ok = [(x["dept"], x["code"], x["qty"]) for x in r["lines"] if x["status"] == "ok"]
        self.assertEqual(ok, [("المطبخ", "PL-001", 3), ("البار", "CP-001", 2)])
        self.assertEqual(r["counts"]["error"], 2)                          # قسم مجهول + مادة مجهولة. السطر بكمية صفر ينتخطى بصمت
        self.assertEqual(r["counts"]["qty"], 5)

    def test_rejects_file_without_breakage_columns(self):
        with self.assertRaises(ValueError):
            importer.parse_breakage(book([("ورقة", [["أ", "ب"], [1, 2]])]), self.by_code, self.by_name, self.depts)

    def test_apply_moves_reduce_department_stock_and_report_totals(self):
        n = self.s.add_movements([
            dict(d=date(2026, 11, 28), item="صحن", typ="كسر وضياع", qty=3, frm="المطبخ", note="تكسير شهر 11"),
            dict(d=date(2026, 11, 28), item="كأس", typ="كسر وضياع", qty=2, frm="المطبخ", note="تكسير شهر 11"),
            dict(d=date(2026, 11, 28), item="كأس", typ="كسر وضياع", qty=4, frm="البار", note="تكسير شهر 11"),
        ])
        self.assertEqual(n, 3)
        d = self.s.all()
        it = {i["name"]: i for i in d["items"]}
        self.assertEqual(it["صحن"]["holdings"]["المطبخ"], 17)
        self.assertEqual(it["كأس"]["holdings"]["البار"], 6)
        rep = reports.build("breakage", d, "2026-11-01", "2026-11-30")
        titles = [t["title"] for t in rep["tables"]]
        self.assertEqual(titles[0], "ملخص التكسير حسب القسم")
        self.assertIn("تكسير قسم المطبخ", titles)
        self.assertEqual(rep["cards"][0]["value"], 9)                                       # 3 + 2 + 4 قطعة
        rate = d["settings"]["rate"]
        want = 3 * 1500 + 2 * 2 * rate + 4 * 2 * rate
        self.assertAlmostEqual(rep["cards"][1]["value"], want)
        kitchen = next(t for t in rep["tables"] if t["title"] == "تكسير قسم المطبخ")
        self.assertEqual(kitchen["rows"][0][1], "كأس")                                      # الأعلى قيمة أولاً (2 × 3000 > 3 × 1500)
        self.assertEqual(kitchen["rows"][-1][1], "الإجمالي")
        only = reports.build("breakage", d, "2026-11-01", "2026-11-30", dept="البار")
        self.assertEqual(only["cards"][0]["value"], 4)
        self.assertEqual(len(only["tables"]), 1)

    def test_period_filter_and_empty_report(self):
        self.s.add_movements([dict(d=date(2026, 10, 5), item="صحن", typ="كسر وضياع", qty=1, frm="المطبخ")])
        rep = reports.build("breakage", self.s.all(), "2026-11-01", "2026-11-30")
        self.assertEqual(rep["cards"][0]["value"], 0)
        self.assertEqual(rep["tables"][0]["rows"], [])
        xlsx = reports.build_xlsx(reports.build("breakage", self.s.all()), "مطعم")
        self.assertTrue(xlsx[:2] == b"PK")

    def test_add_movements_rejects_unknown_item_without_partial_write(self):
        before = len(self.s.all()["moves"])
        with self.assertRaises(ValueError):
            self.s.add_movements([dict(d=date(2026, 11, 1), item="صحن", typ="كسر وضياع", qty=1, frm="المطبخ"),
                                  dict(d=date(2026, 11, 1), item="مجهول", typ="كسر وضياع", qty=1, frm="المطبخ")])
        self.assertEqual(len(self.s.all()["moves"]), before)


if __name__ == "__main__":
    unittest.main()
