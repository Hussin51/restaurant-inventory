import io, os, shutil, sys, tempfile, unittest
from datetime import date, timedelta
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from core import Store
import reports, storage

ROOT = os.path.join(os.path.dirname(__file__), "..")
TODAY = date.today()
d = lambda n: (TODAY - timedelta(days=n)).isoformat()


class ReportTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.mkdtemp()
        data = storage.Data(cls.tmp)
        data.ensure(os.path.join(ROOT, "template.xlsx"))
        s = Store(data.xlsx)
        s.add_supplier("مخازن الرافدين", "0770")
        s.add_supplier("شركة النور", "0780")
        s.add_item("PL-001", "صحن", "صحون وكاسات", "قطعة", 100, 1500, 60, carton=60, shad=12, usage="تقديم", supplier="مخازن الرافدين")
        s.add_item("CT-001", "ملعقة", "أدوات مائدة", "قطعة", 300, 500, 100, carton=200)
        s.add_item("KT-001", "قدر", "معدات مطبخ", "قطعة", 4, 45000, 3, supplier="شركة النور")
        s.add_invoice(dict(type="شراء", date=d(20), party="مخازن الرافدين", lines=[dict(code="PL-001", unit="كارتون", qty=2, price=90000)]))
        s.add_invoice(dict(type="شراء", date=d(3), party="شركة النور", lines=[dict(code="KT-001", qty=2, price=44000)]))
        s.add_invoice(dict(type="إعارة", date=d(10), due=d(2), to="المطبخ", lines=[dict(code="PL-001", unit="شد", qty=5)]))
        s.add_invoice(dict(type="نقل", date=d(4), to="البار", lines=[dict(code="CT-001", qty=50)]))
        s.add_invoice(dict(type="إرجاع", date=d(1), frm="المطبخ", lines=[dict(code="PL-001", qty=20)]))
        s.add_movement(date.fromisoformat(d(2)), "صحن", "كسر وضياع", 6, frm="المطبخ", note="غسيل")
        cls.data = s.all()

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.tmp, ignore_errors=True)

    def table(self, rep, title):
        return next(t for t in rep["tables"] if t["title"] == title)

    def test_movement_report_all_time(self):
        r = reports.movement(self.data)
        self.assertEqual(r["cards"][0]["value"], 2)                          # فاتورتي شراء
        self.assertEqual(r["cards"][1]["value"], 180000 + 88000)             # قيمة المشتريات
        self.assertEqual(self.table(r, "المشتريات حسب المورد")["rows"][0][0], "مخازن الرافدين")
        types = {row[0]: row for row in self.table(r, "الإخراج والإرجاع حسب النوع")["rows"]}
        self.assertEqual(types["إعارة"][2], 60)                              # 5 شد × 12
        self.assertEqual(types["كسر وضياع"][2], 6)
        self.assertEqual(types["كسر وضياع"][3], 6 * 1500)
        depts = {row[0]: row for row in self.table(r, "حركة الأقسام")["rows"]}
        self.assertEqual(depts["المطبخ"][1:], [60, 0, 20, 6])                # إعارة، نقل، رجّع، كسر
        self.assertEqual(depts["البار"][1:], [0, 50, 0, 0])
        self.assertEqual(self.table(r, "الكسر والضياع بالتفصيل")["rows"][0][3], "المطبخ")

    def test_period_filter(self):
        r = reports.movement(self.data, d(5), d(0))                          # آخر 5 أيام: الشراء الثاني فقط
        self.assertEqual(r["cards"][0]["value"], 1)
        self.assertEqual(r["subtitle"], f"من {d(5)} إلى {d(0)}")
        r2 = reports.movement(self.data, "2000-01-01", "2000-01-02")
        self.assertEqual(r2["cards"][0]["value"], 0)

    def test_department_report(self):
        r = reports.department(self.data, "المطبخ")
        row = next(x for x in self.table(r, "المواد")["rows"] if x[0] == "PL-001")
        self.assertEqual(row[2:6], [60, 20, 6, 34])                          # استلم 60، رجّع 20، كسر 6، عنده 60-20-6
        loans = self.table(r, "الإعارات المفتوحة عند القسم")["rows"]
        self.assertEqual(len(loans), 1)
        self.assertIn("متأخرة", loans[0][2])
        with self.assertRaises(ValueError):
            reports.department(self.data, "")

    def test_supplier_statement(self):
        r = reports.supplier_statement(self.data, "مخازن الرافدين")
        rows = self.table(r, "الفواتير")["rows"]
        self.assertEqual((len(rows), rows[0][3], rows[0][4], rows[0][5], rows[0][6]), (1, "دينار", 180000, 180000, 180000))
        self.assertEqual(r["cards"][1]["value"], 180000)
        self.assertIn("0770", r["subtitle"])
        with self.assertRaises(ValueError):
            reports.supplier_statement(self.data, "مجهول")

    def test_reorder_list_groups_by_supplier_and_rounds_to_cartons(self):
        # صحن: رصيد 100+120-60+20-6 = 174 > 60 فما يدخل. ملعقة: 300-50 = 250 فوق الحد. قدر: 4+2 = 6 فوق 3
        self.assertEqual(reports.reorder(self.data)["cards"][0]["value"], 0)
        import copy
        data = copy.deepcopy(self.data)
        for i in data["items"]:
            if i["code"] == "PL-001":
                i["balance"] = 30                                            # تحت الحد 60: المطلوب 120-30 = 90 قطعة = كارتونين
            if i["code"] == "KT-001":
                i["balance"] = 1                                             # حد 3 × 2 = 6 - 1 = 5 قطع (بدون كارتون)
        r = reports.reorder(data)
        self.assertEqual(r["cards"][0]["value"], 2)
        t1 = self.table(r, "المورد: مخازن الرافدين (0770)")
        self.assertEqual(t1["rows"][0][5], "2 كارتون (120 قطعة)")
        self.assertEqual(t1["rows"][0][6], 120 * 1500)
        t2 = self.table(r, "المورد: شركة النور (0780)")
        self.assertEqual(t2["rows"][0][5], "5 قطعة")

    def test_location_report(self):
        r = reports.location(self.data)
        t = r["tables"][0]
        self.assertEqual(t["columns"][:3], ["الكود", "المادة", "بالمخزن"])
        row = next(x for x in t["rows"] if x[0] == "PL-001")
        cols = t["columns"]
        self.assertEqual(row[cols.index("المطبخ")], 34)
        self.assertEqual(row[-1], row[2] + sum(row[3:-1]))                    # الإجمالي = المخزن + الأقسام
        self.assertEqual(reports.build("location", self.data)["kind"], "location")

    def test_packs_text(self):
        it = dict(carton=60, shad=12)
        self.assertEqual(packs_text := reports.packs_text(it, 209), "3 كارتون و2 شد و5 قطعة")
        self.assertEqual(reports.packs_text(dict(carton=0, shad=0), 5), "")
        self.assertEqual(reports.packs_text(it, 0), "0 قطعة")

    def test_excel_export_opens(self):
        from openpyxl import load_workbook
        for kind, kw in (("movement", {}), ("dept", dict(dept="المطبخ")), ("supplier", dict(supplier="شركة النور")), ("reorder", {})):
            rep = reports.build(kind, self.data, **kw)
            wb = load_workbook(io.BytesIO(reports.build_xlsx(rep, "مستلزمات مطعم بيت الكل", os.path.join(ROOT, "assets", "logo-xl.png"))))
            ws = wb["التقرير"]
            self.assertEqual(ws["A1"].value, "مستلزمات مطعم بيت الكل")
            self.assertEqual(ws["A2"].value, rep["title"])
            self.assertEqual(len(ws._images), 1)

    def test_unknown_kind(self):
        with self.assertRaises(ValueError):
            reports.build("nope", self.data)


if __name__ == "__main__":
    unittest.main()
