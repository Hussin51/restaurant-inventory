import io, os, shutil, sys, tempfile, unittest
from datetime import date, timedelta
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from openpyxl import load_workbook
from core import Store, norm_currency
import importer, reports, storage

ROOT = os.path.join(os.path.dirname(__file__), "..")


class CurrencyTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.data = storage.Data(self.tmp)
        self.data.ensure(os.path.join(ROOT, "template.xlsx"))
        self.s = Store(self.data.xlsx)
        self.s.set_settings("مطعم", None, 1500)
        self.s.add_supplier("مورد", "07")

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def item(self, code):
        return next(i for i in self.s.items() if i["code"] == code)

    def test_normalize(self):
        self.assertEqual([norm_currency(x) for x in ("", None, "دينار", "IQD", "دولار", "$", "USD")],
                         ["دينار", "دينار", "دينار", "دينار", "دولار", "دولار", "دولار"])
        with self.assertRaises(ValueError):
            norm_currency("يورو", True)

    def test_default_rate_and_rate_setting(self):
        self.assertEqual(self.s.settings()["rate"], 1500)
        self.s.set_settings("مطعم", None, 1480)
        self.assertEqual(self.s.settings()["rate"], 1480)
        for bad in (0, -5, "abc", float("nan")):
            with self.assertRaises(ValueError):
                self.s.set_settings("مطعم", None, bad)
        self.assertEqual(self.s.settings()["rate"], 1480)

    def test_item_price_in_dollars_is_valued_in_dinar(self):
        self.s.add_item("A-1", "دولاري", "", "قطعة", 10, 2, 1, currency="دولار")
        self.s.add_item("A-2", "دينارنا", "", "قطعة", 10, 500, 1)
        a, b = self.item("A-1"), self.item("A-2")
        self.assertEqual((a["cur"], a["price"], a["price_iqd"], a["value"], a["value_usd"]), ("دولار", 2, 3000, 30000, 20))
        self.assertEqual((b["cur"], b["value"], b["value_usd"]), ("دينار", 5000, 0))
        self.s.set_settings("مطعم", None, 1400)
        self.assertEqual(self.item("A-1")["value"], 28000)                      # القيمة تتبع سعر الصرف الحالي
        with self.assertRaises(ValueError):
            self.s.add_item("A-3", "x", "", "قطعة", 1, 1, 1, currency="يورو")

    def test_usd_purchase_snapshots_the_rate(self):
        self.s.add_item("A-1", "مادة", "", "قطعة", 0, 0, 0, carton=10)
        no = self.s.add_invoice(dict(type="شراء", party="مورد", currency="دولار", update_prices=True,
                                     lines=[dict(code="A-1", unit="كارتون", qty=2, price=30)]))
        inv = self.s.all()["invoices"][0]
        self.assertEqual((no, inv["cur"], inv["rate"], inv["total"], inv["total_iqd"]), ("PUR-0001", "دولار", 1500, 60, 90000))
        it = self.item("A-1")
        self.assertEqual((it["cur"], it["price"], it["balance"]), ("دولار", 3, 20))           # آخر سعر شراء بعملة الفاتورة
        self.s.set_settings("مطعم", None, 1000)
        inv = self.s.all()["invoices"][0]
        self.assertEqual((inv["rate"], inv["total_iqd"]), (1500, 90000))                       # الفاتورة القديمة ما تتغير
        self.s.add_invoice(dict(type="شراء", party="مورد", currency="دينار", lines=[dict(code="A-1", qty=10, price=4000)]))
        inv2 = self.s.all()["invoices"][1]
        self.assertEqual((inv2["cur"], inv2["rate"], inv2["total_iqd"]), ("دينار", 1, 40000))

    def test_loans_and_transfers_ignore_currency(self):
        self.s.add_item("A-1", "مادة", "", "قطعة", 10, 1, 1)
        self.s.add_invoice(dict(type="نقل", to="البار", currency="دولار", lines=[dict(code="A-1", qty=1)]))
        self.assertEqual(self.s.all()["invoices"][0]["cur"], "دينار")

    def test_invalid_invoice_currency(self):
        self.s.add_item("A-1", "مادة", "", "قطعة", 10, 1, 1)
        with self.assertRaises(ValueError):
            self.s.add_invoice(dict(type="شراء", party="مورد", currency="يورو", lines=[dict(code="A-1", qty=1, price=1)]))

    def test_migration_of_a_pre_currency_workbook(self):
        wb = load_workbook(self.data.xlsx)
        wb["الأصناف"]["R4"] = None
        wb["الفواتير"]["K4"] = None
        wb["إعدادات"]["A7"] = None
        wb["إعدادات"]["B7"] = None
        wb["الأصناف"]["L5"] = '=IF(B5="","",I5*N(F5))'
        wb.save(self.data.xlsx)
        st = Store(self.data.xlsx)
        self.assertEqual(st.settings()["rate"], 1500)                        # القراءة تشتغل مباشرة بعد الترحيل بالذاكرة
        st.set_settings("مطعم")                                              # أول كتابة تحفظ الترحيل بالملف
        wb = load_workbook(self.data.xlsx)
        self.assertEqual((wb["الأصناف"]["R4"].value, wb["الفواتير"]["K4"].value), ("عملة السعر", "العملة"))
        self.assertIn("$B$7", wb["الأصناف"]["L5"].value)

    def test_reports_show_iqd_with_original_dollars(self):
        self.s.add_item("A-1", "مادة", "", "قطعة", 0, 0, 0)
        self.s.add_invoice(dict(type="شراء", party="مورد", currency="دولار", lines=[dict(code="A-1", qty=10, price=2)]))
        self.s.add_invoice(dict(type="شراء", party="مورد", currency="دينار", lines=[dict(code="A-1", qty=10, price=1000)]))
        d = self.s.all()
        r = reports.movement(d)
        self.assertEqual(r["cards"][1]["value"], 10 * 2 * 1500 + 10 * 1000)
        self.assertIn("20", r["cards"][1]["sub"])
        row = r["tables"][0]["rows"][0]
        self.assertEqual(row, ["مورد", 2, 40000, 10000, 20])
        st = reports.supplier_statement(d, "مورد")
        self.assertEqual([c["value"] for c in st["cards"]][:4], [2, 40000, 10000, 20])
        self.assertEqual(st["tables"][1]["rows"][1][6], 40000)                                   # المجموع التراكمي بالدينار

    def test_import_currency_column(self):
        wb = load_workbook(io.BytesIO(importer.build_template(["أدوات"])))
        ws = wb["المواد"]
        for j, v in enumerate(["U-1", "دولاري", "أدوات", None, None, None, 2, None, None, None, None, 5, None, "دولار"], 1):
            ws.cell(6, j, v)
        for j, v in enumerate(["U-2", "ديناري", "أدوات", None, None, None, 500, None, None, None, None, 5, None, None], 1):
            ws.cell(7, j, v)
        for j, v in enumerate(["U-3", "غلط", "أدوات", None, None, None, 1, None, None, None, None, 5, None, "يورو"], 1):
            ws.cell(8, j, v)
        b = io.BytesIO(); wb.save(b)
        r = importer.parse(b.getvalue())
        self.assertEqual([x["status"] for x in r["rows"]], ["new", "new", "error"])
        self.s.import_items(r["rows"])
        self.assertEqual((self.item("U-1")["cur"], self.item("U-2")["cur"]), ("دولار", "دينار"))
        self.assertEqual(self.item("U-1")["value"], 5 * 2 * 1500)


if __name__ == "__main__":
    unittest.main()


class SupplierCurrencyTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.data = storage.Data(self.tmp)
        self.data.ensure(os.path.join(ROOT, "template.xlsx"))
        self.s = Store(self.data.xlsx)

    def test_supplier_currency_default_and_roundtrip(self):
        self.s.add_supplier("أ", "1")
        self.s.add_supplier("ب", "2", "دولار")
        sups = {x["name"]: x["currency"] for x in self.s.all()["suppliers"]}
        self.assertEqual(sups, {"أ": "دينار", "ب": "دولار"})

    def test_update_currency_and_reject_invalid(self):
        self.s.add_supplier("أ", "1")
        self.s.update_supplier("أ", "5", "دولار")
        self.assertEqual(self.s.all()["suppliers"][0]["currency"], "دولار")
        with self.assertRaises(ValueError):
            self.s.update_supplier("أ", "5", "يورو")
        with self.assertRaises(ValueError):
            self.s.add_supplier("ج", "", "ين")

    def test_delete_keeps_other_supplier_currency(self):
        self.s.add_supplier("أ", "1", "دولار")
        self.s.add_supplier("ب", "2", "دينار")
        self.s.add_supplier("ج", "3", "دولار")
        self.s.delete_supplier("ب")
        self.assertEqual([(x["name"], x["currency"]) for x in self.s.all()["suppliers"]], [("أ", "دولار"), ("ج", "دولار")])

    def test_statement_lists_items_with_currency_and_last_price(self):
        self.s.add_supplier("مورد", "1", "دولار")
        self.s.add_item("A-1", "كأس", "صحون وكاسات", "قطعة", 0, 0, 0)
        self.s.add_invoice(dict(type="شراء", party="مورد", currency="دولار", date="2026-01-05", lines=[dict(code="A-1", unit="قطعة", qty=10, price=2)]))
        self.s.add_invoice(dict(type="شراء", party="مورد", currency="دولار", date="2026-02-05", lines=[dict(code="A-1", unit="قطعة", qty=5, price=3)]))
        import reports
        r = reports.build("supplier", self.s.all(), supplier="مورد")
        t = r["tables"][0]
        self.assertEqual(t["title"], "شنو اشتريت من هذا المورد")
        row = t["rows"][0]
        self.assertEqual(row[0], "A-1")
        self.assertEqual(row[3], 3)                 # آخر سعر
        self.assertEqual(row[4], "دولار")
        self.assertEqual(row[5], 35)                # 10*2 + 5*3 بعملتها
        self.assertIn("عملة التعامل: دولار", r["subtitle"])
