import os, shutil, sys, tempfile, unittest, zipfile, io
from datetime import date, timedelta
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from core import Store, FileBusy
import storage

ROOT = os.path.join(os.path.dirname(__file__), "..")
TEMPLATE = os.path.join(ROOT, "template.xlsx")
TODAY = date.today()
NEXT_WEEK = (TODAY + timedelta(days=7)).isoformat()


class CoreTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.data = storage.Data(self.tmp)
        self.data.ensure(TEMPLATE)
        self.s = Store(self.data.xlsx)
        self.s.add_item("PL-001", "صحن", "صحون وكاسات", "قطعة", 100, 1500, 20, carton=60, shad=12, usage="تقديم الأكل")
        self.s.add_item("CT-001", "ملعقة", "أدوات مائدة", "قطعة", 50, 250, 10)
        self.s.add_supplier("مورد", "0770")

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def item(self, code):
        return next(i for i in self.s.items() if i["code"] == code)

    def loan(self, code="PL-001", unit="قطعة", qty=1, dept="المطبخ", due=NEXT_WEEK, on=None):
        return self.s.add_invoice(dict(type="إعارة", to=dept, due=due, date=on, lines=[dict(code=code, unit=unit, qty=qty)]))

    def ret(self, code="PL-001", unit="قطعة", qty=1, dept="المطبخ"):
        return self.s.add_invoice(dict(type="إرجاع", frm=dept, lines=[dict(code=code, unit=unit, qty=qty)]))

    # ---------- الأساسيات ----------
    def test_default_name_and_departments(self):
        st = self.s.settings()
        self.assertEqual(st["name"], "مستلزمات مطعم بيت الكل")
        self.assertEqual(st["depts"], ["البار", "المطبخ", "الصالة", "الأراكيل"])

    def test_unique_code_and_name(self):
        with self.assertRaises(ValueError):
            self.s.add_item("pl-001", "ثاني", "", "قطعة", 0, 0, 0)
        with self.assertRaises(ValueError):
            self.s.add_item("X-1", "صحن", "", "قطعة", 0, 0, 0)

    def test_usage_field(self):
        self.assertEqual(self.item("PL-001")["usage"], "تقديم الأكل")
        self.s.update_item("PL-001", "صحون وكاسات", "قطعة", 1500, 20, 60, 12, usage="كاسة شوربة")
        self.assertEqual(self.item("PL-001")["usage"], "كاسة شوربة")

    # ---------- الفواتير ----------
    def test_purchase_in_cartons(self):
        no = self.s.add_invoice(dict(type="شراء", party="مورد", update_prices=True,
                                     lines=[dict(code="PL-001", unit="كارتون", qty=2, price=90000)]))
        self.assertEqual(no, "PUR-0001")
        i = self.item("PL-001")
        self.assertEqual(i["balance"], 220)
        self.assertEqual(i["price"], 1500)
        inv = self.s.all()["invoices"][0]
        self.assertEqual((inv["total"], inv["to"]), (180000, "المخزن"))
        self.assertEqual(inv["lines"][0]["qty"], 120)

    def test_loan_goes_from_warehouse_to_department(self):
        no = self.loan(unit="شد", qty=3)
        self.assertEqual(no, "LON-0001")
        inv = self.s.all()["invoices"][0]
        self.assertEqual((inv["frm"], inv["to"], inv["party"]), ("المخزن", "المطبخ", "المطبخ"))
        self.assertEqual(str(inv["due"]), NEXT_WEEK)
        self.assertEqual(self.item("PL-001")["balance"], 64)

    def test_return_comes_from_department_to_warehouse(self):
        self.loan(unit="شد", qty=3)
        self.ret(unit="شد", qty=1)
        inv = self.s.all()["invoices"][-1]
        self.assertEqual((inv["frm"], inv["to"]), ("المطبخ", "المخزن"))
        self.assertEqual(self.item("PL-001")["balance"], 76)

    def test_transfer_moves_stock_from_warehouse_to_department(self):
        self.s.add_invoice(dict(type="نقل", to="البار", lines=[dict(code="CT-001", qty=5)]))
        it = self.item("CT-001")
        self.assertEqual((it["balance"], it["holdings"], it["held"], it["total"]), (45, {"البار": 5}, 5, 50))
        inv = self.s.all()["invoices"][0]
        self.assertEqual((inv["frm"], inv["to"]), ("المخزن", "البار"))

    def test_stock_model_scenario(self):
        """سيناريو أسبوع كامل: المتبقي بالمخزن وعند الأقسام."""
        s = self.s
        s.add_invoice(dict(type="شراء", party="مورد", lines=[dict(code="PL-001", unit="كارتون", qty=2, price=1)]))       # +120 -> 220
        s.add_invoice(dict(type="نقل", to="البار", lines=[dict(code="PL-001", qty=24)]))                                  # -24
        s.add_invoice(dict(type="نقل", to="المطبخ", lines=[dict(code="PL-001", unit="كارتون", qty=1)]))                   # -60
        s.add_invoice(dict(type="إعارة", due=NEXT_WEEK, to="الصالة", lines=[dict(code="PL-001", qty=12)]))                # -12
        s.add_invoice(dict(type="إرجاع", frm="الصالة", lines=[dict(code="PL-001", qty=8)]))                               # +8
        s.add_movement(TODAY, "صحن", "كسر وضياع", 3, frm="المطبخ")                                                        # المطبخ -3 (المخزن ما يتغير)
        s.add_movement(TODAY, "صحن", "كسر وضياع", 2)                                                                      # المخزن -2
        s.add_movement(TODAY, "ملعقة", "نقل", 10, frm="البار", to="المطبخ")                                              # بين قسمين: المخزن ما يتغير
        it = self.item("PL-001")
        self.assertEqual(it["balance"], 100 + 120 - 24 - 60 - 12 + 8 - 2)
        self.assertEqual(it["holdings"], {"البار": 24, "المطبخ": 57, "الصالة": 4})
        self.assertEqual(it["total"], it["balance"] + 24 + 57 + 4)
        ct = self.item("CT-001")
        self.assertEqual((ct["balance"], ct["holdings"]), (50, {"البار": -10, "المطبخ": 10}))                            # البار ما استلم شي: يطلع سالب (تنبيه)
        mv = {m["type"] + m["to"]: m["wh"] for m in s.movements()}
        self.assertEqual(mv["نقلالمطبخ"] if "نقلالمطبخ" in mv else 0, 0)

    def test_unit_not_defined(self):
        with self.assertRaises(ValueError):
            self.loan(code="CT-001", unit="كارتون")

    def test_validations(self):
        L = [dict(code="PL-001", qty=1)]
        bad = [dict(type="شراء", party="مجهول", lines=L),
               dict(type="نقل", to="مطبخ غلط", lines=L),
               dict(type="نقل", to="المخزن", lines=L),
               dict(type="إعارة", to="المطبخ", lines=L),                                   # بدون تاريخ استرجاع
               dict(type="إعارة", to="المطبخ", due=(TODAY - timedelta(days=1)).isoformat(), lines=L),
               dict(type="إعارة", to="غلط", due=NEXT_WEEK, lines=L),
               dict(type="إرجاع", frm="المخزن", lines=L),
               dict(type="إعارة", to="المطبخ", due=NEXT_WEEK, lines=[dict(code="NOPE", qty=1)]),
               dict(type="إعارة", to="المطبخ", due=NEXT_WEEK, lines=[dict(code="PL-001", qty=0)]),
               dict(type="إعارة", to="المطبخ", due=NEXT_WEEK, lines=[])]
        for inv in bad:
            with self.assertRaises(ValueError, msg=str(inv)):
                self.s.add_invoice(inv)
        self.assertEqual(self.s.all()["invoices"], [])

    # ---------- الإعارات المفتوحة والتنبيه ----------
    def test_open_loans_status_and_fifo_returns(self):
        self.loan(qty=10, dept="المطبخ", due=(TODAY - timedelta(days=2)).isoformat(), on=(TODAY - timedelta(days=9)).isoformat())      # متأخرة
        self.loan(qty=5, dept="المطبخ", due=TODAY.isoformat(), on=(TODAY - timedelta(days=3)).isoformat())   # اليوم
        self.loan(qty=4, dept="البار", due=NEXT_WEEK)                                       # قادمة
        loans = {l["no"]: l for l in self.s.all()["loans"]}
        self.assertEqual(loans["LON-0001"]["status"], "متأخرة")
        self.assertEqual(loans["LON-0001"]["days"], -2)
        self.assertEqual(loans["LON-0002"]["status"], "اليوم")
        self.assertEqual(loans["LON-0003"]["status"], "قادمة")
        self.ret(qty=12, dept="المطبخ")                                                     # يطفي الأولى كلها و2 من الثانية
        loans = {l["no"]: l for l in self.s.all()["loans"]}
        self.assertNotIn("LON-0001", loans)
        self.assertEqual(loans["LON-0002"]["lines"][0]["remaining"], 3)
        self.ret(qty=3, dept="المطبخ")
        self.ret(qty=4, dept="البار")
        self.assertEqual(self.s.all()["loans"], [])

    # ---------- الجرد الفعلي ----------
    def test_set_opening_balance(self):
        self.s.set_opening([dict(code="PL-001", pieces=209), dict(code="CT-001", pieces=300)])
        self.assertEqual(self.item("PL-001")["balance"], 209)
        self.assertEqual(self.item("CT-001")["opening"], 300)
        with self.assertRaises(ValueError):
            self.s.set_opening([dict(code="NOPE", pieces=1)])
        with self.assertRaises(ValueError):
            self.s.set_opening([dict(code="PL-001", pieces=-1)])

    def test_adjustment_creates_movements(self):
        self.loan(qty=10)                                              # الرصيد 90
        n = self.s.adjust([dict(code="PL-001", pieces=95), dict(code="CT-001", pieces=50)])   # +5، وملعقة بدون فرق
        self.assertEqual(n, 1)
        self.assertEqual(self.item("PL-001")["balance"], 95)
        m = self.s.movements()[-1]
        self.assertEqual((m["type"], m["qty"]), ("تسوية زيادة", 5))
        self.s.adjust([dict(code="PL-001", pieces=90)])
        self.assertEqual(self.s.movements()[-1]["type"], "تسوية نقص")
        self.assertEqual(self.item("PL-001")["balance"], 90)

    def test_departments_setting(self):
        self.s.set_settings("مطعم", "البار، المطبخ، الكاشير")
        self.assertEqual(self.s.settings()["depts"], ["البار", "المطبخ", "الكاشير"])
        self.s.add_invoice(dict(type="نقل", to="الكاشير", lines=[dict(code="CT-001", qty=1)]))
        with self.assertRaises(ValueError):
            self.s.add_invoice(dict(type="نقل", to="الأراكيل", lines=[dict(code="CT-001", qty=1)]))

    # ---------- حركة سريعة ----------
    def test_quick_movement_units(self):
        self.s.add_movement(TODAY, "صحن", "كسر وضياع", 1, unit="شد")
        self.assertEqual(self.item("PL-001")["balance"], 88)

    # ---------- النسخ الاحتياطي ----------
    def test_backup_restore_roundtrip(self):
        self.s.add_movement(TODAY, "ملعقة", "شراء", 10)
        z = self.data.backup_bytes()
        self.s.add_movement(TODAY, "ملعقة", "شراء", 100)
        self.assertEqual(self.item("CT-001")["balance"], 160)
        self.data.restore_bytes(z)
        self.assertEqual(self.item("CT-001")["balance"], 60)
        self.assertTrue(any(f.startswith("before-restore") for f in os.listdir(self.data.backups)))

    def test_cloud_backup(self):
        cloud = tempfile.mkdtemp()
        try:
            self.assertEqual(self.data.cloud_backup(), None)
            self.data.set_cloud(cloud)
            target = os.path.join(cloud, storage.CLOUD_SUB)
            names = os.listdir(target)
            self.assertIn("جرد_المطعم_آخر_نسخة.zip", names)
            self.assertEqual(len(names), 2)
            self.s.add_movement(TODAY, "ملعقة", "شراء", 10)
            self.data.cloud_backup()
            with open(os.path.join(target, "جرد_المطعم_آخر_نسخة.zip"), "rb") as fh:
                z = fh.read()
            self.data.restore_bytes(z)
            self.assertEqual(self.item("CT-001")["balance"], 60)
            with self.assertRaises(ValueError):
                self.data.set_cloud(os.path.join(cloud, "no-such"))
            self.data.set_cloud("")
            self.assertEqual(self.data.cloud_dir(), "")
        finally:
            shutil.rmtree(cloud, ignore_errors=True)

    def test_rejects_bad_numbers(self):
        for kw in (dict(opening=-1), dict(price=-5), dict(minimum=-2), dict(opening=float("nan")), dict(opening="abc")):
            args = dict(code="Z-1", name="Z", cat="", unit="قطعة", opening=0, price=0, minimum=0)
            args.update(kw)
            with self.assertRaises(ValueError):
                self.s.add_item(**args)
        with self.assertRaises(ValueError):
            self.s.add_movement(TODAY, "صحن", "شراء", -3)
        with self.assertRaises(ValueError):
            self.s.add_movement(TODAY, "صحن", "شراء", float("inf"))
        self.assertEqual(self.item("PL-001")["balance"], 100)

    def test_text_starting_with_equals_is_not_a_formula(self):
        self.s.add_item("E-1", "=1+1", "", "قطعة", 0, 0, 0, usage="=SUM(1)")
        self.assertEqual(self.item("E-1")["name"], "1+1")
        self.assertFalse(self.item("E-1")["usage"].startswith("="))
        self.s.add_movement(TODAY, "صحن", "كسر وضياع", 1, note="=HYPERLINK(\"x\")")
        self.assertFalse(self.s.movements()[-1]["note"].startswith("="))

    def test_save_is_atomic_when_file_locked(self):
        import core
        with open(self.data.xlsx, "rb") as fh:
            before = fh.read()
        real = os.replace

        def locked(a, b):
            raise PermissionError("locked")
        core.os.replace = locked
        try:
            with self.assertRaises(FileBusy):
                self.s.add_item("L-1", "مقفول", "", "قطعة", 0, 0, 0)
        finally:
            core.os.replace = real
        with open(self.data.xlsx, "rb") as fh:
            self.assertEqual(fh.read(), before)
        self.assertFalse(os.path.exists(self.data.xlsx + ".tmp"))

    def test_logo_survives_saving(self):
        with zipfile.ZipFile(self.data.xlsx) as z:
            self.assertTrue(any(n.startswith("xl/media/") for n in z.namelist()))

    def test_app_window_browser_lookup(self):
        import app
        old = os.environ.get("INVENTORY_BROWSER")
        try:
            os.environ["INVENTORY_BROWSER"] = os.path.join(self.tmp, "nope.exe")
            self.assertIsNone(app.find_app_browser())
            os.environ["INVENTORY_BROWSER"] = sys.executable
            self.assertEqual(app.find_app_browser(), sys.executable)
        finally:
            if old is None:
                os.environ.pop("INVENTORY_BROWSER", None)
            else:
                os.environ["INVENTORY_BROWSER"] = old

    def test_image_name_collision(self):
        import app
        self.assertNotEqual(app.safe_code("A/B"), app.safe_code("A_B"))
        self.assertEqual(app.safe_code("PL-001"), "PL-001")

    def test_restore_rejects_garbage(self):
        with self.assertRaises(ValueError):
            self.data.restore_bytes(b"not a zip")
        b = io.BytesIO()
        with zipfile.ZipFile(b, "w") as z:
            z.writestr("x.txt", "hi")
        with self.assertRaises(ValueError):
            self.data.restore_bytes(b.getvalue())


if __name__ == "__main__":
    unittest.main()
