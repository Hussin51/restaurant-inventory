import os, shutil, sys, tempfile, unittest
from datetime import date
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from openpyxl import load_workbook
from core import Store
import clothes, reports, storage

ROOT = os.path.join(os.path.dirname(__file__), "..")
TODAY = date.today().isoformat()


class ClothesTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.data = storage.Data(self.tmp)
        self.data.ensure(os.path.join(ROOT, "template.xlsx"))
        self.s = Store(self.data.xlsx)
        clothes.add_items(self.s, "قميص", "أبيض", [dict(size="M", opening=10, min=3), dict(size="L", opening=5, min=3), dict(size="XL", opening=0, min=2)])
        clothes.add_items(self.s, "بنطلون", "", [dict(size="32", opening=4, min=2), dict(size="30", opening=4, min=2)])

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def cl(self):
        return self.s.all()["clothes"]

    def item(self, name, size):
        return next(i for i in self.cl()["items"] if i["name"] == name and i["size"] == size)

    def test_items_codes_and_sizes_sorted(self):
        items = self.cl()["items"]
        self.assertEqual([i["code"] for i in items if i["name"] == "قميص"], ["CL-001", "CL-002", "CL-003"])
        self.assertEqual([i["size"] for i in items if i["name"] == "قميص"], ["M", "L", "XL"])        # ترتيب المقاسات
        self.assertEqual([i["size"] for i in items if i["name"] == "بنطلون"], ["30", "32"])
        with self.assertRaises(ValueError):
            clothes.add_items(self.s, " قميص ", "أبيض", [dict(size="m")])                                # مكرر بنفس المقاس
        with self.assertRaises(ValueError):
            clothes.add_items(self.s, "", "", [dict(size="M")])

    def test_receive_and_issue_update_balance(self):
        clothes.receive(self.s, "CL-001", 5, TODAY, "طلبية")
        self.assertEqual(self.item("قميص", "M")["balance"], 15)
        no = clothes.issue(self.s, "علي حسن", TODAY, "", [dict(code="CL-001", qty=2), dict(code="CL-002", qty=1)])
        self.assertEqual(no, "CLO-0001")
        self.assertEqual((self.item("قميص", "M")["balance"], self.item("قميص", "L")["balance"]), (13, 4))
        d = self.cl()["docs"][0]
        self.assertEqual((d["employee"], d["total"], len(d["lines"])), ("علي حسن", 3, 2))
        with self.assertRaises(ValueError):
            clothes.issue(self.s, "", TODAY, "", [dict(code="CL-001", qty=1)])
        with self.assertRaises(ValueError):
            clothes.issue(self.s, "س", TODAY, "", [])
        with self.assertRaises(ValueError):
            clothes.issue(self.s, "س", TODAY, "", [dict(code="NOPE", qty=1)])
        with self.assertRaises(ValueError):
            clothes.issue(self.s, "س", TODAY, "", [dict(code="CL-001", qty=0)])

    def test_same_employee_reuses_existing_spelling(self):
        clothes.issue(self.s, "  علي   حسن ", TODAY, "", [dict(code="CL-001", qty=1)])
        clothes.issue(self.s, "علي حسن", TODAY, "موافقة الإدارة", [dict(code="CL-001", qty=1)])
        clothes.issue(self.s, "أحمد", TODAY, "", [dict(code="CL-004", qty=1)])
        clothes.issue(self.s, "احمد", TODAY, "", [dict(code="CL-004", qty=1)])                     # أ/ا نفس الاسم
        emps = {e["name"]: e["count"] for e in self.cl()["employees"]}
        self.assertEqual(emps, {"أحمد": 2, "علي حسن": 2})

    def test_employee_statement(self):
        clothes.issue(self.s, "علي حسن", "2026-01-10", "", [dict(code="CL-001", qty=2), dict(code="CL-004", qty=1)])
        clothes.issue(self.s, "علي حسن", "2026-11-10", "موافقة الإدارة", [dict(code="CL-001", qty=2)])
        clothes.issue(self.s, "سعد", "2026-11-11", "", [dict(code="CL-002", qty=1)])
        r = reports.employee(self.s.all(), "علي   حسن")
        self.assertEqual([c["value"] for c in r["cards"]], [2, 5])
        docs = r["tables"][0]["rows"]
        self.assertEqual([x[1] for x in docs], ["2026-01-10", "2026-11-10"])
        self.assertIn("قميص M أبيض × 2", docs[0][2])
        self.assertEqual(docs[1][4], "موافقة الإدارة")
        summ = {tuple(x[:3]): x[3] for x in r["tables"][1]["rows"]}
        self.assertEqual(summ[("قميص", "M", "أبيض")], 4)
        self.assertIn("أول استلام 2026-01-10", r["subtitle"])
        with self.assertRaises(ValueError):
            reports.employee(self.s.all(), "ما موجود")
        self.assertEqual(reports.employee(self.s.all(), "علي حسن", "2026-11-01", "2026-11-30")["cards"][0]["value"], 1)

    def test_stock_report_and_status(self):
        clothes.issue(self.s, "س", TODAY, "", [dict(code="CL-002", qty=3)])                          # L: 5 -> 2 (الحد 3) ناقص
        self.assertEqual(self.item("قميص", "L")["status"], "ناقص - اطلب")
        self.assertEqual(self.item("قميص", "XL")["status"], "نفد")
        r = reports.build("clothes", self.s.all())
        self.assertEqual(r["cards"][0]["value"], 5)

    def test_delete_doc_restores_stock_and_keeps_others(self):
        a = clothes.issue(self.s, "علي", TODAY, "", [dict(code="CL-001", qty=2)])
        b = clothes.issue(self.s, "سعد", TODAY, "", [dict(code="CL-001", qty=1), dict(code="CL-002", qty=1)])
        c = clothes.issue(self.s, "نور", TODAY, "", [dict(code="CL-004", qty=1)])
        clothes.delete_doc(self.s, b)
        self.assertEqual((self.item("قميص", "M")["balance"], self.item("قميص", "L")["balance"]), (8, 5))
        d = self.cl()
        self.assertEqual([x["no"] for x in d["docs"]], [a, c])
        self.assertEqual([m["doc"] for m in d["moves"]], [a, c])
        ws = load_workbook(self.data.xlsx)["استلامات_الملابس"]
        self.assertIn("A6", ws["E6"].value)
        self.assertIsNone(ws["A7"].value)
        self.assertEqual(clothes.issue(self.s, "جديد", TODAY, "", [dict(code="CL-004", qty=1)]), "CLO-0004")
        with self.assertRaises(ValueError):
            clothes.delete_doc(self.s, "NOPE")

    def test_delete_item_and_movement_rules(self):
        clothes.issue(self.s, "علي", TODAY, "", [dict(code="CL-001", qty=1)])
        with self.assertRaises(ValueError):
            clothes.delete_item(self.s, "CL-001")                                                      # موجود بورقة
        clothes.receive(self.s, "CL-005", 3, TODAY)
        mv = [m for m in self.cl()["moves"] if not m["doc"]][0]
        clothes.delete_move(self.s, mv["row"])
        with self.assertRaises(ValueError):
            clothes.delete_move(self.s, self.cl()["moves"][0]["row"])                                  # تابعة لورقة
        clothes.receive(self.s, "CL-005", 3, TODAY)
        clothes.delete_item(self.s, "CL-005")                                                          # حركته الحرة تنحذف معه
        self.assertTrue(all(i["code"] != "CL-005" for i in self.cl()["items"]))
        self.assertTrue(all(m["code"] != "CL-005" for m in self.cl()["moves"]))

    def test_adjust(self):
        self.assertEqual(clothes.adjust(self.s, "CL-001", 12, TODAY), 1)
        self.assertEqual(self.item("قميص", "M")["balance"], 12)
        self.assertEqual(clothes.adjust(self.s, "CL-001", 12, TODAY), 0)
        self.assertEqual(self.cl()["moves"][-1]["type"], "تسوية زيادة")

    def test_old_workbook_gets_the_clothes_sheets(self):
        wb = load_workbook(self.data.xlsx)
        for n in ("الملابس", "حركة_الملابس", "استلامات_الملابس", "بنود_الاستلام"):
            del wb[n]
        wb.save(self.data.xlsx)
        self.assertEqual(Store(self.data.xlsx).all()["clothes"]["items"], [])                          # القراءة تنشئها بالذاكرة بدون خطأ
        clothes.add_items(Store(self.data.xlsx), "قميص", "", [dict(size="S", opening=1)])
        self.assertIn("الملابس", load_workbook(self.data.xlsx).sheetnames)

    def test_formula_text_neutralised(self):
        clothes.add_items(self.s, "=SUM(1)", "=x", [dict(size="=M", opening=1)])
        it = [i for i in self.cl()["items"] if i["name"].startswith("SUM")][0]
        self.assertEqual((it["name"], it["color"], it["size"]), ("SUM(1)", "x", "M"))


if __name__ == "__main__":
    unittest.main()


class ClothesPurchaseTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.data = storage.Data(self.tmp)
        self.data.ensure(os.path.join(ROOT, "template.xlsx"))
        self.s = Store(self.data.xlsx)
        self.s.set_settings("م", None, 1500)
        self.s.add_supplier("مورد ملابس", "0770")
        clothes.add_items(self.s, "قميص", "أبيض", [dict(size="M", opening=0, min=3), dict(size="L", opening=0, min=3)])

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def cl(self):
        return self.s.all()["clothes"]

    def test_purchase_adds_stock_and_totals(self):
        no = clothes.buy(self.s, "مورد ملابس", TODAY, "طلبية الصيف", "دينار", [dict(code="CL-001", qty=10, price=5000), dict(code="CL-002", qty=6, price=5500)])
        self.assertEqual(no, "CLP-0001")
        p = self.cl()["purchases"][0]
        self.assertEqual((p["party"], p["total"], p["total_iqd"], p["pieces"]), ("مورد ملابس", 83000, 83000, 16))
        self.assertEqual([i["balance"] for i in self.cl()["items"]], [10, 6])
        self.assertEqual(self.cl()["moves"][0]["doc"], "CLP-0001")

    def test_dollar_purchase_snapshots_the_rate(self):
        clothes.buy(self.s, "مورد ملابس", TODAY, "", "دولار", [dict(code="CL-001", qty=10, price=4)])
        self.s.set_settings("م", None, 1000)
        p = self.cl()["purchases"][0]
        self.assertEqual((p["cur"], p["rate"], p["total"], p["total_iqd"]), ("دولار", 1500, 40, 60000))

    def test_validations(self):
        L = [dict(code="CL-001", qty=1, price=1)]
        for kw in (dict(supplier="مجهول", lines=L), dict(supplier="", lines=L), dict(supplier="مورد ملابس", lines=[]),
                   dict(supplier="مورد ملابس", lines=[dict(code="CL-001", qty=0, price=1)]),
                   dict(supplier="مورد ملابس", lines=[dict(code="CL-001", qty=1, price=-1)]),
                   dict(supplier="مورد ملابس", lines=[dict(code="X", qty=1, price=1)]),
                   dict(supplier="مورد ملابس", currency="يورو", lines=L)):
            args = dict(note="", currency="دينار", d=TODAY)
            args.update(kw)
            with self.assertRaises(ValueError):
                clothes.buy(self.s, args["supplier"], args["d"], args["note"], args["currency"], args["lines"])
        self.assertEqual(self.cl()["purchases"], [])

    def test_delete_purchase_restores_stock_and_blocks_item_delete(self):
        a = clothes.buy(self.s, "مورد ملابس", TODAY, "", "دينار", [dict(code="CL-001", qty=10, price=1)])
        b = clothes.buy(self.s, "مورد ملابس", TODAY, "", "دينار", [dict(code="CL-002", qty=4, price=2)])
        with self.assertRaises(ValueError):
            clothes.delete_item(self.s, "CL-001")
        clothes.delete_purchase(self.s, a)
        self.assertEqual([i["balance"] for i in self.cl()["items"]], [0, 4])
        self.assertEqual([p["no"] for p in self.cl()["purchases"]], [b])
        ws = load_workbook(self.data.xlsx)["مشتريات_الملابس"]
        self.assertIn("A5", ws["G5"].value)
        self.assertEqual(clothes.buy(self.s, "مورد ملابس", TODAY, "", "دينار", [dict(code="CL-001", qty=1, price=1)]), "CLP-0003")
        with self.assertRaises(ValueError):
            clothes.delete_purchase(self.s, "NOPE")

    def test_purchases_flow_into_supplier_statement_and_movement_report(self):
        self.s.add_item("A-1", "مادة", "", "قطعة", 0, 0, 0)
        self.s.add_invoice(dict(type="شراء", party="مورد ملابس", lines=[dict(code="A-1", qty=1, price=1000)]))
        clothes.buy(self.s, "مورد ملابس", TODAY, "", "دينار", [dict(code="CL-001", qty=10, price=5000)])
        d = self.s.all()
        st = reports.supplier_statement(d, "مورد ملابس")
        self.assertEqual(st["cards"][0]["value"], 2)
        self.assertEqual(st["cards"][1]["value"], 51000)
        r = reports.movement(d)
        self.assertEqual(r["cards"][1]["value"], 51000)
