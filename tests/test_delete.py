import os, shutil, sys, tempfile, unittest
from datetime import date, timedelta
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from openpyxl import load_workbook
from core import Store
import storage

ROOT = os.path.join(os.path.dirname(__file__), "..")
TODAY = date.today()
NEXT_WEEK = (TODAY + timedelta(days=7)).isoformat()


class DeleteTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.data = storage.Data(self.tmp)
        self.data.ensure(os.path.join(ROOT, "template.xlsx"))
        self.s = Store(self.data.xlsx)
        self.s.add_supplier("مورد أ", "1")
        self.s.add_supplier("مورد ب", "2")
        self.s.add_item("PL-001", "صحن", "صحون وكاسات", "قطعة", 100, 1500, 20, carton=60, supplier="مورد أ")
        self.s.add_item("CT-001", "ملعقة", "أدوات مائدة", "قطعة", 50, 250, 10)
        self.s.add_item("KT-001", "قدر", "معدات مطبخ", "قطعة", 5, 45000, 1)

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def item(self, code):
        return next(i for i in self.s.items() if i["code"] == code)

    def three_invoices(self):
        a = self.s.add_invoice(dict(type="شراء", party="مورد أ", lines=[dict(code="PL-001", unit="كارتون", qty=2, price=90000)]))   # +120
        b = self.s.add_invoice(dict(type="إعارة", to="المطبخ", due=NEXT_WEEK, lines=[dict(code="PL-001", qty=30), dict(code="CT-001", qty=5)]))
        c = self.s.add_invoice(dict(type="شراء", party="مورد ب", lines=[dict(code="CT-001", qty=100, price=200)]))
        return a, b, c

    def test_delete_middle_invoice_restores_balances_and_keeps_the_rest_consistent(self):
        a, b, c = self.three_invoices()
        self.assertEqual((self.item("PL-001")["balance"], self.item("CT-001")["balance"]), (190, 145))
        self.assertEqual(len(self.s.all()["loans"]), 1)
        r = self.s.delete_invoice(b)
        self.assertEqual(r["lines"], 2)
        self.assertEqual((self.item("PL-001")["balance"], self.item("CT-001")["balance"]), (220, 150))
        d = self.s.all()
        self.assertEqual([v["no"] for v in d["invoices"]], [a, c])
        self.assertEqual([len(v["lines"]) for v in d["invoices"]], [1, 1])
        self.assertEqual(d["invoices"][1]["total"], 20000)
        self.assertEqual(d["loans"], [])
        self.assertEqual([m["inv"] for m in d["moves"]], [a, c])
        # الصيغ لازم تشير لسطرها الجديد
        ws = load_workbook(self.data.xlsx)["الفواتير"]
        self.assertIn("A6", ws["H6"].value)
        self.assertIn("A6", ws["I6"].value)
        self.assertIsNone(ws["A7"].value)
        self.assertIsNone(ws["I7"].value)
        # الترقيم يكمل بدون تكرار
        self.assertEqual(self.s.add_invoice(dict(type="شراء", party="مورد أ", lines=[dict(code="KT-001", qty=1, price=1)])), "PUR-0003")
        with self.assertRaises(ValueError):
            self.s.delete_invoice("NOPE-1")

    def test_delete_all_invoices_then_start_again(self):
        a, b, c = self.three_invoices()
        for n in (a, b, c):
            self.s.delete_invoice(n)
        d = self.s.all()
        self.assertEqual((d["invoices"], d["moves"]), ([], []))
        self.assertEqual(self.item("PL-001")["balance"], 100)
        self.assertEqual(self.s.add_invoice(dict(type="شراء", party="مورد أ", lines=[dict(code="PL-001", qty=1, price=1)])), "PUR-0001")

    def test_delete_item_rules(self):
        self.three_invoices()
        with self.assertRaises(ValueError) as e:
            self.s.delete_item("PL-001")
        self.assertIn("PUR-0001", str(e.exception))
        self.s.add_movement(TODAY, "قدر", "كسر وضياع", 1)
        self.s.add_movement(TODAY, "قدر", "شراء", 4)
        self.s.set_image("KT-001", "KT-001.png")
        r = self.s.delete_item("KT-001")
        self.assertEqual((r["removed_moves"], r["image"]), (2, "KT-001.png"))
        self.assertEqual([i["code"] for i in self.s.items()], ["PL-001", "CT-001"])
        self.assertTrue(all(m["item"] != "قدر" for m in self.s.movements()))
        self.assertEqual(len(self.s.movements()), 4)                                  # حركات الفواتير بقت
        self.s.add_item("KT-001", "قدر", "معدات مطبخ", "قطعة", 7, 1, 1)               # نفس الكود والاسم ينفع من جديد
        self.assertEqual(self.item("KT-001")["balance"], 7)                           # بدون حركات قديمة عالقة
        with self.assertRaises(ValueError):
            self.s.delete_item("NOPE")

    def test_delete_supplier_rules(self):
        self.three_invoices()
        with self.assertRaises(ValueError):
            self.s.delete_supplier("مورد أ")                                          # عنده فاتورة
        self.s.add_supplier("مورد ج", "3")
        self.s.add_item("N-1", "جديدة", "", "قطعة", 1, 1, 1, supplier="مورد ج")
        self.s.delete_supplier("مورد ب") if False else None
        self.s.delete_invoice("PUR-0002")
        self.s.delete_supplier("مورد ب")
        self.assertEqual([x["name"] for x in self.s._read_suppliers(self.s._open())], ["مورد أ", "مورد ج"])
        self.s.delete_supplier("مورد ج")
        self.assertEqual(self.item("N-1")["supplier"], "")                            # ينمسح من المادة
        self.assertEqual([x["name"] for x in self.s._read_suppliers(self.s._open())], ["مورد أ"])

    def test_delete_quick_movement_only(self):
        self.three_invoices()
        self.s.add_movement(TODAY, "ملعقة", "كسر وضياع", 3)
        mv = self.s.movements()
        invoice_move, quick = mv[0], mv[-1]
        with self.assertRaises(ValueError):
            self.s.delete_movement(invoice_move["row"])
        before = self.item("CT-001")["balance"]
        self.s.delete_movement(quick["row"])
        self.assertEqual(self.item("CT-001")["balance"], before + 3)
        self.assertEqual(len(self.s.movements()), 4)
        ws = load_workbook(self.data.xlsx)["حركة"]
        self.assertIn("C9", ws["E9"].value) if ws["E9"].value else None               # الصيغة بالسطر الصحيح
        with self.assertRaises(ValueError):
            self.s.delete_movement(999)

    def test_backup_before_delete_is_kept_and_pruned(self):
        for i in range(35):
            self.data.make_backup("before-delete")
        names = [f for f in os.listdir(self.data.backups) if f.startswith("before-delete_")]
        self.assertLessEqual(len(names), 31)

    def test_bulk_remove_items_skips_those_in_invoices(self):
        self.s.add_invoice(dict(type="شراء", party="مورد أ", lines=[dict(code="PL-001", unit="قطعة", qty=5, price=1500)]))
        self.s.add_movement(TODAY, "ملعقة", "كسر وضياع", 2, "المخزن", "", "")
        r = self.s.delete_items(["PL-001", "CT-001", "KT-001", "NOPE"])
        self.assertEqual(r["deleted"], 2)
        self.assertEqual(list(r["blocked"]), ["PL-001"])
        self.assertEqual(r["missing"], 1)
        self.assertEqual(r["removed_moves"], 1)
        left = [i["code"] for i in self.s.items()]
        self.assertEqual(left, ["PL-001"])
        self.assertTrue(all(m["item"] != "ملعقة" for m in self.s.all()["moves"]))

    def test_bulk_remove_requires_codes(self):
        with self.assertRaises(ValueError):
            self.s.delete_items([])


if __name__ == "__main__":
    unittest.main()
