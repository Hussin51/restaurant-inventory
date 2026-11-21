import os, shutil, sys, tempfile, unittest
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from core import Store
import reports, storage

ROOT = os.path.join(os.path.dirname(__file__), "..")


class DiscountTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.data = storage.Data(self.tmp)
        self.data.ensure(os.path.join(ROOT, "template.xlsx"))
        self.s = Store(self.data.xlsx)
        self.s.add_supplier("مورد", "1", "دولار")
        self.s.add_item("A-1", "صحن", "صحون وكاسات", "قطعة", 0, 0, 0)
        self.s.add_item("B-1", "كأس", "صحون وكاسات", "قطعة", 0, 0, 0)

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def buy(self, disc=0, cur="دولار"):
        return self.s.add_invoice(dict(type="شراء", party="مورد", currency=cur, discount=disc, date="2026-11-02",
                                       lines=[dict(code="A-1", unit="قطعة", qty=10, price=10), dict(code="B-1", unit="قطعة", qty=20, price=5)]))

    def test_discount_reduces_net_total_and_keeps_gross(self):
        no = self.buy(40)
        v = next(x for x in self.s.all()["invoices"] if x["no"] == no)
        self.assertEqual((v["gross"], v["discount"], v["total"]), (200, 40, 160))
        self.assertAlmostEqual(v["total_iqd"], 160 * v["rate"])
        self.assertAlmostEqual(v["lines"][0]["net"], 80)                    # 100 × (160/200)

    def test_no_discount_is_unchanged(self):
        no = self.buy()
        v = next(x for x in self.s.all()["invoices"] if x["no"] == no)
        self.assertEqual((v["gross"], v["discount"], v["total"]), (200, 0, 200))

    def test_discount_validation(self):
        with self.assertRaises(ValueError):
            self.buy(999)
        with self.assertRaises(ValueError):
            self.buy(-1)

    def test_supplier_statement_uses_net_and_shows_discount(self):
        self.buy(40)
        d = self.s.all()
        st = reports.supplier_statement(d, "مورد")
        self.assertEqual(st["cards"][3]["value"], 160)                       # دولار أصلي بعد الخصم
        inv = st["tables"][1]
        self.assertEqual(inv["rows"][0][4], 160)
        self.assertEqual(inv["rows"][0][8], 40)
        items = st["tables"][0]["rows"]
        self.assertAlmostEqual(sum(r[5] for r in items), 160)                # الخصم موزّع على المواد
        mv = reports.movement(d, "", "")
        self.assertAlmostEqual(mv["cards"][1]["value"], 160 * d["settings"]["rate"])

    def test_delete_invoice_keeps_other_discounts(self):
        a, b = self.buy(10), self.buy(30)
        self.s.delete_invoice(a)
        left = self.s.all()["invoices"]
        self.assertEqual([(x["no"], x["discount"], x["total"]) for x in left], [(b, 30, 170)])

    def test_loans_have_no_discount(self):
        self.buy()
        self.s.add_invoice(dict(type="إعارة", to="المطبخ", due="2099-01-01", lines=[dict(code="A-1", unit="قطعة", qty=2)]))
        loan = next(x for x in self.s.all()["invoices"] if x["type"] == "إعارة")
        self.assertEqual(loan["discount"], 0)


if __name__ == "__main__":
    unittest.main()
