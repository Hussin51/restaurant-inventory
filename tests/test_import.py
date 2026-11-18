import io, os, shutil, sys, tempfile, time, unittest
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from openpyxl import load_workbook
from core import Store, CATEGORIES
import importer, storage

ROOT = os.path.join(os.path.dirname(__file__), "..")
HEAD = ["الكود *", "اسم المادة *", "التصنيف", "الاستخدام", "قطع بالكارتون", "قطع بالشد", "سعر القطعة", "الحد الأدنى (قطعة)",
        "الحد الأدنى (كارتون)", "الموجود: كارتون", "الموجود: شد", "الموجود: قطعة", "المورد المعتاد"]


def fill(rows):
    wb = load_workbook(io.BytesIO(importer.build_template(CATEGORIES)))
    ws = wb["المواد"]
    for i, r in enumerate(rows, 6):
        for j, v in enumerate(r, 1):
            c = ws.cell(i, j, v)
            if isinstance(v, str) and v.startswith("="):
                c.data_type = "s"                                          # نص حرفي مو معادلة
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


class ImportTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.data = storage.Data(self.tmp)
        self.data.ensure(os.path.join(ROOT, "template.xlsx"))
        self.s = Store(self.data.xlsx)

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def parse(self, rows):
        its = self.s.items()
        return importer.parse(fill(rows), {i["code"].casefold(): i["name"] for i in its}, {i["name"]: i["code"] for i in its},
                              [x["name"] for x in self.s._read_suppliers(self.s._open())])

    def test_template_headers_and_example_ignored(self):
        wb = load_workbook(io.BytesIO(importer.build_template(CATEGORIES)))
        self.assertEqual([c.value for c in wb["المواد"][4]][:13], HEAD)
        self.assertIn("تعليمات", wb.sheetnames)
        self.assertEqual(self.parse([])["counts"]["total"], 0)               # سطر المثال ما ينحسب

    def test_import_new_items_with_cartons_and_sorting(self):
        r = self.parse([["PL-010", "كاسة ماء", "صحون وكاسات", "ماء", 48, 6, 600, None, 1, 2, 1, 5, "مخازن الرافدين"],
                        ["CT-002", "شوكة", "أدوات مائدة", "أكل", 100, None, 250, 100, None, None, None, 300, None],
                        ["PL-002", "كاسة صوص", "صحون وكاسات", None, None, None, 500, None, None, None, None, 10, None],
                        ["KT-001", "قدر", "معدات مطبخ", None, None, None, 45000, 3, None, None, None, 4, None]])
        self.assertEqual(r["counts"]["new"], 4)
        self.assertEqual(r["counts"]["new_suppliers"], ["مخازن الرافدين"])
        res = self.s.import_items(r["rows"])
        self.assertEqual((res["added"], res["new_suppliers"]), (4, 1))
        items = self.s.items()
        self.assertEqual([i["code"] for i in items], ["PL-002", "PL-010", "CT-002", "KT-001"])      # حسب التصنيف ثم الكود
        w = next(i for i in items if i["code"] == "PL-010")
        self.assertEqual((w["opening"], w["min"], w["carton"], w["shad"], w["supplier"]), (2 * 48 + 1 * 6 + 5 + 0, 48, 48, 6, "مخازن الرافدين"))
        self.assertEqual(w["opening"], 107)
        self.assertEqual(self.s._read_suppliers(self.s._open())[0]["name"], "مخازن الرافدين")

    def test_errors_are_reported_per_row(self):
        r = self.parse([["", "بدون كود", None, None, None, None, None, None, None, None, None, None, None],
                        ["A-1", "", None, None, None, None, None, None, None, None, None, None, None],
                        ["A-2", "سالب", None, None, None, None, -5, None, None, None, None, None, None],
                        ["A-3", "كارتون بدون تعبئة", None, None, None, None, 1, None, None, 2, None, None, None],
                        ["A-4", "نص", None, None, None, None, "abc", None, None, None, None, None, None],
                        ["A-5", "جيد", None, None, None, None, "1,500", None, None, None, None, "٥", None],
                        ["A-5", "مكرر الكود", None, None, None, None, 1, None, None, None, None, None, None],
                        ["A-6", "جيد", None, None, None, None, 1, None, None, None, None, None, None]])
        st = {x["row"]: x for x in r["rows"]}
        self.assertEqual(r["counts"]["error"], 7)
        self.assertEqual(r["counts"]["new"], 1)
        self.assertEqual(st[6]["status"], "error")
        self.assertIn("الكود مطلوب", st[6]["message"])
        self.assertIn("سالب", st[8]["message"])
        self.assertIn("قطع بالكارتون", st[9]["message"])
        self.assertIn("مو رقم", st[10]["message"])
        self.assertEqual((st[11]["status"], st[11]["price"], st[11]["opening"]), ("new", 1500, 5))   # فواصل وأرقام عربية
        self.assertIn("مكرر", st[12]["message"])
        self.assertIn("الاسم مكرر", st[13]["message"])
        res = self.s.import_items(r["rows"])                                  # الأخطاء ما تنستورد
        self.assertEqual(res["added"], 1)

    def test_existing_codes_skip_or_update(self):
        self.s.add_item("PL-001", "صحن", "صحون وكاسات", "قطعة", 100, 1500, 20, usage="قديم")
        rows = [["pl-001", "صحن", "صحون وكاسات", "جديد", None, None, 2000, None, None, None, None, 999, None],
                ["PL-002", "صحن ثاني", None, None, None, None, 1, None, None, None, None, 1, None]]
        r = self.parse(rows)
        self.assertEqual([x["status"] for x in r["rows"]], ["exists", "new"])
        res = self.s.import_items(r["rows"], "skip")
        self.assertEqual((res["added"], res["skipped"], res["updated"]), (1, 1, 0))
        self.assertEqual(next(i for i in self.s.items() if i["code"] == "PL-001")["price"], 1500)
        res = self.s.import_items(self.parse(rows)["rows"], "update")
        it = next(i for i in self.s.items() if i["code"] == "PL-001")
        self.assertEqual((res["updated"], it["price"], it["usage"], it["opening"], it["min"]), (2, 2000, "جديد", 100, 20))   # الرصيد الافتتاحي ما يتغير

    def test_name_used_by_other_code(self):
        self.s.add_item("PL-001", "صحن", "", "قطعة", 1, 1, 1)
        r = self.parse([["PL-009", "صحن", None, None, None, None, 1, None, None, None, None, None, None]])
        self.assertEqual(r["rows"][0]["status"], "error")
        self.assertIn("الاسم مستخدم", r["rows"][0]["message"])

    def test_phone_template_roundtrip(self):
        wb = load_workbook(io.BytesIO(importer.build_phone_template()))
        ws = wb["المواد"]
        self.assertEqual([c.value for c in ws[1]], ["الكود *", "اسم المادة *", "الاستخدام", "قطع بالكارتون", "الموجود: كارتون", "الموجود: قطعة"])
        for j, v in enumerate(["P-1", "كاسة", "ماء", 24, 3, 7], 1):
            ws.cell(3, j, v)
        ws.cell(4, 1, "P-2"); ws.cell(4, 2, "ملعقة"); ws.cell(4, 6, 40)
        b = io.BytesIO(); wb.save(b)
        r = importer.parse(b.getvalue())
        self.assertEqual(r["counts"]["total"], 2)                                  # سطر المثال ما ينحسب
        self.assertEqual([x["opening"] for x in r["rows"]], [3 * 24 + 7, 40])
        self.s.import_items(r["rows"])
        self.assertEqual(self.s.items()[0]["carton"], 24)

    def test_header_aliases_and_missing_headers(self):
        from openpyxl import Workbook
        wb = Workbook(); ws = wb.active
        ws.append(["كود", "الاسم", "السعر", "الكمية"]); ws.append(["X-1", "مادة", 100, 7])
        b = io.BytesIO(); wb.save(b)
        r = importer.parse(b.getvalue())
        self.assertEqual((r["rows"][0]["code"], r["rows"][0]["price"], r["rows"][0]["opening"]), ("X-1", 100, 7))
        wb = Workbook(); wb.active.append(["a", "b"]); b = io.BytesIO(); wb.save(b)
        with self.assertRaises(ValueError):
            importer.parse(b.getvalue())
        with self.assertRaises(ValueError):
            importer.parse(b"not excel")

    def test_formula_text_is_neutralised(self):
        r = self.parse([["F-1", "=1+1", None, "=SUM(1)", None, None, 1, None, None, None, None, None, None]])
        self.s.import_items(r["rows"])
        it = self.s.items()[0]
        self.assertEqual((it["name"], it["usage"]), ("1+1", "SUM(1)"))

    def test_import_1000_items_is_fast_and_limit(self):
        rows = [[f"M-{n:04d}", f"مادة {n}", "أدوات مائدة", None, None, None, 100, None, None, None, None, n, None] for n in range(1, 1001)]
        t = time.time()
        r = self.parse(rows)
        res = self.s.import_items(r["rows"])
        self.assertEqual(res["added"], 1000)
        self.assertLess(time.time() - t, 30)
        self.assertEqual(len(self.s.items()), 1000)
        r2 = self.parse([["M-9999", "زيادة", None, None, None, None, 1, None, None, None, None, None, None]])
        with self.assertRaises(ValueError):
            self.s.import_items(r2["rows"])

    def test_manual_sort(self):
        self.s.add_item("KT-2", "ب", "معدات مطبخ", "قطعة", 1, 1, 1)
        self.s.add_item("PL-10", "ا", "صحون وكاسات", "قطعة", 1, 1, 1)
        self.s.add_item("PL-9", "ج", "صحون وكاسات", "قطعة", 1, 1, 1)
        self.s.set_image("PL-9", "PL-9.png")
        self.s.sort_items()
        its = self.s.items()
        self.assertEqual([i["code"] for i in its], ["PL-9", "PL-10", "KT-2"])
        self.assertEqual(its[0]["image"], "PL-9.png")                         # الصورة تتحرك مع مادتها


if __name__ == "__main__":
    unittest.main()


class ImportImagesTests(unittest.TestCase):
    def test_embedded_images_are_extracted_by_row(self):
        import io
        from PIL import Image
        from openpyxl import load_workbook
        from openpyxl.drawing.image import Image as XImage
        wb = load_workbook(io.BytesIO(importer.build_template(["صحون وكاسات"])))
        ws = wb["المواد"]
        for r, col in ((6, (255, 0, 0)), (8, (0, 0, 255))):
            buf = io.BytesIO(); Image.new("RGB", (20, 20), col).save(buf, "PNG"); buf.seek(0)
            ws.add_image(XImage(buf), f"O{r}")
        out = io.BytesIO(); wb.save(out)
        imgs = importer.extract_images(out.getvalue())
        self.assertEqual(sorted(r for r in imgs if r > 4), [6, 8])      # صف 1 = شعار القالب
        self.assertEqual(imgs[6][1], "image/png")
        self.assertTrue(imgs[6][0].startswith(b"\x89PNG"))

    def test_no_images_is_fine(self):
        self.assertEqual([r for r in importer.extract_images(importer.build_template(["x"])) if r > 4], [])
        self.assertEqual(importer.extract_images(b"not an excel file"), {})


class InvoiceImportTests(unittest.TestCase):
    def make(self, rows, header):
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.append(header)
        for r in rows:
            ws.append(r)
        out = io.BytesIO()
        wb.save(out)
        return out.getvalue()

    def test_template_roundtrip(self):
        r = importer.parse_invoice(importer.invoice_template(), {"pl-001": {"code": "PL-001", "name": "صحن تقديم"}}, {})
        self.assertEqual(r["counts"], {"exists": 1, "new": 0, "error": 0, "total": 1})
        self.assertEqual((r["lines"][0]["qty"], r["lines"][0]["price"], r["lines"][0]["unit"]), (120, 2500, "قطعة"))

    def test_supplier_style_headers_and_statuses(self):
        raw = self.make([["9673", "Board", 10, 38.35], ["NEW-1", "New thing", 3, 2], [None, "No code", 5, 1.5], ["X", "Bad", "abc", 1], [None, None, None, None]],
                        ["كود المادة", "تفاصيل المادة", "الكمية المباعة", "سعر البيع"])
        r = importer.parse_invoice(raw, {"9673": {"code": "9673", "name": "Board"}}, {})
        st = [x["status"] for x in r["lines"]]
        self.assertEqual(st, ["exists", "new", "new", "error"])
        self.assertEqual(r["lines"][2]["code"], "")                 # كود تلقائي بوقت الإنشاء

    def test_serial_number_column_is_not_a_code(self):
        raw = self.make([[1, "A thing", 4, 2.0], [2, "B thing", 2, 3.0]], ["Item No.", "Product Description", "Qty", "U.Price"])
        r = importer.parse_invoice(raw)
        self.assertTrue(all(x["code"] == "" for x in r["lines"]))

    def test_name_match_when_no_code_column(self):
        raw = self.make([["صحن تقديم", 5, 100]], ["المادة", "الكمية", "السعر"])
        r = importer.parse_invoice(raw, {}, {"صحن تقديم": {"code": "PL-001", "name": "صحن تقديم"}})
        self.assertEqual((r["lines"][0]["code"], r["lines"][0]["status"]), ("PL-001", "exists"))

    def test_rejects_file_without_columns(self):
        raw = self.make([["a", "b"]], ["x", "y"])
        with self.assertRaises(ValueError):
            importer.parse_invoice(raw)
