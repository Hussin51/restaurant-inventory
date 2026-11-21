import io, os, sys, unittest
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
import importer, pdfimport

DATA = os.path.join(os.path.dirname(__file__), "data")


def pdf(name):
    with open(os.path.join(DATA, name), "rb") as f:
        return f.read()


class PdfImportTests(unittest.TestCase):
    def test_is_pdf(self):
        self.assertTrue(pdfimport.is_pdf(pdf("invoice_grid.pdf")))
        self.assertFalse(pdfimport.is_pdf(b"PK\x03\x04"))

    def test_grid_invoice_rows_discount_and_currency(self):
        r = pdfimport.extract(pdf("invoice_grid.pdf"), with_images=False)
        self.assertEqual([(x["code"], x["qty"], x["price"]) for x in r["rows"]], [("GR-001", 10, 2.5), ("GR-002", 24, 1.25), ("GR-003", 0, 9.0)])
        self.assertEqual(r["rows"][0]["name"], "Test Plate 20 cm")
        self.assertEqual(r["discount"], 5.0)
        self.assertEqual(r["currency"], "دولار")

    def test_borderless_offer_rows_codes_and_images(self):
        r = pdfimport.extract(pdf("offer_text.pdf"))
        self.assertEqual([(x["code"], x["qty"], x["price"]) for x in r["rows"]], [("AT10-001", 60, 2.0), ("AT10-002", 4, 3.5)])
        self.assertTrue(all(x["image"] and x["image"][:4] == b"\x89PNG" for x in r["rows"]))
        self.assertNotIn("AT10-001", r["rows"][0]["name"])
        self.assertEqual(r["discount"], 10.0)

    def test_invoice_import_from_pdf_skips_zero_quantity_and_reads_meta(self):
        known = {"gr-001": {"code": "GR-001", "name": "Test Plate 20 cm"}}
        r = importer.parse_invoice(pdf("invoice_grid.pdf"), known, {})
        self.assertEqual([(x["code"], x["status"]) for x in r["lines"]], [("GR-001", "exists"), ("GR-002", "new")])   # سطر الكمية صفر انتخطى
        self.assertEqual((r["discount"], r["currency"]), (5.0, "دولار"))

    def test_items_import_from_pdf_with_images(self):
        r = importer.parse(pdf("offer_text.pdf"), {}, {}, [])
        self.assertEqual([(x["code"], x["price"], x["currency"], x["status"]) for x in r["rows"]], [("AT10-001", 2, "دولار", "new"), ("AT10-002", 3.5, "دولار", "new")])
        imgs = importer.extract_images(pdf("offer_text.pdf"))
        self.assertEqual(sorted(imgs), [x["row"] for x in r["rows"]])           # كل مادة لها صورتها بسطرها

    def test_not_a_table_pdf_is_a_clean_error(self):
        body = b"BT /F1 12 Tf 72 700 Td (Just a letter, no table here) Tj ET"
        objs = [b"<< /Type /Catalog /Pages 2 0 R >>", b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
                b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>",
                b"<< /Length %d >>\nstream\n" % len(body) + body + b"\nendstream", b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"]
        out, offs = b"%PDF-1.4\n", []
        for i, o in enumerate(objs, 1):
            offs.append(len(out))
            out += b"%d 0 obj\n" % i + o + b"\nendobj\n"
        xref = len(out)
        out += b"xref\n0 %d\n0000000000 65535 f \n" % (len(objs) + 1) + b"".join(b"%010d 00000 n \n" % o for o in offs)
        out += b"trailer\n<< /Size %d /Root 1 0 R >>\nstartxref\n%d\n%%%%EOF" % (len(objs) + 1, xref)
        with self.assertRaises(ValueError):
            importer.parse_invoice(out, {}, {})

    def test_excel_discount_and_currency_rows_are_detected(self):
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.append(["الكود", "المادة", "الكمية", "السعر"])
        ws.append(["A-1", "مادة", 5, 10])
        ws.append([None, "الخصم", None, 7.5])
        ws.append([None, "العملة", "USD", None])
        out = io.BytesIO()
        wb.save(out)
        r = importer.parse_invoice(out.getvalue(), {}, {})
        self.assertEqual([x["code"] for x in r["lines"]], ["A-1"])
        self.assertEqual((r["discount"], r["currency"]), (7.5, "دولار"))


if __name__ == "__main__":
    unittest.main()
