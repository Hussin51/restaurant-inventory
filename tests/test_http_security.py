import json, os, shutil, sys, tempfile, threading, unittest, urllib.error, urllib.request
from http.server import ThreadingHTTPServer
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
import app, storage
from core import Store

ROOT = os.path.join(os.path.dirname(__file__), "..")


class HttpSecurityTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.mkdtemp()
        app.DATA = storage.Data(cls.tmp)
        app.DATA.ensure(os.path.join(ROOT, "template.xlsx"))
        app.STORE, app.IMG_DIR = Store(app.DATA.xlsx), app.DATA.images
        cls.srv = ThreadingHTTPServer(("127.0.0.1", 0), app.H)
        app.PORT = cls.port = cls.srv.server_address[1]
        threading.Thread(target=cls.srv.serve_forever, daemon=True).start()
        os.environ.pop("INVENTORY_DEV", None)

    @classmethod
    def tearDownClass(cls):
        cls.srv.shutdown()
        shutil.rmtree(cls.tmp, ignore_errors=True)

    def req(self, path, body=None, headers=None):
        h = {"Content-Type": "application/json"}
        h.update(headers or {})
        r = urllib.request.Request(f"http://127.0.0.1:{self.port}{path}", data=json.dumps(body).encode() if body is not None else None, headers=h)
        try:
            with urllib.request.urlopen(r, timeout=30) as x:
                return x.status
        except urllib.error.HTTPError as e:
            return e.code

    def test_page_and_api_work_for_the_app_itself(self):
        self.assertEqual(self.req("/"), 200)
        self.assertEqual(self.req("/api/ping"), 200)
        self.assertEqual(self.req("/api/data"), 200)
        self.assertEqual(self.req("/api/settings", {"name": "x"}, {"X-Requested-With": "jard", "Origin": f"http://127.0.0.1:{self.port}"}), 200)

    def test_cross_site_posts_are_rejected(self):
        self.assertEqual(self.req("/api/item_delete", {"code": "x"}), 403)                                   # بدون ترويسة البرنامج
        self.assertEqual(self.req("/api/item_delete", {"code": "x"}, {"X-Requested-With": "jard", "Origin": "http://evil.example"}), 403)
        self.assertEqual(self.req("/api/quit", {}, {"Origin": "http://evil.example"}), 403)

    def test_foreign_host_header_is_rejected(self):
        self.assertEqual(self.req("/api/data", None, {"Host": "evil.example"}), 403)
        self.assertEqual(self.req("/api/backup", None, {"Host": "evil.example:80"}), 403)

    def test_bad_json_is_a_clean_400(self):
        r = urllib.request.Request(f"http://127.0.0.1:{self.port}/api/settings", data=b"not json",
                                   headers={"X-Requested-With": "jard", "Content-Type": "application/json"})
        with self.assertRaises(urllib.error.HTTPError) as e:
            urllib.request.urlopen(r, timeout=10)
        self.assertEqual(e.exception.code, 400)

    def test_view_only_mode_blocks_writes_but_allows_restore_and_toggle(self):
        ok = {"X-Requested-With": "jard", "Origin": f"http://127.0.0.1:{self.port}"}
        self.assertEqual(self.req("/api/viewonly", {"on": True}, ok), 200)
        try:
            self.assertTrue(app.DATA.viewonly())
            self.assertEqual(self.req("/api/supplier", {"name": "م", "phone": "1"}, ok), 403)
            self.assertEqual(self.req("/api/settings", {"name": "x"}, ok), 403)
            self.assertEqual(self.req("/api/items_delete", {"codes": ["x"]}, ok), 403)
            self.assertEqual(self.req("/api/data"), 200)                                  # القراءة شغالة
            self.assertEqual(self.req("/api/restore", {"data": ""}, ok), 400)             # مسموح (يرفض الملف الفاضي بس مو بسبب الوضع)
            d = json.loads(urllib.request.urlopen(f"http://127.0.0.1:{self.port}/api/data", timeout=30).read())
            self.assertTrue(d["viewonly"])
        finally:
            self.assertEqual(self.req("/api/viewonly", {"on": False}, ok), 200)
        self.assertFalse(app.DATA.viewonly())
        self.assertEqual(self.req("/api/supplier", {"name": "م", "phone": "1"}, ok), 200)

    def test_view_only_flag_stops_cloud_copy(self):
        d = storage.Data(tempfile.mkdtemp())
        d.ensure(os.path.join(ROOT, "template.xlsx"))
        cloud = tempfile.mkdtemp()
        d.set_cloud(cloud)
        d.set_viewonly(True)
        self.assertIsNone(d.cloud_backup())
        d.set_viewonly(False)
        self.assertIsNotNone(d.cloud_backup())

    def test_pdf_endpoint_needs_the_app_header_and_never_crashes(self):
        body = {"html": "<p>x</p>", "size": "A4"}
        self.assertEqual(self.req("/api/pdf", body), 403)                                                    # بدون ترويسة البرنامج
        self.assertEqual(self.req("/api/pdf", body, {"X-Requested-With": "jard", "Origin": "http://evil.example"}), 403)
        ok = {"X-Requested-With": "jard", "Origin": f"http://127.0.0.1:{self.port}"}
        from unittest import mock
        with mock.patch.object(app, "find_app_browser", return_value=None), mock.patch.object(app, "FALLBACK_BROWSERS", ()):
            self.assertEqual(self.req("/api/pdf", body, ok), 400)                                           # ما فيه متصفح: رسالة واضحة مو انهيار
        self.assertEqual(self.req("/pdfpage/doesnotexist"), 404)


if __name__ == "__main__":
    unittest.main()
