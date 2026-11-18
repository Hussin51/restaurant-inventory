import hashlib, http.server, json, os, tempfile, threading, unittest
import updater


class H(http.server.BaseHTTPRequestHandler):
    exe = b"MZfake-installer"
    ver = "9.9.9"
    sha = None

    def do_GET(self):
        if self.path.endswith("/version.json"):
            body = json.dumps(dict(version=H.ver, sha256=H.sha or hashlib.sha256(H.exe).hexdigest(), notes="x")).encode()
        elif self.path.endswith("/RestaurantInventory-Setup.exe"):
            body = H.exe
        else:
            self.send_response(404); self.end_headers(); return
        self.send_response(200); self.send_header("Content-Length", str(len(body))); self.end_headers(); self.wfile.write(body)

    def log_message(self, *a): pass


class UpdaterTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.srv = http.server.HTTPServer(("127.0.0.1", 0), H)
        threading.Thread(target=cls.srv.serve_forever, daemon=True).start()
        os.environ["INVENTORY_UPDATE_URL"] = f"http://127.0.0.1:{cls.srv.server_port}"

    @classmethod
    def tearDownClass(cls):
        cls.srv.shutdown(); os.environ.pop("INVENTORY_UPDATE_URL", None)

    def setUp(self):
        H.ver, H.sha = "9.9.9", None

    def test_versions_compare_numerically(self):
        self.assertTrue(updater.vtuple("1.10.0") > updater.vtuple("1.9.2"))
        self.assertEqual(updater.vtuple("v1.8.1"), (1, 8, 1))

    def test_available_when_newer(self):
        r = updater.check("1.9.0")
        self.assertTrue(r["available"]); self.assertEqual(r["latest"], "9.9.9")

    def test_not_available_when_same_or_older(self):
        H.ver = "1.9.0"
        self.assertFalse(updater.check("1.9.0")["available"])
        H.ver = "1.2.0"
        self.assertFalse(updater.check("1.9.0")["available"])

    def test_offline_is_silent(self):
        os.environ["INVENTORY_UPDATE_URL"] = "http://127.0.0.1:1"
        try:
            r = updater.check("1.9.0", timeout=1)
            self.assertFalse(r["available"])
        finally:
            os.environ["INVENTORY_UPDATE_URL"] = f"http://127.0.0.1:{self.srv.server_port}"

    def test_download_verifies_hash(self):
        p = updater.download_installer()
        self.assertEqual(open(p, "rb").read(), H.exe)

    def test_download_rejects_wrong_hash(self):
        H.sha = "0" * 64
        with self.assertRaises(ValueError):
            updater.download_installer()

    def test_rejects_info_without_hash(self):
        with self.assertRaises(ValueError):
            updater.download_installer({"version": "9.9.9"})


if __name__ == "__main__":
    unittest.main()
