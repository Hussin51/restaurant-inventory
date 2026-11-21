import hashlib, http.server, json, os, tempfile, threading, unittest
import updater


class H(http.server.BaseHTTPRequestHandler):
    exe = b"MZfake-installer"
    ver = "9.9.9"
    sha = None
    file = None          # اسم المثبّت بالنسخة الجديدة (مثل RestaurantInventory-Setup-9.9.9.exe)
    stale = 0            # كم تنزيل أول يرجّع ملف قديم (كاش) قبل ما يطلع الصحيح
    hits = 0

    def do_GET(self):
        if self.path.endswith("/version.json"):
            d = dict(version=H.ver, sha256=H.sha or hashlib.sha256(H.exe).hexdigest(), notes="x")
            if H.file:
                d["file"] = H.file
            body = json.dumps(d).encode()
        elif self.path.endswith("/" + (H.file or "RestaurantInventory-Setup.exe")):
            H.hits += 1
            body = b"STALE-OLD-INSTALLER" if H.hits <= H.stale else H.exe
        elif self.path.endswith("/RestaurantInventory-Setup.exe"):
            body = b"STALE-OLD-INSTALLER"                    # الاسم الثابت لسه فيه الملف القديم
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
        H.ver, H.sha, H.file, H.stale, H.hits = "9.9.9", None, None, 0, 0

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
        p = updater.download_installer(wait=0)
        self.assertEqual(open(p, "rb").read(), H.exe)

    def test_download_rejects_wrong_hash(self):
        H.sha = "0" * 64
        with self.assertRaises(ValueError) as e:
            updater.download_installer(wait=0)
        self.assertIn("بصمة", str(e.exception))

    def test_rejects_info_without_hash(self):
        with self.assertRaises(ValueError):
            updater.download_installer({"version": "9.9.9"})

    def test_versioned_installer_name_is_used(self):
        H.file = "RestaurantInventory-Setup-9.9.9.exe"           # الاسم الثابت يرجّع ملف قديم، والاسم بالرقم هو الصحيح
        p = updater.download_installer(wait=0)
        self.assertEqual(open(p, "rb").read(), H.exe)

    def test_unsafe_file_name_falls_back_to_default(self):
        self.assertEqual(updater._installer_name({"file": "../evil.exe"}), "RestaurantInventory-Setup.exe")
        self.assertEqual(updater._installer_name({"file": "RestaurantInventory-Setup-1.9.11.exe"}), "RestaurantInventory-Setup-1.9.11.exe")

    def test_stale_download_is_retried_then_succeeds(self):
        H.stale = 1                                              # أول تنزيل ملف قديم (كاش)، الثاني الصحيح
        p = updater.download_installer(wait=0)
        self.assertEqual(open(p, "rb").read(), H.exe)
        self.assertEqual(H.hits, 2)

    def test_installer_gets_a_clean_pyinstaller_environment(self):
        os.environ["_PYI_APPLICATION_HOME_DIR"] = "C:/old/_MEI123"
        os.environ["_MEIPASS2"] = "C:/old/_MEI123"
        try:
            env = updater._clean_env()
        finally:
            os.environ.pop("_PYI_APPLICATION_HOME_DIR", None)
            os.environ.pop("_MEIPASS2", None)
        self.assertFalse(any(k.startswith("_PYI") or k == "_MEIPASS2" for k in env))
        self.assertEqual(env["PYINSTALLER_RESET_ENVIRONMENT"], "1")


if __name__ == "__main__":
    unittest.main()
