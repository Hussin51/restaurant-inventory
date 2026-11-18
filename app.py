"""برنامج جرد المطعم: يشغّل صفحة محلية بالمتصفح ويقرأ/يكتب بملف الإكسل نفسه."""
import json, os, sys, threading, webbrowser, socket, argparse, subprocess, base64, re, mimetypes, shlex, tempfile, time
import urllib.request
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from datetime import date
from core import Store, FileBusy, CATEGORIES, TYPES, INV_TYPES, UNITS, WAREHOUSE
import storage
import reports
import importer
import clothes
import updater
from urllib.parse import urlparse, parse_qs, quote, unquote

VERSION = "1.9.3"
BASE = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
LOCK = threading.Lock()
PENDING = False
LAST_WRITE = 0.0
WRITE_PATHS = ("/api/movement", "/api/invoice", "/api/item", "/api/item_update", "/api/image", "/api/supplier", "/api/supplier_update", "/api/settings", "/api/restore", "/api/count", "/api/import_apply", "/api/sort_items", "/api/item_delete", "/api/items_delete", "/api/invoice_delete", "/api/supplier_delete", "/api/movement_delete", "/api/cl_item", "/api/cl_receive", "/api/cl_issue", "/api/cl_adjust", "/api/cl_doc_delete", "/api/cl_item_delete", "/api/cl_move_delete", "/api/cl_buy", "/api/cl_buy_delete")
STORE = None
PORT = 0
DATA = None
IMG_DIR = None
EXT = {"image/jpeg": "jpg", "image/jpg": "jpg", "image/png": "png", "image/webp": "webp", "image/gif": "gif"}
MAX_IMG = 8 * 1024 * 1024


def to_json(o):
    return json.dumps(o, ensure_ascii=False, default=str).encode("utf-8")


def safe_code(code):
    """اسم ملف آمن للكود. إذا تغيّر الكود نضيف بصمة حتى ما تتصادم صورتين."""
    code = code.strip()
    base = re.sub(r"[^\w\-]+", "_", code, flags=re.UNICODE) or "item"
    if base != code:
        import hashlib
        base += "_" + hashlib.md5(code.encode("utf8")).hexdigest()[:6]
    return base


def save_image(code, data, ctype):
    ext = EXT.get(ctype.split(";")[0].strip().lower())
    if not ext:
        raise ValueError("نوع الصورة غير مدعوم (jpg / png / webp / gif)")
    if len(data) > MAX_IMG:
        raise ValueError("الصورة كبيرة (أكثر من 8 ميغا)")
    os.makedirs(IMG_DIR, exist_ok=True)
    base = safe_code(code)
    for f in os.listdir(IMG_DIR):                      # امسح الصورة القديمة لنفس الكود
        if os.path.splitext(f)[0] == base:
            os.remove(os.path.join(IMG_DIR, f))
    fn = f"{base}.{ext}"
    with open(os.path.join(IMG_DIR, fn), "wb") as fh:
        fh.write(data)
    return fn


def remove_image_files(code):
    base = safe_code(code)
    if os.path.isdir(IMG_DIR):
        for f in os.listdir(IMG_DIR):
            if os.path.splitext(f)[0] == base:
                try:
                    os.remove(os.path.join(IMG_DIR, f))
                except OSError:
                    pass


def fetch_url(url):
    if not re.match(r"^https?://", url, re.I):
        raise ValueError("الرابط لازم يبدأ بـ http أو https")
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    try:
        with urllib.request.urlopen(req, timeout=15) as r:
            ctype = r.headers.get("Content-Type", "")
            data = r.read(MAX_IMG + 1)
    except Exception:
        raise ValueError("ما كدرت أنزّل الصورة من الرابط. تأكد من الإنترنت أو انسخ الصورة نفسها والصقها")
    if not ctype.lower().startswith("image/"):
        raise ValueError("الرابط مو رابط صورة مباشر. اضغط يمين على الصورة واختار «نسخ عنوان الصورة»")
    return data, ctype


class H(BaseHTTPRequestHandler):
    def log_message(self, *a): pass

    def _blocked(self, post):
        """حماية: يقبل الطلبات من صفحة البرنامج نفسها فقط (يمنع المواقع الثانية من التحكم بالبرنامج)."""
        if os.environ.get("INVENTORY_DEV"):
            return False
        if (self.headers.get("Host") or "").lower() not in (f"127.0.0.1:{PORT}", f"localhost:{PORT}"):
            return True
        if post:
            o = self.headers.get("Origin")
            if o and o not in (f"http://127.0.0.1:{PORT}", f"http://localhost:{PORT}"):
                return True
            if self.headers.get("X-Requested-With") != "jard":
                return True
        return False

    def send(self, code, body, ctype="application/json; charset=utf-8"):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if self._blocked(False):
            return self.send(403, to_json(dict(error="طلب مرفوض")))
        path = self.path.split("?")[0]
        qs = {k: v[0] for k, v in parse_qs(urlparse(self.path).query).items()}
        if path in ("/logo.png", "/logo-round.png", "/favicon.png"):
            fp = os.path.join(BASE, "web", path[1:])
            self.send(200, open(fp, "rb").read(), "image/png") if os.path.isfile(fp) else self.send(404, b"")
        elif path == "/api/items_template":
            phone = qs.get("kind") == "phone"
            z = importer.build_phone_template() if phone else importer.build_template(CATEGORIES)
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", "attachment; filename=\"" + ("phone_count_template" if phone else "items_import_template") + ".xlsx\"; filename*=UTF-8''" + quote("قالب_الجرد_بالموبايل.xlsx" if phone else "قالب_استيراد_المواد.xlsx"))
            self.send_header("Content-Length", str(len(z)))
            self.end_headers()
            self.wfile.write(z)
        elif path in ("/api/report", "/api/report_xlsx"):
            try:
                with LOCK:
                    data = STORE.all()
                rep = reports.build(qs.get("kind", ""), data, qs.get("from", ""), qs.get("to", ""), qs.get("dept", ""), qs.get("supplier", ""), qs.get("factor", 2), qs.get("employee", ""))
                rep["restaurant"] = data["settings"]["name"]
                if path == "/api/report":
                    self.send(200, to_json(rep))
                else:
                    z = reports.build_xlsx(rep, rep["restaurant"], os.path.join(BASE, "assets", "logo-xl.png"))
                    fn = f"{rep['title']}_{__import__('datetime').date.today()}.xlsx"
                    ascii_name = f"report_{qs.get('kind', 'x')}_{__import__('datetime').date.today()}.xlsx"
                    self.send_response(200)
                    self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    self.send_header("Content-Disposition", "attachment; filename=\"" + ascii_name + "\"; filename*=UTF-8''" + quote(fn))
                    self.send_header("Content-Length", str(len(z)))
                    self.end_headers()
                    self.wfile.write(z)
            except ValueError as e:
                self.send(400, to_json(dict(error=str(e))))
            except Exception as e:
                self.send(500, to_json(dict(error=str(e))))
        elif path == "/api/ping":
            self.send(200, to_json(dict(app="jard-almat3am", version=VERSION)))
        elif path == "/api/update_check":
            self.send(200, to_json(updater.check(VERSION)))
        elif path == "/api/data":
            try:
                with LOCK:
                    d = STORE.all()
                d.update(categories=CATEGORIES + sorted({i["cat"] for i in d["items"] if i["cat"] and i["cat"] not in CATEGORIES}), types=TYPES, inv_types=list(INV_TYPES), depts=d["settings"]["depts"], warehouse=WAREHOUSE, units=UNITS, file=DATA.xlsx_name,
                         data_dir=DATA.dir, version=VERSION, backups=DATA.list_backups(),
                         cloud=dict(dir=DATA.cloud_dir(), last=DATA.last_cloud, error=DATA.cloud_error, detected=storage.detect_cloud_dirs()))
                self.send(200, to_json(d))
            except Exception as e:
                self.send(500, to_json(dict(error=f"ما كدرت أقرأ ملف الإكسل: {e}")))
        elif path == "/api/backup":
            with LOCK:
                z = DATA.backup_bytes()
            from datetime import datetime
            fn = f"جرد_المطعم_{datetime.now():%Y-%m-%d_%H%M}.zip"
            ascii_name = f"inventory_backup_{datetime.now():%Y-%m-%d_%H%M}.zip"
            self.send_response(200)
            self.send_header("Content-Type", "application/zip")
            self.send_header("Content-Disposition", "attachment; filename=\"" + ascii_name + "\"; filename*=UTF-8''" + quote(fn))
            self.send_header("Content-Length", str(len(z)))
            self.end_headers()
            self.wfile.write(z)
        elif path in ("/", "/index.html"):
            self.send(200, open(os.path.join(BASE, "web", "index.html"), "rb").read(), "text/html; charset=utf-8")
        elif path.startswith("/img/"):
            fn = os.path.basename(path[5:])
            fn = os.path.basename(unquote(fn))
            fp = os.path.join(IMG_DIR, fn)
            if os.path.isfile(fp):
                self.send(200, open(fp, "rb").read(), mimetypes.guess_type(fp)[0] or "image/jpeg")
            else:
                self.send(404, b"")
        else:
            self.send(404, b"{}")

    def do_POST(self):
        if self._blocked(True):
            return self.send(403, to_json(dict(error="طلب مرفوض")))
        n = int(self.headers.get("Content-Length") or 0)
        if n > 120 * 1024 * 1024:
            return self.send(413, to_json(dict(error="الملف كبير")))
        try:
            b = json.loads(self.rfile.read(n) or b"{}")
            if not isinstance(b, dict):
                raise ValueError
        except ValueError:
            return self.send(400, to_json(dict(error="طلب غير صالح")))
        p = self.path
        try:
            res = {}
            if p == "/api/movement":
                q = float(b["qty"]); q = int(q) if q == int(q) else q
                with LOCK:
                    from datetime import datetime
                    d = datetime.strptime(b.get("date") or date.today().isoformat(), "%Y-%m-%d").date()
                    STORE.add_movement(d, b["item"], b["type"], q, b.get("frm", ""), b.get("to", ""), b.get("note", ""), unit=b.get("unit", "قطعة"))
            elif p == "/api/update_install":
                try:
                    info = updater.fetch_info()
                    if updater.vtuple(info["version"]) <= updater.vtuple(VERSION):
                        raise ValueError("عندك آخر نسخة")
                    path = updater.download_installer(info)
                except OSError:
                    raise ValueError("ما كدرت أتصل بموقع التحديث، تأكد من الإنترنت وجرّب مرة ثانية")
                with LOCK:
                    DATA.make_backup("before-update")
                updater.launch(path)
                res["ok"] = True
                threading.Timer(2.0, shutdown).start()
            elif p == "/api/invoice":
                with LOCK:
                    res["no"] = STORE.add_invoice(b)
            elif p == "/api/item":
                with LOCK:
                    STORE.add_item(b["code"], b["name"], b.get("cat", ""), "قطعة", float(b.get("opening") or 0),
                                   float(b.get("price") or 0), float(b.get("min") or 0),
                                   carton=int(float(b.get("carton") or 0)), shad=int(float(b.get("shad") or 0)),
                                   usage=b.get("usage", ""), supplier=b.get("supplier", ""), currency=b.get("currency", ""))
            elif p == "/api/item_update":
                with LOCK:
                    STORE.update_item(b["code"], b.get("cat", ""), "قطعة", float(b.get("price") or 0), float(b.get("min") or 0),
                                      carton=int(float(b.get("carton") or 0)), shad=int(float(b.get("shad") or 0)),
                                      usage=b.get("usage", ""), supplier=b.get("supplier", ""), currency=b.get("currency"))
            elif p == "/api/count":
                with LOCK:
                    counts = b.get("counts") or []
                    if not counts:
                        raise ValueError("ما كو كميات للاعتماد")
                    if b.get("mode") == "opening":
                        res["n"] = STORE.set_opening(counts)
                    else:
                        res["n"] = STORE.adjust(counts, b.get("date"), b.get("note") or "تسوية جرد")
            elif p == "/api/image":
                code = b["code"]
                with LOCK:
                    if b.get("remove"):
                        STORE.set_image(code, "")
                        base = safe_code(code)
                        if os.path.isdir(IMG_DIR):
                            for f in os.listdir(IMG_DIR):
                                if os.path.splitext(f)[0] == base:
                                    os.remove(os.path.join(IMG_DIR, f))
                    else:
                        if b.get("data"):
                            m = re.match(r"^data:(image/[\w.+-]+);base64,(.*)$", b["data"], re.S)
                            if not m:
                                raise ValueError("صورة غير صالحة")
                            data, ctype = base64.b64decode(m.group(2)), m.group(1)
                        else:
                            data, ctype = fetch_url(b["url"].strip())
                        res["image"] = fn = save_image(code, data, ctype)
                        STORE.set_image(code, fn)
            elif p == "/api/supplier":
                with LOCK:
                    STORE.add_supplier(b["name"], b.get("phone", ""), b.get("currency", ""))
            elif p == "/api/supplier_update":
                with LOCK:
                    STORE.update_supplier(b["name"], b.get("phone", ""), b.get("currency"))
            elif p == "/api/settings":
                with LOCK:
                    STORE.set_settings(b.get("name", ""), b.get("depts"), b.get("rate"))
            elif p in ("/api/import_preview", "/api/import_apply"):
                raw = base64.b64decode((b.get("data") or "").split(",")[-1] or b"")
                with LOCK:
                    items = STORE.items()
                    sups = [x["name"] for x in STORE._read_suppliers(STORE._open())]
                    parsed = importer.parse(raw, {i["code"].casefold(): i["name"] for i in items}, {i["name"]: i["code"] for i in items}, sups)
                    if p == "/api/import_preview":
                        res.update(parsed)
                    else:
                        DATA.make_backup("before-import")
                        res.update(STORE.import_items(parsed["rows"], b.get("mode", "skip")))
                        res["errors"] = parsed["counts"]["error"]
                        imgs, n_img = importer.extract_images(raw), 0
                        have = {i["code"].casefold(): i.get("image") for i in STORE.items()}
                        for x in parsed["rows"]:
                            if x["status"] == "error" or x["row"] not in imgs:
                                continue
                            was_new = x["status"] == "new"
                            if not was_new and (b.get("mode") != "update" or have.get(x["code"].casefold())):
                                continue                                   # الموجود: صورة جديدة فقط بوضع التحديث وإذا ما عنده صورة
                            try:
                                data_i, ctype = imgs[x["row"]]
                                STORE.set_image(x["code"], save_image(x["code"], data_i, ctype))
                                n_img += 1
                            except Exception:
                                pass
                        res["images"] = n_img
            elif p == "/api/items_delete":
                with LOCK:
                    DATA.make_backup("before-delete")
                    r = STORE.delete_items(b.get("codes") or [])
                    for fn in r.pop("images", []):
                        base = os.path.splitext(fn)[0]
                        if os.path.isdir(IMG_DIR):
                            for f in os.listdir(IMG_DIR):
                                if os.path.splitext(f)[0] == base:
                                    os.remove(os.path.join(IMG_DIR, f))
                    res.update(r)
            elif p in ("/api/item_delete", "/api/invoice_delete", "/api/supplier_delete", "/api/movement_delete"):
                with LOCK:
                    DATA.make_backup("before-delete")                      # نسخة احتياطية قبل أي حذف (تكدر ترجّع منها)
                    if p == "/api/item_delete":
                        r = STORE.delete_item(b["code"])
                        remove_image_files(b["code"])
                        res.update(r)
                    elif p == "/api/invoice_delete":
                        res.update(STORE.delete_invoice(b["no"]))
                    elif p == "/api/supplier_delete":
                        STORE.delete_supplier(b["name"])
                    else:
                        STORE.delete_movement(int(b["row"]))
            elif p == "/api/cl_item":
                with LOCK:
                    res["n"] = clothes.add_items(STORE, b.get("name", ""), b.get("color", ""), b.get("variants") or [])
            elif p == "/api/cl_receive":
                with LOCK:
                    clothes.receive(STORE, b["code"], b.get("qty"), b.get("date"), b.get("note", ""))
            elif p == "/api/cl_issue":
                with LOCK:
                    res["no"] = clothes.issue(STORE, b.get("employee", ""), b.get("date"), b.get("note", ""), b.get("lines") or [])
            elif p == "/api/cl_adjust":
                with LOCK:
                    res["n"] = clothes.adjust(STORE, b["code"], b.get("pieces"), b.get("date"))
            elif p == "/api/cl_buy":
                with LOCK:
                    res["no"] = clothes.buy(STORE, b.get("supplier", ""), b.get("date"), b.get("note", ""), b.get("currency", ""), b.get("lines") or [])
            elif p in ("/api/cl_doc_delete", "/api/cl_item_delete", "/api/cl_move_delete", "/api/cl_buy_delete"):
                with LOCK:
                    DATA.make_backup("before-delete")
                    if p == "/api/cl_buy_delete":
                        clothes.delete_purchase(STORE, b["no"])
                    elif p == "/api/cl_doc_delete":
                        clothes.delete_doc(STORE, b["no"])
                    elif p == "/api/cl_item_delete":
                        clothes.delete_item(STORE, b["code"])
                    else:
                        clothes.delete_move(STORE, int(b["row"]))
            elif p == "/api/sort_items":
                with LOCK:
                    res["n"] = STORE.sort_items()
            elif p == "/api/restore":
                with LOCK:
                    res["images"] = DATA.restore_bytes(base64.b64decode(b["data"].split(",")[-1]))
            elif p == "/api/cloud":
                with LOCK:
                    res["dir"] = DATA.set_cloud(b.get("path", ""))
            elif p == "/api/cloud_now":
                with LOCK:
                    if not DATA.cloud_dir():
                        raise ValueError("اختار مجلد السحابة أول")
                    DATA.cloud_backup()
            elif p == "/api/openfolder":
                try: os.startfile(DATA.dir)
                except AttributeError: subprocess.Popen(["xdg-open", DATA.dir])
            elif p == "/api/open":
                try: os.startfile(STORE.path)
                except AttributeError: subprocess.Popen(["xdg-open", STORE.path])
            elif p == "/api/quit":
                self.send(200, b"{}"); threading.Timer(0.3, shutdown).start(); return
            else:
                return self.send(404, b"{}")
            res["ok"] = True
            if p in WRITE_PATHS:
                global PENDING, LAST_WRITE
                PENDING, LAST_WRITE = True, __import__("time").time()
            self.send(200, to_json(res))
        except FileBusy as e:
            self.send(409, to_json(dict(error=str(e))))
        except (ValueError, KeyError) as e:
            msg = "تأكد من الأرقام والحقول المطلوبة" if isinstance(e, KeyError) or "could not" in str(e) else str(e)
            self.send(400, to_json(dict(error=msg)))
        except OSError as e:
            self.send(500, to_json(dict(error="فشل النسخ للسحابة: " + str(e))))
        except Exception as e:
            self.send(500, to_json(dict(error=str(e))))


def shutdown():
    """إغلاق البرنامج: آخر نسخة للسحابة إذا في تعديلات جديدة، ثم الخروج."""
    global PENDING
    try:
        if PENDING and DATA is not None and LOCK.acquire(timeout=5):
            try:
                DATA.cloud_backup_safe()
            finally:
                LOCK.release()
    except Exception:
        pass
    os._exit(0)


def find_app_browser():
    """Edge أو Chrome لفتح البرنامج بنافذة مستقلة (بدون شريط عنوان وتبويبات)."""
    exe = os.environ.get("INVENTORY_BROWSER")
    if exe:
        return exe if os.path.exists(exe) else None
    if os.name == "nt":
        for base in (os.environ.get("ProgramFiles(x86)"), os.environ.get("ProgramFiles"), os.environ.get("LOCALAPPDATA")):
            if not base:
                continue
            for rel in (r"Microsoft\Edge\Application\msedge.exe", r"Google\Chrome\Application\chrome.exe"):
                fp = os.path.join(base, rel)
                if os.path.exists(fp):
                    return fp
    return None


def open_ui(url):
    """يفتح نافذة البرنامج ويرجع العملية (أو None إذا فتح المتصفح العادي)."""
    exe = find_app_browser()
    if not exe:
        webbrowser.open(url)
        return None
    prof = os.path.join(os.environ.get("LOCALAPPDATA") or tempfile.gettempdir(), "RestaurantInventory", "browser")
    try:
        os.makedirs(prof, exist_ok=True)
        args = [exe, f"--app={url}", f"--user-data-dir={prof}", "--window-size=1400,900", "--no-first-run",
                "--no-default-browser-check", "--disable-background-mode"] + shlex.split(os.environ.get("INVENTORY_BROWSER_ARGS", ""))
        return subprocess.Popen(args)
    except OSError:
        webbrowser.open(url)
        return None


def port_busy(p):
    with socket.socket() as s:
        return s.connect_ex(("127.0.0.1", p)) == 0


def is_ours(p):
    """هل الي شغّال على المنفذ هو البرنامج نفسه؟"""
    try:
        with urllib.request.urlopen(f"http://127.0.0.1:{p}/api/ping", timeout=2) as r:
            return json.loads(r.read()).get("app") == "jard-almat3am"
    except Exception:
        return False


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("file", nargs="?", help="ملف إكسل محدد (اختياري). الافتراضي: مجلد المستندات\\جرد المطعم")
    ap.add_argument("--port", type=int, default=8765)
    ap.add_argument("--no-browser", action="store_true")
    a = ap.parse_args()
    if sys.stdout is None:                      # نسخة exe بدون نافذة: ما كو شاشة طباعة
        sys.stdout = open(os.devnull, "w")
    if sys.stderr is None:
        sys.stderr = open(os.devnull, "w")
    port = a.port
    while port_busy(port):
        if is_ours(port):                      # البرنامج شغّال من قبل: بس افتح الصفحة
            open_ui(f"http://127.0.0.1:{port}/"); sys.exit(0)
        port += 1                              # المنفذ مشغول ببرنامج ثاني: جرب الي بعده
        if port > a.port + 20:
            sys.exit(1)
    a.port = port
    PORT = port
    url = f"http://127.0.0.1:{a.port}/"
    DATA = storage.Data(os.path.dirname(os.path.abspath(a.file)), os.path.basename(a.file)) if a.file else storage.Data(storage.default_dir())
    DATA.ensure(os.path.join(BASE, "template.xlsx"))
    STORE, IMG_DIR = Store(DATA.xlsx), DATA.images
    try: DATA.auto_backup_if_needed()
    except Exception: pass
    DATA.cloud_backup_safe()
    def _bg_loop():
        import time
        global PENDING
        last_daily = time.time()
        while True:
            time.sleep(15)
            try:
                with LOCK:
                    if PENDING and time.time() - LAST_WRITE >= 20:     # بعد 20 ثانية من آخر تعديل
                        PENDING = False
                        DATA.cloud_backup_safe()
                    if time.time() - last_daily > 6 * 3600:
                        last_daily = time.time()
                        DATA.auto_backup_if_needed()
                        DATA.cloud_backup_safe()
            except Exception:
                pass
    threading.Thread(target=_bg_loop, daemon=True).start()
    srv = ThreadingHTTPServer(("127.0.0.1", a.port), H)
    if not a.no_browser:
        def _ui():
            t0 = time.time()
            proc = open_ui(url)
            if proc is None:
                return                            # متصفح عادي: البرنامج يظل شغّال لحد «إغلاق البرنامج»
            proc.wait()
            if time.time() - t0 > 8:              # النافذة انسكّرت: سكّر البرنامج
                shutdown()
        threading.Timer(0.8, lambda: threading.Thread(target=_ui, daemon=True).start()).start()
    srv.serve_forever()
