"""مكان البيانات والنسخ الاحتياطي. البيانات تنحفظ بمجلد المستندات (منفصل عن البرنامج)."""
import io, json, os, re, shutil, zipfile
from datetime import datetime

XLSX = "جرد_المطعم.xlsx"
IMAGES = "صور_المواد"
BACKUPS = "نسخ_احتياطية"
FOLDER = "جرد المطعم"
KEEP_AUTO = 30
KEEP_CLOUD_DAYS = 14
CLOUD_SUB = "جرد المطعم - نسخ احتياطية"
CFG = "config.json"
IMG_EXT = (".jpg", ".png", ".webp", ".gif")
MAX_ENTRY = 10 * 1024 * 1024


def documents_dir():
    if os.name == "nt":                     # يحترم تحويل المستندات لـ OneDrive
        try:
            import ctypes
            buf = ctypes.create_unicode_buffer(260)
            ctypes.windll.shell32.SHGetFolderPathW(None, 5, None, 0, buf)
            if buf.value:
                return buf.value
        except Exception:
            pass
    return os.path.join(os.path.expanduser("~"), "Documents")


def default_dir():
    return os.environ.get("INVENTORY_DATA") or os.path.join(documents_dir(), FOLDER)


def detect_cloud_dirs():
    """مجلدات OneDrive / Google Drive / Dropbox الموجودة على الجهاز."""
    found = []

    def add(name, p):
        if p and os.path.isdir(p) and os.path.abspath(p) not in [f["path"] for f in found]:
            found.append({"name": name, "path": os.path.abspath(p)})
    for k in ("OneDriveConsumer", "OneDriveCommercial", "OneDrive"):
        add("OneDrive", os.environ.get(k))
    home = os.path.expanduser("~")
    add("OneDrive", os.path.join(home, "OneDrive"))
    add("Dropbox", os.path.join(home, "Dropbox"))
    add("Google Drive", os.path.join(home, "Google Drive"))
    if os.name == "nt":
        for L in "DEFGHIJKLMNOPQRSTUVWXYZ":
            add("Google Drive", f"{L}:\\My Drive")
    return found


class Data:
    def __init__(self, folder, xlsx_name=XLSX):
        self.dir = os.path.abspath(folder)
        self.xlsx = os.path.join(self.dir, xlsx_name)
        self.xlsx_name = xlsx_name
        self.images = os.path.join(self.dir, IMAGES)
        self.backups = os.path.join(self.dir, BACKUPS)

    # ---------- نسخة سحابية (نسخ لمجلد خدمة سحابية موجودة على الجهاز) ----------
    last_cloud = ""
    cloud_error = ""

    def _cfg(self):
        try:
            with open(os.path.join(self.dir, CFG), encoding="utf8") as fh:
                return json.load(fh)
        except Exception:
            return {}

    def _save_cfg(self, d):
        with open(os.path.join(self.dir, CFG), "w", encoding="utf8") as fh:
            json.dump(d, fh, ensure_ascii=False, indent=1)

    def cloud_dir(self):
        return self._cfg().get("cloud_dir", "")

    def set_cloud(self, path):
        path = (path or "").strip().strip('"')
        cfg = self._cfg()
        if not path:
            cfg.pop("cloud_dir", None)
            self._save_cfg(cfg)
            self.last_cloud = self.cloud_error = ""
            return ""
        if not os.path.isdir(path):
            raise ValueError("المجلد غير موجود. تأكد من المسار")
        target = os.path.join(path, CLOUD_SUB)
        try:
            os.makedirs(target, exist_ok=True)
            t = os.path.join(target, ".test")
            with open(t, "w") as fh:
                fh.write("ok")
            os.remove(t)
        except OSError:
            raise ValueError("ما أكدر أكتب بهذا المجلد. اختار مجلد ثاني")
        cfg["cloud_dir"] = path
        self._save_cfg(cfg)
        self.cloud_backup()
        return path

    def _vo_path(self):
        return os.path.join(self.dir, "view-only.flag")

    def auto_update(self):
        """التحديث التلقائي مفعّل افتراضياً، ويتعطل بعلامة محلية."""
        return not os.path.exists(os.path.join(self.dir, "no-auto-update.flag"))

    def set_auto_update(self, on):
        f = os.path.join(self.dir, "no-auto-update.flag")
        if on and os.path.exists(f):
            os.remove(f)
        elif not on:
            os.makedirs(self.dir, exist_ok=True)
            with open(f, "w", encoding="utf8") as fh:
                fh.write("1")

    def viewonly(self):
        """وضع المشاهدة فقط (للحاسبة الثانية): علامة محلية خارج ملف الإكسل وخارج النسخ الاحتياطية."""
        return os.path.exists(self._vo_path())

    def set_viewonly(self, on):
        if on:
            os.makedirs(self.dir, exist_ok=True)
            with open(self._vo_path(), "w", encoding="utf8") as f:
                f.write("1")
        elif os.path.exists(self._vo_path()):
            os.remove(self._vo_path())

    def cloud_backup(self):
        """يكتب آخر نسخة + نسخة اليوم بمجلد السحابة. يرفع OSError إذا فشل."""
        path = self.cloud_dir()
        if not path or self.viewonly():                    # نسخة المشاهدة ما تكتب فوق نسخة السحابة
            return None
        target = os.path.join(path, CLOUD_SUB)
        os.makedirs(target, exist_ok=True)
        data = self.backup_bytes()

        def put(name):
            fp, tmp = os.path.join(target, name), os.path.join(target, name + ".tmp")
            with open(tmp, "wb") as fh:
                fh.write(data)
            os.replace(tmp, fp)
        put("جرد_المطعم_آخر_نسخة.zip")
        put(f"جرد_المطعم_{datetime.now():%Y-%m-%d}.zip")
        dated = sorted(f for f in os.listdir(target) if re.fullmatch(r"جرد_المطعم_\d{4}-\d{2}-\d{2}\.zip", f))
        for f in dated[:-KEEP_CLOUD_DAYS]:
            try:
                os.remove(os.path.join(target, f))
            except OSError:
                pass
        self.last_cloud = f"{datetime.now():%Y-%m-%d %H:%M}"
        self.cloud_error = ""
        return self.last_cloud

    def cloud_backup_safe(self):
        try:
            return self.cloud_backup()
        except Exception as e:
            self.cloud_error = "فشل النسخ للسحابة: " + str(e)
            return None

    def ensure(self, template):
        os.makedirs(self.dir, exist_ok=True)
        os.makedirs(self.images, exist_ok=True)
        os.makedirs(self.backups, exist_ok=True)
        if not os.path.exists(self.xlsx):
            shutil.copyfile(template, self.xlsx)

    # ---------- نسخ احتياطي ----------
    def _write_zip(self, fh):
        with zipfile.ZipFile(fh, "w", zipfile.ZIP_DEFLATED) as z:
            if os.path.exists(self.xlsx):
                z.write(self.xlsx, XLSX)
            if os.path.isdir(self.images):
                for f in sorted(os.listdir(self.images)):
                    if f.lower().endswith(IMG_EXT):
                        z.write(os.path.join(self.images, f), f"{IMAGES}/{f}")

    def backup_bytes(self):
        b = io.BytesIO()
        self._write_zip(b)
        return b.getvalue()

    def make_backup(self, tag="auto"):
        os.makedirs(self.backups, exist_ok=True)
        name = f"{tag}_{datetime.now():%Y-%m-%d_%H%M%S}.zip"
        with open(os.path.join(self.backups, name), "wb") as fh:
            self._write_zip(fh)
        self._prune()
        return name

    def _prune(self):
        for prefix, keep in (("auto_", KEEP_AUTO), ("before-delete_", 30), ("before-import_", 10)):
            fs = sorted(f for f in os.listdir(self.backups) if f.startswith(prefix) and f.endswith(".zip"))
            for f in fs[:-keep]:
                try:
                    os.remove(os.path.join(self.backups, f))
                except OSError:
                    pass

    def auto_backup_if_needed(self):
        today = f"{datetime.now():%Y-%m-%d}"
        if os.path.isdir(self.backups) and any(f.startswith(f"auto_{today}") for f in os.listdir(self.backups)):
            return None
        if not os.path.exists(self.xlsx):
            return None
        return self.make_backup("auto")

    def list_backups(self, n=8):
        if not os.path.isdir(self.backups):
            return []
        fs = sorted((f for f in os.listdir(self.backups) if f.endswith(".zip")), reverse=True)[:n]
        return [dict(name=f, size=os.path.getsize(os.path.join(self.backups, f))) for f in fs]

    def restore_bytes(self, data):
        """استرجاع نسخة (zip). يحفظ نسخة من الوضع الحالي أولاً."""
        if not zipfile.is_zipfile(io.BytesIO(data)):
            raise ValueError("هذا مو ملف نسخة احتياطية صحيح")
        z = zipfile.ZipFile(io.BytesIO(data))
        names = z.namelist()
        if XLSX not in names:
            raise ValueError("النسخة ما فيها ملف الجرد")
        xbytes = z.read(XLSX)
        try:                                              # تأكد إنه ملفنا
            from openpyxl import load_workbook
            wb = load_workbook(io.BytesIO(xbytes))
            if "الأصناف" not in wb.sheetnames or "حركة" not in wb.sheetnames:
                raise ValueError
        except Exception:
            raise ValueError("ملف الجرد داخل النسخة مو سليم")
        imgs = []
        for n in names:
            m = re.fullmatch(rf"{IMAGES}/([^/\\]+)", n)
            if m and m.group(1).lower().endswith(IMG_EXT):
                info = z.getinfo(n)
                if info.file_size <= MAX_ENTRY:
                    imgs.append((m.group(1), z.read(n)))
        self.make_backup("before-restore")
        try:
            tmp = self.xlsx + ".tmp"
            with open(tmp, "wb") as fh:
                fh.write(xbytes)
            os.replace(tmp, self.xlsx)
        except PermissionError:
            raise ValueError("ملف الإكسل مفتوح. سكّره وارجع جرب")
        os.makedirs(self.images, exist_ok=True)
        for f in os.listdir(self.images):
            if f.lower().endswith(IMG_EXT):
                os.remove(os.path.join(self.images, f))
        for f, b in imgs:
            with open(os.path.join(self.images, f), "wb") as fh:
                fh.write(b)
        return len(imgs)
