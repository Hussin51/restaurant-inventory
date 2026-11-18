"""تحديث البرنامج من مستودع عام للتحديثات: يفحص رقم آخر نسخة، وينزّل المثبّت ويتحقق من بصمته ثم يشغّله."""
import hashlib, json, os, subprocess, tempfile, urllib.request

REPO = os.environ.get("INVENTORY_UPDATE_REPO", "Hussin51/restaurant-inventory")
TAG = "inventory-latest"


def base_url():
    return os.environ.get("INVENTORY_UPDATE_URL") or f"https://github.com/{REPO}/releases/download/{TAG}"


def vtuple(v):
    out = []
    for p in str(v).strip().lstrip("vV").split("."):
        digits = "".join(ch for ch in p if ch.isdigit())
        out.append(int(digits or 0))
    return tuple(out)


def fetch_info(timeout=6):
    req = urllib.request.Request(base_url() + "/version.json", headers={"User-Agent": "restaurant-inventory"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        info = json.loads(r.read().decode("utf-8"))
    if not isinstance(info, dict) or "version" not in info:
        raise ValueError("ملف التحديث غير صالح")
    return info


def check(current, timeout=6):
    """يرجع dict: available, current, latest, notes. أي فشل بالاتصال = لا يوجد تحديث (بدون إزعاج)."""
    try:
        info = fetch_info(timeout)
        latest = str(info["version"])
        return dict(available=vtuple(latest) > vtuple(current), current=current, latest=latest, notes=str(info.get("notes", "")))
    except Exception as e:
        return dict(available=False, current=current, latest="", notes="", error=str(e))


def download_installer(info=None, timeout=120):
    """ينزّل المثبّت لمجلد مؤقت ويتحقق من SHA-256. يرجع مسار الملف."""
    info = info or fetch_info()
    want = str(info.get("sha256", "")).lower()
    if len(want) != 64:
        raise ValueError("ملف التحديث ما فيه بصمة للتحقق، تم الإلغاء")
    d = tempfile.mkdtemp(prefix="inv-update-")
    path = os.path.join(d, "RestaurantInventory-Setup.exe")
    h = hashlib.sha256()
    req = urllib.request.Request(base_url() + "/RestaurantInventory-Setup.exe", headers={"User-Agent": "restaurant-inventory"})
    with urllib.request.urlopen(req, timeout=timeout) as r, open(path, "wb") as f:
        while True:
            b = r.read(1 << 16)
            if not b:
                break
            h.update(b)
            f.write(b)
    if h.hexdigest().lower() != want:
        os.remove(path)
        raise ValueError("بصمة الملف المنزَّل ما تطابق، تم إلغاء التحديث")
    return path


def launch(path):
    """يشغّل المثبّت بصمت (يغلق البرنامج القديم ويعيد تشغيل الجديد)."""
    if os.name != "nt":
        raise ValueError("التحديث التلقائي يشتغل على Windows فقط")
    flags = getattr(subprocess, "DETACHED_PROCESS", 0) | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
    subprocess.Popen([path, "/SILENT", "/CLOSEAPPLICATIONS", "/NORESTART", "/SUPPRESSMSGBOXES", "/RELAUNCH=1"], creationflags=flags, close_fds=True)
