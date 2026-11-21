"""تحديث البرنامج من مستودع عام للتحديثات: يفحص رقم آخر نسخة، وينزّل المثبّت ويتحقق من بصمته ثم يشغّله."""
import hashlib, json, os, re, subprocess, tempfile, time, urllib.request

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


def _installer_name(info):
    """اسم ملف المثبّت: النسخ الجديدة تنشر اسم بالرقم (مثل RestaurantInventory-Setup-1.9.11.exe) حتى ما يصير تعارض مع ملف قديم محفوظ بالكاش."""
    n = str(info.get("file", "")) if isinstance(info, dict) else ""
    return n if re.fullmatch(r"RestaurantInventory-Setup[A-Za-z0-9._-]*\.exe", n) else "RestaurantInventory-Setup.exe"


def _download_once(info, timeout):
    want = str(info.get("sha256", "")).lower()
    if len(want) != 64:
        raise ValueError("ملف التحديث ما فيه بصمة للتحقق، تم الإلغاء")
    d = tempfile.mkdtemp(prefix="inv-update-")
    path = os.path.join(d, "RestaurantInventory-Setup.exe")
    h = hashlib.sha256()
    req = urllib.request.Request(base_url() + "/" + _installer_name(info), headers={"User-Agent": "restaurant-inventory", "Cache-Control": "no-cache"})
    with urllib.request.urlopen(req, timeout=timeout) as r, open(path, "wb") as f:
        while True:
            b = r.read(1 << 16)
            if not b:
                break
            h.update(b)
            f.write(b)
    if h.hexdigest().lower() != want:
        os.remove(path)
        return None
    return path


def download_installer(info=None, timeout=120, retries=2, wait=6):
    """ينزّل المثبّت لمجلد مؤقت ويتحقق من SHA-256. إذا ما طابقت (ملف قديم بالكاش أو النشر لسه شغال) يعيد قراءة معلومات النسخة ويحاول مرة ثانية."""
    info = info or fetch_info()
    for k in range(retries + 1):
        path = _download_once(info, timeout)
        if path:
            return path
        if k < retries:
            time.sleep(wait)
            info = fetch_info()
    raise ValueError("بصمة الملف المنزَّل ما تطابق (ممكن النسخة الجديدة لسه تنرفع). انتظر دقيقتين وجرّب مرة ثانية، أو نزّل المثبّت من صفحة الإصدارات")


def launch(path):
    """يشغّل المثبّت بصمت (يغلق البرنامج القديم ويعيد تشغيل الجديد)."""
    if os.name != "nt":
        raise ValueError("التحديث التلقائي يشتغل على Windows فقط")
    flags = getattr(subprocess, "DETACHED_PROCESS", 0) | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
    subprocess.Popen([path, "/SILENT", "/CLOSEAPPLICATIONS", "/NORESTART", "/SUPPRESSMSGBOXES", "/RELAUNCH=1"], creationflags=flags, close_fds=True)
