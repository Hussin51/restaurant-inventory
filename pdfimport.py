"""قراءة فواتير وعروض الأسعار (PDF) كجدول مواد: الكود، الاسم، الكمية، السعر، الصورة، وخصم الفاتورة.
أي ملف PDF تقدر تقرأه هنا يتحول لنفس شكل ملف الإكسل، فيمر على نفس المعاينة والتحقق."""
import io
import re
import unicodedata

_AR = re.compile(r"[؀-ۿݐ-ݿﭐ-﷿ﹰ-﻿]")


def is_pdf(data):
    return bytes(data[:5]) == b"%PDF-"


def _fix_rtl(text):
    """نص عربي من PDF يطلع بترتيب بصري معكوس: نعكس حروف كل كلمة عربية وترتيب الكلمات العربية المتتالية، ثم نوحّد أشكال الحروف."""
    if not text or not _AR.search(text):
        return text or ""
    out_lines = []
    for line in text.split("\n"):
        toks = line.split(" ")
        n_ar = sum(1 for t in toks if _AR.search(t))
        if n_ar and n_ar * 2 >= len([t for t in toks if t]):          # سطر عربي بالأساس: القراءة من اليمين، فنعكس ترتيب كل الكلمات
            toks = [t[::-1] if _AR.search(t) else t for t in reversed(toks)]
            out_lines.append(unicodedata.normalize("NFKC", " ".join(toks)))
            continue
        toks = [t[::-1] if _AR.search(t) else t for t in toks]
        res, run = [], []
        for t in toks:
            if _AR.search(t):
                run.append(t)
            else:
                res.extend(reversed(run)); run = []
                res.append(t)
        res.extend(reversed(run))
        out_lines.append(" ".join(res))
    return unicodedata.normalize("NFKC", "\n".join(out_lines))


def _clean(text):
    t = re.sub(r"\(cid:\d+\)", ".", str(text or ""))
    return _fix_rtl(t)


def _num(text):
    """رقم من نص مثل «$ 1,110.00» أو «12.70» أو «1 500». يرجّع None إذا ما فيه رقم."""
    t = _clean(text).replace("$", " ").replace(",", "").strip()
    t = re.sub(r"(?<=\d)\s+(?=\d{3}\b)", "", t) if re.fullmatch(r"\d{1,3}(\s\d{3})+", t) else t
    m = re.search(r"-?\d+(?:\.\d+)?", t)
    return float(m.group()) if m else None


def _join_lines(text):
    lines = [l.strip() for l in _clean(text).split("\n") if l.strip()]
    out = ""
    for l in lines:
        out = (out + l) if out.endswith("-") and not _AR.match(l) else (out + " " + l if out else l)
    return re.sub(r"\s+", " ", out).strip()


# ---------------------------------------------------------------- مفاتيح الأعمدة
def _aliases():
    from importer import _norm, _INV_ALIAS
    extra = {"الكمية المطلوبة": "qty_req", "المبلغ": "total", "المجموع": "total", "t.price": "total", "total": "total", "t.price (usd)": "total",
             "الكمية المباعة": "qty", "pcs/box": "pcs", "pcs/bo": "pcs", "item no": "idx", "item no.": "idx", "no": "idx", "#": "idx",
             "brand": "brand", "باركود": "barcode", "picture": "pic", "الصورة": "pic", "unit": "unit"}
    al = dict(_INV_ALIAS)
    for k, v in extra.items():
        al[_norm(k)] = v
    al.pop("", None)                                       # رموز مثل # و \ و - تتحول لنص فاضي بعد التطبيع
    al.pop(_norm("pcs"), None)                             # «PCS» داخل الجدول وحدة مو عنوان كمية
    return _norm, al


def _hdr_key(text, _norm, al):
    return al.get(_norm(text))


# ---------------------------------------------------------------- الاستراتيجية أ: جداول بخطوط (فواتير بشبكة)
def _from_tables(page, _norm, al):
    out = []
    for tb in page.find_tables():
        data = tb.extract()
        if not data or len(data) < 2 or len(data[0]) < 3:
            continue
        hdr_i, cand = None, {}
        for i, row in enumerate(data[:3]):
            f = {}
            for j, c in enumerate(row):
                k = _hdr_key(_join_lines(c), _norm, al)
                if k:
                    f.setdefault(k, []).append(j)
            if "qty" in f and ("code" in f or "name" in f):
                hdr_i, cand = i, f
                break
        if hdr_i is None:
            continue
        body = data[hdr_i + 1:]

        def score(k, j):
            vals = [_join_lines(r[j]) if j < len(r) else "" for r in body]
            if k == "name":
                return sum(len(re.sub(r"[\d\W_]", "", v)) for v in vals)
            if k == "code":
                return sum(1 for v in vals if re.fullmatch(r"[A-Za-z0-9][\w\-\s]{2,}", v) and re.search(r"\d|[A-Za-z]{3}", v))
            return sum(1 for v in vals if _num(v) is not None)
        ncol = max(len(r) for r in data)
        allc = list(range(ncol))
        cols = {}
        for k, js in cand.items():
            pool = js
            if k == "name":
                pool = sorted(set(js) | {j + 1 for j in js} | {j - 1 for j in js if j > 0})      # عمود الوصف أحياناً ترويسته مدموجة مع اللي جنبه
                pool = [j for j in pool if 0 <= j < ncol]
            elif k == "code":
                pool = sorted(set(js) | {j + 1 for j in js} | {j - 1 for j in js if j > 0})
                pool = [j for j in pool if 0 <= j < ncol]
            cols[k] = max(pool, key=lambda j: score(k, j))
        if "name" in cols and "code" in cols and cols["name"] == cols["code"]:
            others = [j for j in allc if j != cols["code"]]
            cols["name"] = max(others, key=lambda j: score("name", j))
        for i in range(hdr_i + 1, len(data)):
            row = data[i]
            get = lambda k: row[cols[k]] if k in cols and cols[k] < len(row) else None
            code = _join_lines(get("code")).replace(" ", "")
            name = _join_lines(get("name"))
            qty, price = _num(get("qty")), _num(get("price"))
            if qty is None and price is None:
                continue
            if not code and not name:
                continue
            bbox = tb.rows[i].bbox if i < len(tb.rows) else None
            out.append(dict(code=code, name=name, qty=qty, price=price, bbox=bbox))
    return out


# ---------------------------------------------------------------- الاستراتيجية ب: كلمات بمواقعها (عروض بدون خطوط)
def _from_words(page, _norm, al):
    words = page.extract_words(x_tolerance=1.5, y_tolerance=2, keep_blank_chars=False)
    for w in words:
        w["t"] = _clean(w["text"])
    lines = {}
    for w in words:
        lines.setdefault(round(w["top"] / 3.5), []).append(w)
    L = [sorted(v, key=lambda w: w["x0"]) for _, v in sorted(lines.items())]
    # ترويسة: سطر فيه كمية وسعر واسم/كود
    hdr_i, hits = None, []
    for i, l in enumerate(L):
        found = []
        for w in l:
            k = _hdr_key(w["t"], _norm, al)
            if k:
                found.append((k, w))
        for a, b in zip(l, l[1:]):                                   # عبارتين متتاليتين (مثل Product Description)
            k = _hdr_key(a["t"] + " " + b["t"], _norm, al)
            if k:
                found.append((k, dict(a, x1=b["x1"])))
        keys = {k for k, _ in found}
        if "qty" in keys and "price" in keys and ("name" in keys or "code" in keys):
            hdr_i, hits = i, found
            break
    if hdr_i is None:
        return []
    top = L[hdr_i][0]["top"]
    for l in L[hdr_i + 1:hdr_i + 3]:                                  # الكلمة الثانية للترويسة المقسومة سطرين (Zad / Code)
        if l[0]["top"] - top < 16:
            for w in l:
                k = _hdr_key(w["t"], _norm, al)
                if k and k not in {x for x, _ in hits}:
                    hits.append((k, w))
    cols = {}
    for k, w in hits:
        cols.setdefault(k, w)
    numeric = {"qty", "price", "total", "pcs", "qty_req", "unit", "idx"}
    num_cols = {k: (w["x0"] + w["x1"]) / 2 for k, w in cols.items() if k in numeric}
    txt_cols = sorted(((w["x0"], k) for k, w in cols.items() if k not in numeric))

    def col_of(w):
        if _num(w["t"]) is not None and re.fullmatch(r"[$]?[\d,.\s]+", w["t"]) and num_cols:
            cx = (w["x0"] + w["x1"]) / 2
            k = min(num_cols, key=lambda k: abs(num_cols[k] - cx))
            if abs(num_cols[k] - cx) <= 34 or not txt_cols:
                return k
        if w["t"] in ("PCS", "PC", "قطعة") and "unit" in num_cols:
            return "unit"
        c = None
        for x0, k in txt_cols:
            if w["x0"] >= x0 - 6:
                c = k
        return c
    body = L[hdr_i + 1:]
    while body and body[0][0]["top"] - top < 16 and all(_hdr_key(w["t"], _norm, al) for w in body[0]):
        body = body[1:]                                               # باقي الترويسة (سطر ثاني)
    for i, l in enumerate(body):                                      # نهاية الجدول: سطر المجموع
        txt = " ".join(w["t"] for w in l)
        if re.search(r"\bPage\s*\d+\s*(of|/)\s*\d+|www\.", txt, re.I) or (any(re.fullmatch(r"(total|discount|net|المجموع|الخصم)[:]?", w["t"].lower()) for w in l) and not any(col_of(w) == "qty" for w in l if re.fullmatch(r"\d+", w["t"]))):
            body = body[:i]
            break
    cells = []
    for l in body:
        d = {}
        for w in l:
            k = col_of(w)
            if k:
                d.setdefault(k, []).append(w["t"])
        cells.append((l[0]["top"], l[-1]["bottom"], d))
    anchors = [i for i, (_, _, d) in enumerate(cells) if "qty" in d and "price" in d and _num(" ".join(d["qty"])) is not None and _num(" ".join(d["price"])) is not None]
    if not anchors:
        return []
    first_text = any("name" in cells[i][2] for i in range(0, anchors[0]))
    def split_between(i0, i1):
        """حد الفصل بين صفين: أكبر فراغ عمودي بين أسطر الصفين (وإلا النص)."""
        tops = sorted({cells[i0][0], cells[i1][0], *[t for (t, _, _) in cells if cells[i0][0] < t < cells[i1][0]]})
        gaps = [(tops[k + 1] - tops[k], k) for k in range(len(tops) - 1)]
        g, k = max(gaps)
        return (tops[k] + tops[k + 1]) / 2 if g >= 14 else (cells[i0][0] + cells[i1][0]) / 2
    spans = []
    for n, a in enumerate(anchors):
        if first_text:                                                 # الوصف يبدأ قبل سطر الكمية والسعر (مثل Zad): نقسم بالنص بين الصفوف
            lo = 0 if n == 0 else split_between(anchors[n - 1], a)
            hi = 1e9 if n == len(anchors) - 1 else split_between(a, anchors[n + 1])
        else:                                                          # سطر عنوان الصف ثم التفاصيل تحته (مثل Atlas)
            lo = cells[a][0] - 2
            hi = 1e9 if n == len(anchors) - 1 else cells[anchors[n + 1]][0] - 2
        spans.append((lo, hi))
    out = []
    code_rx = re.compile(r"\b[A-Za-z]{1,4}\d{1,3}-\d{2,4}\b")
    for n, a in enumerate(anchors):
        lo, hi = spans[n]
        rows = [(t, b, d) for (t, b, d) in cells if lo <= t < hi]
        name = " ".join(" ".join(d.get("name", [])) for _, _, d in rows).strip()
        code = " ".join(" ".join(d.get("code", [])) for _, _, d in rows).replace(" ", "")
        if not code:
            blob = " ".join(w for _, _, d in rows for ws in d.values() for w in ws)
            m = code_rx.findall(blob)
            code = m[-1] if m else ""
        d = cells[a][2]
        out.append(dict(code=code, name=re.sub(r"\s+", " ", name), qty=_num(" ".join(d["qty"])), price=_num(" ".join(d["price"])),
                        bbox=(0, lo if lo > 0 else cells[a][0] - 12, page.width, hi if hi < 1e8 else page.height - 20)))
    return out


# ---------------------------------------------------------------- الخصم والمجموع
def _footer(pages_text):
    disc = tot = None
    for line in pages_text.split("\n"):
        low = line.lower()
        if re.search(r"discount|الخصم", low) and disc is None:
            nums = re.findall(r"\d[\d,]*(?:\.\d+)?", _clean(line).replace(",", ""))
            if nums:
                disc = float(nums[0])
    return disc


def _image_for(page, bbox, used):
    """أكبر صورة داخل مستطيل الصف (بدون الشعار) كـ PNG."""
    if not bbox:
        return None
    x0, y0, x1, y1 = bbox
    best, area = None, 0
    for im in page.images:
        cy = (im["top"] + im["bottom"]) / 2
        w, h = im["x1"] - im["x0"], im["bottom"] - im["top"]
        if y0 <= cy <= y1 and w > 8 and h > 8 and id(im) not in used and not (w > 150 and h < 60):
            if w * h > area:
                best, area = im, w * h
    if best is None:
        return None
    used.add(id(best))
    try:
        from PIL import Image, ImageChops
        bb = (best["x0"] + 1, best["top"] + 1, best["x1"] - 1, best["bottom"] - 1)
        img = page.crop(bb).to_image(resolution=260).original.convert("RGB")
        bg = Image.new("RGB", img.size, (255, 255, 255))
        box = ImageChops.difference(img, bg).convert("L").point(lambda v: 255 if v > 18 else 0).getbbox()
        if box:
            img = img.crop((max(0, box[0] - 4), max(0, box[1] - 4), min(img.width, box[2] + 4), min(img.height, box[3] + 4)))
        img.thumbnail((260, 260))
        buf = io.BytesIO()
        img.save(buf, "PNG")
        return buf.getvalue()
    except Exception:
        return None


def _pick_code(text):
    """كود المادة من خلية قد تحتوي كلام زيادة (ترويسة أو عربي)."""
    toks = [t for t in re.findall(r"[A-Za-z0-9][A-Za-z0-9\-_/.]*", _clean(text)) if re.search(r"\d", t)]
    return max(toks, key=len) if toks else _clean(text).replace(" ", "")


def _split_name(name):
    """اسم لاتيني + وصف عربي بنفس الخلية: الاسم بالأساس لاتيني والعربي استخدام، أو العكس."""
    toks = name.split()
    ar = [t for t in toks if _AR.search(t)]
    lat = [t for t in toks if not _AR.search(t)]
    latin_txt = " ".join(lat)
    if ar and len(re.findall(r"[A-Za-z]", latin_txt)) >= 10 and not re.fullmatch(r"\(.*\)", latin_txt.strip()):
        return re.sub(r"\s+", " ", latin_txt).strip(), " ".join(ar)
    if ar and lat:                                                     # عربي + (BARREL): الاسم عربي والباقي بالنهاية
        return " ".join(ar) + " " + latin_txt, ""
    return name, ""


def _currency(text):
    t = text or ""
    if re.search(r"\bUSD\b|\$|دولار", t):
        return "دولار"
    if re.search(r"\bIQD\b|دينار", t):
        return "دينار"
    return ""


_CACHE = {}


def extract_cached(data, with_images=True):
    import hashlib
    k = (hashlib.md5(bytes(data)).hexdigest(), with_images)
    if k not in _CACHE:
        if len(_CACHE) > 3:
            _CACHE.pop(next(iter(_CACHE)))
        _CACHE[k] = extract(data, with_images)
    return _CACHE[k]


def extract(data, with_images=True):
    """يرجّع {rows:[{code,name,qty,price,image}], discount}. يرفع ValueError إذا ما لقى جدول مواد."""
    import pdfplumber
    _norm, al = _aliases()
    rows, text, used = [], [], set()
    try:
        pdf = pdfplumber.open(io.BytesIO(data))
    except Exception:
        raise ValueError("هذا مو ملف PDF صحيح")
    with pdf:
        for page in pdf.pages:
            text.append(page.extract_text() or "")
            recs = _from_tables(page, _norm, al) or _from_words(page, _norm, al)
            for r in recs:
                r["image"] = _image_for(page, r.pop("bbox", None), used) if with_images else None
                r["code"] = _pick_code(r["code"]) if r["code"] and not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9\-_/.]*", r["code"]) else r["code"]
                if r["code"]:
                    r["name"] = re.sub(r"\s*" + re.escape(r["code"]) + r"\s*", " ", r["name"])            # الكود ما يتكرر داخل الاسم
                r["name"], r["usage"] = _split_name(re.sub(r"\s+", " ", r["name"]).strip())
                rows.append(r)
    if not rows:
        raise ValueError("ما لقيت جدول مواد بملف الـ PDF (الكود أو المادة + الكمية + السعر). جرّب ملف إكسل")
    alltext = "\n".join(text)
    return {"rows": rows, "discount": _footer(alltext), "currency": _currency(alltext)}
