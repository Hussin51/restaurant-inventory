"""استيراد المواد من ملف إكسل (قالب الاستيراد) + بناء القالب الفاضي."""
import io
import math
import os
import re
from core import norm_currency

SHEET = "المواد"
EXAMPLE_CODES = {"مثال", "مثال-001", "example"}

# (المفتاح الداخلي، العنوان بالقالب، العرض، ملاحظة بالتعليمات، أسماء بديلة مقبولة)
COLUMNS = [
    ("code", "الكود *", 14, "كود المادة. يكتبه انت ولا يتكرر. مثلاً PL-001", ["الكود", "كود", "رمز", "code"]),
    ("name", "اسم المادة *", 28, "اسم المادة. لا يتكرر", ["اسم المادة", "الاسم", "اسم", "المادة", "name"]),
    ("category", "التصنيف", 18, "صحون وكاسات / أدوات مائدة / معدات مطبخ / أثاث وأجهزة (أو أي تصنيف تكتبه)", ["التصنيف", "الفئة", "category"]),
    ("usage", "الاستخدام", 26, "وين تنستخدم المادة. مثلاً كاسة شوربة", ["الاستخدام", "استخدام", "usage"]),
    ("carton", "قطع بالكارتون", 14, "كم قطعة بالكارتون (اتركه فاضي إذا ماكو)", ["قطع بالكارتون", "عدد القطع بالكارتون", "قطعة بالكارتون", "carton"]),
    ("shad", "قطع بالشد", 12, "كم قطعة بالشد (اتركه فاضي إذا ماكو)", ["قطع بالشد", "عدد القطع بالشد", "قطعة بالشد", "shad"]),
    ("price", "سعر القطعة", 13, "سعر القطعة الوحدة. إذا تعرف سعر الكارتون قسمه على عدد القطع", ["سعر القطعة", "السعر", "سعر", "price"]),
    ("min_p", "الحد الأدنى (قطعة)", 17, "تحت هذا العدد يطلع تنبيه «ناقص». بالقطعة", ["الحد الأدنى (قطعة)", "الحد الأدنى", "حد أدنى", "min"]),
    ("min_c", "الحد الأدنى (كارتون)", 18, "نفس الحد الأدنى بس بالكارتون (يتجمع مع الي قبله)", ["الحد الأدنى (كارتون)"]),
    ("q_c", "الموجود: كارتون", 15, "كم كارتون موجود هسه بالجرد", ["الموجود: كارتون", "الكمية: كارتون", "كارتون", "عدد الكراتين"]),
    ("q_s", "الموجود: شد", 13, "كم شد موجود هسه (زيادة على الكراتين)", ["الموجود: شد", "الكمية: شد", "شد"]),
    ("q_p", "الموجود: قطعة", 15, "كم قطعة موجودة هسه (زيادة على الكراتين والشدات)", ["الموجود: قطعة", "الكمية: قطعة", "الكمية", "الموجود", "قطعة"]),
    ("supplier", "المورد المعتاد", 20, "اسم المورد (إذا مو مسجل ينضاف لحاله بدون هاتف)", ["المورد المعتاد", "المورد", "supplier"]),
    ("currency", "عملة السعر", 12, "عملة «سعر القطعة»: دينار أو دولار (الفاضي = دينار)", ["عملة السعر", "العملة", "currency"]),
]
_DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹٫٬،", "01234567890123456789.,,")


def _norm(s):
    s = str(s or "").strip().lower().translate(_DIGITS)
    s = re.sub("[ً-ْـ]", "", s)                       # التشكيل والتطويل
    s = s.translate(str.maketrans("أإآىةؤئ", "ااايهوي"))
    return re.sub(r"[^0-9a-zء-ي]", "", s)


_ALIAS = {}
for _k, _t, _w, _n, _al in COLUMNS:
    for _a in [_t.replace("*", "")] + _al:
        _ALIAS.setdefault(_norm(_a), _k)


def _num(v, what, row_errs, integer=False):
    if v is None or (isinstance(v, str) and not v.strip()):
        return None
    if isinstance(v, bool):
        row_errs.append(f"{what}: قيمة غير صحيحة")
        return None
    if isinstance(v, str):
        v = v.strip().translate(_DIGITS).replace(",", "")
        try:
            v = float(v)
        except ValueError:
            row_errs.append(f"{what}: «{v}» مو رقم")
            return None
    try:
        v = float(v)
    except (TypeError, ValueError):
        row_errs.append(f"{what}: مو رقم")
        return None
    if not math.isfinite(v) or v < 0:
        row_errs.append(f"{what}: لازم رقم مو سالب")
        return None
    if integer and v != int(v):
        row_errs.append(f"{what}: لازم رقم صحيح")
        return None
    return int(v) if v == int(v) else v


def _txt(v):
    if v is None:
        return ""
    if isinstance(v, float) and v == int(v):
        v = int(v)
    t = str(v).strip()
    while t.startswith("="):
        t = t[1:].lstrip()
    return t


def parse(data, existing_codes=None, existing_names=None, suppliers=None):
    """يقرا ملف الإكسل ويرجع {rows, counts}. ما يغيّر شي.
    existing_codes: {code.casefold(): name}، existing_names: {name: code}، suppliers: set بالأسماء."""
    from openpyxl import load_workbook
    existing_codes, existing_names = existing_codes or {}, existing_names or {}
    sup_known = {s.casefold() for s in (suppliers or [])}
    try:
        wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    except Exception:
        raise ValueError("هذا مو ملف إكسل (.xlsx) صحيح")
    ws = wb[SHEET] if SHEET in wb.sheetnames else wb[wb.sheetnames[0]]
    rows_iter = list(ws.iter_rows(values_only=True))
    hdr_i, cols = None, {}
    for i, row in enumerate(rows_iter[:15]):
        found = {}
        for j, v in enumerate(row):
            k = _ALIAS.get(_norm(v))
            if k and k not in found:
                found[k] = j
        if "code" in found and "name" in found:
            hdr_i, cols = i, found
            break
    if hdr_i is None:
        raise ValueError("ما لقيت عناوين الأعمدة (الكود، اسم المادة). استخدم القالب من زر «قالب إكسل»")
    out, seen_codes, seen_names = [], {}, {}
    new_suppliers = set()
    for i, row in enumerate(rows_iter[hdr_i + 1:], start=hdr_i + 2):
        get = lambda k: row[cols[k]] if k in cols and cols[k] < len(row) else None
        if all(c is None or str(c).strip() == "" for c in row):
            continue
        code, name = _txt(get("code")), _txt(get("name"))
        if code.casefold() in {x.casefold() for x in EXAMPLE_CODES}:
            continue                                                  # سطر المثال
        errs = []
        if not code:
            errs.append("الكود مطلوب")
        if not name:
            errs.append("اسم المادة مطلوب")
        carton = _num(get("carton"), "قطع بالكارتون", errs, True)
        shad = _num(get("shad"), "قطع بالشد", errs, True)
        price = _num(get("price"), "سعر القطعة", errs)
        min_p, min_c = _num(get("min_p"), "الحد الأدنى (قطعة)", errs), _num(get("min_c"), "الحد الأدنى (كارتون)", errs)
        q_c, q_s, q_p = _num(get("q_c"), "الموجود: كارتون", errs), _num(get("q_s"), "الموجود: شد", errs), _num(get("q_p"), "الموجود: قطعة", errs)
        for v, f, lab in ((min_c, carton, "الحد الأدنى بالكارتون"), (q_c, carton, "الموجود بالكارتون")):
            if v and not f:
                errs.append(f"{lab}: حدد «قطع بالكارتون» أول")
        if q_s and not shad:
            errs.append("الموجود بالشد: حدد «قطع بالشد» أول")
        key = code.casefold()
        if code and key in seen_codes:
            errs.append(f"الكود مكرر بالملف (سطر {seen_codes[key]})")
        elif code:
            seen_codes[key] = i
        if name and name in seen_names:
            errs.append(f"الاسم مكرر بالملف (سطر {seen_names[name]})")
        elif name:
            seen_names[name] = i
        status = "error" if errs else "new"
        if not errs and key in existing_codes:
            status = "exists"
            if existing_codes[key] != name:
                pass                                                  # وضع التحديث ما يغيّر الاسم
        elif not errs and name in existing_names:
            errs.append(f"الاسم مستخدم لمادة ثانية (كودها {existing_names[name]})")
            status = "error"
        cur_raw = _txt(get("currency"))
        try:
            currency = norm_currency(cur_raw, True)
        except ValueError:
            currency = "دينار"
            errs.append("عملة السعر لازم «دينار» أو «دولار»")
            status = "error"
        sup = _txt(get("supplier"))
        if sup and sup.casefold() not in sup_known and not errs:
            new_suppliers.add(sup)
        opening = (q_c or 0) * (carton or 0) + (q_s or 0) * (shad or 0) + (q_p or 0)
        minimum = (min_p or 0) + (min_c or 0) * (carton or 0)
        out.append(dict(row=i, code=code, name=name, category=_txt(get("category")) or "غير مصنف", usage=_txt(get("usage")),
                        carton=carton or 0, shad=shad or 0, price=price or 0, opening=opening, min=minimum, supplier=sup, currency=currency,
                        has=dict(currency=bool(cur_raw), price=price is not None, min=(min_p is not None or min_c is not None), carton=carton is not None,
                                 shad=shad is not None, usage=bool(_txt(get("usage"))), supplier=bool(sup), category=bool(_txt(get("category")))),
                        status=status, message="؛ ".join(errs)))
    counts = {s: sum(1 for r in out if r["status"] == s) for s in ("new", "exists", "error")}
    counts["total"] = len(out)
    counts["new_suppliers"] = sorted(new_suppliers)
    return {"rows": out, "counts": counts}


def extract_images(data):
    """صور مدمجة بملف الإكسل (مربوطة بسطر المادة): {رقم السطر: (bytes, 'image/png')}. أي فشل = بدون صور."""
    out = {}
    try:
        from openpyxl import load_workbook
        wb = load_workbook(io.BytesIO(data))
        ws = wb[SHEET] if SHEET in wb.sheetnames else wb[wb.sheetnames[0]]
        for im in getattr(ws, "_images", []):
            try:
                row = im.anchor._from.row + 1
                raw = im._data()
                fmt = (getattr(im, "format", "") or "png").lower()
                out.setdefault(row, (raw, "image/jpeg" if fmt in ("jpg", "jpeg") else "image/" + fmt))
            except Exception:
                continue
    except Exception:
        pass
    return out


# ---------------------------------------------------------------- استيراد فاتورة شراء من إكسل
INV_COLS = {
    "code": ["الكود", "كود", "كود المادة", "رمز", "رقم المادة", "code", "item code", "item no", "item no.", "sku", "zad code", "mould no", "model"],
    "name": ["المادة", "اسم المادة", "تفاصيل المادة", "الاسم", "الوصف", "البيان", "name", "description", "product description", "item", "item name", "product"],
    "qty": ["الكمية", "الكمية المباعة", "الكميه", "العدد", "qty", "quantity", "qtty", "pcs", "sales qtty", "sales qtty."],
    "price": ["السعر", "سعر البيع", "سعر القطعة", "سعر الوحدة", "سعر", "price", "unit price", "u.price", "u price"],
    "unit": ["الوحدة", "unit"],
}
_INV_ALIAS = {}
for _k, _al in INV_COLS.items():
    for _a in _al:
        _INV_ALIAS.setdefault(_norm(_a), _k)


def invoice_template():
    """قالب بسيط لفاتورة شراء: الكود، المادة، الكمية، سعر القطعة."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    wb = Workbook()
    ws = wb.active
    ws.title = "الفاتورة"
    ws.sheet_view.rightToLeft = True
    hs = ["الكود", "المادة", "الكمية", "سعر القطعة", "الوحدة"]
    for j, h in enumerate(hs, 1):
        c = ws.cell(1, j, h)
        c.font, c.fill, c.alignment = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="1F4E78"), Alignment(horizontal="center")
        ws.column_dimensions[c.column_letter].width = [16, 34, 12, 14, 12][j - 1]
    for j, v in enumerate(["PL-001", "صحن تقديم", 120, 2500, "قطعة"], 1):
        ws.cell(2, j, v).alignment = Alignment(horizontal="center")
    ws["G1"] = "الوحدة اختيارية (قطعة أو كارتون أو شد). إذا تركتها فاضية تُعتبر قطعة."
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def parse_invoice(data, by_code=None, by_name=None):
    """يقرأ فاتورة شراء بإكسل (أي ترتيب أعمدة): يرجّع {lines, counts, header}. ما يغيّر شي.
    by_code: {code.casefold(): item}، by_name: {name.casefold(): item}."""
    from openpyxl import load_workbook
    by_code, by_name = by_code or {}, by_name or {}
    try:
        wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    except Exception:
        raise ValueError("هذا مو ملف إكسل (.xlsx) صحيح")
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    hdr, cols = None, {}
    for i, row in enumerate(rows[:40]):
        found = {}
        for j, v in enumerate(row):
            k = _INV_ALIAS.get(_norm(v))
            if k and k not in found:
                found[k] = j
        if "qty" in found and ("code" in found or "name" in found):
            hdr, cols = i, found
            break
    if hdr is None:
        raise ValueError("ما لقيت أعمدة الفاتورة (الكود أو المادة + الكمية). استخدم «قالب فاتورة شراء» أو تأكد من العناوين")
    if "code" in cols and "name" in cols:                           # عمود «الكود» اللي فيه أرقام تسلسلية 1,2,3 = ترقيم سطور مو كود
        vals = [rows[r][cols["code"]] if cols["code"] < len(rows[r]) else None for r in range(hdr + 1, len(rows))]
        vals = [str(v).strip().rstrip(".0") for v in vals if v is not None and str(v).strip() != ""]
        if vals and all(v.isdigit() for v in vals) and [int(v) for v in vals][:5] == list(range(1, min(5, len(vals)) + 1)):
            del cols["code"]
    out = []
    for i, row in enumerate(rows[hdr + 1:], start=hdr + 2):
        get = lambda k: row[cols[k]] if k in cols and cols[k] < len(row) else None
        code, name = _txt(get("code")), _txt(get("name"))
        errs = []
        qty = _num(get("qty"), "الكمية", errs)
        if not code and not name:
            continue
        if qty is None and not errs:
            if all(c is None or str(c).strip() == "" for c in (get("qty"), get("price"))):
                continue                                            # سطر بدون كمية وسعر (عنوان فرعي أو سطر فاضي)
        price = _num(get("price"), "السعر", errs)
        if qty is not None and qty <= 0:
            errs.append("الكمية لازم أكبر من صفر")
        unit = _txt(get("unit")) or "قطعة"
        if unit not in ("قطعة", "كارتون", "شد"):
            unit = "قطعة"
        it = by_code.get(code.casefold()) if code else None
        if it is None and name:
            it = by_name.get(name.casefold())
        status = "error" if (errs and any("بدون كود" not in e for e in errs)) or qty is None else ("exists" if it else "new")
        if status == "new" and not code:
            errs.append("بدون كود: ينعمل كود تلقائي")
        out.append(dict(row=i, code=(it["code"] if it else code), name=(it["name"] if it else name), qty=qty or 0, price=price or 0, unit=unit,
                        status=status, message="؛ ".join(errs)))
    counts = {s: sum(1 for r in out if r["status"] == s) for s in ("exists", "new", "error")}
    counts["total"] = len(out)
    return {"lines": out, "counts": counts}


# ---------------------------------------------------------------- استيراد التكسير (نهاية الشهر)
BK_COLS = {
    "dept": ["القسم", "قسم", "department", "dept"],
    "code": INV_COLS["code"],
    "name": INV_COLS["name"],
    "qty": ["الكمية", "الكسر", "كمية الكسر", "المكسور", "التكسير", "العدد", "qty", "quantity", "broken"],
    "date": ["التاريخ", "date"],
    "note": ["ملاحظة", "ملاحظات", "السبب", "note", "reason"],
}
_BK_ALIAS = {}
for _k, _al in BK_COLS.items():
    for _a in _al:
        _BK_ALIAS.setdefault(_norm(_a), _k)


def breakage_template(depts):
    """قالب التكسير: ورقة وحدة بعمود القسم (أو ورقة لكل قسم باسمه)."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    wb = Workbook()
    ws = wb.active
    ws.title = "التكسير"
    ws.sheet_view.rightToLeft = True
    hs = ["القسم", "الكود", "المادة", "الكمية", "ملاحظة"]
    for j, h in enumerate(hs, 1):
        c = ws.cell(1, j, h)
        c.font, c.fill, c.alignment = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="1F4E78"), Alignment(horizontal="center")
        ws.column_dimensions[c.column_letter].width = [16, 16, 34, 12, 28][j - 1]
    for j, v in enumerate([(depts or ["المطبخ"])[0], "PL-001", "صحن تقديم", 3, "كسر بالغسيل"], 1):
        ws.cell(2, j, v).alignment = Alignment(horizontal="center")
    ws["G1"] = "اكتب القسم بكل سطر، أو سوّي ورقة لكل قسم باسم القسم. الكود أو اسم المادة (واحد منهم يكفي). الكمية بالقطعة."
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def parse_breakage(data, by_code=None, by_name=None, depts=None):
    """يقرأ ملف التكسير (أي ترتيب أعمدة؛ القسم من عمود «القسم» أو من اسم الورقة). يرجّع {lines, counts}. ما يغيّر شي."""
    from openpyxl import load_workbook
    by_code, by_name = by_code or {}, by_name or {}
    dmap = {_norm(d): d for d in (depts or [])}
    try:
        wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    except Exception:
        raise ValueError("هذا مو ملف إكسل (.xlsx) صحيح")
    out, found_any = [], False
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        hdr, cols = None, {}
        for i, row in enumerate(rows[:40]):
            f = {}
            for j, v in enumerate(row):
                k = _BK_ALIAS.get(_norm(v))
                if k and k not in f:
                    f[k] = j
            if "qty" in f and ("code" in f or "name" in f):
                hdr, cols = i, f
                break
        if hdr is None:
            continue
        found_any = True
        sheet_dept = dmap.get(_norm(ws.title), "")
        for i, row in enumerate(rows[hdr + 1:], start=hdr + 2):
            get = lambda k: row[cols[k]] if k in cols and cols[k] < len(row) else None
            code, name = _txt(get("code")), _txt(get("name"))
            if not code and not name:
                continue
            errs = []
            qty = _num(get("qty"), "الكمية", errs)
            if qty is None and not errs:
                continue                                                    # سطر بدون كمية = ما كسر
            if qty == 0 and not errs:
                continue                                                    # صفر = ما كسر شي
            dept_raw = _txt(get("dept"))
            dept = dmap.get(_norm(dept_raw), "") if dept_raw else sheet_dept
            if not dept:
                errs.append("القسم مو معروف" + (f": «{dept_raw}»" if dept_raw else " (اكتب عمود القسم أو سمّي الورقة باسم القسم)"))
            it = by_code.get(code.casefold()) if code else None
            if it is None and name:
                it = by_name.get(name.casefold())
            if it is None:
                errs.append("المادة مو موجودة بالبرنامج")
            date_v = get("date")
            ds = ""
            if hasattr(date_v, "strftime"):
                ds = date_v.strftime("%Y-%m-%d")
            elif _txt(date_v):
                ds = _txt(date_v)[:10]
            out.append(dict(sheet=ws.title, row=i, dept=dept, code=(it["code"] if it else code), name=(it["name"] if it else name), qty=qty or 0,
                            date=ds, note=_txt(get("note")), status="error" if errs else "ok", message="؛ ".join(errs)))
    if not found_any:
        raise ValueError("ما لقيت أعمدة التكسير (الكود أو المادة + الكمية). استخدم «قالب التكسير»")
    counts = {s: sum(1 for r in out if r["status"] == s) for s in ("ok", "error")}
    counts["total"] = len(out)
    counts["qty"] = sum(r["qty"] for r in out if r["status"] == "ok")
    return {"lines": out, "counts": counts}


def build_template(categories):
    """ملف إكسل فاضي للتعبئة: ورقة «المواد» وورقة «تعليمات» (فيها مثال وشرح الأعمدة)."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.worksheet.datavalidation import DataValidation
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET
    ws.sheet_view.rightToLeft = True
    F = "Arial"
    thin = Side(style="thin", color="BFBFBF")
    bd = Border(left=thin, right=thin, top=thin, bottom=thin)
    hfill = PatternFill("solid", fgColor="1F4E78")
    ex_fill = PatternFill("solid", fgColor="FFF2CC")
    ws["A1"] = "قالب استيراد المواد - مستلزمات مطعم بيت الكل"
    ws["A1"].font = Font(name=F, bold=True, size=14, color="1F4E78")
    ws["A2"] = "عبّي من السطر 5 نازل. الأعمدة الي عليها * مطلوبة. السطر الأصفر مثال وما ينستورد. اقرا ورقة «تعليمات»."
    ws["A2"].font = Font(name=F, italic=True, size=10)
    hr = 4
    for j, (k, title, w, note, al) in enumerate(COLUMNS, 1):
        c = ws.cell(hr, j, title)
        c.font, c.fill, c.border = Font(name=F, bold=True, color="FFFFFF"), hfill, bd
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[c.column_letter].width = w
    ws.row_dimensions[hr].height = 32
    example = ["مثال", "كاسة شوربة", categories[0] if categories else "صحون وكاسات", "تقديم الشوربة", 48, 12, 1500, 0, 1, 2, 3, 5, "", "دينار"]
    for j, v in enumerate(example, 1):
        c = ws.cell(hr + 1, j, v)
        c.font, c.fill, c.border = Font(name=F, italic=True, color="7F7F7F"), ex_fill, bd
        c.alignment = Alignment(horizontal="center")
    for r in range(hr + 2, hr + 1002):
        for j in range(1, len(COLUMNS) + 1):
            c = ws.cell(r, j)
            c.border = bd
            c.font = Font(name=F, color="0000FF")
            c.alignment = Alignment(horizontal="center")
    ws.freeze_panes = ws.cell(hr + 1, 3)
    dv = DataValidation(type="list", formula1='"' + ",".join(categories) + '"', allow_blank=True, showErrorMessage=False)
    ws.add_data_validation(dv)
    dv.add(f"C{hr + 2}:C{hr + 1001}")
    cd = DataValidation(type="list", formula1='"دينار,دولار"', allow_blank=True)
    ws.add_data_validation(cd)
    cd.add(f"N{hr + 2}:N{hr + 1001}")
    nd = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True)
    nd.error = "اكتب رقم مو سالب"
    ws.add_data_validation(nd)
    for col in "EFGHIJKL":
        nd.add(f"{col}{hr + 2}:{col}{hr + 1001}")
    t = wb.create_sheet("تعليمات")
    t.sheet_view.rightToLeft = True
    t["A1"] = "شلون تعبّي القالب"
    t["A1"].font = Font(name=F, bold=True, size=14, color="1F4E78")
    tips = ["1) عبّي ورقة «المواد» من السطر 5. كل مادة بسطر. لا تغيّر أسماء العناوين ولا تمسح الأعمدة.",
            "2) الكود والاسم مطلوبين ولا يتكرران. الكود تكتبه انت (مثلاً PL-001) وما يتغير بعدين.",
            "3) الكميات كلها بالقطعة داخل البرنامج. إذا عدّيت بالكراتين اكتب بعمود «الموجود: كارتون» وحدد «قطع بالكارتون»، والبرنامج يحسب المجموع.",
            "4) «الموجود» (الكارتون + الشد + القطعة) هو الجرد الي عدّيته، ويصير الرصيد الافتتاحي.",
            "5) الحد الأدنى: تحته يطلع تنبيه «ناقص» وتدخل المادة بقائمة طلب الشراء.",
            "6) الصور: إذا دخّلت صورة داخل الإكسل (Insert ← Pictures) وخليتها على سطر المادة، تنستورد وياها. غير هيج تضيفها بعدين من نافذة المادة.",
            "7) سعر القطعة يكون بعملة المادة (عمود «عملة السعر»: دينار أو دولار، والفاضي = دينار). البرنامج يحوّل الدولار للدينار بسعر الصرف المحدد بالإعدادات.",
            "8) بعد ما تخلص، من البرنامج: الجرد والمواد ← استيراد من إكسل ← اختار الملف ← راجع الأخطاء ← استيراد."]
    for i, x in enumerate(tips, 3):
        t.cell(i, 1, x).font = Font(name=F, size=11)
    r0 = len(tips) + 5
    t.cell(r0, 1, "شرح الأعمدة").font = Font(name=F, bold=True, size=12, color="1F4E78")
    for i, (k, title, w, note, al) in enumerate(COLUMNS, r0 + 1):
        t.cell(i, 1, title.replace(" *", "")).font = Font(name=F, bold=True)
        t.cell(i, 2, note).font = Font(name=F)
    t.column_dimensions["A"].width = 26
    t.column_dimensions["B"].width = 100
    logo = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "logo-xl.png")
    if os.path.exists(logo):
        try:
            from openpyxl.drawing.image import Image as XImage
            img = XImage(logo)
            img.width = img.height = 62
            ws.add_image(img, "H1")
            ws.row_dimensions[1].height = 34
        except Exception:
            pass
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def build_phone_template():
    """قالب مبسّط للموبايل: العناوين بأول سطر وستة أعمدة بس (الكود، الاسم، الاستخدام، التعبئة، الموجود بالكارتون والقطع)."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.worksheet.datavalidation import DataValidation
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET
    ws.sheet_view.rightToLeft = True
    F = "Arial"
    thin = Side(style="thin", color="BFBFBF")
    bd = Border(left=thin, right=thin, top=thin, bottom=thin)
    cols = [("الكود *", 11), ("اسم المادة *", 26), ("الاستخدام", 18), ("قطع بالكارتون", 11), ("الموجود: كارتون", 11), ("الموجود: قطعة", 11)]
    for j, (t, w) in enumerate(cols, 1):
        c = ws.cell(1, j, t)
        c.font, c.fill, c.border = Font(name=F, bold=True, color="FFFFFF", size=12), PatternFill("solid", fgColor="1F4E78"), bd
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[c.column_letter].width = w
    ws.row_dimensions[1].height = 34
    for j, v in enumerate(["مثال", "كاسة شوربة", "تقديم الشوربة", 48, 2, 5], 1):
        c = ws.cell(2, j, v)
        c.font, c.fill, c.border, c.alignment = Font(name=F, italic=True, color="7F7F7F", size=12), PatternFill("solid", fgColor="FFF2CC"), bd, Alignment(horizontal="center")
    for r in range(3, 603):
        for j in range(1, 7):
            c = ws.cell(r, j)
            c.border, c.font, c.alignment = bd, Font(name=F, color="0000FF", size=12), Alignment(horizontal="center")
    ws.freeze_panes = "C2"
    nd = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True)
    nd.error = "اكتب رقم مو سالب"
    ws.add_data_validation(nd)
    nd.add("D3:F602")
    t = wb.create_sheet("تعليمات")
    t.sheet_view.rightToLeft = True
    t.column_dimensions["A"].width = 90
    lines = ["قالب الجرد المبسّط (للموبايل)",
             "1) كل مادة بسطر: الكود والاسم مطلوبين. الكود تكتبه انت ولا يتكرر.",
             "2) «قطع بالكارتون»: كم قطعة بالكارتون (اتركه فاضي إذا المادة بدون كارتون).",
             "3) «الموجود»: الي عدّيته هسه. اكتب الكراتين وبعدها القطع الزايدة.",
             "4) السطر الأصفر مثال وما ينستورد.",
             "5) بعد ما تخلص: ابعث الملف لجهاز الكمبيوتر (واتساب/إيميل/درايف)، وبالبرنامج: الجرد والمواد ← استيراد من إكسل.",
             "6) الأسعار والتصنيف والحد الأدنى والموردين تكملها بعدين من البرنامج (أو بالقالب الكامل)."]
    for i, x in enumerate(lines, 1):
        t.cell(i, 1, x).font = Font(name=F, bold=(i == 1), size=13 if i == 1 else 11)
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()
