"""قسم الملابس: مخزون بالمقاسات، إدخال كميات، تسليم لموظف (ورقة استلام)، وكشف لكل موظف.
كل شي ينكتب بنفس ملف الإكسل بأوراق منفصلة عن مواد المخزن."""
import re
from datetime import date, datetime

from openpyxl.styles import Font

import core
from core import CUR_IQD, _pos, _s, _t, _d, _first_empty, _make_sheet, _style_row, _F  # noqa: F401

CL_ITEMS, CL_MOVES, CL_DOCS, CL_LINES = "الملابس", "حركة_الملابس", "استلامات_الملابس", "بنود_الاستلام"
CL_BUYS, CL_BLINES = "مشتريات_الملابس", "بنود_مشتريات_الملابس"
BUY_PREFIX = "CLP"
C0, C_MAX = 5, 504                      # صفوف الملابس (500 مقاس/لون)
M_END = 20004
DOC_PREFIX = "CLO"
SIZE_ORDER = ["XXS", "XS", "S", "M", "L", "XL", "XXL", "XXXL", "2XL", "3XL", "4XL"]


# ---------------------------------------------------------------- الإعداد
def ensure_schema(wb):
    if CL_ITEMS not in wb.sheetnames:
        ws = _make_sheet(wb, CL_ITEMS, "مخزون الملابس",
                         ["الكود", "الاسم", "المقاس", "اللون", "الكمية الافتتاحية", "الحد الأدنى", "إجمالي الداخل", "إجمالي التسليم", "الرصيد", "الحالة"],
                         [12, 24, 12, 14, 16, 12, 14, 14, 12, 16])
        mv = f"'{CL_MOVES}'"
        rng = lambda col: f"{mv}!${col}$5:${col}${M_END}"
        for r in range(C0, C_MAX + 1):
            ws.cell(r, 7, f'=IF(A{r}="","",SUMIFS({rng("D")},{rng("B")},A{r},{rng("C")},"شراء")+SUMIFS({rng("D")},{rng("B")},A{r},{rng("C")},"تسوية زيادة"))')
            ws.cell(r, 8, f'=IF(A{r}="","",SUMIFS({rng("D")},{rng("B")},A{r},{rng("C")},"تسليم")+SUMIFS({rng("D")},{rng("B")},A{r},{rng("C")},"تسوية نقص"))')
            ws.cell(r, 9, f'=IF(A{r}="","",E{r}+G{r}-H{r})')
            ws.cell(r, 10, f'=IF(A{r}="","",IF(I{r}<=0,"نفد",IF(I{r}<=F{r},"ناقص - اطلب","متوفر")))')
            _style_row(ws, r, 10, formula_cols=(7, 8, 9, 10))
    if CL_MOVES not in wb.sheetnames:
        _make_sheet(wb, CL_MOVES, "حركة الملابس", ["التاريخ", "الكود", "النوع", "الكمية", "الموظف", "رقم الورقة", "ملاحظات"], [14, 12, 14, 10, 24, 14, 30])
    if CL_DOCS not in wb.sheetnames:
        _make_sheet(wb, CL_DOCS, "ورقات استلام الملابس", ["رقم الورقة", "التاريخ", "الموظف", "ملاحظات", "عدد القطع"], [14, 14, 26, 34, 12])
    if CL_LINES not in wb.sheetnames:
        _make_sheet(wb, CL_LINES, "بنود ورقات الاستلام", ["رقم الورقة", "الكود", "الاسم", "المقاس", "اللون", "الكمية"], [14, 12, 24, 12, 14, 10])
    if CL_BUYS not in wb.sheetnames:
        _make_sheet(wb, CL_BUYS, "فواتير شراء الملابس", ["رقم الفاتورة", "التاريخ", "المورد", "العملة", "سعر الصرف", "ملاحظات", "الإجمالي", "الإجمالي بالدينار"], [14, 14, 24, 12, 12, 30, 16, 18])
    if CL_BLINES not in wb.sheetnames:
        _make_sheet(wb, CL_BLINES, "بنود فواتير شراء الملابس", ["رقم الفاتورة", "الكود", "الاسم", "المقاس", "اللون", "الكمية", "سعر القطعة", "المجموع"], [14, 12, 24, 12, 14, 10, 14, 16])


# ---------------------------------------------------------------- أدوات
def _norm(s):
    t = re.sub("[ً-ْـ]", "", _s(s)).lower()
    t = t.translate(str.maketrans("أإآىة", "ااايه"))
    return re.sub(r"\s+", " ", t).strip()


def _clean(s):
    return re.sub(r"\s+", " ", _t(s)).strip()


def _size_key(size):
    s = _s(size).upper()
    if s in SIZE_ORDER:
        return (0, SIZE_ORDER.index(s), "")
    try:
        return (1, float(s), "")
    except ValueError:
        return (2, 0, s)


def _status(bal, mn):
    return "نفد" if bal <= 0 else ("ناقص - اطلب" if bal <= mn else "متوفر")


# ---------------------------------------------------------------- قراءة
def store_rate(wb):
    return core.Store._rate(wb)


def read(wb):
    ensure_schema(wb)
    ws, mw, dw, lw = wb[CL_ITEMS], wb[CL_MOVES], wb[CL_DOCS], wb[CL_LINES]
    moves, r = [], 5
    while mw.cell(r, 2).value not in (None, ""):
        moves.append(dict(row=r, date=_d(mw.cell(r, 1).value), code=_s(mw.cell(r, 2).value), type=_s(mw.cell(r, 3).value),
                          qty=core._num(mw.cell(r, 4).value), employee=_s(mw.cell(r, 5).value), doc=_s(mw.cell(r, 6).value),
                          note=_s(mw.cell(r, 7).value)))
        r += 1
    delta = {}
    for m in moves:
        sign = 1 if m["type"] in ("شراء", "تسوية زيادة") else -1 if m["type"] in ("تسليم", "تسوية نقص") else 0
        d = delta.setdefault(m["code"].casefold(), [0, 0])
        d[0 if sign > 0 else 1] += m["qty"] if sign else 0
    items = []
    for r in range(C0, C_MAX + 1):
        code = _s(ws.cell(r, 1).value)
        if not code:
            continue
        opening, mn = core._num(ws.cell(r, 5).value), core._num(ws.cell(r, 6).value)
        din, dout = delta.get(code.casefold(), [0, 0])
        bal = opening + din - dout
        items.append(dict(code=code, name=_s(ws.cell(r, 2).value), size=_s(ws.cell(r, 3).value), color=_s(ws.cell(r, 4).value),
                          opening=opening, min=mn, received=din, issued=dout, balance=bal, status=_status(bal, mn), row=r))
    by_code = {i["code"].casefold(): i for i in items}
    lines, r = {}, 5
    while lw.cell(r, 1).value not in (None, ""):
        lines.setdefault(_s(lw.cell(r, 1).value), []).append(
            dict(code=_s(lw.cell(r, 2).value), name=_s(lw.cell(r, 3).value), size=_s(lw.cell(r, 4).value),
                 color=_s(lw.cell(r, 5).value), qty=core._num(lw.cell(r, 6).value)))
        r += 1
    docs, r = [], 5
    while dw.cell(r, 1).value not in (None, ""):
        no = _s(dw.cell(r, 1).value)
        ln = lines.get(no, [])
        docs.append(dict(no=no, date=_d(dw.cell(r, 2).value), employee=_s(dw.cell(r, 3).value), note=_s(dw.cell(r, 4).value),
                         lines=ln, total=sum(l["qty"] for l in ln)))
        r += 1
    bl, r = {}, 5
    bw, blw = wb[CL_BUYS], wb[CL_BLINES]
    while blw.cell(r, 1).value not in (None, ""):
        bl.setdefault(_s(blw.cell(r, 1).value), []).append(
            dict(code=_s(blw.cell(r, 2).value), name=_s(blw.cell(r, 3).value), size=_s(blw.cell(r, 4).value), color=_s(blw.cell(r, 5).value),
                 qty=core._num(blw.cell(r, 6).value), price=core._num(blw.cell(r, 7).value)))
        r += 1
    purchases, r = [], 5
    cur_rate = store_rate(wb)
    while bw.cell(r, 1).value not in (None, ""):
        no = _s(bw.cell(r, 1).value)
        cur = core.norm_currency(bw.cell(r, 4).value)
        rate = (core._num(bw.cell(r, 5).value) or cur_rate) if cur == core.CUR_USD else 1.0
        ln = bl.get(no, [])
        for l in ln:
            l["total"] = l["qty"] * l["price"]
            l["total_iqd"] = l["total"] * rate
        purchases.append(dict(no=no, date=_d(bw.cell(r, 2).value), party=_s(bw.cell(r, 3).value), cur=cur, rate=rate, note=_s(bw.cell(r, 6).value),
                              lines=ln, total=sum(l["total"] for l in ln), total_iqd=sum(l["total"] for l in ln) * rate,
                              pieces=sum(l["qty"] for l in ln)))
        r += 1
    emps = {}
    for dd in docs:
        e = emps.setdefault(_norm(dd["employee"]), dict(name=dd["employee"], count=0, total=0, first=dd["date"], last=dd["date"]))
        e["count"] += 1
        e["total"] += dd["total"]
        e["first"] = min(str(e["first"]), str(dd["date"]))
        e["last"] = max(str(e["last"]), str(dd["date"]))
    items.sort(key=lambda i: (_norm(i["name"]), _norm(i["color"]), _size_key(i["size"])))
    return dict(items=items, docs=docs, moves=moves, purchases=purchases, employees=sorted(emps.values(), key=lambda e: _norm(e["name"])))


# ---------------------------------------------------------------- كتابة
def _next_code(items):
    mx = 0
    for i in items:
        m = re.fullmatch(r"CL-(\d+)", i["code"])
        if m:
            mx = max(mx, int(m.group(1)))
    return f"CL-{mx + 1:03d}"


def _write_move(wb, d, code, typ, qty, employee="", doc="", note=""):
    ws = wb[CL_MOVES]
    r = _first_empty(ws, 2, 5)
    ws.cell(r, 1, d).number_format = "yyyy-mm-dd"
    for c, v in ((2, code), (3, typ), (4, qty), (5, employee or None), (6, doc or None), (7, _t(note) or None)):
        ws.cell(r, c, v)
    _style_row(ws, r, 7)


def add_items(store, name, color, variants):
    """variants: [{size, opening, min}]. يضيف سطر لكل مقاس."""
    name, color = _clean(name), _clean(color)
    if not name:
        raise ValueError("اكتب اسم الملبس")
    if not variants:
        raise ValueError("اكتب مقاس واحد على الأقل (أو اتركه فاضي لملبس بدون مقاس)")
    wb = store._open()
    cur = read(wb)["items"]
    taken = {(_norm(i["name"]), _norm(i["size"]), _norm(i["color"])) for i in cur}
    seen = set()
    rows = []
    for v in variants:
        size = _clean(v.get("size"))
        key = (_norm(name), _norm(size), _norm(color))
        if key in taken or key in seen:
            raise ValueError(f"موجود من قبل: {name} {size} {color}".strip())
        seen.add(key)
        rows.append((size, _pos(v.get("opening"), "الكمية"), _pos(v.get("min"), "الحد الأدنى")))
    ws = wb[CL_ITEMS]
    code = _next_code(cur)
    n = int(code.split("-")[1])
    for size, opening, mn in rows:
        r = _first_empty(ws, 1, C0, C_MAX)
        for c, v in ((1, f"CL-{n:03d}"), (2, name), (3, size or None), (4, color or None), (5, opening), (6, mn)):
            ws.cell(r, c, v)
        n += 1
    store._save(wb)
    return len(rows)


def _item(wb, code):
    for i in read(wb)["items"]:
        if i["code"].casefold() == _s(code).casefold():
            return i
    raise ValueError(f"الملبس غير موجود: {code}")


def receive(store, code, qty, d=None, note=""):
    qty = _pos(qty, "الكمية", strict=True)
    d = datetime.strptime(d, "%Y-%m-%d").date() if isinstance(d, str) and d else (d or date.today())
    wb = store._open()
    it = _item(wb, code)
    _write_move(wb, d, it["code"], "شراء", int(qty) if qty == int(qty) else qty, note=note)
    store._save(wb)


def adjust(store, code, pieces, d=None):
    pieces = _pos(pieces, "الكمية")
    d = datetime.strptime(d, "%Y-%m-%d").date() if isinstance(d, str) and d else (d or date.today())
    wb = store._open()
    it = _item(wb, code)
    diff = pieces - it["balance"]
    if abs(diff) < 1e-9:
        return 0
    diff = int(diff) if diff == int(diff) else diff
    _write_move(wb, d, it["code"], "تسوية زيادة" if diff > 0 else "تسوية نقص", abs(diff), note="تسوية جرد")
    store._save(wb)
    return 1


def issue(store, employee, d, note, lines):
    """يسجل تسليم ملابس لموظف (ورقة استلام). الموظف ينضاف أول مرة، وإذا الاسم موجود يستخدم نفس كتابته."""
    emp = _clean(employee)
    if not emp:
        raise ValueError("اكتب اسم الموظف")
    if not lines:
        raise ValueError("اختار ملبس واحد على الأقل")
    d = datetime.strptime(d, "%Y-%m-%d").date() if isinstance(d, str) and d else (d or date.today())
    wb = store._open()
    data = read(wb)
    for e in data["employees"]:
        if _norm(e["name"]) == _norm(emp):
            emp = e["name"]                                              # نفس الموظف (يمنع تكرار الاسم بكتابة مختلفة)
            break
    by = {i["code"].casefold(): i for i in data["items"]}
    out = []
    for l in lines:
        it = by.get(_s(l.get("code")).casefold())
        if it is None:
            raise ValueError(f"الملبس غير موجود: {l.get('code')}")
        q = _pos(l.get("qty"), "كمية الملبس", strict=True)
        out.append((it, int(q) if q == int(q) else q))
    mx = 0
    for dd in data["docs"]:
        m = re.fullmatch(rf"{DOC_PREFIX}-(\d+)", dd["no"])
        if m:
            mx = max(mx, int(m.group(1)))
    no = f"{DOC_PREFIX}-{mx + 1:04d}"
    dw, lw = wb[CL_DOCS], wb[CL_LINES]
    r = _first_empty(dw, 1, 5)
    for c, v in ((1, no), (2, d), (3, emp), (4, _t(note) or None)):
        dw.cell(r, c, v)
    dw.cell(r, 2).number_format = "yyyy-mm-dd"
    dw.cell(r, 5, f"=SUMIFS('{CL_LINES}'!$F$5:$F$50000,'{CL_LINES}'!$A$5:$A$50000,A{r})")
    _style_row(dw, r, 5, formula_cols=(5,))
    lr = _first_empty(lw, 1, 5)
    for it, q in out:
        for c, v in ((1, no), (2, it["code"]), (3, it["name"]), (4, it["size"] or None), (5, it["color"] or None), (6, q)):
            lw.cell(lr, c, v)
        _style_row(lw, lr, 6)
        _write_move(wb, d, it["code"], "تسليم", q, employee=emp, doc=no, note=note)
        lr += 1
    store._save(wb)
    return no


def buy(store, supplier, d, note, currency, lines):
    """فاتورة شراء ملابس من مورد (بسعر القطعة وعملتها). تزيد رصيد الملابس."""
    sup = _clean(supplier)
    if not sup:
        raise ValueError("اختار المورد")
    if not lines:
        raise ValueError("اختار ملبس واحد على الأقل")
    d = datetime.strptime(d, "%Y-%m-%d").date() if isinstance(d, str) and d else (d or date.today())
    cur = core.norm_currency(currency, True)
    wb = store._open()
    known = {x["name"].casefold(): x["name"] for x in store._read_suppliers(wb)}
    if sup.casefold() not in known:
        raise ValueError("المورد غير مسجل. ضيفه أول من شاشة الموردين")
    sup = known[sup.casefold()]
    rate = store._rate(wb) if cur == core.CUR_USD else 1.0
    data = read(wb)
    by = {i["code"].casefold(): i for i in data["items"]}
    out = []
    for l in lines:
        it = by.get(_s(l.get("code")).casefold())
        if it is None:
            raise ValueError(f"الملبس غير موجود: {l.get('code')}")
        q = _pos(l.get("qty"), "كمية الملبس", strict=True)
        out.append((it, int(q) if q == int(q) else q, _pos(l.get("price"), "سعر القطعة")))
    mx = 0
    for p in data["purchases"]:
        m = re.fullmatch(rf"{BUY_PREFIX}-(\d+)", p["no"])
        if m:
            mx = max(mx, int(m.group(1)))
    no = f"{BUY_PREFIX}-{mx + 1:04d}"
    bw, blw = wb[CL_BUYS], wb[CL_BLINES]
    r = _first_empty(bw, 1, 5)
    for c, v in ((1, no), (2, d), (3, sup), (4, cur), (5, rate), (6, _t(note) or None)):
        bw.cell(r, c, v)
    bw.cell(r, 2).number_format = "yyyy-mm-dd"
    bw.cell(r, 7, f"=SUMIFS('{CL_BLINES}'!$H$5:$H$50000,'{CL_BLINES}'!$A$5:$A$50000,A{r})").number_format = "#,##0"
    bw.cell(r, 8, f'=G{r}*IF(D{r}="{core.CUR_USD}",E{r},1)').number_format = "#,##0"
    _style_row(bw, r, 8, formula_cols=(7, 8))
    lr = _first_empty(blw, 1, 5)
    for it, q, price in out:
        for c, v in ((1, no), (2, it["code"]), (3, it["name"]), (4, it["size"] or None), (5, it["color"] or None), (6, q), (7, price)):
            blw.cell(lr, c, v)
        blw.cell(lr, 8, f"=F{lr}*G{lr}").number_format = "#,##0"
        _style_row(blw, lr, 8, formula_cols=(8,))
        _write_move(wb, d, it["code"], "شراء", q, doc=no, note=f"شراء من {sup}" + (f" - {_t(note)}" if _t(note) else ""))
        lr += 1
    store._save(wb)
    return no


def delete_purchase(store, no):
    no = _s(no)
    wb = store._open()
    data = read(wb)
    if no not in [p["no"] for p in data["purchases"]]:
        raise ValueError("الفاتورة غير موجودة")

    class Fbuys:
        cols = (7, 8)

        def __call__(self, rr):
            wb[CL_BUYS].cell(rr, 7).value = f"=SUMIFS('{CL_BLINES}'!$H$5:$H$50000,'{CL_BLINES}'!$A$5:$A$50000,A{rr})"
            wb[CL_BUYS].cell(rr, 8).value = f'=G{rr}*IF(D{rr}="{core.CUR_USD}",E{rr},1)'

    class Fblines:
        cols = (8,)

        def __call__(self, rr):
            wb[CL_BLINES].cell(rr, 8).value = f"=F{rr}*G{rr}"
    _rewrite(wb[CL_BUYS], 5, (1, 2, 3, 4, 5, 6), 0, lambda row: _s(row[0]) != no, Fbuys())
    _rewrite(wb[CL_BLINES], 5, (1, 2, 3, 4, 5, 6, 7), 0, lambda row: _s(row[0]) != no, Fblines())
    _rewrite(wb[CL_MOVES], 5, (2, 1, 3, 4, 5, 6, 7), 0, lambda row: _s(row[5]) != no)
    store._save(wb)


# ---------------------------------------------------------------- حذف
def _rewrite(ws, start, cols, rows_before, keep, formulas=None):
    """يعيد كتابة جدول بعد استبعاد صفوف (الصفوف متتالية وتنتهي بأول فاضي بالعمود الأول)."""
    data, r = [], start
    first = cols[0]
    while ws.cell(r, first).value not in (None, ""):
        data.append([ws.cell(r, c).value for c in cols])
        r += 1
    kept = [row for row in data if keep(row)]
    for i in range(len(data)):
        rr = start + i
        if i < len(kept):
            for c, v in zip(cols, kept[i]):
                ws.cell(rr, c).value = v
            if formulas:
                formulas(rr)
        else:
            for c in cols:
                ws.cell(rr, c).value = None
            if formulas:
                for c in formulas.cols:
                    ws.cell(rr, c).value = None
    return len(data) - len(kept)


def delete_doc(store, no):
    no = _s(no)
    wb = store._open()
    data = read(wb)
    if no not in [d["no"] for d in data["docs"]]:
        raise ValueError("الورقة غير موجودة")

    class Fdocs:
        cols = (5,)

        def __call__(self, rr):
            wb[CL_DOCS].cell(rr, 5).value = f"=SUMIFS('{CL_LINES}'!$F$5:$F$50000,'{CL_LINES}'!$A$5:$A$50000,A{rr})"
    _rewrite(wb[CL_DOCS], 5, (1, 2, 3, 4), 0, lambda row: _s(row[0]) != no, Fdocs())
    _rewrite(wb[CL_LINES], 5, (1, 2, 3, 4, 5, 6), 0, lambda row: _s(row[0]) != no)
    _rewrite(wb[CL_MOVES], 5, (2, 1, 3, 4, 5, 6, 7), 0, lambda row: _s(row[5]) != no)
    store._save(wb)


def delete_move(store, row):
    wb = store._open()
    ws = wb[CL_MOVES]
    if row < 5 or ws.cell(row, 2).value in (None, ""):
        raise ValueError("الحركة غير موجودة")
    if _s(ws.cell(row, 6).value):
        raise ValueError("هذي الحركة تابعة لورقة استلام. احذف الورقة بدل الحركة")
    cols = (2, 1, 3, 4, 5, 6, 7)
    data, r = [], 5
    while ws.cell(r, 2).value not in (None, ""):
        data.append((r, [ws.cell(r, c).value for c in cols]))
        r += 1
    kept = [v for rr, v in data if rr != row]
    for i in range(len(data)):
        rr = 5 + i
        for c, v in zip(cols, kept[i] if i < len(kept) else [None] * len(cols)):
            ws.cell(rr, c).value = v
    store._save(wb)


def delete_item(store, code):
    wb = store._open()
    data = read(wb)
    it = _item(wb, code)
    docs = sorted({l_no for d in data["docs"] for l_no in [d["no"]] if any(x["code"].casefold() == it["code"].casefold() for x in d["lines"])})
    docs += sorted(p["no"] for p in data["purchases"] if any(x["code"].casefold() == it["code"].casefold() for x in p["lines"]))
    if docs:
        raise ValueError("الملبس موجود بأوراق (" + "، ".join(docs[:5]) + ("..." if len(docs) > 5 else "") + "). احذف الأوراق أول")
    cols = (2, 1, 3, 4, 5, 6, 7)
    ws = wb[CL_MOVES]
    data_m, r = [], 5
    while ws.cell(r, 2).value not in (None, ""):
        data_m.append([ws.cell(r, c).value for c in cols])
        r += 1
    kept = [v for v in data_m if _s(v[0]).casefold() != it["code"].casefold()]
    for i in range(len(data_m)):
        for c, v in zip(cols, kept[i] if i < len(kept) else [None] * len(cols)):
            ws.cell(5 + i, c).value = v
    iw = wb[CL_ITEMS]
    for c in range(1, 7):
        iw.cell(it["row"], c).value = None
    store._save(wb)
