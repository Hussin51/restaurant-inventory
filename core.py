"""منطق قراءة وكتابة ملف إكسل الجرد (بدون واجهة)."""
import math, os, re
from datetime import date, datetime
from copy import copy
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

ITEMS_SHEET, LOG_SHEET = "الأصناف", "حركة"
SUP_SHEET, INV_SHEET, LINES_SHEET, SET_SHEET = "الموردين", "الفواتير", "بنود_الفواتير", "إعدادات"
I0, I_MAX = 5, 1004         # صفوف الأصناف (1000 صنف)
L0 = 5                       # أول صف بسجل الحركة
R0 = 5                       # أول صف بالموردين والفواتير والبنود
CATEGORIES = ["صحون وكاسات", "أدوات مائدة", "معدات مطبخ", "أثاث وأجهزة"]
TYPES = {"شراء": "in", "إرجاع": "in", "إعارة": "out", "كسر وضياع": "out", "نقل": "move", "تسوية زيادة": "in", "تسوية نقص": "out"}
INV_TYPES = {"شراء": "PUR", "إرجاع": "RET", "إعارة": "LON", "نقل": "TRF"}
WAREHOUSE = "المخزن"                                  # المصدر الثابت
DEPTS = ["البار", "المطبخ", "الصالة", "الأراكيل"]      # الأقسام الافتراضية (تتغير من الإعدادات)
DEFAULT_NAME = "مستلزمات مطعم بيت الكل"
CUR_IQD, CUR_USD = "دينار", "دولار"
DEFAULT_RATE = 1500.0                                  # دينار لكل 1 دولار
DIR_FORMULA = ('=IF(C{r}="","",IF(OR(C{r}="شراء",C{r}="إرجاع"),IF(OR(G{r}="",G{r}="المخزن"),"دخول","نقل"),'
               'IF(C{r}="تسوية زيادة","دخول",IF(C{r}="تسوية نقص","خروج",'
               'IF(OR(C{r}="إعارة",C{r}="كسر وضياع"),IF(OR(F{r}="",F{r}="المخزن"),"خروج","نقل"),'
               'IF(AND(OR(F{r}="",F{r}="المخزن"),G{r}<>"",G{r}<>"المخزن"),"خروج",'
               'IF(AND(F{r}<>"",F{r}<>"المخزن",OR(G{r}="",G{r}="المخزن")),"دخول","نقل")))))))')


def move_effects(t, frm, to, q):
    """أثر الحركة على كل مكان: {المخزن أو اسم القسم: التغيير}. المخزن مفتاحه WAREHOUSE."""
    f, d = _s(frm), _s(to)
    wh = lambda x: x in ("", WAREHOUSE)
    e = {}

    def add(loc, v):
        e[loc] = e.get(loc, 0) + v
    if t == "شراء":
        add(WAREHOUSE if wh(d) else d, q)
    elif t == "إرجاع":
        if f and not wh(f):
            add(f, -q)
        add(WAREHOUSE if wh(d) else d, q)
    elif t == "إعارة":
        add(WAREHOUSE if wh(f) else f, -q)
        if d and not wh(d):
            add(d, q)
    elif t == "نقل":
        src, dst = (WAREHOUSE if wh(f) else f), (WAREHOUSE if wh(d) else d)
        if src != dst:
            add(src, -q)
            add(dst, q)
    elif t == "كسر وضياع":
        add(WAREHOUSE if wh(f) else f, -q)
    elif t == "تسوية زيادة":
        add(WAREHOUSE, q)
    elif t == "تسوية نقص":
        add(WAREHOUSE, -q)
    return e
UNITS = ["قطعة", "شد", "كارتون"]

_F = "Arial"
_HDR_FILL = PatternFill("solid", fgColor="1F4E78")
_HDR_FONT = Font(name=_F, bold=True, color="FFFFFF", size=11)
_THIN = Side(style="thin", color="BFBFBF")
_BD = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_CTR = Alignment(horizontal="center", vertical="center")


class FileBusy(Exception):
    pass


def _num(v):
    return float(v) if isinstance(v, (int, float)) else 0.0


def _s(v):
    return "" if v is None else str(v).strip()


def norm_currency(v, strict=False):
    """يرجّع «دينار» أو «دولار». الفاضي = دينار. strict=True يرفض أي قيمة ثانية."""
    t = str(v or "").strip().lower().replace("$", "usd")
    if t in ("", "دينار", "دينار عراقي", "iqd", "iq", "عراقي", "د.ع", "دع"):
        return CUR_IQD
    if t in ("دولار", "دولار امريكي", "دولار أمريكي", "usd", "dollar", "us"):
        return CUR_USD
    if strict:
        raise ValueError("العملة لازم «دينار» أو «دولار»")
    return CUR_IQD


def _t(v):
    """نص للكتابة بالإكسل: يمنع اعتباره معادلة إذا بدا بـ ="""
    v = _s(v)
    while v.startswith("="):
        v = v[1:].lstrip()
    return v


def _pos(v, what, strict=False):
    """رقم سليم (مو سالب ولا NaN). strict=True يعني لازم أكبر من صفر."""
    try:
        x = float(v or 0)
    except (TypeError, ValueError):
        raise ValueError(f"{what}: اكتب رقم صحيح")
    if not math.isfinite(x) or x < 0 or (strict and x == 0):
        raise ValueError(f"{what}: لازم رقم {'أكبر من صفر' if strict else 'مو سالب'}")
    return x


def _first_empty(ws, col, start, stop=None):
    r = start
    while ws.cell(r, col).value not in (None, ""):
        r += 1
        if stop and r > stop:
            raise ValueError("الجدول ممتلي")
    return r


def _d(v):
    return v.date() if isinstance(v, datetime) else v


def _make_sheet(wb, name, title, headers, widths):
    ws = wb.create_sheet(name)
    ws.sheet_view.rightToLeft = True
    ws["A1"] = title
    ws["A1"].font = Font(name=_F, bold=True, size=16, color="1F4E78")
    ws["A2"] = "تنكتب من البرنامج. لا تغيّر ترتيب الأعمدة."
    ws["A2"].font = Font(name=_F, italic=True, size=10)
    for c, (h, w) in enumerate(zip(headers, widths), 1):
        x = ws.cell(4, c, h)
        x.font, x.fill, x.border, x.alignment = _HDR_FONT, _HDR_FILL, _BD, _CTR
        ws.column_dimensions[ws.cell(4, c).column_letter].width = w
    ws.freeze_panes = "A5"
    return ws


def _style_row(ws, r, ncols, formula_cols=()):
    for c in range(1, ncols + 1):
        x = ws.cell(r, c)
        x.border, x.alignment = _BD, _CTR
        x.font = Font(name=_F, size=11, color="000000" if c in formula_cols else "0000FF")


class Store:
    def __init__(self, path):
        self.path = path

    # ---------- فتح وحفظ + ترحيل الملف القديم ----------
    def _open(self):
        wb = load_workbook(self.path)
        self._ensure_schema(wb)
        return wb

    def _save(self, wb):
        """حفظ آمن: يكتب بملف مؤقت ثم يبدّل، فلو انقطع الحفظ ما يخرب الملف الأصلي."""
        tmp = self.path + ".tmp"
        try:
            wb.save(tmp)
            os.replace(tmp, self.path)
        except PermissionError:
            try:
                os.remove(tmp)
            except OSError:
                pass
            raise FileBusy("ملف الإكسل مفتوح. سكّره وارجع جرب.")

    def _ensure_schema(self, wb):
        ws = wb[ITEMS_SHEET]
        if ws["A4"].value != "الكود":
            ws["A4"] = "الكود"
            for r in range(I0, I_MAX + 1):
                a, b = ws.cell(r, 1), ws.cell(r, 2)
                if isinstance(a.value, (int, float)):
                    a.value = f"M{int(a.value):03d}" if b.value not in (None, "") else None
                a.font = Font(name=_F, size=11, color="0000FF")
        if ws["M4"].value != "الصورة":
            ws["M4"] = "الصورة"
            ws["M4"].font, ws["M4"].fill, ws["M4"].border, ws["M4"].alignment = _HDR_FONT, _HDR_FILL, _BD, _CTR
            ws.column_dimensions["M"].width = 20
            for r in range(I0, I_MAX + 1):
                x = ws.cell(r, 13)
                x.border, x.alignment, x.font = _BD, _CTR, Font(name=_F, size=10, color="808080")
        for col, hd, w in ((14, "قطع بالكارتون", 15), (15, "قطع بالشد", 13)):
            c = ws.cell(4, col)
            if c.value != hd:
                c.value = hd
                c.font, c.fill, c.border, c.alignment = _HDR_FONT, _HDR_FILL, _BD, _CTR
                ws.column_dimensions[c.column_letter].width = w
                for r in range(I0, I_MAX + 1):
                    x = ws.cell(r, col)
                    x.border, x.alignment, x.font = _BD, _CTR, Font(name=_F, size=11, color="0000FF")
        for r in range(I0, I_MAX + 1):
            if ws.cell(r, 4).value == "حبة":
                ws.cell(r, 4).value = "قطعة"
        ws["E4"], ws["F4"], ws["J4"] = "الكمية الافتتاحية (قطعة)", "سعر القطعة", "الحد الأدنى (قطعة)"
        ws.column_dimensions["A"].width = 12
        lg = wb[LOG_SHEET]
        lg["D4"] = "الكمية (قطعة)"
        if lg["I4"].value != "رقم الفاتورة":
            lg["I4"] = "رقم الفاتورة"
            lg["I4"].font, lg["I4"].fill, lg["I4"].border, lg["I4"].alignment = _HDR_FONT, _HDR_FILL, _BD, _CTR
            lg.column_dimensions["I"].width = 16
            for r in range(L0, L0 + 3000):
                x = lg.cell(r, 9)
                x.border, x.alignment, x.font = _BD, _CTR, Font(name=_F, size=11, color="0000FF")
        if SUP_SHEET not in wb.sheetnames:
            _make_sheet(wb, SUP_SHEET, "الموردين", ["الاسم", "رقم الهاتف", "عملة التعامل"], [30, 22, 16])
        sw = wb[SUP_SHEET]
        if sw["C4"].value != "عملة التعامل":
            sw["C4"] = "عملة التعامل"
            sw["C4"].font, sw["C4"].fill, sw["C4"].border, sw["C4"].alignment = _HDR_FONT, _HDR_FILL, _BD, _CTR
            sw.column_dimensions["C"].width = 16
        if INV_SHEET not in wb.sheetnames:
            _make_sheet(wb, INV_SHEET, "الفواتير", ["رقم الفاتورة", "التاريخ", "النوع", "الجهة", "من قسم", "إلى قسم", "ملاحظات", "عدد البنود", "الإجمالي", "تاريخ الاسترجاع", "العملة", "سعر الصرف", "الإجمالي بالدينار", "الخصم"],
                        [16, 14, 12, 26, 14, 14, 30, 12, 16, 16, 12, 12, 18, 12])
        if LINES_SHEET not in wb.sheetnames:
            _make_sheet(wb, LINES_SHEET, "بنود الفواتير", ["رقم الفاتورة", "الكود", "الصنف", "الكمية (قطعة)", "سعر القطعة", "المجموع", "الوحدة", "العدد", "سعر الوحدة"], [16, 12, 28, 14, 14, 16, 12, 10, 14])
        hs = wb[INV_SHEET]
        if hs["J4"].value != "تاريخ الاسترجاع":
            c = hs.cell(4, 10, "تاريخ الاسترجاع")
            c.font, c.fill, c.border, c.alignment = _HDR_FONT, _HDR_FILL, _BD, _CTR
            hs.column_dimensions["J"].width = 16
        ws = wb[ITEMS_SHEET]
        if ws["P4"].value != "الاستخدام":
            c = ws.cell(4, 16, "الاستخدام")
            c.font, c.fill, c.border, c.alignment = _HDR_FONT, _HDR_FILL, _BD, _CTR
            ws.column_dimensions["P"].width = 28
            for r in range(I0, I_MAX + 1):
                x = ws.cell(r, 16)
                x.border, x.alignment, x.font = _BD, _CTR, Font(name=_F, size=11, color="0000FF")
        if ws["Q4"].value != "المورد المعتاد":
            c = ws.cell(4, 17, "المورد المعتاد")
            c.font, c.fill, c.border, c.alignment = _HDR_FONT, _HDR_FILL, _BD, _CTR
            ws.column_dimensions["Q"].width = 22
            for r in range(I0, I_MAX + 1):
                x = ws.cell(r, 17)
                x.border, x.alignment, x.font = _BD, _CTR, Font(name=_F, size=11, color="0000FF")
        ls = wb[LINES_SHEET]
        if ls["G4"].value != "الوحدة":
            for col, hd, w in ((7, "الوحدة", 12), (8, "العدد", 10), (9, "سعر الوحدة", 14)):
                c = ls.cell(4, col, hd)
                c.font, c.fill, c.border, c.alignment = _HDR_FONT, _HDR_FILL, _BD, _CTR
                ls.column_dimensions[c.column_letter].width = w
            ls["D4"], ls["E4"] = "الكمية (قطعة)", "سعر القطعة"
        if SET_SHEET not in wb.sheetnames:
            s = _make_sheet(wb, SET_SHEET, "الإعدادات", ["الإعداد", "القيمة"], [22, 34])
            s["A5"], s["B5"] = "اسم المطعم", DEFAULT_NAME
            _style_row(s, 5, 2)
        st = wb[SET_SHEET]
        if st["A6"].value in (None, ""):
            st["A6"], st["B6"] = "الأقسام (افصل بفاصلة)", "،".join(DEPTS)
            _style_row(st, 6, 2)
        if st["A7"].value in (None, ""):
            st["A7"], st["B7"] = "سعر صرف الدولار (دينار لكل 1 دولار)", DEFAULT_RATE
            _style_row(st, 7, 2)
        iw = wb[ITEMS_SHEET]
        if iw["R4"].value != "عملة السعر":                      # ترحيل العملات
            c = iw.cell(4, 18, "عملة السعر")
            c.font, c.fill, c.border, c.alignment = _HDR_FONT, _HDR_FILL, _BD, _CTR
            iw.column_dimensions["R"].width = 14
            iw["F4"], iw["L4"] = "سعر القطعة (بعملتها)", "قيمة المخزون (دينار)"
            for r in range(I0, I_MAX + 1):
                x = iw.cell(r, 18)
                x.border, x.alignment, x.font = _BD, _CTR, Font(name=_F, size=11, color="0000FF")
                iw.cell(r, 12, f"=IF(B{r}=\"\",\"\",I{r}*N(F{r})*IF(R{r}=\"{CUR_USD}\",'{SET_SHEET}'!$B$7,1))")
        if st["A8"].value in (None, ""):                       # نموذج الأرصدة: المخزن + الأقسام
            st["A8"], st["B8"] = "نموذج الأرصدة", "2"
            lw = wb[LOG_SHEET]
            last = L0
            while lw.cell(last, 2).value not in (None, ""):
                last += 1
            for r in range(L0, max(L0 + 3000, last + 200)):
                lw.cell(r, 5, DIR_FORMULA.format(r=r))
        import clothes                                         # قسم الملابس (أوراق مستقلة)
        clothes.ensure_schema(wb)
        hv = wb[INV_SHEET]
        if hv["K4"].value != "العملة":
            for col, hd, w in ((11, "العملة", 12), (12, "سعر الصرف", 12), (13, "الإجمالي بالدينار", 18), (14, "الخصم", 12)):
                c = hv.cell(4, col, hd)
                c.font, c.fill, c.border, c.alignment = _HDR_FONT, _HDR_FILL, _BD, _CTR
                hv.column_dimensions[c.column_letter].width = w
        if hv["N4"].value != "الخصم":
            c = hv.cell(4, 14, "الخصم")
            c.font, c.fill, c.border, c.alignment = _HDR_FONT, _HDR_FILL, _BD, _CTR
            hv.column_dimensions["N"].width = 12

    # ---------- الإعدادات ----------
    @staticmethod
    def _depts(wb):
        raw = _s(wb[SET_SHEET]["B6"].value)
        out = []
        for d in raw.replace(",", "،").split("،"):
            d = d.strip()
            if d and d != WAREHOUSE and d not in out:
                out.append(d)
        return out or list(DEPTS)

    @staticmethod
    def _rate(wb):
        try:
            r = float(wb[SET_SHEET]["B7"].value)
        except (TypeError, ValueError):
            return DEFAULT_RATE
        return r if math.isfinite(r) and r > 0 else DEFAULT_RATE

    def settings(self):
        wb = self._open()
        return {"name": _s(wb[SET_SHEET]["B5"].value) or DEFAULT_NAME, "depts": self._depts(wb), "rate": self._rate(wb)}

    def set_settings(self, name, depts=None, rate=None):
        wb = self._open()
        wb[SET_SHEET]["B5"] = _t(name) or DEFAULT_NAME
        if rate not in (None, ""):
            wb[SET_SHEET]["B7"] = _pos(rate, "سعر صرف الدولار", strict=True)
        if depts is not None:
            ds = []
            for d in str(depts).replace(",", "،").split("،"):
                d = _t(d)
                if d and d != WAREHOUSE and d not in ds:
                    ds.append(d)
            if not ds:
                raise ValueError("اكتب قسم واحد على الأقل")
            wb[SET_SHEET]["B6"] = "،".join(ds)
        self._save(wb)

    # ---------- قراءة ----------
    def _read_moves(self, wb):
        ws = wb[LOG_SHEET]
        out, r = [], L0
        while ws.cell(r, 2).value not in (None, ""):
            m = dict(row=r, date=_d(ws.cell(r, 1).value), item=ws.cell(r, 2).value, type=ws.cell(r, 3).value,
                     qty=_num(ws.cell(r, 4).value), frm=_s(ws.cell(r, 6).value), to=_s(ws.cell(r, 7).value),
                     note=_s(ws.cell(r, 8).value), inv=_s(ws.cell(r, 9).value))
            m["wh"] = move_effects(m["type"], m["frm"], m["to"], m["qty"]).get(WAREHOUSE, 0)     # أثرها على رصيد المخزن
            out.append(m)
            r += 1
        return out

    def _read_items(self, wb, moves):
        ws = wb[ITEMS_SHEET]
        locs = {}                                              # المادة ← {المكان: الكمية المضافة}
        for m in moves:
            d = locs.setdefault(m["item"], {})
            for loc, v in move_effects(m["type"], m["frm"], m["to"], m["qty"]).items():
                d[loc] = d.get(loc, 0) + v
        res = []
        rate = self._rate(wb)
        for r in range(I0, I_MAX + 1):
            name = ws.cell(r, 2).value
            if name in (None, ""):
                continue
            opening, price, mn = _num(ws.cell(r, 5).value), _num(ws.cell(r, 6).value), _num(ws.cell(r, 10).value)
            cur = norm_currency(ws.cell(r, 18).value)
            price_iqd = price * rate if cur == CUR_USD else price
            mine = locs.get(name, {})
            bal = opening + mine.get(WAREHOUSE, 0)                   # المتبقي بالمخزن
            holdings = {k: v for k, v in mine.items() if k != WAREHOUSE and abs(v) > 1e-9}
            held = sum(holdings.values())
            status = "نفد" if bal <= 0 else ("ناقص - اطلب" if bal <= mn else "متوفر")
            res.append(dict(code=_s(ws.cell(r, 1).value), name=name, cat=_s(ws.cell(r, 3).value), unit=_s(ws.cell(r, 4).value),
                            opening=opening, price=price, cur=cur, price_iqd=price_iqd, min=mn, balance=bal, status=status,
                            value=bal * price_iqd, value_usd=bal * price if cur == CUR_USD else 0,
                            holdings=holdings, held=held, total=bal + held, value_total=(bal + held) * price_iqd,
                            image=_s(ws.cell(r, 13).value), row=r,
                            carton=int(_num(ws.cell(r, 14).value)), shad=int(_num(ws.cell(r, 15).value)),
                            usage=_s(ws.cell(r, 16).value), supplier=_s(ws.cell(r, 17).value)))
        return res

    def _read_suppliers(self, wb):
        ws = wb[SUP_SHEET]
        out, r = [], R0
        while ws.cell(r, 1).value not in (None, ""):
            out.append(dict(name=_s(ws.cell(r, 1).value), phone=_s(ws.cell(r, 2).value), currency=norm_currency(ws.cell(r, 3).value)))
            r += 1
        return out

    def _read_invoices(self, wb):
        hs, ls = wb[INV_SHEET], wb[LINES_SHEET]
        lines, r = {}, R0
        while ls.cell(r, 1).value not in (None, ""):
            lines.setdefault(_s(ls.cell(r, 1).value), []).append(
                dict(code=_s(ls.cell(r, 2).value), name=_s(ls.cell(r, 3).value), qty=_num(ls.cell(r, 4).value),
                     price=_num(ls.cell(r, 5).value), unit=_s(ls.cell(r, 7).value) or "قطعة",
                     count=_num(ls.cell(r, 8).value) or _num(ls.cell(r, 4).value),
                     uprice=_num(ls.cell(r, 9).value) if ls.cell(r, 7).value else _num(ls.cell(r, 5).value)))
            r += 1
        out, r = [], R0
        cur_rate = self._rate(wb)
        while hs.cell(r, 1).value not in (None, ""):
            no = _s(hs.cell(r, 1).value)
            ln = lines.get(no, [])
            cur = norm_currency(hs.cell(r, 11).value)
            rate = (_num(hs.cell(r, 12).value) or cur_rate) if cur == CUR_USD else 1.0
            for l in ln:
                l["total"] = l["qty"] * l["price"]
                l["total_iqd"] = l["total"] * rate
            gross = sum(l["total"] for l in ln)
            disc = min(_num(hs.cell(r, 14).value), gross) if cur is not None else 0.0
            ratio = (gross - disc) / gross if gross else 1.0
            for l in ln:                                          # الخصم موزّع على البنود بالنسبة (للتقارير حسب المادة)
                l["net"], l["net_iqd"] = l["total"] * ratio, l["total_iqd"] * ratio
            out.append(dict(no=no, date=_d(hs.cell(r, 2).value), type=_s(hs.cell(r, 3).value), party=_s(hs.cell(r, 4).value),
                            frm=_s(hs.cell(r, 5).value), to=_s(hs.cell(r, 6).value), note=_s(hs.cell(r, 7).value),
                            due=_d(hs.cell(r, 10).value), cur=cur, rate=rate, lines=ln, gross=gross, gross_iqd=gross * rate,
                            discount=disc, total=gross - disc, total_iqd=(gross - disc) * rate))
            r += 1
        return out

    def all(self):
        wb = self._open()
        moves = self._read_moves(wb)
        items = self._read_items(wb, moves)
        invoices = self._read_invoices(wb)
        import clothes
        return dict(items=items, moves=moves, suppliers=self._read_suppliers(wb), invoices=invoices, clothes=clothes.read(wb),
                    loans=self.open_loans(invoices),
                    settings={"name": _s(wb[SET_SHEET]["B5"].value) or DEFAULT_NAME, "depts": self._depts(wb), "rate": self._rate(wb)})

    @staticmethod
    def open_loans(invoices, today=None):
        """الإعارات الي لسا ما رجعت كلها. الإرجاع يطفي أقدم إعارة أول (FIFO) لنفس القسم والمادة."""
        today = today or date.today()
        returned = {}
        for v in invoices:
            if v["type"] == "إرجاع":
                for l in v["lines"]:
                    k = (v["frm"], l["code"])
                    returned[k] = returned.get(k, 0) + l["qty"]
        out = []
        for v in sorted((x for x in invoices if x["type"] == "إعارة"), key=lambda x: (str(x["date"]), x["no"])):
            lines = []
            for l in v["lines"]:
                k = (v["to"], l["code"])
                give = returned.get(k, 0)
                take = min(give, l["qty"])
                returned[k] = give - take
                if l["qty"] - take > 0:
                    lines.append(dict(code=l["code"], name=l["name"], qty=l["qty"], remaining=l["qty"] - take))
            if lines:
                due = v.get("due")
                days = (due - today).days if isinstance(due, date) else None
                status = "متأخرة" if days is not None and days < 0 else ("اليوم" if days == 0 else "قادمة")
                out.append(dict(no=v["no"], date=v["date"], dept=v["to"], due=due, days=days, status=status, lines=lines))
        return out

    def items(self):
        wb = self._open()
        return self._read_items(wb, self._read_moves(wb))

    def movements(self):
        return self._read_moves(self._open())

    # ---------- حركة منفردة (تسجيل سريع) ----------
    def _item_row(self, wb, name=None, code=None):
        ws = wb[ITEMS_SHEET]
        for r in range(I0, I_MAX + 1):
            if ws.cell(r, 2).value in (None, ""):
                continue
            if (name is not None and ws.cell(r, 2).value == name) or (code is not None and _s(ws.cell(r, 1).value).casefold() == code.casefold()):
                return r
        return None

    def _factor(self, wb, row, unit):
        unit = _s(unit) or "قطعة"
        if unit == "قطعة":
            return 1
        if unit not in UNITS:
            raise ValueError("وحدة غير صحيحة")
        ws = wb[ITEMS_SHEET]
        f = int(_num(ws.cell(row, 14 if unit == "كارتون" else 15).value))
        if f < 1:
            raise ValueError(f"المادة «{ws.cell(row, 2).value}» ما إلها {unit} محدد. حدده من تعديل المادة")
        return f

    def add_movement(self, d, item, typ, qty, frm="", to="", note="", inv="", unit="قطعة"):
        if typ not in TYPES:
            raise ValueError("نوع الحركة غير صحيح")
        qty = _pos(qty, "الكمية", strict=True)
        wb = self._open()
        row = self._item_row(wb, name=item)
        if row is None:
            raise ValueError("الصنف غير موجود")
        qty = qty * self._factor(wb, row, unit)
        qty = int(qty) if qty == int(qty) else qty
        self._write_move(wb, d, item, typ, qty, frm, to, note, inv)
        self._save(wb)

    def add_breakage(self, d, lines, frm="", note=""):
        """تسجيل تكسير لعدة مواد مرة وحدة (كل سطر حركة «كسر وضياع»)، بحفظ واحد."""
        wb = self._open()
        prepared = []
        for l in lines or []:
            item = _t(l.get("item"))
            if not item and not l.get("qty"):
                continue
            qty = _pos(l.get("qty"), "الكمية", strict=True)
            row = self._item_row(wb, name=item)
            if row is None:
                raise ValueError(f"المادة «{item}» غير موجودة")
            qty = qty * self._factor(wb, row, l.get("unit") or "قطعة")
            prepared.append((item, int(qty) if qty == int(qty) else qty, _t(l.get("note"))))
        if not prepared:
            raise ValueError("ضيف مادة وحدة على الأقل")
        for item, qty, n in prepared:
            self._write_move(wb, d, item, "كسر وضياع", qty, frm, "", n or note, "")
        self._save(wb)
        return len(prepared)

    def _write_move(self, wb, d, item, typ, qty, frm, to, note, inv):
        ws = wb[LOG_SHEET]
        r = _first_empty(ws, 2, L0)
        ws.cell(r, 1, d).number_format = "yyyy-mm-dd"
        for c, v in ((2, item), (3, typ), (4, qty), (6, frm), (7, to), (8, note), (9, inv)):
            v = _t(v) if isinstance(v, str) else v
            ws.cell(r, c, v if v != "" else None)
        if ws.cell(r, 5).value in (None, ""):
            ws.cell(r, 5, DIR_FORMULA.format(r=r))
            _style_row(ws, r, 9, formula_cols=(5,))

    # ---------- الأصناف ----------
    def add_item(self, code, name, cat, unit, opening, price, minimum, carton=0, shad=0, usage="", supplier="", currency=""):
        code, name = _t(code), _t(name)
        opening, price, minimum = _pos(opening, "الكمية الموجودة"), _pos(price, "السعر"), _pos(minimum, "الحد الأدنى")
        carton, shad = int(_pos(carton, "قطع الكارتون")), int(_pos(shad, "قطع الشد"))
        if not code:
            raise ValueError("اكتب كود المادة")
        if not name:
            raise ValueError("اكتب اسم الصنف")
        wb = self._open()
        if self._item_row(wb, code=code) is not None:
            raise ValueError("هذا الكود مستخدم لمادة ثانية")
        if self._item_row(wb, name=name) is not None:
            raise ValueError("هذا الاسم موجود من قبل")
        ws = wb[ITEMS_SHEET]
        r = _first_empty(ws, 2, I0, I_MAX)
        for c, v in ((1, code), (2, name), (3, _t(cat)), (4, _t(unit) or "قطعة"), (5, opening), (6, price), (10, minimum),
                     (14, int(carton) or None), (15, int(shad) or None), (16, _t(usage) or None), (17, self._sup(wb, supplier)), (18, norm_currency(currency, True))):
            ws.cell(r, c, v)
        self._save(wb)

    def _sup(self, wb, name):
        """اسم المورد المعتاد (لازم يكون مسجل) أو None."""
        name = _t(name)
        if not name:
            return None
        for sp in self._read_suppliers(wb):
            if sp["name"].casefold() == name.casefold():
                return sp["name"]
        raise ValueError("المورد المعتاد غير مسجل. ضيفه أول من شاشة الموردين")

    def update_item(self, code, cat, unit, price, minimum, carton=0, shad=0, usage="", supplier="", currency=None):
        price, minimum = _pos(price, "السعر"), _pos(minimum, "الحد الأدنى")
        carton, shad = int(_pos(carton, "قطع الكارتون")), int(_pos(shad, "قطع الشد"))
        wb = self._open()
        r = self._item_row(wb, code=code)
        if r is None:
            raise ValueError("المادة غير موجودة")
        ws = wb[ITEMS_SHEET]
        for c, v in ((3, _t(cat)), (4, _t(unit) or "قطعة"), (6, price), (10, minimum), (14, carton or None), (15, shad or None), (16, _t(usage) or None), (17, self._sup(wb, supplier)),
                     (18, norm_currency(currency, True) if currency is not None else ws.cell(r, 18).value)):
            ws.cell(r, c, v)
        self._save(wb)

    def set_image(self, code, filename):
        wb = self._open()
        r = self._item_row(wb, code=code)
        if r is None:
            raise ValueError("المادة غير موجودة")
        wb[ITEMS_SHEET].cell(r, 13, filename or None)
        self._save(wb)

    def set_images(self, mapping):
        """يحدد صور عدة مواد دفعة وحدة (حفظ واحد): {code: filename}. المواد المجهولة تنتخطى."""
        wb = self._open()
        n = 0
        for code, fn in mapping.items():
            r = self._item_row(wb, code=code)
            if r is not None:
                wb[ITEMS_SHEET].cell(r, 13, fn or None)
                n += 1
        if n:
            self._save(wb)
        return n

    # ---------- الموردين ----------
    def add_supplier(self, name, phone, currency=""):
        name = _t(name)
        currency = norm_currency(currency, True)
        if not name:
            raise ValueError("اكتب اسم المورد")
        wb = self._open()
        if name.casefold() in [s["name"].casefold() for s in self._read_suppliers(wb)]:
            raise ValueError("هذا المورد موجود من قبل")
        ws = wb[SUP_SHEET]
        r = _first_empty(ws, 1, R0)
        ws.cell(r, 1, name); ws.cell(r, 2, _t(phone)); ws.cell(r, 3, currency)
        _style_row(ws, r, 3)
        self._save(wb)

    def update_supplier(self, name, phone, currency=None):
        wb = self._open()
        ws = wb[SUP_SHEET]
        r = R0
        while ws.cell(r, 1).value not in (None, ""):
            if _s(ws.cell(r, 1).value) == name:
                ws.cell(r, 2, _t(phone))
                if currency is not None:
                    ws.cell(r, 3, norm_currency(currency, True))
                    _style_row(ws, r, 3)
                self._save(wb)
                return
            r += 1
        raise ValueError("المورد غير موجود")

    # ---------- الفواتير ----------
    def add_invoice(self, inv):
        typ = inv.get("type")
        if typ not in INV_TYPES:
            raise ValueError("نوع الفاتورة غير صحيح")
        raw = inv.get("lines") or []
        if not raw:
            raise ValueError("أضف مادة وحدة على الأقل")
        d = inv.get("date")
        if isinstance(d, str):
            d = datetime.strptime(d, "%Y-%m-%d").date()
        d = d or date.today()
        party, frm, to, note = _t(inv.get("party")), _t(inv.get("frm")), _t(inv.get("to")), _t(inv.get("note"))
        wb = self._open()
        depts = self._depts(wb)
        cur = norm_currency(inv.get("currency"), True) if typ == "شراء" else CUR_IQD
        rate = self._rate(wb) if cur == CUR_USD else 1.0
        due = inv.get("due")
        if isinstance(due, str):
            due = datetime.strptime(due, "%Y-%m-%d").date() if due else None
        if typ == "شراء":                                   # مورد ← المخزن
            if not party:
                raise ValueError("اختار المورد")
            if party.casefold() not in [s["name"].casefold() for s in self._read_suppliers(wb)]:
                raise ValueError("المورد غير مسجل. ضيفه أول")
            frm, to, due = "", WAREHOUSE, None
        elif typ == "إرجاع":                                 # قسم ← المخزن
            if frm not in depts:
                raise ValueError("اختار القسم الي رجّع المواد")
            party, to, due = frm, WAREHOUSE, None
        elif typ == "إعارة":                                 # المخزن ← قسم (وتاريخ استرجاع)
            if to not in depts:
                raise ValueError("اختار القسم المستعير")
            if not due:
                raise ValueError("حدد تاريخ الاسترجاع")
            if due < d:
                raise ValueError("تاريخ الاسترجاع قبل تاريخ الإعارة")
            party, frm = to, WAREHOUSE
        else:                                                # نقل: المخزن ← قسم
            if to not in depts:
                raise ValueError("اختار القسم الوجهة")
            party, frm, due = "", WAREHOUSE, None
        lines = []
        for l in raw:
            row = self._item_row(wb, code=_s(l.get("code")))
            if row is None:
                raise ValueError(f"الكود غير موجود: {l.get('code')}")
            cnt = _pos(l.get("qty"), "كمية المادة", strict=True)
            unit = _s(l.get("unit")) or "قطعة"
            f = self._factor(wb, row, unit)
            q = cnt * f
            q = int(q) if q == int(q) else q
            cnt = int(cnt) if cnt == int(cnt) else cnt
            up = _pos(l.get("price"), "سعر المادة")
            lines.append((row, _s(l.get("code")), wb[ITEMS_SHEET].cell(row, 2).value, q, up / f, unit, cnt, up))
        # رقم الفاتورة
        pre = INV_TYPES[typ]
        hs, ls = wb[INV_SHEET], wb[LINES_SHEET]
        mx, r = 0, R0
        while hs.cell(r, 1).value not in (None, ""):
            v = _s(hs.cell(r, 1).value)
            if v.startswith(pre + "-"):
                try:
                    mx = max(mx, int(v.split("-")[1]))
                except ValueError:
                    pass
            r += 1
        no = f"{pre}-{mx + 1:04d}"
        disc = _pos(inv.get("discount"), "الخصم") if typ == "شراء" else 0
        if disc and disc > sum(l[3] * l[4] for l in lines) + 1e-9:
            raise ValueError("الخصم أكبر من إجمالي الفاتورة")
        hr = r
        for c, v in ((1, no), (2, d), (3, typ), (4, party), (5, frm), (6, to), (7, note), (10, due),
                     (11, cur if typ == "شراء" else None), (12, rate if typ == "شراء" else None)):
            hs.cell(hr, c, v if v not in ("", None) else None)
        hs.cell(hr, 14, disc or None).number_format = "#,##0.00"
        hs.cell(hr, 13, f'=(I{hr}-N{hr})*IF(K{hr}="{CUR_USD}",L{hr},1)').number_format = "#,##0"
        hs.cell(hr, 2).number_format = "yyyy-mm-dd"
        hs.cell(hr, 10).number_format = "yyyy-mm-dd"
        rng = f"'{LINES_SHEET}'!$A$5:$A$50000"
        hs.cell(hr, 8, f"=COUNTIF({rng},A{hr})")
        hs.cell(hr, 9, f"=SUMIFS('{LINES_SHEET}'!$F$5:$F$50000,{rng},A{hr})").number_format = "#,##0"
        _style_row(hs, hr, 14, formula_cols=(8, 9, 13))
        lr = _first_empty(ls, 1, R0)
        for row, code, name, q, price, unit, cnt, up in lines:
            ls.cell(lr, 1, no); ls.cell(lr, 2, code); ls.cell(lr, 3, name); ls.cell(lr, 4, q)
            ls.cell(lr, 5, price if typ == "شراء" else 0)
            ls.cell(lr, 6, f"=D{lr}*E{lr}").number_format = "#,##0"
            ls.cell(lr, 7, unit); ls.cell(lr, 8, cnt); ls.cell(lr, 9, up if typ == "شراء" else 0)
            _style_row(ls, lr, 9, formula_cols=(3, 6))
            lr += 1
            if typ == "شراء" and inv.get("update_prices") and price > 0:
                wb[ITEMS_SHEET].cell(row, 6, price)                       # آخر سعر شراء، بعملة الفاتورة
                wb[ITEMS_SHEET].cell(row, 18, cur)
            mf, mt = (party, to) if typ == "شراء" else (frm, to)
            self._write_move(wb, d, name, typ, q, mf, mt, note, no)
        self._save(wb)
        return no

    # ---------- الجرد الفعلي: رصيد افتتاحي أو تسوية ----------
    def set_opening(self, counts):
        """يحدد الرصيد الافتتاحي (بالقطعة) للمواد المذكورة."""
        wb = self._open()
        n = 0
        for c in counts:
            row = self._item_row(wb, code=_s(c.get("code")))
            if row is None:
                raise ValueError(f"الكود غير موجود: {c.get('code')}")
            wb[ITEMS_SHEET].cell(row, 5, _pos(c.get("pieces"), "الكمية"))
            n += 1
        self._save(wb)
        return n

    def adjust(self, counts, d=None, note="تسوية جرد"):
        """يسجل حركة تسوية (زيادة أو نقص) لتصير أرصدة البرنامج = الجرد الفعلي."""
        if isinstance(d, str):
            d = datetime.strptime(d, "%Y-%m-%d").date() if d else None
        d = d or date.today()
        wb = self._open()
        bal = {i["code"].casefold(): i for i in self._read_items(wb, self._read_moves(wb))}
        n = 0
        for c in counts:
            it = bal.get(_s(c.get("code")).casefold())
            if it is None:
                raise ValueError(f"الكود غير موجود: {c.get('code')}")
            diff = _pos(c.get("pieces"), "الكمية") - it["balance"]
            if abs(diff) < 1e-9:
                continue
            diff = int(diff) if diff == int(diff) else diff
            self._write_move(wb, d, it["name"], "تسوية زيادة" if diff > 0 else "تسوية نقص", abs(diff), WAREHOUSE, "", _t(note), "")
            n += 1
        self._save(wb)
        return n

    # ---------- استيراد المواد وترتيبها ----------
    @staticmethod
    def _natural(code):
        return [(0, int(x), "") if x.isdigit() else (1, 0, x.casefold()) for x in re.findall(r"\d+|\D+", str(code))]

    def _sort_items(self, wb):
        """يرتّب المواد فعلياً بورقة الإكسل: حسب التصنيف (الافتراضية أول) ثم الكود."""
        ws = wb[ITEMS_SHEET]
        cols = (1, 2, 3, 4, 5, 6, 10, 13, 14, 15, 16, 17, 18)
        rows = []
        for r in range(I0, I_MAX + 1):
            if ws.cell(r, 2).value in (None, ""):
                continue
            rows.append([ws.cell(r, c).value for c in cols])
        order = {c: i for i, c in enumerate(CATEGORIES)}
        rows.sort(key=lambda v: (order.get(_s(v[2]), len(order)), _s(v[2]).casefold(), self._natural(_s(v[0]))))
        for i, vals in enumerate(rows):
            for c, v in zip(cols, vals):
                ws.cell(I0 + i, c).value = v
        return len(rows)

    def sort_items(self):
        wb = self._open()
        n = self._sort_items(wb)
        self._save(wb)
        return n

    def import_items(self, rows, mode="skip", sort=True):
        """rows: من importer.parse (الحالات new/exists). mode: skip (تخطي الموجود) أو update (تحديث الموجود)."""
        wb = self._open()
        ws = wb[ITEMS_SHEET]
        by_code, nxt = {}, I0
        for r in range(I0, I_MAX + 1):
            if ws.cell(r, 2).value not in (None, ""):
                by_code[_s(ws.cell(r, 1).value).casefold()] = r
                nxt = r + 1
        known = {sp["name"].casefold(): sp["name"] for sp in self._read_suppliers(wb)}
        sup_ws = wb[SUP_SHEET]
        added = updated = skipped = new_sup = 0
        for x in rows:
            if x["status"] in ("error", "skip"):
                continue
            key = x["code"].casefold()
            sup = None
            if x["supplier"]:
                sup = known.get(x["supplier"].casefold())
                if sup is None:
                    sup = _t(x["supplier"])
                    rr = _first_empty(sup_ws, 1, R0)
                    sup_ws.cell(rr, 1, sup)
                    sup_ws.cell(rr, 3, x.get("currency") or CUR_IQD)
                    _style_row(sup_ws, rr, 3)
                    known[sup.casefold()] = sup
                    new_sup += 1
            if key in by_code:
                if mode != "update":
                    skipped += 1
                    continue
                r, h = by_code[key], x["has"]
                for c, v, ok in ((3, _t(x["category"]), h["category"]), (6, x["price"], h["price"]), (10, x["min"], h["min"]),
                                 (14, x["carton"] or None, h["carton"]), (15, x["shad"] or None, h["shad"]),
                                 (16, _t(x["usage"]) or None, h["usage"]), (17, sup, h["supplier"]),
                                 (18, x.get("currency") or CUR_IQD, h.get("currency", False))):
                    if ok:
                        ws.cell(r, c, v)
                updated += 1
                continue
            if nxt > I_MAX:
                raise ValueError("وصلت الحد الأعلى (1000 مادة)")
            r, nxt = nxt, nxt + 1
            for c, v in ((1, _t(x["code"])), (2, _t(x["name"])), (3, _t(x["category"])), (4, "قطعة"), (5, x["opening"]), (6, x["price"]),
                         (10, x["min"]), (14, x["carton"] or None), (15, x["shad"] or None), (16, _t(x["usage"]) or None), (17, sup),
                         (18, x.get("currency") or CUR_IQD)):
                ws.cell(r, c, v)
            by_code[key] = r
            added += 1
        if sort:
            self._sort_items(wb)
        self._save(wb)
        return dict(added=added, updated=updated, skipped=skipped, new_suppliers=new_sup)

    # ---------- الحذف ----------
    def _rewrite_log(self, wb, keep):
        """يعيد كتابة سجل الحركات بعد استبعاد الحركات الي keep(row, values) ترجع لها False (يحافظ على الصيغ والترتيب)."""
        ws = wb[LOG_SHEET]
        cols = (1, 2, 3, 4, 6, 7, 8, 9)
        rows, r = [], L0
        while ws.cell(r, 2).value not in (None, ""):
            rows.append((r, [ws.cell(r, c).value for c in cols]))
            r += 1
        kept = [v for rr, v in rows if keep(rr, v)]
        for i, v in enumerate(kept):
            rr = L0 + i
            for c, x in zip(cols, v):
                ws.cell(rr, c).value = x
            ws.cell(rr, 5).value = DIR_FORMULA.format(r=rr)
            ws.cell(rr, 1).number_format = "yyyy-mm-dd"
        for rr in range(L0 + len(kept), L0 + len(rows)):
            for c in cols:
                ws.cell(rr, c).value = None
            ws.cell(rr, 5).value = DIR_FORMULA.format(r=rr) if rr < L0 + 3000 else None
        return len(rows) - len(kept)

    def delete_movement(self, row):
        wb = self._open()
        ws = wb[LOG_SHEET]
        if ws.cell(row, 2).value in (None, "") or row < L0:
            raise ValueError("الحركة غير موجودة")
        if _s(ws.cell(row, 9).value):
            raise ValueError("هذي الحركة تابعة لفاتورة. احذف الفاتورة بدل الحركة")
        self._rewrite_log(wb, lambda rr, v: rr != row)
        self._save(wb)

    def delete_item(self, code):
        """يحذف المادة وحركاتها. يرفض إذا المادة موجودة بفواتير (نحافظ على سجل الفواتير)."""
        wb = self._open()
        row = self._item_row(wb, code=_s(code))
        if row is None:
            raise ValueError("المادة غير موجودة")
        ws = wb[ITEMS_SHEET]
        name, image = ws.cell(row, 2).value, _s(ws.cell(row, 13).value)
        ls, nos, r = wb[LINES_SHEET], [], R0
        while ls.cell(r, 1).value not in (None, ""):
            if _s(ls.cell(r, 2).value).casefold() == _s(code).casefold() and _s(ls.cell(r, 1).value) not in nos:
                nos.append(_s(ls.cell(r, 1).value))
            r += 1
        if nos:
            raise ValueError("المادة موجودة بفواتير (" + "، ".join(nos[:5]) + ("..." if len(nos) > 5 else "") + "). احذف هالفواتير أول")
        removed = self._rewrite_log(wb, lambda rr, v: v[1] != name)
        for c in (1, 2, 3, 4, 5, 6, 10, 13, 14, 15, 16, 17, 18):
            ws.cell(row, c).value = None
        self._save(wb)
        return dict(removed_moves=removed, image=image)

    def delete_items(self, codes):
        """حذف مجموعة مواد دفعة وحدة (حفظ واحد). المواد الموجودة بفواتير تتخطى وتنرجع بقائمة blocked."""
        wb = self._open()
        want = {_s(c).casefold() for c in codes if _s(c)}
        if not want:
            raise ValueError("ما حددت مواد")
        ws = wb[ITEMS_SHEET]
        rows = {}
        for r in range(I0, I_MAX + 1):
            code = _s(ws.cell(r, 1).value)
            if ws.cell(r, 2).value not in (None, "") and code.casefold() in want:
                rows[code.casefold()] = r
        used, r, ls = {}, R0, wb[LINES_SHEET]
        while ls.cell(r, 1).value not in (None, ""):
            k = _s(ls.cell(r, 2).value).casefold()
            if k in rows:
                used.setdefault(k, [])
                if _s(ls.cell(r, 1).value) not in used[k]:
                    used[k].append(_s(ls.cell(r, 1).value))
            r += 1
        ok = [k for k in rows if k not in used]
        names = {ws.cell(rows[k], 2).value for k in ok}
        images = [_s(ws.cell(rows[k], 13).value) for k in ok]
        removed = self._rewrite_log(wb, lambda rr, v: v[1] not in names) if names else 0
        for k in ok:
            for c in (1, 2, 3, 4, 5, 6, 10, 13, 14, 15, 16, 17, 18):
                ws.cell(rows[k], c).value = None
        if ok:
            self._save(wb)
        return dict(deleted=len(ok), blocked={_s(ws.cell(rows[k], 1).value): v for k, v in used.items()}, removed_moves=removed,
                    missing=len(want) - len(rows), images=[i for i in images if i])

    def delete_supplier(self, name):
        wb = self._open()
        sups = self._read_suppliers(wb)
        if name not in [x["name"] for x in sups]:
            raise ValueError("المورد غير موجود")
        nos = [v["no"] for v in self._read_invoices(wb) if v["type"] == "شراء" and v["party"] == name]
        if nos:
            raise ValueError("المورد عنده فواتير (" + "، ".join(nos[:5]) + ("..." if len(nos) > 5 else "") + "). ما ينحذف")
        ws = wb[SUP_SHEET]
        keep = [x for x in sups if x["name"] != name]
        for i in range(len(sups)):
            rr = R0 + i
            ws.cell(rr, 1).value = keep[i]["name"] if i < len(keep) else None
            ws.cell(rr, 2).value = (keep[i]["phone"] or None) if i < len(keep) else None
            ws.cell(rr, 3).value = keep[i]["currency"] if i < len(keep) else None
        iw = wb[ITEMS_SHEET]
        for r in range(I0, I_MAX + 1):
            if _s(iw.cell(r, 17).value) == name:
                iw.cell(r, 17).value = None
        self._save(wb)

    def delete_invoice(self, no):
        """يحذف الفاتورة وبنودها وحركاتها، فترجع الأرصدة مثل قبل الفاتورة."""
        no = _s(no)
        wb = self._open()
        hs, ls = wb[INV_SHEET], wb[LINES_SHEET]
        heads, r = [], R0
        while hs.cell(r, 1).value not in (None, ""):
            heads.append([hs.cell(r, c).value for c in (1, 2, 3, 4, 5, 6, 7, 10, 11, 12, 14)])
            r += 1
        if no not in [_s(h[0]) for h in heads]:
            raise ValueError("الفاتورة غير موجودة")
        kept = [h for h in heads if _s(h[0]) != no]
        for i in range(len(heads)):
            rr = R0 + i
            if i < len(kept):
                for c, x in zip((1, 2, 3, 4, 5, 6, 7, 10, 11, 12, 14), kept[i]):
                    hs.cell(rr, c).value = x
                hs.cell(rr, 8).value = f"=COUNTIF('{LINES_SHEET}'!$A$5:$A$50000,A{rr})"
                hs.cell(rr, 9).value = f"=SUMIFS('{LINES_SHEET}'!$F$5:$F$50000,'{LINES_SHEET}'!$A$5:$A$50000,A{rr})"
                hs.cell(rr, 13).value = f'=(I{rr}-N{rr})*IF(K{rr}="{CUR_USD}",L{rr},1)'
                hs.cell(rr, 2).number_format = hs.cell(rr, 10).number_format = "yyyy-mm-dd"
            else:
                for c in range(1, 15):
                    hs.cell(rr, c).value = None
        lines, r = [], R0
        while ls.cell(r, 1).value not in (None, ""):
            lines.append([ls.cell(r, c).value for c in (1, 2, 3, 4, 5, 7, 8, 9)])
            r += 1
        lkept = [l for l in lines if _s(l[0]) != no]
        for i in range(len(lines)):
            rr = R0 + i
            if i < len(lkept):
                for c, x in zip((1, 2, 3, 4, 5, 7, 8, 9), lkept[i]):
                    ls.cell(rr, c).value = x
                ls.cell(rr, 6).value = f"=D{rr}*E{rr}"
            else:
                for c in range(1, 10):
                    ls.cell(rr, c).value = None
        self._rewrite_log(wb, lambda rr, v: _s(v[7]) != no)
        self._save(wb)
        return dict(lines=len(lines) - len(lkept))
