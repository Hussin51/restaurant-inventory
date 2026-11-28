"""تكسير الأقسام: تسجيل مواد متكسرة (اسم المادة والكمية والقسم) لأجل التقرير فقط.
ما يأثر على رصيد المخزن ولا الأقسام. ينكتب بورقة مستقلة بنفس ملف الإكسل، وسعر القطعة ينحفظ وقت التسجيل."""
from datetime import date

import core
from core import CUR_USD, _pos, _s, _t, _d, _first_empty, _make_sheet, _style_row, norm_currency

BK = "تكسير_الأقسام"
B0 = 5
COLS = ["التاريخ", "الكود", "المادة", "الكمية (قطعة)", "القسم", "ملاحظات", "سعر القطعة (دينار)", "القيمة (دينار)"]


def ensure_schema(wb):
    if BK not in wb.sheetnames:
        _make_sheet(wb, BK, "تكسير الأقسام (للتقرير فقط، ما ينقص من المخزن)", COLS, [14, 12, 28, 14, 20, 30, 18, 18])


def read(wb):
    ws = wb[BK]
    out, r = [], B0
    while ws.cell(r, 2).value not in (None, ""):
        qty = core._num(ws.cell(r, 4).value)
        price = core._num(ws.cell(r, 7).value)
        out.append(dict(row=r, date=_d(ws.cell(r, 1).value), code=_s(ws.cell(r, 2).value), item=_s(ws.cell(r, 3).value), qty=qty,
                        dept=_s(ws.cell(r, 5).value), note=_s(ws.cell(r, 6).value), price=price, value=qty * price))
        r += 1
    return out


def add(store, d, dept, note, lines):
    dept = _t(dept)
    if not dept:
        raise ValueError("اكتب أو اختار القسم")
    wb = store._open()
    iws = wb[core.ITEMS_SHEET]
    rate = store._rate(wb)
    prepared = []
    for l in lines or []:
        item = _t(l.get("item"))
        if not item and not l.get("qty"):
            continue
        qty = _pos(l.get("qty"), "الكمية", strict=True)
        row = store._item_row(wb, name=item)
        if row is None:
            raise ValueError(f"المادة «{item}» غير موجودة")
        qty = qty * store._factor(wb, row, l.get("unit") or "قطعة")
        price = core._num(iws.cell(row, 6).value)
        if norm_currency(iws.cell(row, 18).value) == CUR_USD:
            price *= rate
        prepared.append((iws.cell(row, 1).value, iws.cell(row, 2).value, int(qty) if qty == int(qty) else qty, _t(l.get("note")), price))
    if not prepared:
        raise ValueError("ضيف مادة وحدة على الأقل")
    ws = wb[BK]
    for code, name, qty, n, price in prepared:
        r = _first_empty(ws, 2, B0)
        ws.cell(r, 1, d).number_format = "yyyy-mm-dd"
        for c, v in ((2, code), (3, name), (4, qty), (5, dept), (6, n or _t(note) or None), (7, price)):
            ws.cell(r, c, v)
        ws.cell(r, 8, f"=D{r}*G{r}").number_format = "#,##0"
        ws.cell(r, 7).number_format = "#,##0"
        _style_row(ws, r, 8, formula_cols=(8,))
    store._save(wb)
    return len(prepared)


def delete(store, row):
    wb = store._open()
    ws = wb[BK]
    if row < B0 or ws.cell(row, 2).value in (None, ""):
        raise ValueError("التسجيل غير موجود")
    cols = (1, 2, 3, 4, 5, 6, 7)
    data, r = [], B0
    while ws.cell(r, 2).value not in (None, ""):
        data.append((r, [ws.cell(r, c).value for c in cols]))
        r += 1
    kept = [v for rr, v in data if rr != row]
    for i in range(len(data)):
        rr = B0 + i
        if i < len(kept):
            for c, v in zip(cols, kept[i]):
                ws.cell(rr, c).value = v
            ws.cell(rr, 8).value = f"=D{rr}*G{rr}"
        else:
            for c in range(1, 9):
                ws.cell(rr, c).value = None
    store._save(wb)
