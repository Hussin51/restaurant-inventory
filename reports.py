"""التقارير. تُحسب من بيانات Store.all() وترجع هيكل واحد (بطاقات + جداول) للعرض والطباعة وتصدير الإكسل."""
import io
import math
import os

KINDS = ("movement", "reorder", "dept", "supplier", "location", "employee", "clothes")


def _iso(d):
    return str(d)[:10] if d else ""


def _in(d, a, b):
    d = _iso(d)
    return (not a or d >= a) and (not b or d <= b)


def packs_text(it, pcs):
    """5 كارتون و3 شد و7 قطعة (إذا للمادة تعبئة)."""
    if not it or not (it.get("carton") or it.get("shad")):
        return ""
    r = round(abs(pcs), 3)
    parts = []
    for name, f in sorted([("كارتون", it.get("carton") or 0), ("شد", it.get("shad") or 0)], key=lambda x: -x[1]):
        if f > 0:
            c = math.floor(r / f + 1e-9)
            if c:
                parts.append(f"{c:g} {name}")
                r -= c * f
    if r > 1e-9 or not parts:
        parts.append(f"{r:g} قطعة")
    return ("-" if pcs < 0 else "") + " و".join(parts)


def qty_text(it, pcs):
    t = packs_text(it, pcs)
    return f"{pcs:g} قطعة" + (f" ({t})" if t else "")


def _items(data):
    return {i["name"]: i for i in data["items"]}


def _val(items, name, qty):
    return qty * (items.get(name, {}).get("price_iqd") or 0)          # بالدينار (الدولار محوّل بسعر الصرف الحالي)


def _table(title, columns, rows, numeric=()):
    return {"title": title, "columns": columns, "rows": rows, "numeric": list(numeric)}


def _period(a, b):
    if a and b:
        return f"من {a} إلى {b}"
    if a:
        return f"من {a}"
    if b:
        return f"لحد {b}"
    return "كل الفترة"


# ---------------------------------------------------------------- مشتريات وإخراج
def _cl_purchases(data):
    """فواتير شراء الملابس بنفس شكل فواتير الشراء (حتى تدخل بالتقارير وكشف المورد)."""
    out = []
    for p in data.get("clothes", {}).get("purchases", []):
        out.append(dict(no=p["no"], date=p["date"], type="شراء", party=p["party"], cur=p["cur"], rate=p["rate"], note=p["note"], total=p["total"],
                        total_iqd=p["total_iqd"], lines=[dict(code=l["code"], name=" ".join(x for x in (l["name"], l["size"], l["color"]) if x), qty=l["qty"],
                                                              total=l["total"], total_iqd=l["total_iqd"]) for l in p["lines"]]))
    return out


def movement(data, a="", b=""):
    items = _items(data)
    inv = [v for v in data["invoices"] + _cl_purchases(data) if v["type"] == "شراء" and _in(v["date"], a, b)]
    moves = [m for m in data["moves"] if _in(m["date"], a, b)]

    by_sup, by_item = {}, {}
    for v in inv:
        usd = v["cur"] == "دولار"
        s = by_sup.setdefault(v["party"], [0, 0.0, 0.0, 0.0])           # عدد، إجمالي بالدينار، دينار أصلي، دولار أصلي
        s[0] += 1
        s[1] += v["total_iqd"]
        s[3 if usd else 2] += v["total"]
        for l in v["lines"]:
            x = by_item.setdefault(l["code"], [l["name"], 0.0, 0.0, 0.0])
            x[1] += l["qty"]
            x[2] += l.get("net_iqd", l["total_iqd"])
            x[3] += l.get("net", l["total"]) if usd else 0
    total_buy = sum(v["total_iqd"] for v in inv)
    usd_buy = sum(v["total"] for v in inv if v["cur"] == "دولار")

    types = {t: [0, 0.0, 0.0] for t in ("إعارة", "نقل", "كسر وضياع", "تسوية نقص", "إرجاع")}
    depts, broken = {}, []
    for m in moves:
        t = m["type"]
        if t in types:
            x = types[t]
            x[0] += 1
            x[1] += m["qty"]
            x[2] += _val(items, m["item"], m["qty"])
        d = depts.setdefault(m["to"] if t in ("إعارة", "نقل") else m["frm"] if t in ("إرجاع", "كسر وضياع") else "", [0, 0, 0, 0])
        if t == "إعارة":
            d[0] += m["qty"]
        elif t == "نقل":
            d[1] += m["qty"]
        elif t == "إرجاع":
            d[2] += m["qty"]
        elif t == "كسر وضياع":
            d[3] += m["qty"]
        if t == "كسر وضياع":
            broken.append([_iso(m["date"]), m["item"], m["qty"], m["frm"] or "—", m["note"] or "", _val(items, m["item"], m["qty"])])
    depts.pop("", None)
    out_qty = types["إعارة"][1] + types["نقل"][1] + types["كسر وضياع"][1] + types["تسوية نقص"][1]
    return {
        "kind": "movement", "title": "تقرير المشتريات والإخراج", "subtitle": _period(a, b),
        "cards": [
            {"label": "فواتير الشراء", "value": len(inv)},
            {"label": "قيمة المشتريات (دينار)", "value": total_buy, "money": True, "sub": f"منها {usd_buy:,.0f} دولار أصلي" if usd_buy else ""},
            {"label": "إجمالي الإخراج (قطعة)", "value": out_qty},
            {"label": "قيمة الكسر والضياع", "value": types["كسر وضياع"][2], "money": True, "tone": "bad" if types["كسر وضياع"][2] else ""},
        ],
        "tables": [
            _table("المشتريات حسب المورد", ["المورد", "عدد الفواتير", "الإجمالي (دينار)", "منه دينار أصلي", "منه دولار أصلي"],
                   [[k, *v] for k, v in sorted(by_sup.items(), key=lambda x: -x[1][1])], (1, 2, 3, 4)),
            _table("المشتريات حسب المادة", ["الكود", "المادة", "الكمية", "الإجمالي (دينار)", "منه دولار أصلي"],
                   [[k, v[0], qty_text(items.get(v[0]), v[1]), v[2], v[3]] for k, v in sorted(by_item.items(), key=lambda x: -x[1][2])], (3, 4)),
            _table("الإخراج والإرجاع حسب النوع", ["النوع", "عدد الحركات", "القطع", "القيمة التقديرية"],
                   [[t, v[0], v[1], v[2]] for t, v in types.items() if v[0]], (1, 2, 3)),
            _table("حركة الأقسام", ["القسم", "إعارة (قطعة)", "نقل (قطعة)", "رجّع للمخزن", "كسر وضياع"],
                   [[k, *v] for k, v in sorted(depts.items())], (1, 2, 3, 4)),
            _table("الكسر والضياع بالتفصيل", ["التاريخ", "المادة", "القطع", "القسم", "ملاحظة", "القيمة التقديرية"], broken, (2, 5)),
        ],
        "note": f"المشتريات بالدولار تتحول للدينار بسعر الصرف وقت كل فاتورة. قيمة الإخراج والكسر تقديرية حسب سعر القطعة الحالي (الدولار بسعر الصرف الحالي {data.get('settings', {}).get('rate', 0):,.0f}).",
    }


# ---------------------------------------------------------------- تقرير قسم
def breakage(data, a="", b=""):
    items = _items(data)
    rows = [m for m in data["moves"] if m["type"] == "كسر وضياع" and _in(m["date"], a, b)]
    by_item, by_loc, detail = {}, {}, []
    for m in sorted(rows, key=lambda x: (_iso(x["date"]), x.get("row", 0))):
        it = items.get(m["item"], {})
        price = it.get("price_iqd") or 0
        val = m["qty"] * price
        x = by_item.setdefault(m["item"], [it.get("code", ""), 0, 0.0])
        x[1] += m["qty"]
        x[2] += val
        loc = m["frm"] or "المخزن"
        y = by_loc.setdefault(loc, [0, 0.0])
        y[0] += m["qty"]
        y[1] += val
        detail.append([_iso(m["date"]), it.get("code", ""), m["item"], qty_text(it, m["qty"]), price, val, loc, m["note"] or ""])
    total_q = sum(v[1] for v in by_item.values())
    total_v = sum(v[2] for v in by_item.values())
    top = max(by_item.items(), key=lambda kv: kv[1][2])[0] if by_item else ""
    return {
        "kind": "breakage", "title": "تقرير التكسير", "subtitle": _period(a, b),
        "cards": [
            {"label": "عدد المواد المتكسرة", "value": len(by_item)},
            {"label": "إجمالي القطع المتكسرة", "value": total_q, "tone": "bad" if total_q else ""},
            {"label": "قيمة التكسير (دينار)", "value": total_v, "money": True, "tone": "bad" if total_v else ""},
            {"label": "عدد السجلات", "value": len(rows), "sub": f"الأكثر خسارة: {top}" if top else ""},
        ],
        "tables": [
            _table("التكسير حسب المادة", ["الكود", "المادة", "الكمية المتكسرة", "سعر القطعة (دينار)", "القيمة (دينار)"],
                   [[v[0], k, qty_text(items.get(k), v[1]), items.get(k, {}).get("price_iqd") or 0, v[2]]
                    for k, v in sorted(by_item.items(), key=lambda kv: -kv[1][2])], (3, 4)),
            _table("التكسير حسب القسم / الجهة", ["القسم", "القطع", "القيمة (دينار)"],
                   [[k, *v] for k, v in sorted(by_loc.items(), key=lambda kv: -kv[1][1])], (1, 2)),
            _table("التفاصيل", ["التاريخ", "الكود", "المادة", "الكمية", "سعر القطعة (دينار)", "القيمة (دينار)", "القسم", "ملاحظة"], detail, (4, 5)),
        ],
        "note": f"القيمة = الكمية المتكسرة × سعر القطعة الحالي بالدينار (مواد الدولار بسعر الصرف الحالي {data.get('settings', {}).get('rate', 0):,.0f}).",
    }


def department(data, dept, a="", b=""):
    if not dept:
        raise ValueError("اختار القسم")
    items = _items(data)
    code_of = {i["name"]: i["code"] for i in data["items"]}
    rec, ret, brk = {}, {}, {}
    for m in data["moves"]:
        if not _in(m["date"], a, b):
            continue
        if m["type"] in ("إعارة", "نقل") and m["to"] == dept:
            rec[m["item"]] = rec.get(m["item"], 0) + m["qty"]
        elif m["type"] == "إرجاع" and m["frm"] == dept:
            ret[m["item"]] = ret.get(m["item"], 0) + m["qty"]
        elif m["type"] == "كسر وضياع" and m["frm"] == dept:
            brk[m["item"]] = brk.get(m["item"], 0) + m["qty"]
    loans = [l for l in data.get("loans", []) if l["dept"] == dept]
    held = {i["name"]: i["holdings"].get(dept, 0) for i in data["items"] if abs(i["holdings"].get(dept, 0)) > 1e-9}
    names = sorted(set(rec) | set(ret) | set(brk) | set(held), key=lambda n: code_of.get(n, ""))
    rows = [[code_of.get(n, ""), n, rec.get(n, 0), ret.get(n, 0), brk.get(n, 0), held.get(n, 0), _val(items, n, rec.get(n, 0))] for n in names]
    loan_rows = [[l["no"], _iso(l["due"]), ("متأخرة %d يوم" % abs(l["days"]) if l["status"] == "متأخرة" else "موعدها اليوم" if l["status"] == "اليوم" else "بعد %d يوم" % l["days"]) if l["days"] is not None else "",
                  "، ".join(f"{x['name']} × {x['remaining']:g}" for x in l["lines"])] for l in loans]
    return {
        "kind": "dept", "title": f"تقرير قسم {dept}", "subtitle": _period(a, b),
        "cards": [
            {"label": "قطع استلمها من المخزن", "value": sum(rec.values())},
            {"label": "قطع رجّعها", "value": sum(ret.values())},
            {"label": "كسر وضياع (قطعة)", "value": sum(brk.values()), "tone": "bad" if brk else ""},
            {"label": "موجود عند القسم الآن (قطعة)", "value": sum(held.values())},
            {"label": "إعارات مفتوحة", "value": len(loans), "tone": "warn" if loans else ""},
        ],
        "tables": [
            _table("المواد", ["الكود", "المادة", "استلم", "رجّع", "كسر وضياع", "الموجود عنده الآن", "قيمة المستلم"], rows, (2, 3, 4, 5, 6)),
            _table("الإعارات المفتوحة عند القسم", ["رقم الإعارة", "موعد الاسترجاع", "الحالة", "المواد المتبقية"], loan_rows),
        ],
        "note": "«الموجود عنده الآن» = كل الي استلمه القسم − الي رجّعه − الكسر المسجل عليه، من أول يوم (مو ضمن الفترة). باقي الأرقام ضمن الفترة المحددة.",
    }


# ---------------------------------------------------------------- كشف حساب مورد
def supplier_statement(data, name, a="", b=""):
    if not name:
        raise ValueError("اختار المورد")
    sup = next((s for s in data["suppliers"] if s["name"] == name), None)
    if sup is None:
        raise ValueError("المورد غير موجود")
    items = _items(data)
    inv = sorted((v for v in data["invoices"] + _cl_purchases(data) if v["type"] == "شراء" and v["party"] == name and _in(v["date"], a, b)),
                 key=lambda v: (_iso(v["date"]), v["no"]))
    run, rows, by_item, usd, iqd = 0.0, [], {}, 0.0, 0.0
    for v in inv:
        run += v["total_iqd"]
        if v["cur"] == "دولار":
            usd += v["total"]
        else:
            iqd += v["total"]
        rows.append([v["no"], _iso(v["date"]), len(v["lines"]), v["cur"], v["total"], v["total_iqd"], run, v["note"] or "", v.get("discount", 0.0)])
        for l in v["lines"]:
            x = by_item.setdefault((l["code"], v["cur"]), dict(name=l["name"], qty=0.0, own=0.0, iqd=0.0, price=0.0, last=""))
            x["qty"] += l["qty"]
            x["own"] += l.get("net", l.get("total", 0.0))
            x["iqd"] += l.get("net_iqd", l["total_iqd"])
            if _iso(v["date"]) >= x["last"]:
                x["last"], x["price"] = _iso(v["date"]), l.get("price", 0.0)
    pref = sup.get("currency") or "دينار"
    item_rows = [[k[0], x["name"], qty_text(items.get(x["name"]), x["qty"]), x["price"], k[1], x["own"], x["iqd"]]
                 for k, x in sorted(by_item.items(), key=lambda kv: -kv[1]["iqd"])]
    return {
        "kind": "supplier", "title": f"كشف حساب المورد: {name}",
        "subtitle": _period(a, b) + f" · عملة التعامل: {pref}" + (f" · الهاتف {sup['phone']}" if sup["phone"] else ""),
        "cards": [{"label": "عدد الفواتير", "value": len(inv)}, {"label": "إجمالي المشتريات (دينار)", "value": run, "money": True},
                  {"label": "منه دينار أصلي", "value": iqd, "money": True}, {"label": "منه دولار أصلي", "value": usd, "money": True},
                  {"label": "عدد المواد المشتراة", "value": len({k[0] for k in by_item})}],
        "tables": [
            _table("شنو اشتريت من هذا المورد", ["الكود", "المادة", "الكمية", "آخر سعر للقطعة", "العملة", "الإجمالي (بعملتها)", "بالدينار"], item_rows, (3, 5, 6)),
            _table("الفواتير", ["رقم الفاتورة", "التاريخ", "عدد المواد", "العملة", "الصافي بعد الخصم (بعملتها)", "بالدينار", "المجموع التراكمي (دينار)", "ملاحظات", "الخصم (بعملتها)"], rows, (2, 4, 5, 6, 8)),
        ],
        "note": "الكشف يعرض المشتريات فقط. المدفوعات غير مسجلة بالبرنامج. فواتير الدولار تتحول للدينار بسعر الصرف وقت الفاتورة. الأسعار والإجماليات «بعملتها» هي بعملة الفاتورة الأصلية.",
    }


# ---------------------------------------------------------------- قائمة طلب شراء
def reorder(data, factor=2.0):
    try:
        factor = float(factor)
    except (TypeError, ValueError):
        factor = 2.0
    if not math.isfinite(factor) or factor < 1:
        factor = 2.0
    phones = {s["name"]: s["phone"] for s in data["suppliers"]}
    groups, total_cost, n = {}, 0.0, 0
    for i in sorted(data["items"], key=lambda x: x["code"]):
        if not i["min"] or i["balance"] > i["min"]:
            continue
        need = max(0, math.ceil(i["min"] * factor - i["balance"]))
        if need <= 0:
            continue
        cartons = math.ceil(need / i["carton"]) if i["carton"] else 0
        order_pcs = cartons * i["carton"] if cartons else need
        cost = order_pcs * (i["price_iqd"] or 0)
        total_cost += cost
        n += 1
        ask = f"{cartons} كارتون ({order_pcs:g} قطعة)" if cartons else f"{order_pcs:g} قطعة"
        groups.setdefault(i["supplier"] or "", []).append(
            [i["code"], i["name"], i["usage"] or "", i["balance"], i["min"], ask, cost])
    tables = []
    for sup in sorted(groups, key=lambda x: (x == "", x)):
        title = f"المورد: {sup}" + (f" ({phones.get(sup)})" if phones.get(sup) else "") if sup else "بدون مورد معتاد"
        tables.append(_table(title, ["الكود", "المادة", "الاستخدام", "الرصيد (قطعة)", "الحد الأدنى", "المطلوب", "التكلفة التقديرية"], groups[sup], (3, 4, 6)))
    return {
        "kind": "reorder", "title": "قائمة طلب شراء", "subtitle": f"تغطية حتى {factor:g} × الحد الأدنى",
        "cards": [{"label": "مواد تحتاج طلب", "value": n, "tone": "warn" if n else ""}, {"label": "التكلفة التقديرية", "value": total_cost, "money": True}],
        "tables": tables or [_table("ما كو مواد تحتاج طلب", ["—"], [])],
        "note": "المطلوب يُقرَّب لأقرب كارتون كامل إذا للمادة كارتون. الحد الأدنى لازم يكون محدد للمادة حتى تدخل القائمة. التكلفة بالدينار (مواد الدولار بسعر الصرف الحالي).",
    }


def location(data):
    """توزيع كل مادة: بالمخزن وعند كل قسم."""
    depts = list(data["settings"].get("depts", []))
    for i in data["items"]:
        for k in i["holdings"]:
            if k not in depts:
                depts.append(k)
    rows = []
    for i in sorted(data["items"], key=lambda x: x["code"]):
        if not i["balance"] and not i["held"]:
            continue
        rows.append([i["code"], i["name"], i["balance"], *[i["holdings"].get(d, 0) for d in depts], i["total"]])
    wh_val = sum(i["value"] for i in data["items"])
    tot_val = sum(i["value_total"] for i in data["items"])
    return {
        "kind": "location", "title": "توزيع المواد (المخزن والأقسام)", "subtitle": f"بتاريخ {_iso(__import__('datetime').date.today())}",
        "cards": [{"label": "مواد", "value": len(rows)}, {"label": "قيمة المخزن (دينار)", "value": wh_val, "money": True},
                  {"label": "قيمة الي عند الأقسام (دينار)", "value": tot_val - wh_val, "money": True}],
        "tables": [_table("الكميات بالقطعة", ["الكود", "المادة", "بالمخزن", *depts, "الإجمالي"], rows, tuple(range(2, 4 + len(depts))))],
        "note": "«بالمخزن» هو المتبقي القابل للصرف. الأقسام: ما استلمته (نقل وإعارة) ناقص ما رجّعته والكسر المسجل عليها.",
    }


def _cl_text(l):
    return f"{l['name']}{(' ' + l['size']) if l['size'] else ''}{(' ' + l['color']) if l['color'] else ''} × {l['qty']:g}"


def employee(data, name, a="", b=""):
    """كشف موظف: كل ورقات الاستلام."""
    import clothes
    name = (name or "").strip()
    if not name:
        raise ValueError("اكتب اسم الموظف")
    docs = [d for d in data["clothes"]["docs"] if clothes._norm(d["employee"]) == clothes._norm(name) and _in(d["date"], a, b)]
    if not docs:
        raise ValueError("ما كو استلامات بهذا الاسم")
    docs.sort(key=lambda d: (_iso(d["date"]), d["no"]))
    shown = docs[0]["employee"]
    summary = {}
    for d in docs:
        for l in d["lines"]:
            k = (l["name"], l["size"], l["color"])
            summary[k] = summary.get(k, 0) + l["qty"]
    return {
        "kind": "employee", "title": f"كشف استلام الملابس: {shown}",
        "subtitle": _period(a, b) + f" · أول استلام {_iso(docs[0]['date'])} · آخر استلام {_iso(docs[-1]['date'])}",
        "cards": [{"label": "عدد الاستلامات", "value": len(docs)}, {"label": "مجموع القطع", "value": sum(d["total"] for d in docs)}],
        "tables": [
            _table("ورقات الاستلام", ["رقم الورقة", "التاريخ", "الملابس", "القطع", "ملاحظات"],
                   [[d["no"], _iso(d["date"]), "، ".join(_cl_text(l) for l in d["lines"]), d["total"], d["note"] or ""] for d in docs], (3,)),
            _table("المجموع حسب الملبس", ["الملبس", "المقاس", "اللون", "القطع"],
                   [[k[0], k[1] or "—", k[2] or "—", v] for k, v in sorted(summary.items())], (3,)),
        ],
        "note": "الاستلام الجديد بعد مدة يتسجل بورقة جديدة، وتكتب بالملاحظات «موافقة الإدارة» إذا تريد.",
    }


def clothes_stock(data):
    items = data["clothes"]["items"]
    return {
        "kind": "clothes", "title": "مخزون الملابس", "subtitle": f"بتاريخ {_iso(__import__('datetime').date.today())}",
        "cards": [{"label": "أنواع (مقاس/لون)", "value": len(items)}, {"label": "مجموع القطع بالمخزن", "value": sum(i["balance"] for i in items)},
                  {"label": "ناقصة أو نافدة", "value": sum(1 for i in items if i["status"] != "متوفر"), "tone": "warn"}],
        "tables": [_table("الملابس", ["الكود", "الاسم", "المقاس", "اللون", "المتبقي", "الحد الأدنى", "الحالة"],
                          [[i["code"], i["name"], i["size"] or "—", i["color"] or "—", i["balance"], i["min"], i["status"]] for i in items], (4, 5))],
        "note": "",
    }


def build(kind, data, a="", b="", dept="", supplier="", factor=2.0, employee_name=""):
    if kind == "movement":
        return movement(data, a, b)
    if kind == "dept":
        return department(data, dept, a, b)
    if kind == "supplier":
        return supplier_statement(data, supplier, a, b)
    if kind == "reorder":
        return reorder(data, factor)
    if kind == "location":
        return location(data)
    if kind == "employee":
        return employee(data, employee_name, a, b)
    if kind == "clothes":
        return clothes_stock(data)
    if kind == "breakage":
        return breakage(data, a, b)
    raise ValueError("نوع التقرير غير معروف")


# ---------------------------------------------------------------- تصدير إكسل
def build_xlsx(rep, restaurant, logo_path=None):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    wb = Workbook()
    ws = wb.active
    ws.title = "التقرير"
    ws.sheet_view.rightToLeft = True
    F = "Arial"
    hf, hfill = Font(name=F, bold=True, color="FFFFFF"), PatternFill("solid", fgColor="1F4E78")
    thin = Side(style="thin", color="BFBFBF")
    bd = Border(left=thin, right=thin, top=thin, bottom=thin)
    ws["A1"], ws["A2"], ws["A3"] = restaurant, rep["title"], rep["subtitle"]
    ws["A1"].font, ws["A2"].font, ws["A3"].font = Font(name=F, bold=True, size=16, color="1F4E78"), Font(name=F, bold=True, size=13), Font(name=F, size=11, color="555555")
    r = 5
    for c in rep["cards"]:
        ws.cell(r, 1, c["label"]).font = Font(name=F, bold=True)
        v = ws.cell(r, 2, c["value"])
        v.font = Font(name=F, bold=True, size=12)
        v.number_format = "#,##0"
        if c.get("sub"):
            ws.cell(r, 3, c["sub"]).font = Font(name=F, size=10, color="777777")
        r += 1
    r += 1
    width = {}
    for t in rep["tables"]:
        ws.cell(r, 1, t["title"]).font = Font(name=F, bold=True, size=12, color="1F4E78")
        r += 1
        for j, col in enumerate(t["columns"], 1):
            x = ws.cell(r, j, col)
            x.font, x.fill, x.border, x.alignment = hf, hfill, bd, Alignment(horizontal="center", vertical="center", wrap_text=True)
            width[j] = max(width.get(j, 10), len(str(col)) + 4)
        r += 1
        for row in t["rows"]:
            for j, v in enumerate(row, 1):
                x = ws.cell(r, j, v)
                x.font, x.border = Font(name=F), bd
                x.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                if isinstance(v, (int, float)):
                    x.number_format = "#,##0"
                width[j] = min(48, max(width.get(j, 10), len(str(v)) + 3))
            r += 1
        r += 1
    if rep.get("note"):
        ws.cell(r, 1, rep["note"]).font = Font(name=F, italic=True, size=10, color="777777")
    for j, w in width.items():
        ws.column_dimensions[ws.cell(1, j).column_letter].width = w
    ws.column_dimensions["A"].width = max(ws.column_dimensions["A"].width or 10, 26)
    if logo_path and os.path.exists(logo_path):
        try:
            from openpyxl.drawing.image import Image as XImage
            img = XImage(logo_path)
            img.width = img.height = 60
            ws.add_image(img, "F1")
            ws.row_dimensions[1].height = 30
        except Exception:
            pass
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()
