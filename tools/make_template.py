"""يبني template.xlsx النظيف (1000 صنف، سجل حركات حتى 20000 سطر) من ملف الإكسل الأساسي."""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
from copy import copy
from openpyxl.worksheet.datavalidation import DataValidation
from core import *

LOG_END = 20004
src, dst = sys.argv[1], sys.argv[2]
import shutil; shutil.copy(src, dst)
s = Store(dst)
wb = s._open()
ws, lg = wb[ITEMS_SHEET], wb[LOG_SHEET]
rng_b, rng_c, rng_d, rng_e = (f"'{LOG_SHEET}'!${c}${L0}:${c}${LOG_END}" for c in "BCDE")
ref = 204
for r in range(I0, I_MAX + 1):
    if r > ref:                                   # انسخ تنسيق صف 204 للصفوف الجديدة
        for c in range(1, 16):
            a, b = ws.cell(ref, c), ws.cell(r, c)
            b.font, b.border, b.alignment, b.number_format, b.fill = copy(a.font), copy(a.border), copy(a.alignment), a.number_format, copy(a.fill)
        ws.cell(r, 1).value = None
    for c in (1, 2, 3, 4, 5, 6, 10, 13, 14, 15):
        ws.cell(r, c).value = None
    sm = lambda d: f'=IF(B{r}="","",SUMIFS({rng_d},{rng_b},B{r},{rng_e},"{d}"))'
    ws.cell(r, 7, sm("دخول")); ws.cell(r, 8, sm("خروج"))
    ws.cell(r, 9, f'=IF(B{r}="","",E{r}+G{r}-H{r})')
    ws.cell(r, 11, f'=IF(B{r}="","",IF(I{r}<=0,"نفد",IF(I{r}<=J{r},"ناقص - اطلب","متوفر")))')
    ws.cell(r, 12, f'=IF(B{r}="","",I{r}*N(F{r})*IF(R{r}="دولار",\'{SET_SHEET}\'!$B$7,1))')
    ws.cell(r, 6).number_format = ws.cell(r, 12).number_format = "#,##0"
# نطاقات التنسيق الشرطي وقوائم الاختيار
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill
ws.conditional_formatting._cf_rules.clear()
f = lambda c: PatternFill("solid", bgColor=c, fgColor=c)
rg = f"K{I0}:K{I_MAX}"
ws.conditional_formatting.add(rg, FormulaRule(formula=[f'K{I0}="نفد"'], fill=f("F8CBAD")))
ws.conditional_formatting.add(rg, FormulaRule(formula=[f'LEFT(K{I0},4)="ناقص"'], fill=f("FFE699")))
ws.conditional_formatting.add(rg, FormulaRule(formula=[f'K{I0}="متوفر"'], fill=f("C6E0B4")))
ws.data_validations.dataValidation = []
dcur = DataValidation(type="list", formula1='"دينار,دولار"', allow_blank=True)
ws.add_data_validation(dcur); dcur.add(f"R{I0}:R{I_MAX}")
d = DataValidation(type="list", formula1='"' + ",".join(CATEGORIES) + '"', allow_blank=True)
ws.add_data_validation(d); d.add(f"C{I0}:C{I_MAX}")
lg.data_validations.dataValidation = []
v = [(DataValidation(type="list", formula1=f"='{ITEMS_SHEET}'!$B${I0}:$B${I_MAX}", allow_blank=True), "B"),
     (DataValidation(type="list", formula1='"شراء,إرجاع,إعارة,كسر وضياع,نقل"', allow_blank=True), "C"),
     (DataValidation(type="decimal", operator="greaterThan", formula1="0", allow_blank=True), "D")]
for dv, col in v:
    lg.add_data_validation(dv); dv.add(f"{col}{L0}:{col}{LOG_END}")
for r in range(L0, L0 + 3000):                      # صيغة الاتجاه (تشمل التسوية)
    lg.cell(r, 5, DIR_FORMULA.format(r=r))
v2 = DataValidation(type="list", formula1='"' + ",".join(TYPES) + '"', allow_blank=True)
lg.data_validations.dataValidation = [x for x in lg.data_validations.dataValidation if "شراء" not in str(x.formula1)]
lg.add_data_validation(v2); v2.add(f"C{L0}:C{LOG_END}")
# ورقة الملخص: صيغ تغطي 1000 صنف وتعليمات محدّثة
sm = wb["ملخص"]
A_ = f"'{ITEMS_SHEET}'"
sm["B3"] = f"=COUNTA({A_}!B{I0}:B{I_MAX})"
sm["B4"] = f'=COUNTIF({A_}!K{I0}:K{I_MAX},"ناقص*")'
sm["B5"] = f'=COUNTIF({A_}!K{I0}:K{I_MAX},"نفد")'
sm["B6"] = f"=SUM({A_}!L{I0}:L{I_MAX})"
for r in range(9, 13):
    sm[f"B{r}"] = f"=SUMIFS({A_}!$L${I0}:$L${I_MAX},{A_}!$C${I0}:$C${I_MAX},A{r})"
for r in range(15, 24):
    sm[f"A{r}"] = None
tips = ["هذا الملف يقراه ويكتب بيه البرنامج (مستلزمات مطعم بيت الكل). سكّره قبل لا تحفظ من البرنامج.",
        "الأصناف: الكود والاسم والاستخدام والكمية الافتتاحية والسعر وقطع الكارتون. الرصيد والحالة والقيمة تنحسب لحالها.",
        "حركة: كل حركات المخزن. الاتجاه: شراء وإرجاع وتسوية زيادة = دخول | إعارة وكسر وضياع وتسوية نقص = خروج | نقل = بدون تغيير الرصيد.",
        "الفواتير وبنود_الفواتير والموردين: تنكتب من البرنامج. لا تغيّر ترتيب الأعمدة.",
        "الكميات بالقطعة دائماً. أحمر = نفد، أصفر = ناقص، أخضر = متوفر.",
        "الأسعار: كل مادة بعملتها (دينار أو دولار). قيمة المخزون تتحول للدينار حسب سعر الصرف بورقة «إعدادات» (الخلية B7)."]
from openpyxl.styles import Font as F_
for i, t in enumerate(tips):
    sm.cell(15 + i, 1, f"{i + 1}) {t}").font = F_(name="Arial", size=11)
# شعار المطعم داخل ملف الإكسل
from openpyxl.drawing.image import Image as XImage
LOGO = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "assets", "logo-xl.png")
for name, anchor in ((ITEMS_SHEET, "L1"), (INV_SHEET, "I1"), (SUP_SHEET, "C1"), (LOG_SHEET, "F1"), ("ملخص", "C1")):
    w_ = wb[name]
    w_._images = []
    img = XImage(LOGO); img.width = img.height = 62
    w_.add_image(img, anchor)
    w_.row_dimensions[1].height = 34
    w_.row_dimensions[2].height = 18
wb[SET_SHEET]["B5"] = DEFAULT_NAME
s._save(wb)
print("template ok:", len(s.items()), "items")
