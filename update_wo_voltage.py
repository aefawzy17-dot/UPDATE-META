# -*- coding: utf-8 -*-
"""
update_wo_voltage.py
====================
يربط شيت Master Production Plan بالويب (Metals Dashboard):
يقرأ كل أمر شغل (W.O) وفئة الكابل بتاعته (Low Voltage / Medium Voltage / ...)
ويكتب الملف wo_voltage.js اللي الويب بيقرأ منه عشان يعرض LV / MV
بين البرسنتج (%) والمقاس (Size) في جدول الماكينات.

الاستخدام (الأسهل):
    حط أحدث شيت ماستر بلان (.xlsx) جوه مجلد  Master_Plan  (جنب index.html)
    وبعدين شغّل  تحديث_الماستر_بلان.bat  أو  python update_wo_voltage.py
    (بياخد أحدث ملف .xlsx في المجلد حسب تاريخ التعديل، وكمان بيتشغل تلقائي مع update_daily_data.py)

أو حدد الملف يدوياً:
    python update_wo_voltage.py "Master Production Plan Nov-2026.xlsx"

المطلوب:  pip install openpyxl

ترتيب الأولوية لمصدر التصنيف لكل W.O:
    1) شيت Orders        -> عمود Cable Category   (الأدق: أوردرات خطة الشهر)
    2) شيت Backlog 2     -> عمود family           (لأوردرات أقدم مش في خطة الشهر)
    3) شيت 90 for ADV.   -> عمود family
أي W.O مش موجود في الثلاثة بيفضل على التصنيف القديم الموجود في بيانات الويب.
"""

import os
import sys
import re
import glob
import json
import datetime

if sys.platform.startswith('win'):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

try:
    import openpyxl
except ImportError:
    sys.exit("openpyxl غير مثبت. نفّذ:  pip install openpyxl")

# Cable Category (شيت Orders) -> كود قصير بنفس أكواد family في باقي الشيتات
CATEGORY_TO_CODE = {
    'low voltage': 'LV',
    'medium voltage': 'MV',
    'high voltage': 'HV',
    'control cables': 'CC',
    'building wires': 'BC',
    'bare copper': 'BARE',
    'overhead': 'OH',
    'solar': 'SOLAR',
    'solid': 'SOLID',
    'telecom': 'TC',
}

WO_RE = re.compile(r'^\d{1,5}/\d{2}/\d{1,3}')


def clean_wo(v):
    if v is None:
        return None
    s = re.sub(r'\s+', ' ', str(v)).strip()
    return s if WO_RE.match(s) else None


def header_index(ws, wanted):
    """يرجّع {اسم العمود: رقم} من أول صف في الشيت (مش بالرقم الثابت عشان لو الأعمدة اتحركت)."""
    first = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
    names = [re.sub(r'\s+', ' ', str(c)).strip().lower() if c is not None else '' for c in first]
    out = {}
    for key, cands in wanted.items():
        for c in cands:
            if c in names:
                out[key] = names.index(c)
                break
    return out


def read_orders(wb, result):
    if 'Orders' not in wb.sheetnames:
        return 0
    ws = wb['Orders']
    idx = header_index(ws, {'wo': ['w.o.', 'w.o', 'wo'],
                            'cat': ['cable category', 'cable\ncategory']})
    if 'wo' not in idx or 'cat' not in idx:
        print("⚠️  شيت Orders: مش لاقي عمودي W.O. و Cable Category")
        return 0
    n = 0
    for r in ws.iter_rows(min_row=2, values_only=True):
        wo = clean_wo(r[idx['wo']])
        cat = r[idx['cat']]
        if not wo or not cat:
            continue
        code = CATEGORY_TO_CODE.get(str(cat).strip().lower())
        if code and wo not in result:
            result[wo] = code
            n += 1
    return n


def read_family_sheet(wb, sheet, wo_col_names, result):
    if sheet not in wb.sheetnames:
        return 0
    ws = wb[sheet]
    idx = header_index(ws, {'wo': wo_col_names, 'fam': ['family']})
    if 'wo' not in idx or 'fam' not in idx:
        print(f"⚠️  شيت {sheet}: مش لاقي عمودي W.O و family")
        return 0
    n = 0
    for r in ws.iter_rows(min_row=2, values_only=True):
        wo = clean_wo(r[idx['wo']])
        fam = r[idx['fam']]
        if not wo or not fam:
            continue
        fam = str(fam).strip().upper()
        if fam and wo not in result:
            result[wo] = fam
            n += 1
    return n


HERE = os.path.dirname(os.path.abspath(__file__))
MASTER_DIR = os.path.join(HERE, 'Master_Plan')


def find_latest_plan():
    """أحدث .xlsx في مجلد Master_Plan (أو في مجلد المشروع لو اسمه Master*Plan*)."""
    cands = [f for f in glob.glob(os.path.join(MASTER_DIR, '*.xlsx')) if not os.path.basename(f).startswith('~$')]
    if not cands:
        cands = [f for f in glob.glob(os.path.join(HERE, 'Master*Plan*.xlsx')) if not os.path.basename(f).startswith('~$')]
    return max(cands, key=os.path.getmtime) if cands else None


def auto_update():
    """بيتندهي من update_daily_data.py: لو في شيت ماستر بلان أحدث من wo_voltage.js بيحدّثه، وغير كده بيعدّي."""
    src = find_latest_plan()
    if not src:
        print("ℹ️  مفيش شيت ماستر بلان في مجلد Master_Plan - هيتم استخدام wo_voltage.js الحالي")
        return False
    out = os.path.join(HERE, 'wo_voltage.js')
    if os.path.exists(out) and os.path.getmtime(out) >= os.path.getmtime(src):
        print(f"ℹ️  wo_voltage.js محدّث بالفعل من {os.path.basename(src)}")
        return False
    main(src)
    return True


def main(src=None):
    here = HERE
    if src is None and len(sys.argv) > 1:
        src = sys.argv[1]
    if src is None:
        src = find_latest_plan()
        if not src:
            sys.exit("حط شيت الماستر بلان (.xlsx) جوه مجلد Master_Plan أو مرّر مساره كـ argument")

    print(f"📖 قراءة: {src}")
    wb = openpyxl.load_workbook(src, read_only=True, data_only=True)

    result = {}
    n1 = read_orders(wb, result)
    n2 = read_family_sheet(wb, 'Backlog 2', ['wo'], result)
    n3 = read_family_sheet(wb, '90 for ADV.', ['work_order'], result)
    wb.close()

    ordered = dict(sorted(result.items()))
    meta = {
        'source': os.path.basename(src),
        'generated': datetime.datetime.now().strftime('%Y-%m-%d %H:%M'),
        'from_orders': n1, 'from_backlog': n2, 'from_adv': n3, 'total': len(ordered),
    }
    out = os.path.join(here, 'wo_voltage.js')
    with open(out, 'w', encoding='utf-8') as f:
        f.write('// Auto-generated by update_wo_voltage.py - لا تعدّل يدوياً\n')
        f.write('var WO_VOLTAGE_META = ' + json.dumps(meta, ensure_ascii=False) + ';\n')
        f.write('var WO_VOLTAGE_MAP = ' + json.dumps(ordered, ensure_ascii=False, separators=(',', ':')) + ';\n')

    counts = {}
    for v in ordered.values():
        counts[v] = counts.get(v, 0) + 1
    print(f"✅ تم كتابة {out}")
    print(f"   Orders: {n1} | Backlog 2: +{n2} | 90 for ADV.: +{n3} | الإجمالي: {len(ordered)} أمر شغل")
    print("   " + " | ".join(f"{k}: {v}" for k, v in sorted(counts.items(), key=lambda x: -x[1])))


if __name__ == '__main__':
    main()
