# -*- coding: utf-8 -*-
"""
merge_local_imports.py
======================
بيرجّع أسماء الأوبريتور لكل المتصفحات.

المشكلة: أسماء الأوبريتور كانت محفوظة بس جوه localStorage بتاع كروم (من الـ Drag & Drop)،
والبيانات الأساسية (metals_data.json / metals_data.js / index.html) مفيهاش حقل operator.

الحل: صدّر البيانات من كروم (الخطوة 1 في الشرح) واحفظها باسم chrome_imports.json جنب الملفات،
وبعدين شغّل:
    python merge_local_imports.py
أو لو اسم الملف مختلف:
    python merge_local_imports.py "المسار\\الملف.json"

بيدمج السجلات في metals_data.json و metals_data.js ويضمّنها في index.html
(بنفس دوال update_daily_data.py) ويحدّث BUILD_VERSION.
"""

import os
import sys
import json

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import update_daily_data as u

if sys.platform.startswith('win'):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

KEEP = ['workorder', 'joborder', 'label_no', 'cable', 'size', 'material', 'familyid',
        'drumno', 'length', 'weight', 'customer', 'shift', 'machine',
        'dateIso', 'dateDisplay', 'operator']


def main():
    src = sys.argv[1] if len(sys.argv) > 1 else 'chrome_imports.json'
    if not os.path.exists(src):
        sys.exit(f"❌ الملف مش موجود: {src}\nصدّره من كروم الأول (راجع الشرح) واحفظه بالاسم ده جنب index.html")

    with open(src, 'r', encoding='utf-8') as f:
        raw = json.load(f)
    if not isinstance(raw, list) or not raw:
        sys.exit("❌ الملف فاضي أو مش بصيغة صحيحة")

    records = [{k: r.get(k, '') for k in KEEP} for r in raw if r.get('dateIso')]
    dates = {r['dateIso'] for r in records}
    with_op = sum(1 for r in records if r['operator'])
    print(f"📥 {len(records)} سجل | {with_op} منهم فيهم اسم أوبريتور | التواريخ: {sorted(dates)}")

    u.merge_and_save_json(records, dates)
    u.update_html_cache_buster()
    print("\n✅ خلاص. ارفع metals_data.json و metals_data.js و index.html (أو شغّل update_daily_data.py --push)")
    print("   وفي أي متصفح تاني اعمل Ctrl+F5.")


if __name__ == '__main__':
    main()
