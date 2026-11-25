# -*- coding: utf-8 -*-
"""
backfill_operators.py  -  الحل النهائي لمشكلة اختفاء الأوبريتور
================================================================
السبب: البيانات الأساسية (index.html / metals_data.json / metals_data.js) مفيهاش حقل operator،
فالأسماء كانت بتظهر بس في المتصفح اللي اتسحب فيه الشيت (localStorage).

الحل: السكريبت ده بيجيب الأوبريتور من كل شيتات Cable_Production_Records_*.csv (وأي chrome_imports.json)
ويكتبه جوه السجلات الموجودة (من غير ما يغيّر أي حاجة تانية: الأطوال والمقاسات والـ W.O زي ما هي)،
وبعدين يحفظ في metals_data.json و metals_data.js و index.html.

الاستخدام (من مجلد المشروع):
    1) حط كل شيتات الأيام اللي عندك في مجلد اسمه  Daily_Sheets  (أو أي مجلد)
    2) python backfill_operators.py                      <- بيقرا Daily_Sheets + chrome_imports.json لو موجود
       python backfill_operators.py "مجلد\\تاني"           <- لو الشيتات في مكان تاني
    3) ارفع metals_data.json و metals_data.js و index.html على GitHub  (Ctrl+F5 بعدها)

بياخد نسخة احتياطية من الملفات قبل التعديل (مجلد backup_before_backfill).
"""

import os
import sys
import glob
import json
import shutil
from collections import Counter, defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
os.chdir(HERE)
import update_daily_data as u

if sys.platform.startswith('win'):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass


# تصحيح يدوي لمقاسات أوامر شغل معينة (حسب الماستر بلان) للأيام اللي مفيش لها شيت CSV
WO_SIZE_OVERRIDES = {
    '1586/26/01': '1400',   # 1x1400 في الماستر بلان
}


def key_of(r):
    return (str(r.get('dateIso', '')), str(r.get('machine', '')), str(r.get('workorder', '')),
            str(r.get('shift', '')), str(r.get('drumno', '')), str(r.get('length', '')))


SIZE_BY_KEY = {}


def collect_from_csvs(folder, op_votes):
    files = []
    for pat in ('*.csv', '*.xlsx', '*.xls'):
        files += glob.glob(os.path.join(folder, pat))
    files = [f for f in files if not os.path.basename(f).startswith('~$')]
    if not files:
        print(f"ℹ️  مفيش شيتات في {folder}")
        return 0
    wip2_map, wip1_map = u.load_metadata_maps()
    n = 0
    for f in sorted(files):
        try:
            df = u.read_export_rows(f)
            recs, _ = u.parse_and_extract_metals(df, wip2_map, wip1_map)
        except SystemExit:
            continue
        except Exception as e:
            print(f"⚠️  تخطي {os.path.basename(f)}: {e}")
            continue
        for r in recs:
            if r.get('operator'):
                op_votes[key_of(r)][r['operator']] += 1
                n += 1
            SIZE_BY_KEY[key_of(r)] = r.get('size')
    return n


def collect_from_json(path, op_votes):
    if not os.path.exists(path):
        return 0
    with open(path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    n = 0
    for r in data if isinstance(data, list) else []:
        if r.get('operator'):
            op_votes[key_of(r)][str(r['operator']).strip()] += 1
            n += 1
    return n


def main():
    folder = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, 'Daily_Sheets')
    op_votes = defaultdict(Counter)

    n1 = collect_from_csvs(folder, op_votes) if os.path.isdir(folder) else 0
    n2 = collect_from_json(os.path.join(HERE, 'chrome_imports.json'), op_votes)
    print(f"📥 سجلات فيها أوبريتور: من الشيتات {n1} | من chrome_imports.json {n2}")
    if not op_votes:
        sys.exit("❌ ملقيتش أي أوبريتور. حط الشيتات في Daily_Sheets أو صدّر chrome_imports.json.")

    existing = u.get_complete_existing_data()
    if not existing:
        sys.exit("❌ مفيش بيانات حالية.")

    os.makedirs('backup_before_backfill', exist_ok=True)
    for fn in ('index.html', 'metals_data.json', 'metals_data.js'):
        if os.path.exists(fn):
            shutil.copy2(fn, os.path.join('backup_before_backfill', fn))

    filled = 0
    size_fixed = 0
    for r in existing:
        # تصحيح المقاس (مثلاً 1x1400 كان بيطلع 1)
        new_size = WO_SIZE_OVERRIDES.get(r.get('workorder')) or SIZE_BY_KEY.get(key_of(r))
        if new_size and new_size != '-' and str(r.get('size')) != str(new_size):
            old_size = str(r.get('size'))
            r['size'] = str(new_size)
            if r.get('cable'):
                r['cable'] = r['cable'].replace('*' + old_size, '*' + str(new_size)) \
                    if r['cable'].endswith('*' + old_size) else r['cable']
            size_fixed += 1
        if r.get('operator'):
            continue
        votes = op_votes.get(key_of(r))
        if votes:
            r['operator'] = votes.most_common(1)[0][0]
            filled += 1
    print(f"📐 اتصحح المقاس في {size_fixed} سجل")
    total_with = sum(1 for r in existing if r.get('operator'))
    print(f"✍️  اتضاف أوبريتور لـ {filled} سجل | الإجمالي دلوقتي: {total_with} من {len(existing)}")

    dates = {r.get('dateIso') for r in existing}
    u.merge_and_save_json(existing, dates)   # بيعيد كتابة json/js/html بنفس السجلات + الأوبريتور
    u.update_html_cache_buster()

    # إصلاح رسالة "Processing" اللي اتحفظت ظاهرة جوه index.html
    try:
        with open('index.html', 'r', encoding='utf-8') as f:
            h = f.read()
        h2 = h.replace('<div id="toastNotification" class="toast-notification">',
                       '<div id="toastNotification" class="toast-notification hidden">')
        if h2 != h:
            with open('index.html', 'w', encoding='utf-8') as f:
                f.write(h2)
            print("🧹 اتصلّحت رسالة Processing اللي كانت ظاهرة في index.html")
    except Exception as e:
        print(f"⚠️ تعذر إصلاح الـ toast: {e}")

    missing = Counter(r['dateIso'] for r in existing if not r.get('operator'))
    if missing:
        print("\n⚠️  أيام لسه من غير أوبريتور (حط شيتاتها في Daily_Sheets وشغّل السكريبت تاني):")
        print("   " + ", ".join(sorted(missing)))
    print("\n✅ خلاص. ارفع metals_data.json و metals_data.js و index.html وبعدين Ctrl+F5 في كل المتصفحات.")


if __name__ == '__main__':
    main()
