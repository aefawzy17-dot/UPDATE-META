# -*- coding: utf-8 -*-
"""
fix_material_from_plan.py
=========================
بيصحح نوع الموصل (AL / CU) في البيانات القديمة حسب عمود Metal في شيت Orders بالماستر بلان.
السبب: كان الكود بيعتبر أي سجل فيه وزن نحاس (عمود CU) إنه كابل نحاس، حتى لو الموصل ألمنيوم
والنحاس ده بتاع الشيلد (CW).

الاستخدام:  python fix_material_from_plan.py        (بياخد أحدث ملف في Master_Plan)
بياخد نسخة احتياطية في backup_before_material_fix. ارفع index.html و metals_data.json و metals_data.js بعدها.
"""
import os, re, sys, shutil
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE); os.chdir(HERE)
import openpyxl
import update_daily_data as u
import update_wo_voltage as w

if sys.platform.startswith('win'):
    try: sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception: pass


def plan_metals(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb['Orders']
    it = ws.iter_rows(values_only=True)
    hdr = [re.sub(r'\s+', ' ', str(c)).strip().lower() if c else '' for c in next(it)]
    iw, im = hdr.index('w.o.'), hdr.index('metal')
    out = {}
    for r in it:
        wo = w.clean_wo(r[iw])
        m = str(r[im] or '').upper()
        if wo and wo not in out:
            if m.startswith('CU') or 'COPPER' in m:
                out[wo] = 'CU'
            elif m.startswith('AL'):
                out[wo] = 'AL'
    wb.close()
    return out


def main():
    src = w.find_latest_plan()
    if not src:
        sys.exit("حط شيت الماستر بلان في مجلد Master_Plan")
    pm = plan_metals(src)
    data = u.get_complete_existing_data()
    os.makedirs('backup_before_material_fix', exist_ok=True)
    for fn in ('index.html', 'metals_data.json', 'metals_data.js'):
        if os.path.exists(fn):
            shutil.copy2(fn, os.path.join('backup_before_material_fix', fn))
    fixed = {}
    for r in data:
        want = pm.get(r.get('workorder'))
        if want and r.get('material') != want:
            old = r['material']
            r['material'] = want
            c = r.get('cable') or ''
            r['cable'] = c.replace(f',{old},', f',{want},') if f',{old},' in c else c
            fixed[(r['workorder'], old, want)] = fixed.get((r['workorder'], old, want), 0) + 1
    for (wo, o, n), cnt in sorted(fixed.items()):
        print(f"  {wo}: {o} -> {n}  ({cnt} سجل)")
    print(f"✅ اتصحح {sum(fixed.values())} سجل")
    u.merge_and_save_json(data, {r.get('dateIso') for r in data})
    u.update_html_cache_buster()
    try:
        with open('index.html', 'r', encoding='utf-8') as f: h = f.read()
        h = h.replace('<div id="toastNotification" class="toast-notification">',
                      '<div id="toastNotification" class="toast-notification hidden">')
        with open('index.html', 'w', encoding='utf-8') as f: f.write(h)
    except Exception: pass


if __name__ == '__main__':
    main()
