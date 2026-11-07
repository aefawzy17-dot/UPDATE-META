# -*- coding: utf-8 -*-
"""
ci_update.py - بيشتغل تلقائي على GitHub Actions (أو يدوي: python ci_update.py)
==========================================================================
1) أي شيت (.csv/.xlsx) في مجلد incoming/  -> بيتدمج في البيانات (بنفس منطق update_daily_data.py)
   وبعدين بيتنقل لمجلد processed/ .
2) لو الماستر بلان (Master_Plan/*.xlsx) اتغير في آخر commit -> بيتجدد wo_voltage.js.
3) بيحدّث index.html + metals_data.json + metals_data.js + رقم النسخة (cache buster).
لو مفيش حاجة جديدة بيطلع من غير ما يغيّر حاجة.
"""
import os, sys, glob, shutil, subprocess

HERE = os.path.dirname(os.path.abspath(__file__))
os.chdir(HERE); sys.path.insert(0, HERE)
import update_daily_data as u

if sys.platform.startswith('win'):
    try: sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception: pass

INCOMING, PROCESSED = 'incoming', 'processed'


def plan_changed():
    if os.environ.get('FORCE_PLAN') == '1':
        return True
    try:
        out = subprocess.run(['git', 'diff', '--name-only', 'HEAD~1', 'HEAD'],
                             capture_output=True, text=True, check=True).stdout.splitlines()
        return any(p.startswith('Master_Plan/') and p.lower().endswith('.xlsx') for p in out)
    except Exception:
        return False


def main():
    changed = False

    # ---- 1) شيتات الإنتاج الجديدة
    files = []
    for pat in ('*.csv', '*.xlsx', '*.xls'):
        files += glob.glob(os.path.join(INCOMING, pat))
    files = sorted(f for f in files if not os.path.basename(f).startswith('~$'))
    if files:
        wip2_map, wip1_map = u.load_metadata_maps()
        os.makedirs(PROCESSED, exist_ok=True)
        failed = []
        for f in files:
            print(f"\n=== {os.path.basename(f)} ===")
            try:
                df = u.read_export_rows(f)
                records, dates = u.parse_and_extract_metals(df, wip2_map, wip1_map)
                if not records:
                    raise ValueError("مفيش أوامر شغل لماكينات المعادن في الشيت ده")
                u.merge_and_save_json(records, dates)
                shutil.move(f, os.path.join(PROCESSED, os.path.basename(f)))
                changed = True
            except Exception as e:
                print(f"❌ فشل {os.path.basename(f)}: {e}")
                failed.append(os.path.basename(f))
        if failed:
            print("\n❌ شيتات فشلت:", ", ".join(failed))
            sys.exit(1)
    else:
        print("ℹ️  مفيش شيتات جديدة في incoming/")

    # ---- 2) الماستر بلان
    if plan_changed():
        try:
            import update_wo_voltage as w
            src = w.find_latest_plan()
            if src:
                w.main(src)
                changed = True
        except SystemExit as e:
            print(f"⚠️ الماستر بلان: {e}")

    # ---- 3) كاش + toast
    if changed:
        u.update_html_cache_buster()
        with open('index.html', 'r', encoding='utf-8') as fh:
            h = fh.read()
        h = h.replace('<div id="toastNotification" class="toast-notification">',
                      '<div id="toastNotification" class="toast-notification hidden">')
        with open('index.html', 'w', encoding='utf-8') as fh:
            fh.write(h)
        print("\n✅ تم التحديث")
    else:
        print("\nℹ️  مفيش تغييرات")


if __name__ == '__main__':
    main()
