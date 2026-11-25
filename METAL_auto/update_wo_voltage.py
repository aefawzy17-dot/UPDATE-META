# -*- coding: utf-8 -*-
"""
update_wo_voltage.py
====================
يربط شيتات Master Production Plan بالويب (Metals Dashboard + Manual Entry):
يقرأ كل أمر شغل (W.O.) وكل بياناته من الماستر بلان:
  - فئة الجهد (LV / MV / HV / BC / CC / ...) والجهد الفعلي (Volt مثل 12/20 KV)
  - المقاس (Size) والمقاس الكامل (cableSize) ومقاس النيوترال (ntSize)
  - المعدن (CU / AL) وعدد الكورات (cores)
  - العميل (Customer) ومواصفة الكابل (cable) والكمية المخططة (plannedKm) والـ TDS والحالة

ويكتب ملفين:
  1) wo_voltage.js     (خريطة الجهد السريعة لكل W.O.)
  2) orders_catalog.js (الكتالوج الكامل لكل أوامر الماستر بلان للبحث والـ Auto-fill)

يعمل محلياً حتى بدون openpyxl (بقارئ ZIP/XML مدمج)، أو مع openpyxl على GitHub Actions.
الاستخدام:
  - التلقائي: حط ملفات الماستر بلان (.xlsx) جوه مجلد Master_Plan وشغّل update_wo_voltage.py
  - اليدوي:   python update_wo_voltage.py "Master Production Plan Nov-2026.xlsx" "Master_Plan\\Job Orders Dec -2026000.xlsx"
"""

import os
import sys
import re
import glob
import json
import datetime
import zipfile
import xml.etree.ElementTree as ET

if sys.platform.startswith('win'):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

try:
    import openpyxl
except ImportError:
    openpyxl = None

HERE = os.path.dirname(os.path.abspath(__file__))
MASTER_DIR = os.path.join(HERE, 'Master_Plan')

# أكواد الفولت/الفاميلي القياسية
CATEGORY_TO_CODE = {
    'low voltage': 'LV',
    'l.v.': 'LV',
    'l.v': 'LV',
    'lv': 'LV',
    'medium voltage': 'MV',
    'm.v.': 'MV',
    'm.v': 'MV',
    'mv': 'MV',
    'high voltage': 'HV',
    'h.v.': 'HV',
    'h.v': 'HV',
    'hv': 'HV',
    'control cables': 'CC',
    'c.c.': 'CC',
    'c.c': 'CC',
    'cc': 'CC',
    'building wires': 'BC',
    'b.w.': 'BC',
    'b.w': 'BC',
    'bc': 'BC',
    'bare copper': 'BARE',
    'air copper': 'BARE',
    'b.cu.': 'BARE',
    'b.cu': 'BARE',
    'bare': 'BARE',
    'overhead': 'OH',
    'o.h.l.': 'OH',
    'o.h.l': 'OH',
    'oh': 'OH',
    'solar': 'SOLAR',
    'solid': 'SOLID',
    'telecom': 'TC',
    'tc': 'TC',
    'instrumentation': 'IC',
    'ic': 'IC',
    'in': 'IN',
    'nd': 'ND',
}

KNOWN_SHORT_CODES = set(CATEGORY_TO_CODE.values())

STD_SIZES = {
    '0.5', '0.75', '1', '1.5', '2.5', '4', '6', '10', '16', '25', '35',
    '50', '70', '75', '78', '95', '120', '150', '185', '240', '300',
    '400', '500', '630', '800', '1000', '1200', '1400', '1600', '2000'
}

WO_RE = re.compile(r'^\d{1,5}/\d{2}/\d{1,3}')


# ---------------------------------------------------------------------------
# قارئ Excel مدمج خفيف (يعمل مع openpyxl أو بدونه عبر zipfile + iterparse)
# ---------------------------------------------------------------------------
def _col_to_idx(ref):
    m = re.match(r'([A-Z]+)', ref or '')
    if not m:
        return -1
    idx = 0
    for ch in m.group(1):
        idx = idx * 26 + (ord(ch) - 64)
    return idx - 1


class _ZipWorkbook:
    def __init__(self, path):
        self._zf = zipfile.ZipFile(path, 'r')
        self._ss = []
        if 'xl/sharedStrings.xml' in self._zf.namelist():
            root = ET.fromstring(self._zf.read('xl/sharedStrings.xml'))
            for si in root:
                self._ss.append(''.join(t.text or '' for t in si.iter() if t.tag.endswith('t')))
        wb_root = ET.fromstring(self._zf.read('xl/workbook.xml'))
        rels_root = ET.fromstring(self._zf.read('xl/_rels/workbook.xml.rels'))
        rmap = {r.attrib['Id']: r.attrib['Target'] for r in rels_root}
        self._sheet_targets = {}
        self.sheetnames = []
        for sh in wb_root.iter():
            if sh.tag.endswith('sheet'):
                name = sh.attrib['name']
                rid = next((v for k, v in sh.attrib.items() if k.endswith('id')), None)
                if not rid or rid not in rmap:
                    continue
                target = rmap[rid].lstrip('/')
                if not target.startswith('xl/'):
                    target = 'xl/' + target
                if target in self._zf.namelist():
                    self.sheetnames.append(name)
                    self._sheet_targets[name] = target

    def iter_rows(self, sheet_name):
        target = self._sheet_targets.get(sheet_name)
        if not target:
            return
        with self._zf.open(target) as sf:
            for _, el in ET.iterparse(sf, events=('end',)):
                if el.tag.endswith('}row') or el.tag == 'row':
                    row_map = {}
                    for c in el:
                        idx = _col_to_idx(c.attrib.get('r', ''))
                        if idx < 0:
                            continue
                        t = c.attrib.get('t')
                        val = None
                        for ch in c:
                            if ch.tag.endswith('}v') or ch.tag == 'v':
                                if t == 's' and ch.text:
                                    si = int(ch.text)
                                    val = self._ss[si] if 0 <= si < len(self._ss) else ''
                                else:
                                    val = ch.text or ''
                            elif ch.tag.endswith('}is') or ch.tag == 'is':
                                val = ''.join(x.text or '' for x in ch.iter() if x.tag.endswith('}t') or x.tag == 't')
                        if val is not None:
                            row_map[idx] = val
                    if row_map:
                        mx = max(row_map.keys())
                        yield tuple(row_map.get(i, None) for i in range(mx + 1))
                    else:
                        yield ()
                    el.clear()

    def close(self):
        self._zf.close()


class _OpenpyxlWorkbook:
    def __init__(self, path):
        self._wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        self.sheetnames = list(self._wb.sheetnames)

    def iter_rows(self, sheet_name):
        ws = self._wb[sheet_name]
        return ws.iter_rows(values_only=True)

    def close(self):
        self._wb.close()


def open_workbook(path):
    if openpyxl is not None:
        return _OpenpyxlWorkbook(path)
    return _ZipWorkbook(path)


# ---------------------------------------------------------------------------
# دوال التنظيف والتوحيد
# ---------------------------------------------------------------------------
def clean_wo(v):
    if v is None:
        return None
    s = re.sub(r'\s+', ' ', str(v)).strip()
    return s if WO_RE.match(s) else None


def clean_str(v):
    if v is None:
        return ''
    s = re.sub(r'\s+', ' ', str(v)).strip()
    return '' if s.lower() in ('nan', 'none', 'null', '#n/a', '#ref!') else s


def clean_num_str(v):
    s = clean_str(v)
    if not s:
        return ''
    try:
        f = float(s)
        if f.is_integer():
            return str(int(f))
        return str(round(f, 3)).rstrip('0').rstrip('.')
    except Exception:
        return s


def clean_voltage_cat(raw):
    s = clean_str(raw)
    if not s:
        return ''
    low = s.lower()
    if low in CATEGORY_TO_CODE:
        return CATEGORY_TO_CODE[low]
    up = s.upper()
    if up in KNOWN_SHORT_CODES:
        return up
    return up if len(up) <= 8 else ''


def clean_metal(raw, cable=''):
    u = clean_str(raw).upper()
    if u in ('C', 'CU') or u.startswith('CU') or 'COPPER' in u:
        return 'CU'
    if u in ('A', 'AL') or u.startswith('AL') or 'ALUMIN' in u or 'AAAC' in u:
        return 'AL'
    c = clean_str(cable).upper()
    if c:
        if re.search(r'(?:^|[\s,/])CU(?:$|[\s,/])', c) or 'COPPER' in c or 'TCU' in c:
            return 'CU'
        if re.search(r'(?:^|[\s,/])AL(?:$|[\s,/])', c) or 'ALUMIN' in c or 'ACSR' in c or 'AAAC' in c:
            return 'AL'
    return ''


def _extract_size_expr_from_cable(cable):
    c = clean_str(cable)
    if not c:
        return ''
    m = re.search(r'\b(\d+(?:\.\d+)?\s*[*xX]\s*(?:\d+(?:\.\d+)?\s*[*xX]\s*)?\d+(?:\.\d+)?(?:\s*\+\s*\d+(?:\.\d+)?)?(?:/\d+(?:\.\d+)?)?)\b', c)
    if m:
        return re.sub(r'\s+', '', m.group(1)).replace('*', 'x').replace('X', 'x')
    m2 = re.search(r',(\d+(?:\.\d+)?)(?:,|$)', c)
    return m2.group(1) if m2 else ''


def clean_size(raw_size, phase_csa='', cable=''):
    pcsa = clean_num_str(phase_csa)
    if pcsa and pcsa not in ('0', '-') and re.match(r'^\d+(?:\.\d+)?$', pcsa):
        return pcsa
    s = clean_str(raw_size)
    if not s or s in ('-', '0'):
        s = _extract_size_expr_from_cable(cable)
    if not s or s in ('-', '0'):
        return ''
    s = re.sub(r'\s*MM2?$', '', s, flags=re.IGNORECASE).replace('²', '').strip()
    if '+' in s:
        s = s.split('+')[0].strip()
    if '/' in s:
        s = s.split('/')[0].strip()
    s = s.replace('*', 'x').replace('X', 'x')
    parts = [p.strip() for p in s.split('x') if p.strip()]
    if parts:
        last = clean_num_str(parts[-1])
        if last in STD_SIZES:
            return last
        for p in parts:
            cp = clean_num_str(p)
            if cp in STD_SIZES:
                return cp
        return last
    return s


def clean_cable_size(raw_size, phase_csa='', nt_csa='', cores='', cable=''):
    s = clean_str(raw_size)
    if s and s not in ('-', '0'):
        s = re.sub(r'\s*MM2?$', '', s, flags=re.IGNORECASE).replace('²', '')
        s = re.sub(r'\s+', '', s).replace('*', 'x').replace('X', 'x')
        return s
    pcsa = clean_num_str(phase_csa)
    nc = extract_cores(cores, '', cable)
    nt = clean_num_str(nt_csa)
    if pcsa and pcsa != '0':
        base = f"{nc}x{pcsa}" if nc else pcsa
        if nt and nt != '0':
            base += f"+{nt}"
        return base
    return _extract_size_expr_from_cable(cable)


def extract_cores(raw_cores, raw_size='', cable=''):
    rc = clean_str(raw_cores)
    if rc and rc != '0':
        m = re.match(r'^(\d+)', rc)
        if m and m.group(1) != '0':
            return m.group(1)
    for txt in (clean_str(raw_size), clean_str(cable)):
        if not txt:
            continue
        m = re.search(r'\b(\d+)\s*[*xX]\s*\d+(?:\.\d+)?', txt)
        if m and m.group(1) != '0':
            return m.group(1)
    return ''


def parse_planned_km(raw_qty, unit=''):
    s = clean_str(raw_qty)
    if not s:
        return None
    try:
        val = float(s)
    except Exception:
        return None
    if val <= 0:
        return None
    if clean_str(unit).upper() == 'M':
        val = val / 1000.0
    val = round(val, 3)
    return int(val) if val.is_integer() else val


def _row_val(row, idx):
    if idx is None or idx < 0 or idx >= len(row):
        return None
    return row[idx]


def header_indices(first_row, wanted):
    """يرجّع {المفتاح: أول رقم عمود مطابق} مع تطبيع المسافات والأسطر الجديدة."""
    names = [re.sub(r'\s+', ' ', str(c)).strip().lower() if c is not None else '' for c in first_row]
    out = {}
    for key, cands in wanted.items():
        for c in cands:
            if c in names:
                out[key] = names.index(c)
                break
    return out


def merge_order_entry(dst, src):
    for k, v in src.items():
        if v is None or v == '' or v == '-':
            continue
        if k == 'plannedKm':
            if v > 0:
                dst[k] = v
        else:
            dst[k] = v


# ---------------------------------------------------------------------------
# استخراج الأوردرات والبيانات الكاملة من شيتات الماستر بلان
# ---------------------------------------------------------------------------
def _read_sheet_records(wb, sheet_name, col_spec, unit_is_m_check=False, allow_nso_wo=False):
    if sheet_name not in wb.sheetnames:
        return {}
    rows_iter = wb.iter_rows(sheet_name)
    first = next(rows_iter, None)
    if not first:
        return {}

    idx = header_indices(first, col_spec)
    wo_cols = []
    for k in ('wo', 'wo2', 'wo3'):
        if k in idx and idx[k] not in wo_cols:
            wo_cols.append(idx[k])
    if not wo_cols:
        return {}

    has_meta = any(k in idx for k in ('cat', 'metal', 'phase_csa', 'size', 'cable', 'desc2', 'volt', 'customer'))
    if not has_meta:
        return {}

    out = {}
    for r in rows_iter:
        if not r:
            continue
        wo = None
        for wc in wo_cols:
            wo = clean_wo(_row_val(r, wc))
            if wo:
                break
        if not wo:
            continue

        cable = clean_str(_row_val(r, idx.get('cable'))) or clean_str(_row_val(r, idx.get('desc2')))
        raw_size = clean_str(_row_val(r, idx.get('size')))
        phase_csa = _row_val(r, idx.get('phase_csa'))
        nt_csa = clean_num_str(_row_val(r, idx.get('nt_csa')))
        raw_cores = _row_val(r, idx.get('cores'))
        raw_metal = _row_val(r, idx.get('metal'))
        raw_cat = _row_val(r, idx.get('cat'))
        volt = clean_str(_row_val(r, idx.get('volt')))
        customer = clean_str(_row_val(r, idx.get('customer')))
        tds = clean_str(_row_val(r, idx.get('tds')))
        status = clean_str(_row_val(r, idx.get('status')))
        unit = _row_val(r, idx.get('unit')) if unit_is_m_check else ''
        planned_km = parse_planned_km(_row_val(r, idx.get('qty')), unit)
        if planned_km is None and 'qty2' in idx:
            planned_km = parse_planned_km(_row_val(r, idx.get('qty2')), unit)

        voltage = clean_voltage_cat(raw_cat)
        metal = clean_metal(raw_metal, cable)
        size = clean_size(raw_size, phase_csa, cable)
        cable_size = clean_cable_size(raw_size, phase_csa, nt_csa, raw_cores, cable)
        cores = extract_cores(raw_cores, raw_size, cable)

        entry = {}
        if voltage:
            entry['voltage'] = voltage
        if volt:
            entry['volt'] = volt
        if size:
            entry['size'] = size
        if cable_size and cable_size != size:
            entry['cableSize'] = cable_size
        if nt_csa and nt_csa != '0':
            entry['ntSize'] = nt_csa
        if metal:
            entry['metal'] = metal
        if cores:
            entry['cores'] = cores
        if customer:
            entry['customer'] = customer
        if cable:
            entry['cable'] = cable
        if planned_km is not None:
            entry['plannedKm'] = planned_km
        if tds:
            entry['tds'] = tds
        if status:
            entry['status'] = status

        if entry:
            dst = out.setdefault(wo, {})
            merge_order_entry(dst, entry)

    return out


def extract_workbook_catalog(src):
    wb = open_workbook(src)
    orders_data = {}
    counts = {'adv': 0, 'backlog': 0, 'orders': 0, 'other': 0}

    try:
        # 1) 90 for ADV. (أقل أولوية داخل نفس الملف)
        adv = _read_sheet_records(wb, '90 for ADV.', {
            'wo': ['work_order', 'w.o.', 'w.o', 'wo'],
            'cat': ['family', 'cable category', 'cable cat.'],
            'metal': ['conductor', 'metal'],
            'phase_csa': ['ph_size', 'phase csa'],
            'size': ['size', 'cable size'],
            'cores': ['ph_n_cores', 'no. of cores'],
            'volt': ['volt'],
            'customer': ['customer'],
            'cable': ['cable', 'desc'],
            'desc2': ['dscrpt', 'description'],
            'qty': ['orderqty', 'planned qty(km)'],
            'unit': ['unit_sales'],
            'tds': ['tds'],
        }, unit_is_m_check=True)
        for wo, ent in adv.items():
            merge_order_entry(orders_data.setdefault(wo, {}), ent)
        counts['adv'] = len(adv)

        # 2) Backlog 2 (col 0: WO, col 28: work_order, col 1: NSO بصيغة W.O.)
        backlog = _read_sheet_records(wb, 'Backlog 2', {
            'wo': ['wo'],
            'wo2': ['work_order'],
            'wo3': ['nso'],
            'cat': ['family', 'cable category', 'cable cat.'],
            'metal': ['conductor', 'metal'],
            'phase_csa': ['ph_size', 'phase csa'],
            'size': ['size', 'cable size'],
            'nt_csa': ['ne_size', 'nt. csa'],
            'cores': ['ph_n_cores', 'no. of cores'],
            'volt': ['volt'],
            'customer': ['customer'],
            'cable': ['cable', 'desc'],
            'desc2': ['dscrpt', 'description'],
            'qty': ['km', 'planned qty(km)'],
            'qty2': ['orderqty'],
            'unit': ['unit_sales'],
            'tds': ['tds'],
            'status': ['ordstatus', 'status'],
        }, unit_is_m_check=True, allow_nso_wo=True)
        for wo, ent in backlog.items():
            merge_order_entry(orders_data.setdefault(wo, {}), ent)
        counts['backlog'] = len(backlog)

        # 3) Orders (الأدق لخطة الشهر)
        orders = _read_sheet_records(wb, 'Orders', {
            'wo': ['w.o.', 'w.o', 'wo'],
            'cat': ['cable category', 'cable cat.', 'family'],
            'metal': ['metal', 'conductor'],
            'phase_csa': ['phase csa'],
            'nt_csa': ['nt. csa'],
            'size': ['cable size', 'size'],
            'cores': ['no. of cores'],
            'volt': ['volt'],
            'customer': ['customer'],
            'cable': ['desc', 'cable'],
            'desc2': ['description'],
            'qty': ['planned qty(km)'],
            'tds': ['tds'],
            'status': ['release to prod. status', 'status'],
        })
        for wo, ent in orders.items():
            merge_order_entry(orders_data.setdefault(wo, {}), ent)
        counts['orders'] = len(orders)

        # 4) أي شيتات إضافية (مثل شيتات التواريخ في Job Orders Dec)
        generic_spec = {
            'wo': ['w.o.', 'w.o', 'wo', 'work_order'],
            'cat': ['cable category', 'cable cat.', 'family'],
            'metal': ['metal', 'conductor'],
            'phase_csa': ['phase csa', 'ph_size'],
            'nt_csa': ['nt. csa', 'ne_size'],
            'size': ['size', 'cable size'],
            'cores': ['no. of cores', 'ph_n_cores'],
            'volt': ['volt'],
            'customer': ['customer'],
            'cable': ['desc', 'cable'],
            'desc2': ['description', 'dscrpt'],
            'qty': ['planned qty(km)', 'km', 'orderqty'],
            'tds': ['tds'],
            'status': ['release to prod. status', 'status', 'ordstatus'],
        }
        other_wos = set()
        for s in wb.sheetnames:
            if s in ('Orders', 'Backlog 2', '90 for ADV.'):
                continue
            extra = _read_sheet_records(wb, s, generic_spec)
            for wo, ent in extra.items():
                merge_order_entry(orders_data.setdefault(wo, {}), ent)
                other_wos.add(wo)
        counts['other'] = len(other_wos)
    finally:
        wb.close()

    return orders_data, counts


# ---------------------------------------------------------------------------
# إدارة الملفات والتحديث التراكمي
# ---------------------------------------------------------------------------
def find_all_plans():
    """كل ملفات .xlsx في مجلد Master_Plan مرتبة من الأقدم للأحدث (الأحدث يكسب)."""
    cands = [f for f in glob.glob(os.path.join(MASTER_DIR, '*.xlsx')) if not os.path.basename(f).startswith('~$')]
    if not cands:
        cands = [f for f in glob.glob(os.path.join(HERE, 'Master*Plan*.xlsx')) if not os.path.basename(f).startswith('~$')]
    return sorted(cands, key=os.path.getmtime)


def find_latest_plan():
    """أحدث .xlsx في مجلد Master_Plan (للتوافق مع ci_update.py)."""
    plans = find_all_plans()
    return plans[-1] if plans else None


def auto_update():
    """بيتندهي من update_daily_data.py: لو في شيت ماستر بلان أحدث من wo_voltage.js أو orders_catalog.js بيحدّثهم."""
    plans = find_all_plans()
    if not plans:
        print("ℹ️  مفيش شيت ماستر بلان في مجلد Master_Plan - هيتم استخدام البيانات الحالية")
        return False
    out_v = os.path.join(HERE, 'wo_voltage.js')
    out_c = os.path.join(HERE, 'orders_catalog.js')
    newest_plan_mtime = max(os.path.getmtime(p) for p in plans)
    if (os.path.exists(out_v) and os.path.exists(out_c)
            and os.path.getmtime(out_v) >= newest_plan_mtime
            and os.path.getmtime(out_c) >= newest_plan_mtime):
        print(f"ℹ️  wo_voltage.js و orders_catalog.js محدّثين بالفعل من {os.path.basename(plans[-1])}")
        return False
    main(plans)
    return True


def _load_js_object(path, var_name):
    if not os.path.exists(path):
        return {}
    try:
        with open(path, 'r', encoding='utf-8') as f:
            txt = f.read()
        m = re.search(rf'var {re.escape(var_name)}\s*=\s*(\{{[\s\S]*?\}});', txt)
        if m:
            obj = json.loads(m.group(1))
            return obj if isinstance(obj, dict) else {}
    except Exception:
        pass
    return {}


def main(src=None):
    if src is None:
        cli_args = [a for a in sys.argv[1:] if not a.startswith('-')]
        sources = cli_args if cli_args else find_all_plans()
    elif isinstance(src, (list, tuple)):
        sources = list(src)
    else:
        sources = [src]

    sources = [s for s in sources if s and os.path.exists(s)]
    if not sources:
        sys.exit("حط شيت الماستر بلان (.xlsx) جوه مجلد Master_Plan أو مرّر مساره كـ argument")

    out_volt = os.path.join(HERE, 'wo_voltage.js')
    out_cat = os.path.join(HERE, 'orders_catalog.js')

    existing_volt = _load_js_object(out_volt, 'WO_VOLTAGE_MAP')
    catalog = _load_js_object(out_cat, 'ORDERS_CATALOG')

    # حفظ أي كود فولت قديم داخل الكتالوج لو الأمر مش موجود فيه وتوحيد الأكواد
    for wo, ent in list(catalog.items()):
        if isinstance(ent, dict) and ent.get('voltage'):
            ent['voltage'] = clean_voltage_cat(ent['voltage']) or ent['voltage']
    for wo, vcode in list(existing_volt.items()):
        norm_v = clean_voltage_cat(vcode) or vcode
        existing_volt[wo] = norm_v
        if norm_v and wo not in catalog:
            catalog[wo] = {'voltage': norm_v}

    total_new_from_files = 0
    last_counts = {'orders': 0, 'backlog': 0, 'adv': 0, 'other': 0}
    source_names = []

    for s in sources:
        print(f"📖 قراءة: {s}")
        data, counts = extract_workbook_catalog(s)
        total_new_from_files += len(data)
        last_counts = counts
        source_names.append(os.path.basename(s))
        for wo, ent in data.items():
            merge_order_entry(catalog.setdefault(wo, {}), ent)
        print(f"   ↳ Orders: {counts['orders']} | Backlog 2: {counts['backlog']} | 90 for ADV.: {counts['adv']} | Other: {counts['other']} (إجمالي الشيت: {len(data)})")

    if total_new_from_files == 0:
        print(f"⚠️  لم يتم العثور على أي أوردرات في الملفات المحددة - تم الإبقاء على البيانات الحالية ({len(catalog)} أمر)")
        return

    ordered_catalog = dict(sorted(catalog.items()))
    volt_map = dict(existing_volt)
    for wo, ent in ordered_catalog.items():
        if ent.get('voltage'):
            volt_map[wo] = ent['voltage']
    ordered_volt = dict(sorted(volt_map.items()))

    now_str = datetime.datetime.now().strftime('%Y-%m-%d %H:%M')
    src_label = ' + '.join(source_names)

    volt_meta = {
        'source': src_label,
        'generated': now_str,
        'from_orders': last_counts['orders'],
        'from_backlog': last_counts['backlog'],
        'from_adv': last_counts['adv'],
        'from_other': last_counts['other'],
        'total': len(ordered_volt),
    }
    with open(out_volt, 'w', encoding='utf-8') as f:
        f.write('// Auto-generated by update_wo_voltage.py - لا تعدّل يدوياً\n')
        f.write('var WO_VOLTAGE_META = ' + json.dumps(volt_meta, ensure_ascii=False) + ';\n')
        f.write('var WO_VOLTAGE_MAP = ' + json.dumps(ordered_volt, ensure_ascii=False, separators=(',', ':')) + ';\n')

    with_size = sum(1 for v in ordered_catalog.values() if v.get('size'))
    with_metal = sum(1 for v in ordered_catalog.values() if v.get('metal'))
    cat_meta = {
        'source': src_label,
        'generated': now_str,
        'total': len(ordered_catalog),
        'withSize': with_size,
        'withMetal': with_metal,
    }
    with open(out_cat, 'w', encoding='utf-8') as f:
        f.write(f'// Auto-generated by update_wo_voltage.py ({len(ordered_catalog)} orders) - لا تعدّل يدوياً\n')
        f.write('var ORDERS_CATALOG_META = ' + json.dumps(cat_meta, ensure_ascii=False) + ';\n')
        f.write('var ORDERS_CATALOG = ' + json.dumps(ordered_catalog, ensure_ascii=False, separators=(',', ':')) + ';\n')
        f.write('if (typeof window !== "undefined") { window.ORDERS_CATALOG = ORDERS_CATALOG; window.ORDERS_CATALOG_META = ORDERS_CATALOG_META; }\n')
        f.write('if (typeof WO_VOLTAGE_MAP !== "undefined" && WO_VOLTAGE_MAP) {\n')
        f.write('  Object.keys(ORDERS_CATALOG).forEach(function (wo) {\n')
        f.write('    if (ORDERS_CATALOG[wo] && ORDERS_CATALOG[wo].voltage) { WO_VOLTAGE_MAP[wo] = ORDERS_CATALOG[wo].voltage; }\n')
        f.write('  });\n')
        f.write('}\n')

    counts_v = {}
    for v in ordered_volt.values():
        counts_v[v] = counts_v.get(v, 0) + 1
    print(f"✅ تم كتابة {out_volt} ({len(ordered_volt)} أمر)")
    print(f"✅ تم كتابة {out_cat} ({len(ordered_catalog)} أمر | مقاس: {with_size} | معدن: {with_metal})")
    print("   " + " | ".join(f"{k}: {v}" for k, v in sorted(counts_v.items(), key=lambda x: -x[1])))


if __name__ == '__main__':
    main()
