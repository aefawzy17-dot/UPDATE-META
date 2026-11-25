# Metals Department Dashboard - Operating Instructions & Data Format

## 🎯 Primary Principle: Single Source of Truth
All dashboard data and updates rely **exclusively on the daily exported production CSV sheet** provided by the department (`Cable_Production_Records_*.csv`). No external files (`wip001_pp.xls`, `wip002_pr.xls`, or external databases) are required.

---

## 📋 Standard Daily Sheet Format (`Cable_Production_Records`)
Any sheet dragged & dropped into the dashboard or passed to the update script must follow this exact structure:

| Column Header | Description | Example Values |
| :--- | :--- | :--- |
| **Record #** | Unique record sequence ID | `808`, `807` |
| **Order No** | Work order number (`W.O`) | `1990/26/02`, `1264/26/04` |
| **Operation** | Process operation | `Stranding`, `Drawing` |
| **Machine** | Target machine (10 core machines) | `RS01`, `RS02`, `RS03`, `RS04`, `RS05`, `RS06`, `CS01`, `CS02`, `DS01`, `CN` (or `CN01`) |
| **Send Card** | Source card identifier | `25050`, empty |
| **Receive Card**| Take-up / receive card number | `3808`, `15372` |
| **Take-Up Drum**| Drum / package serial | `Z-149`, `T-23`, `SR-102` |
| **Conductor Metal** | Primary conductor metal | `AL` (Aluminum), `CU` (Copper) |
| **Size (mm²)** | Wire / conductor cross-section | `4x1x240`, `3x240`, `1x400`, `35*2`, `120` |
| **Conductor Shape** | Conductor geometry | `Circular`, `Compacted`, `Sector` |
| **Stranding Angle** | Geometric angle / shape | `دائري`, `مضغوط` |
| **Cores Count** | Number of cores | `1`, `3`, `4` |
| **Length (m)** | Measured production length in meters | `6300`, `1100`, `2500` |
| **Weight (Ton)**| Total weight in metric tons | `16.366`, `1.152`, `0.581` |
| **CU** | Copper tonnage | `0`, `7.14` |
| **AL** | Aluminum tonnage | `16.366`, `1.152` |
| **Operator Code / Name** | Machine operator name | `حازم حامد وهبه`, `احمد عبده السيد` |
| **Shift Code** | Production shift | `S1`, `S2`, `S3` |
| **Logged By** | Supervisor / engineer | `Ahmed Samy` |
| **Registration Date** | Production record date | `11/1/2026`, `2026-11-01` |

---

## 🚀 How to Update the Dashboard

### Method 1: Instant Drag & Drop (In Browser)
1. Open [`index.html`](file:///C:/Users/mahmoud.mohamed/Desktop/METAL/index.html) in your browser (Google Chrome or Microsoft Edge).
2. Drag and drop any `Cable_Production_Records_*.csv` file directly onto the dashboard window.
3. The dashboard will instantly:
   * Parse the 10 metals machines (`RS01` to `RS06`, `CS01`, `CS02`, `DS01`, `CN`).
   * Clean and normalize wire sizes (`cleanSize`).
   * Add the new day's production to the dataset without overwriting historical records.
   * Update the Machine Productivity table, KPI cards, and charts in real-time.
   * Save the state locally to browser `localStorage`.

### Method 2: Python Script (Local or GitHub Deploy)
Place the new daily file in the project folder and run:
```bash
python update_daily_data.py Cable_Production_Records_2026-11-02.csv
```
Or simply double-click:
```
تحديث_البيانات.bat
```

---

## 📊 Machine Productivity Table Structure
The main machines table displays columns in this exact sequence:
1. **Machine**: Machine identifier (`RS01` to `RS06`, `CS01`, `CS02`, `DS01`, `CN`).
2. **W.O**: Work order numbers executed on this machine during the period (with km and % of the machine's output).
3. **Voltage**: `LV` / `MV` (or `HV`, `CC`, ...) per work order, taken from the **Master Production Plan** (see below). A faded, dashed badge means the W.O was not found in the plan and the old production-data family is shown instead.
4. **Cable Size**: Conductor cross-section(s) produced.
5. **Conductor**: Material badge (`AL` or `CU`).
6. **Target/day**: Standard daily target in km (`Target/day`).
7. **Shift 1**: Shift 1 production in km.
8. **Shift 2**: Shift 2 production in km.
9. **Shift 3**: Shift 3 production in km.
10. **Production km/day**: Total daily production output in **km/day**.
11. **Cable Type**: Cable category / Family ID (`LV`, `MV`, `HV`).

---

## 🔗 Linking the Master Production Plan (LV / MV per order)

The Voltage column is read from `wo_voltage.js`, generated from the plan workbook
(`Orders` → *Cable Category*, then `Backlog 2` / `90 for ADV.` → *family* for older orders).

**Option A – Master_Plan folder (recommended)**: put the latest Master Production Plan `.xlsx` inside the `Master_Plan/` folder
(the newest file by modified date is used), then either double-click `تحديث_الماستر_بلان.bat`
or just run the normal daily update (`تحديث_البيانات.bat`) – it refreshes `wo_voltage.js` automatically when the plan is newer.
Manual: `python update_wo_voltage.py` (needs `pip install openpyxl`). Push `wo_voltage.js` with the other files.

**Option B – Drag & Drop**: drop the Master Production Plan `.xlsx` on the dashboard window.
The map is saved in that browser's `localStorage` (affects only that browser until Option A is run).

