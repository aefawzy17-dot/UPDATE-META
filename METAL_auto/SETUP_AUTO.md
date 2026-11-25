# التحديث التلقائي على GitHub (مرة واحدة بس تعملها)

## الإعداد (مرة واحدة)
1. ارفع كل ملفات الـ zip على الريبو (main) — خلي بالك تكون رافع مجلد `.github` كمان (فيه ملف الـ workflow).
2. على GitHub: Settings ← Actions ← General ← Workflow permissions ← اختار **Read and write permissions** ← Save.
3. Settings ← Pages ← تأكد إن المصدر Deploy from a branch: main.

## كل يوم (30 ثانية)
1. افتح الريبو على GitHub ← مجلد `incoming` ← Add file ← Upload files.
2. اسحب شيت اليوم `Cable_Production_Records_*.csv` ← Commit changes.
3. استنى دقيقة-دقيقتين: GitHub بيدمج الشيت، بيحدّث البيانات والأوبريتور ورقم النسخة، وينقل الشيت لمجلد `processed`. الموقع بيتحدث لوحده.

متابعة التنفيذ: تبويب **Actions** (علامة ✅ = تمام، ❌ = فيه مشكلة وبيوريك السبب).

## الماستر بلان
ارفع ملف الـ .xlsx الجديد في مجلد `Master_Plan` بنفس الطريقة، و`wo_voltage.js` بيتحدث تلقائي.

## ملاحظات
- شيت لنفس اليوم بيستبدل بيانات اليوم ده (مفيش تكرار).
- مفيش حاجة تتثبت على جهازك ولا تشغّل Python.
- متسحبش الشيتات على الداشبورد ولا تنزل index.html منه تاني.
