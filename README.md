# SOP Portable v12

دستورالعمل قابل‌حمل تولید — نسخه خط فرمان (بدون رابط کاربری وب).

> شرکت **پیشگامان پاکران الکترونیک**
> سند: WI-SOP-DRV-001 · بازنگری ۰۲ · تاریخ: ۱۴۰۳/۰۷/۱۲

---

## معرفی

این ابزار یک فایل PDF حرفه‌ای از دستورالعمل استاندارد تولید تولید می‌کند.
کاربر فقط فایل `run.bat` را دابل‌کلیک می‌کند و PDF نهایی در پوشه `output/` ساخته می‌شود.

بدون سرور، بدون رابط گرافیکی، بدون نصب — همه چیز در یک پوشه قابل‌حمل اجرا می‌شود.

---

## ویژگی‌ها

- **خط فرمان خالص** — فقط `run.bat` را اجرا کنید.
- **۴ فاز تولید**: پیش‌مونتاژ، مونتاژ اصلی، آزمون و QC، بسته‌بندی.
- **۳۸ مرحله** استخراج‌شده از سند مبدأ.
- **تصاویر راهنما** — تصویر هر مرحله با نام `step-NN-PP.png` در پوشه `images/` قرار می‌گیرد.
- **پردازش خودکار تصاویر**: EXIF transpose، center-crop به نسبت ۴:۳، فشرده‌سازی JPEG q92.
- **قالب HTML/CSS حرفه‌ای** با فونت **Vazirmatn** (نسخه‌های Light تا Black).
- **کاور رنگی** در ابتدا و صفحه پایانی در انتها.
- **هدر فاز رنگی** برای هر مرحله (A=navy، B=teal، C=purple، D=amber).
- **کد سند هر مرحله**: `WI-{family}-{number}` (مثلاً `WI-ASM-001`).
- **شماره‌گذاری فارسی** مراحل (۱، ۲، ۳، ...).
- **Playwright + Chromium قابل‌حمل** برای رندر PDF.
- **Fallback تکه‌ای** با pypdf اگر V8 به محدودیت رشته برخورد کند.
- **Runtime خودکفا** در `_runtime/` — نیازی به Python نصب‌شده روی سیستم نیست.

---

## ساختار پوشه‌ها

```
v12-repo/
├── app/
│   ├── generate_sop.py      # اسکریپت تولید PDF (خط فرمان)
│   ├── template.html        # قالب Jinja2 (کاور + مراحل + پایانی)
│   ├── styles.css           # استایل CSS
│   ├── logo.png             # لوگوی شرکت
│   └── fonts/               # فونت Vazirmatn (۹ وزن)
├── edit/
│   ├── steps.xlsx           # داده‌های مراحل + DocumentInfo + Phases + Revisions
│   └── content.xlsx         # محتوای پشتیبان (مخفی)
├── images/                  # تصاویر مرحله‌ها (step-NN-PP.png)
├── output/                  # خروجی نهایی (sop.pdf, sop.html, assets/)
├── scripts/
│   └── build_excels.py      # ساخت فایل‌های Excel اولیه از HTML مبدأ
├── _runtime/                # Python + Chromium (setup-runtime.bat دانلود می‌کند)
│   ├── python/              # Python 3.12 embeddable
│   └── chromium/            # Chromium برای Playwright
├── run.bat                  # نقطه ورود کاربر
├── setup-runtime.bat        # نصب Python + ویل‌ها + Chromium
├── .gitignore
└── README.md                # همین فایل
```

---

## راه‌اندازی اولیه

### ۱. نصب Runtime (یکبار)

فایل `setup-runtime.bat` را اجرا کنید. این کار:

- Python 3.12 embeddable را در `_runtime/python/` دانلود و استخراج می‌کند.
- pip را bootstrap می‌کند.
- ویل‌های `openpyxl، Pillow، Jinja2، playwright، pypdf` را در `_runtime/python/libs/` نصب می‌کند.
- Chromium را در `_runtime/chromium/` دانلود می‌کند.

> لازم به اتصال اینترنت است. حدود ۲۰۰ مگابایت دانلود دارد.

### ۲. ساخت فایل‌های Excel اولیه (یکبار)

```bash
_runtime\python\python.exe scripts\build_excels.py
```

این کار `edit/steps.xlsx` و `edit/content.xlsx` را از سند مبدأ می‌سازد.

### ۳. افزودن تصاویر (اختیاری)

تصاویر هر مرحله را در پوشه `images/` با نام `step-NN-PP.png` قرار دهید:

- `NN` = شماره مرحله (دو رقمی، مثلاً `01`)
- `PP` = موقعیت تصویر در گالری (مثلاً `01`, `02`)

مثال: `step-01-01.png`, `step-01-02.png`, `step-15-01.png`

فرمت‌های پشتیبانی‌شده: PNG, JPG, JPEG, WEBP.

---

## استفاده روزانه

فقط فایل `run.bat` را دابل‌کلیک کنید. خروجی در `output/sop.pdf` ساخته می‌شود و File Explorer روی آن باز می‌شود.

---

## ویرایش محتوا

فایل `edit/steps.xlsx` را با Excel باز کنید. چهار شیت دارد:

- **DocumentInfo** — عنوان، کد سند، تاریخ، شرکت و ...
- **Revisions** — تاریخچه بازنگری‌ها.
- **Phases** — تعریف چهار فاز.
- **Steps** — ۳۸ ردیف با ستون‌های:
  - `step_number, phase, station, title`
  - `action_1 .. action_8` (هر اقدام در یک ستون)
  - `tools, materials, acceptance_criteria`
  - `process_family, doc_number, section_responsible`
  - `step_version, step_date`

پس از ویرایش، دوباره `run.bat` را اجرا کنید تا PDF جدید ساخته شود.

---

## تنظیمات PDF

- فرمت: **A4**
- حاشیه: **۱۰ میلی‌متر** چهار طرف
- پس‌زمینه چاپ می‌شود (`print_background=True`).
- صبر ۲ ثانیه برای بارگذاری فونت‌ها.
- در صورت شکست Playwright: fallback تکه‌ای با pypdf.

---

## عیب‌یابی

| مشکل | راه‌حل |
|------|--------|
| `python.exe not found` | `setup-runtime.bat` را اجرا کنید. |
| `chromium not found` | `setup-runtime.bat` را دوباره اجرا کنید تا Chromium دانلود شود. |
| PDF خالی است | مطمئن شوید `edit/steps.xlsx` شیت Steps با حداقل یک ردیف دارد. |
| تصاویر ظاهر نمی‌شوند | نام فایل باید `step-NN-PP.png` باشد. |
| خطای V8 (long SOP) | اسکریپت به‌صورت خودکار از fallback تکه‌ای استفاده می‌کند. |

---

## مجوز

© ۱۴۰۳ پیشگامان پاکران الکترونیک — استفاده داخلی.
