# بات تلگرام تبدیل عکس فاکتور به اکسل با Gemini

این پروژه برای استقرار مستقیم از GitHub روی Railway آماده شده است.

## عملکرد

1. کاربر مجاز عکس فاکتور یا رسید را برای بات تلگرام می‌فرستد.
2. تصویر برای Gemini API ارسال می‌شود.
3. اطلاعات اصلی و ردیف‌های فاکتور به‌صورت ساختاریافته استخراج می‌شوند.
4. پیش‌نمایش برای تأیید نمایش داده می‌شود.
5. پس از تأیید، اطلاعات در فایل اکسل تجمعی ذخیره می‌شوند.
6. فایل اکسل به تلگرام ارسال می‌شود.

## فایل‌های اصلی

```text
bot.py
requirements.txt
railway.json
.python-version
.env.example
.gitignore
data/.gitkeep
```

## متغیرهای لازم در Railway

```text
TELEGRAM_BOT_TOKEN=توکن بات تلگرام
GEMINI_API_KEY=کلید Gemini API
GEMINI_MODEL=gemini-2.5-flash
ALLOWED_USER_ID=شناسه عددی تلگرام
DATA_DIR=/data
LOG_DIR=/data/logs
```

در اجرای اول می‌توان `ALLOWED_USER_ID` را تعریف نکرد. سپس دستور `/myid`
را در بات ارسال کرد و عدد برگشتی را به Variables در Railway افزود.

## ذخیره دائمی

برای حفظ فایل اکسل بعد از استقرار مجدد، یک Volume در Railway به سرویس متصل کن.

Mount Path:

```text
/data
```

فایل نهایی در این مسیر ذخیره می‌شود:

```text
/data/invoices.xlsx
```

## فرمان اجرا

فایل `railway.json` فرمان زیر را اجرا می‌کند:

```text
python bot.py
```

## دستورهای بات

```text
/start
/myid
/export
/cancel
```

## امنیت

- کلید Gemini و توکن تلگرام را داخل GitHub قرار نده.
- آن‌ها را فقط در بخش Variables ریلوی وارد کن.
- فایل `.env` را در مخزن عمومی بارگذاری نکن.
- پس از پیدا کردن شناسه تلگرام، `ALLOWED_USER_ID` را حتماً تنظیم کن.
