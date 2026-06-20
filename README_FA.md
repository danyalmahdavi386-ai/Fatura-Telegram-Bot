# Telegram Diagnostic Bot

این پروژه فقط برای تشخیص مشکل اتصال Telegram و Railway است.

## متغیرهای لازم

فقط این متغیر ضروری است:

```text
TELEGRAM_BOT_TOKEN=توکن جدید بات
```

این متغیر اختیاری است:

```text
ALLOWED_USER_ID=شناسه عددی تلگرام
```

برای آزمایش اول بهتر است `ALLOWED_USER_ID` را حذف کنی.

## استقرار

فایل‌ها را مستقیماً در ریشه یک مخزن GitHub قرار بده:

```text
bot.py
requirements.txt
railway.json
.python-version
.gitignore
```

فرمان اجرا:

```text
python -u bot.py
```

## پیام صحیح در Railway Logs

```text
DIAGNOSTIC_BOT_STARTED
```

## آزمایش در تلگرام

```text
/start
/ping
/version
/myid
```

پاسخ صحیح:

```text
🏓 PONG
نسخه: telegram-diagnostic-v1
```

## نتیجه‌گیری

- اگر `DIAGNOSTIC_BOT_STARTED` در Logs نیست، برنامه در Railway شروع نشده است.
- اگر خطای `InvalidToken` یا `Unauthorized` دیده شود، توکن اشتباه یا لغوشده است.
- اگر خطای `Conflict` دیده شود، نسخه دیگری از همان بات با همان توکن فعال است.
- اگر پیام شروع در Logs وجود دارد و `/ping` پاسخ می‌دهد، اتصال Railway و Telegram سالم است و مشکل فقط در پروژه اصلی است.
