import os
import logging
import requests
from flask import Flask, request, abort, jsonify

logging.basicConfig(level=logging.INFO)
log = logging.getLogger("harbi-bot")

BOT_TOKEN = os.environ.get("BOT_TOKEN", "").strip()
WEBHOOK_SECRET = os.environ.get("WEBHOOK_SECRET", "").strip()
APP_URL = os.environ.get("RENDER_EXTERNAL_URL", "").strip().rstrip("/")
API = f"https://api.telegram.org/bot{BOT_TOKEN}"

app = Flask(__name__)

ABOUT = (
    "🌟 کافی‌نت و لوازم‌التحریر حربی\n\n"
    "💻 خدمات کافی‌نت: تایپ، پرینت، کپی، اسکن و خدمات اینترنتی.\n"
    "✏️ لوازم‌التحریر: انواع دفتر، خودکار، مداد و نوشت‌افزار دانش‌آموزی و اداری.\n\n"
    "❤️ رضایت شما، افتخار ماست."
)
SERVICES = (
    "💻 خدمات کافی‌نت حربی\n\n"
    "• تایپ و آماده‌سازی فایل\n"
    "• پرینت و چاپ مدارک\n"
    "• کپی و اسکن\n"
    "• راهنمایی برای خدمات اینترنتی و ثبت‌نام‌های آنلاین\n\n"
    "برای هماهنگی درباره خدمت موردنظر، پیام خود را همین‌جا بنویسید."
)
STATIONERY = (
    "✏️ لوازم‌التحریر حربی\n\n"
    "• دفتر و کاغذ\n• خودکار و مداد\n• پاک‌کن و تراش\n"
    "• نوشت‌افزار دانش‌آموزی و اداری\n\n"
    "برای استعلام موجودی یا قیمت، نام کالا را در پیام بنویسید."
)
CONTACT = "📞 برای ارتباط با کافی‌نت و لوازم‌التحریر حربی، درخواست خود را در همین گفت‌وگو ارسال کنید."

MAIN_KEYBOARD = {
    "keyboard": [
        [{"text": "ℹ️ درباره ما"}, {"text": "💻 خدمات کافی‌نت"}],
        [{"text": "✏️ لوازم‌التحریر"}, {"text": "📞 تماس با ما"}]
    ],
    "resize_keyboard": True,
    "is_persistent": True
}

def telegram(method, payload):
    if not BOT_TOKEN:
        log.error("BOT_TOKEN environment variable is missing")
        return None
    try:
        response = requests.post(f"{API}/{method}", json=payload, timeout=15)
        response.raise_for_status()
        data = response.json()
        if not data.get("ok"):
            log.error("Telegram API error: %s", data)
        return data
    except requests.RequestException:
        log.exception("Telegram API request failed")
        return None

def send_message(chat_id, text, keyboard=True):
    payload = {"chat_id": chat_id, "text": text}
    if keyboard:
        payload["reply_markup"] = MAIN_KEYBOARD
    telegram("sendMessage", payload)

def handle_update(update):
    message = update.get("message") or update.get("edited_message")
    if not message or not message.get("chat"):
        return
    chat_id = message["chat"]["id"]
    text = (message.get("text") or "").strip()
    if not text:
        send_message(chat_id, "لطفاً پیام متنی یا نام خدمت/کالای موردنظر را ارسال کنید.")
        return

    normalized = text.lower().split("@")[0]
    if normalized in ("/start", "/start start", "شروع"):
        send_message(chat_id, "سلام و خوش آمدید 🌷\nبه ربات کافی‌نت و لوازم‌التحریر حربی خوش آمدید.\nاز منوی زیر یکی از گزینه‌ها را انتخاب کنید.", True)
    elif normalized in ("/help", "راهنما"):
        send_message(chat_id, "از دکمه‌های منو استفاده کنید:\nℹ️ درباره ما\n💻 خدمات کافی‌نت\n✏️ لوازم‌التحریر\n📞 تماس با ما")
    elif normalized in ("/about", "ℹ️ درباره ما", "درباره ما", "درباره"):
        send_message(chat_id, ABOUT)
    elif normalized in ("/services", "💻 خدمات کافی‌نت", "خدمات کافی‌نت", "خدمات"):
        send_message(chat_id, SERVICES)
    elif normalized in ("/stationery", "✏️ لوازم‌التحریر", "لوازم‌التحریر", "لوازم التحریر"):
        send_message(chat_id, STATIONERY)
    elif normalized in ("/contact", "📞 تماس با ما", "تماس با ما", "ارتباط"):
        send_message(chat_id, CONTACT)
    else:
        send_message(chat_id, f"پیام شما دریافت شد: «{text}»\n\nبرای دریافت راهنمایی، یکی از گزینه‌های منو را انتخاب کنید یا نام کالا/خدمت را بنویسید.")

@app.get("/")
def home():
    return jsonify({"status": "ok", "service": "Harbi Telegram Bot"})

@app.get("/health")
def health():
    return jsonify({"status": "healthy", "bot_token_configured": bool(BOT_TOKEN)})

@app.post("/telegram-webhook")
def webhook():
    if WEBHOOK_SECRET:
        supplied = request.headers.get("X-Telegram-Bot-Api-Secret-Token", "")
        if supplied != WEBHOOK_SECRET:
            abort(403)
    update = request.get_json(silent=True)
    if not isinstance(update, dict):
        return jsonify({"ok": False, "error": "invalid update"}), 400
    handle_update(update)
    return jsonify({"ok": True})

if BOT_TOKEN and APP_URL and WEBHOOK_SECRET:
    result = telegram("setWebhook", {
        "url": f"{APP_URL}/telegram-webhook",
        "secret_token": WEBHOOK_SECRET,
        "allowed_updates": ["message"]
    })
    log.info("Webhook setup response: %s", result)
else:
    log.warning("Webhook not configured yet. Set BOT_TOKEN, WEBHOOK_SECRET, and deploy on Render.")
