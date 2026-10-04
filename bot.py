# ============================================================
# AUTO INSTALLER - single file bot
# ============================================================
import sys
import subprocess

REQUIRED_PACKAGES = [
    "pyTelegramBotAPI==4.32.0",
    "openpyxl>=3.1.0",
]

def _ensure_packages():
    """Install/update required packages automatically, then continue."""
    print("Checking required Python packages...")
    try:
        subprocess.check_call([
            sys.executable, "-m", "pip", "install", "--disable-pip-version-check",
            "--no-input", "-U", *REQUIRED_PACKAGES
        ])
        print("All required packages are installed.")
    except Exception as exc:
        print("Automatic package installation failed:", exc)
        print("Try manually: python -m pip install -U pyTelegramBotAPI openpyxl")
        raise

_ensure_packages()

import os
import io
import json
import uuid
import time
import html
from datetime import datetime

import telebot
from telebot import types
from openpyxl import load_workbook, Workbook

# =========================

# ============================================================
# TELEGRAM BUTTON COLORS — INLINE + REPLY KEYBOARD
# ============================================================
# Telegram InlineKeyboardButton supports the predefined styles
# primary (blue), success (green), and danger (red).
def _button_style(text):
    t = str(text or '').lower()
    danger = ('cancel','reject','remove','delete','clear','close','disable','decline','❌','🗑')
    success = ('accept','approve','approved','success','submit','confirm','enable',
               'buy now','deposit','save','done','join','check','on','add','set','upload','enable','✅','🟢','💰','➕')
    if any(x in t for x in danger):
        return 'danger'
    if any(x in t for x in success):
        return 'success'
    return 'primary'

def ColorInlineKeyboardButton(*args, **kwargs):
    # Older pyTelegramBotAPI versions do not support Telegram button styles.
    # Ignore the decorative style argument for broad hosting compatibility.
    kwargs.pop("style", None)
    return types.InlineKeyboardButton(*args, **kwargs)

def ColorKeyboardButton(*args, **kwargs):
    # Reply keyboard buttons do not need a style argument.
    kwargs.pop("style", None)
    return types.KeyboardButton(*args, **kwargs)

def ColorInlineKeyboardMarkup(keyboard=None, *args, **kwargs):
    # Avoid passing a None positional keyboard together with keyword
    # arguments; that causes duplicate-argument errors in telebot.
    if keyboard is None:
        return types.InlineKeyboardMarkup(*args, **kwargs)
    return types.InlineKeyboardMarkup(keyboard, *args, **kwargs)

def ColorReplyKeyboardMarkup(keyboard=None, *args, **kwargs):
    # IMPORTANT: when keyboard is omitted, do not pass None as the first
    # positional argument. This fixes: got multiple values for
    # argument 'resize_keyboard'.
    if keyboard is None:
        return types.ReplyKeyboardMarkup(*args, **kwargs)
    return types.ReplyKeyboardMarkup(keyboard, *args, **kwargs)

# Compatibility aliases: keep the rest of the bot readable while using real telebot types.
types.ColorInlineKeyboardButton = ColorInlineKeyboardButton
types.ColorKeyboardButton = ColorKeyboardButton
types.ColorInlineKeyboardMarkup = ColorInlineKeyboardMarkup
types.ColorReplyKeyboardMarkup = ColorReplyKeyboardMarkup


# CONFIG
# =========================
TOKEN = os.getenv("BOT_TOKEN", "8964138591:AAHaGnoGPR70-t1EoghqvkX7iIAdCiFiLKU").strip()
ADMIN_ID = int(os.getenv("ADMIN_ID", "7817705450"))
REFERRAL_BONUS = 1.00  # Tk reward for each new user who completes required channel checks
if not TOKEN:
    raise RuntimeError("BOT_TOKEN environment variable is missing.")
if ADMIN_ID <= 0:
    raise RuntimeError("ADMIN_ID environment variable is missing or invalid.")

DEFAULT_ADMIN_DEV = "https://t.me/NS_BD_Prime"
DEFAULT_CHANNEL = "https://t.me/AsifCryptoCommunity"

DB_FILE = "database.json"
FILES_DIR = "bot_files"
os.makedirs(FILES_DIR, exist_ok=True)

bot = telebot.TeleBot(TOKEN, parse_mode="HTML")

# BUTTON COLOR NOTE: Telegram styles apply to InlineKeyboardButton.
# ReplyKeyboardButton has no color/style field in the Bot API, so those
# buttons intentionally remain native to avoid runtime errors.

DEFAULT_DB = {
    "users": {},
    "categories": {},
    "proxies": {},
    "meta_products": {},
    "scripts": {},
    "deposits": {},
    "settings": {
        "payments": {
            "bkash": {"name": "bKash", "enabled": True, "number": ""},
            "nagad": {"name": "Nagad", "enabled": True, "number": ""},
            "rocket": {"name": "Rocket", "enabled": True, "number": ""},
            "binance": {"name": "Binance", "enabled": True, "number": ""}
        },
        "admin_dev": DEFAULT_ADMIN_DEV,
        "channel": DEFAULT_CHANNEL,
        "force_join": {
            "channel1": {"enabled": False, "chat_id": "", "url": ""},
            "channel2": {"enabled": False, "chat_id": "", "url": ""}
        }
    }
}

def load_db():
    if not os.path.exists(DB_FILE):
        save_db(DEFAULT_DB)
        return DEFAULT_DB
    try:
        with open(DB_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        for k, v in DEFAULT_DB.items():
            if k not in data:
                data[k] = v
        # Keep newly-added settings when using an older database.json
        if "force_join" not in data.get("settings", {}):
            data.setdefault("settings", {})["force_join"] = DEFAULT_DB["settings"]["force_join"].copy()
        else:
            for fk, fv in DEFAULT_DB["settings"]["force_join"].items():
                data["settings"]["force_join"].setdefault(fk, fv.copy())
        return data
    except Exception:
        return DEFAULT_DB.copy()

def save_db(data):
    tmp = DB_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    os.replace(tmp, DB_FILE)

db = load_db()

def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def is_admin(uid):
    return int(uid) == int(ADMIN_ID)

def money(v):
    return f"{float(v):.2f}"

def get_user(uid):
    uid = str(uid)
    if uid not in db["users"]:
        db["users"][uid] = {
            "id": int(uid), "balance": 0.0, "total_deposit": 0.0,
            "total_purchase": 0.0, "deposits": [], "purchases": [],
            "referred_by": None, "referrals_count": 0,
            "referral_earned": 0.0, "referral_rewarded": False
        }
        save_db(db)
    else:
        # Backward-compatible defaults for users already in database.json
        user = db["users"][uid]
        user.setdefault("referred_by", None)
        user.setdefault("referrals_count", 0)
        user.setdefault("referral_earned", 0.0)
        user.setdefault("referral_rewarded", False)
    return db["users"][uid]

def add_purchase(uid, ptype, name, qty, amount):
    u = get_user(uid)
    u["purchases"].append({
        "id": uuid.uuid4().hex[:10], "type": ptype, "name": name,
        "quantity": qty, "amount": float(amount), "date": now()
    })
    u["total_purchase"] += float(amount)
    save_db(db)

def safe_filename(name):
    name = os.path.basename(name)
    return "".join(c for c in name if c.isalnum() or c in "._-") or "file"

# =========================
# MAIN MENU
# =========================
def main_keyboard(uid):
    kb = types.ColorReplyKeyboardMarkup(resize_keyboard=True)
    # IMPORTANT: pass real KeyboardButton objects instead of plain strings,
    # otherwise telebot creates unstyled buttons automatically.
    kb.row(
        types.ColorKeyboardButton("🛒 Proxy Buy", style="primary"),
        types.ColorKeyboardButton("📱 Meta Account Buy", style="primary"),
    )
    kb.row(
        types.ColorKeyboardButton("🤖 Bot Script Buy", style="primary"),
        types.ColorKeyboardButton("💳 Deposit", style="success"),
    )
    kb.row(
        types.ColorKeyboardButton("💬 Support", style="primary"),
        types.ColorKeyboardButton("💰 My Balance", style="success"),
    )
    kb.row(types.ColorKeyboardButton("👥 Refer & Earn", style="success"))
    if is_admin(uid):
        kb.row(types.ColorKeyboardButton("⚙️ Admin Panel", style="primary"))
    return kb

def force_join_status(uid):
    """Return (joined, missing_channels). Bot must be admin in channels to verify reliably."""
    missing = []
    fj = db.get("settings", {}).get("force_join", {})
    for key in ("channel1", "channel2"):
        ch = fj.get(key, {})
        if not ch.get("enabled"):
            continue
        chat_id = str(ch.get("chat_id", "")).strip()
        if not chat_id:
            continue
        try:
            member = bot.get_chat_member(chat_id, uid)
            if member.status in ("creator", "administrator", "member"):
                continue
            # Telegram can return restricted members; allow them unless explicitly left/kicked.
            if member.status == "restricted" and getattr(member, "is_member", False):
                continue
            missing.append(ch)
        except Exception:
            # If membership cannot be checked, keep the user blocked and tell them to contact admin.
            missing.append(ch)
    return len(missing) == 0, missing

def send_force_join(message):
    ok, missing = force_join_status(message.from_user.id)
    if ok:
        return True
    kb = types.ColorInlineKeyboardMarkup()
    for i, ch in enumerate(missing, 1):
        url = str(ch.get("url", "")).strip()
        if url:
            kb.add(types.ColorInlineKeyboardButton(f"📢 Join Channel {i}", url=url))
    kb.add(types.ColorInlineKeyboardButton("✅ I Joined — Check", callback_data="forcejoin:check"))
    bot.send_message(
        message.chat.id,
        "<b>🔒 Force Join Required</b>\n\n"
        "Please join the required channel(s), then tap <b>I Joined — Check</b>.",
        reply_markup=kb
    )
    return False

def complete_referral(uid):
    """Reward a referrer once, after the new user passes required channel checks."""
    user = get_user(uid)
    referrer_id = str(user.get("referred_by") or "")
    if not referrer_id or user.get("referral_rewarded") or referrer_id == str(uid):
        return
    referrer = db["users"].get(referrer_id)
    if not referrer:
        return
    referrer.setdefault("balance", 0.0)
    referrer.setdefault("referrals_count", 0)
    referrer.setdefault("referral_earned", 0.0)
    referrer["balance"] = float(referrer["balance"]) + REFERRAL_BONUS
    referrer["referrals_count"] += 1
    referrer["referral_earned"] += REFERRAL_BONUS
    user["referral_rewarded"] = True
    save_db(db)
    try:
        bot.send_message(
            int(referrer_id),
            f"🎉 New referral joined! You earned ৳{money(REFERRAL_BONUS)}.\n"
            f"Total referrals: {referrer['referrals_count']}"
        )
    except Exception:
        pass

@bot.message_handler(commands=["start"])
def start(message):
    uid = str(message.from_user.id)
    is_new_user = uid not in db["users"]
    user = get_user(uid)
    # Telegram deep-link format: /start ref_USERID
    if is_new_user and len(message.text.split()) > 1:
        payload = message.text.split(maxsplit=1)[1].strip()
        if payload.startswith("ref_"):
            referrer_id = payload[4:]
            if referrer_id.isdigit() and referrer_id != uid and referrer_id in db["users"]:
                user["referred_by"] = referrer_id
                save_db(db)
    if not send_force_join(message):
        return
    complete_referral(uid)
    bot.send_message(
        message.chat.id,
        "<b>🏠 MAIN MENU</b>\n\nChoose an option:",
        reply_markup=main_keyboard(message.from_user.id)
    )

@bot.message_handler(func=lambda m: m.text == "👥 Refer & Earn")
def referral_menu(message):
    user = get_user(message.from_user.id)
    try:
        bot_username = bot.get_me().username
        link = f"https://t.me/{bot_username}?start=ref_{message.from_user.id}"
    except Exception:
        bot.send_message(message.chat.id, "❌ Referral link could not be created. Try again later.")
        return
    bot.send_message(
        message.chat.id,
        "<b>👥 REFER & EARN</b>\n\n"
        f"Your referral link:\n<code>{link}</code>\n\n"
        f"Reward per valid new referral: <b>৳{money(REFERRAL_BONUS)}</b>\n"
        f"Total referrals: <b>{int(user.get('referrals_count', 0))}</b>\n"
        f"Total earned: <b>৳{money(user.get('referral_earned', 0))}</b>\n\n"
        "A reward is added once when a new user starts through your link and completes the required channel checks."
    )

@bot.callback_query_handler(func=lambda c: c.data == "forcejoin:check")
def forcejoin_check(call):
    ok, missing = force_join_status(call.from_user.id)
    if not ok:
        bot.answer_callback_query(call.id, "❌ You have not joined all required channels yet.", show_alert=True)
        return
    bot.answer_callback_query(call.id, "✅ Verified!")
    complete_referral(call.from_user.id)
    try:
        bot.delete_message(call.message.chat.id, call.message.message_id)
    except Exception:
        pass
    bot.send_message(
        call.message.chat.id,
        "<b>🏠 MAIN MENU</b>\n\nChoose an option:",
        reply_markup=main_keyboard(call.from_user.id)
    )

@bot.message_handler(func=lambda m: m.text == "🛒 Proxy Buy")
def proxy_buy(message):
    kb = types.ColorInlineKeyboardMarkup()
    for cid, c in db["categories"].items():
        kb.add(types.ColorInlineKeyboardButton(
            f"📂 {c['name']}", callback_data=f"proxycat:{cid}"
        ))
    if not kb.keyboard:
        bot.send_message(message.chat.id, "❌ No proxy category available.")
        return
    bot.send_message(message.chat.id, "<b>🛒 PROXY BUY</b>\n\nSelect category:", reply_markup=kb)

@bot.message_handler(func=lambda m: m.text == "📱 Meta Account Buy")
def meta_buy(message):
    kb = types.ColorInlineKeyboardMarkup()
    for mid, p in db["meta_products"].items():
        stock = len(p.get("rows", []))
        if stock:
            kb.add(types.ColorInlineKeyboardButton(
                f"{p['name']} | ৳{money(p['price'])}",
                callback_data=f"meta:{mid}"
            ))
    if not kb.keyboard:
        bot.send_message(message.chat.id, "❌ No Meta product available.")
        return
    bot.send_message(message.chat.id, "<b>📱 META ACCOUNT BUY</b>\n\nSelect product:", reply_markup=kb)

@bot.message_handler(func=lambda m: m.text == "🤖 Bot Script Buy")
def script_buy(message):
    kb = types.ColorInlineKeyboardMarkup()
    for sid, s in db["scripts"].items():
        kb.add(types.ColorInlineKeyboardButton(
            f"🤖 {s['name']}", callback_data=f"script:{sid}"
        ))
    if not kb.keyboard:
        bot.send_message(message.chat.id, "❌ No Bot Script available.")
        return
    bot.send_message(message.chat.id, "<b>🤖 BOT SCRIPT BUY</b>\n\n👇 Select a Bot:", reply_markup=kb)

@bot.message_handler(func=lambda m: m.text == "💬 Support")
def support(message):
    kb = types.ColorInlineKeyboardMarkup()
    kb.add(types.ColorInlineKeyboardButton("👨‍💻 Admin / Dev", url=db["settings"]["admin_dev"]))
    kb.add(types.ColorInlineKeyboardButton("📢 Channel", url=db["settings"]["channel"]))
    bot.send_message(message.chat.id, "<b>💬 SUPPORT</b>", reply_markup=kb)

@bot.message_handler(func=lambda m: m.text == "💰 My Balance")
def balance(message):
    u = get_user(message.from_user.id)
    text = (
        "<b>💰 MY BALANCE</b>\n\n"
        f"Current Balance: <b>৳{money(u['balance'])}</b>\n"
        f"Total Deposit: <b>৳{money(u['total_deposit'])}</b>\n"
        f"Total Purchase: <b>৳{money(u['total_purchase'])}</b>\n\n"
        "<b>📜 Deposit History</b>\n"
    )
    ds = u.get("deposits", [])
    if ds:
        for d in ds[-10:][::-1]:
            text += f"• ৳{money(d['amount'])} | {d['method']} | {d['status']}\n"
    else:
        text += "No deposit history.\n"
    text += "\n<b>🛒 Purchase History</b>\n"
    ps = u.get("purchases", [])
    if ps:
        for p in ps[-10:][::-1]:
            text += f"• {p['name']} × {p['quantity']} | ৳{money(p['amount'])}\n"
    else:
        text += "No purchase history."
    bot.send_message(message.chat.id, text)

# =========================
# PROXY
# =========================
@bot.callback_query_handler(func=lambda c: c.data.startswith("proxycat:"))
def proxy_category(call):
    cid = call.data.split(":", 1)[1]
    kb = types.ColorInlineKeyboardMarkup()
    for pid, p in db["proxies"].items():
        if p["category_id"] == cid and p["stock"] > 0:
            kb.add(types.ColorInlineKeyboardButton(
                f"{p['name']} | {p['size']} | ৳{money(p['price'])} | Stock: {p['stock']}",
                callback_data=f"proxy:{pid}"
            ))
    if not kb.keyboard:
        bot.answer_callback_query(call.id, "No proxy available.")
        return
    bot.send_message(call.message.chat.id, "<b>🌐 PROXY LIST</b>", reply_markup=kb)

@bot.callback_query_handler(func=lambda c: c.data.startswith("proxy:"))
def proxy_details(call):
    pid = call.data.split(":", 1)[1]
    p = db["proxies"].get(pid)
    if not p:
        return
    kb = types.ColorInlineKeyboardMarkup()
    kb.add(types.ColorInlineKeyboardButton(
        f"🛒 BUY NOW | ৳{money(p['price'])}",
        callback_data=f"buyproxy:{pid}"
    ))
    bot.send_message(
        call.message.chat.id,
        f"<b>🌐 {p['name']}</b>\n\n"
        f"Size: {p['size']}\nPrice: ৳{money(p['price'])}\nStock: {p['stock']}",
        reply_markup=kb
    )

@bot.callback_query_handler(func=lambda c: c.data.startswith("buyproxy:"))
def buy_proxy(call):
    pid = call.data.split(":", 1)[1]
    p = db["proxies"].get(pid)
    if not p or not p["inventory"]:
        bot.answer_callback_query(call.id, "Out of stock.")
        return
    u = get_user(call.from_user.id)
    price = float(p["price"])
    if u["balance"] < price:
        bot.answer_callback_query(call.id, "Insufficient balance.")
        return
    item = p["inventory"].pop(0)
    p["stock"] = len(p["inventory"])
    u["balance"] -= price
    add_purchase(call.from_user.id, "proxy", p["name"], 1, price)
    save_db(db)
    bot.send_message(
        call.message.chat.id,
        f"✅ <b>Proxy Purchased</b>\n\n"
        f"Name: {p['name']}\nSize: {p['size']}\n"
        f"Price: ৳{money(price)}\n"
        f"Remaining Balance: ৳{money(u['balance'])}"
    )
    bot.send_message(
        call.message.chat.id,
        "<b>🔐 YOUR PROXY</b>\n\n"
        f"Server: <code>{item['server']}</code>\n"
        f"Port: <code>{item['port']}</code>\n"
        f"Username: <code>{item['username']}</code>\n"
        f"Password: <code>{item['password']}</code>"
    )

# =========================
# META
# =========================
@bot.callback_query_handler(func=lambda c: c.data.startswith("meta:"))
def meta_details(call):
    mid = call.data.split(":", 1)[1]
    p = db["meta_products"].get(mid)
    if not p:
        return
    stock = len(p.get("rows", []))
    bot.send_message(
        call.message.chat.id,
        f"<b>📱 {p['name']}</b>\n\n"
        f"Price: ৳{money(p['price'])} / piece\n"
        f"Stock: {stock}\n\n"
        "Send quantity:"
    )
    bot.register_next_step_handler(call.message, meta_quantity, mid)

def meta_quantity(message, mid):
    if not message.text or not message.text.isdigit() or int(message.text) <= 0:
        bot.send_message(message.chat.id, "❌ Enter a valid quantity.")
        return
    qty = int(message.text)
    p = db["meta_products"].get(mid)
    if not p:
        return
    rows = p.get("rows", [])
    if qty > len(rows):
        bot.send_message(message.chat.id, f"❌ Only {len(rows)} available.")
        return
    total = float(p["price"]) * qty
    u = get_user(message.from_user.id)
    if u["balance"] < total:
        bot.send_message(
            message.chat.id,
            f"❌ Insufficient balance.\nRequired: ৳{money(total)}\nBalance: ৳{money(u['balance'])}"
        )
        return
    selected = rows[:qty]
    p["rows"] = rows[qty:]
    u["balance"] -= total
    add_purchase(message.from_user.id, "meta", p["name"], qty, total)
    save_db(db)

    out = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Meta"
    headers = list(selected[0].keys())
    for col, h in enumerate(headers, 1):
        ws.cell(1, col).value = h
    for r, item in enumerate(selected, 2):
        for col, h in enumerate(headers, 1):
            ws.cell(r, col).value = item.get(h, "")
    wb.save(out)
    out.seek(0)

    bot.send_message(
        message.chat.id,
        f"✅ <b>Purchase Successful</b>\n\n"
        f"Product: {p['name']}\nQuantity: {qty}\n"
        f"Total: ৳{money(total)}\n"
        f"Remaining Balance: ৳{money(u['balance'])}"
    )
    bot.send_document(message.chat.id, out, visible_file_name=f"{p['name']}_{qty}_pcs.xlsx")

# =========================
# BOT SCRIPT
# =========================
@bot.callback_query_handler(func=lambda c: c.data.startswith("script:"))
def script_details(call):
    sid = call.data.split(":", 1)[1]
    s = db["scripts"].get(sid)
    if not s:
        return
    kb = types.ColorInlineKeyboardMarkup()
    kb.add(types.ColorInlineKeyboardButton(
        f"🛒 BUY NOW | ৳{money(s['price'])}",
        callback_data=f"buyscript:{sid}"
    ))
    caption = f"<b>🤖 {s['name']}</b>\n\n💰 Price: ৳{money(s['price'])}"
    if s.get("picture_file_id"):
        bot.send_photo(call.message.chat.id, s["picture_file_id"], caption=caption, reply_markup=kb)
    else:
        bot.send_message(call.message.chat.id, caption, reply_markup=kb)

@bot.callback_query_handler(func=lambda c: c.data.startswith("buyscript:"))
def buy_script(call):
    sid = call.data.split(":", 1)[1]
    s = db["scripts"].get(sid)
    if not s or not os.path.exists(s.get("file_path", "")):
        bot.answer_callback_query(call.id, "File unavailable.")
        return
    u = get_user(call.from_user.id)
    price = float(s["price"])
    if u["balance"] < price:
        bot.answer_callback_query(call.id, "Insufficient balance.")
        return
    u["balance"] -= price
    add_purchase(call.from_user.id, "script", s["name"], 1, price)
    save_db(db)
    bot.send_message(
        call.message.chat.id,
        f"✅ <b>Purchase Successful</b>\n\n"
        f"Script: {s['name']}\nPrice: ৳{money(price)}\n"
        f"Remaining Balance: ৳{money(u['balance'])}"
    )
    with open(s["file_path"], "rb") as f:
        bot.send_document(call.message.chat.id, f, caption=f"🤖 {s['name']}")

# =========================
# DEPOSIT
# =========================
@bot.message_handler(func=lambda m: m.text == "💳 Deposit")
def deposit_menu(message):
    kb = types.ColorInlineKeyboardMarkup()
    for key, p in db["settings"]["payments"].items():
        if p["enabled"]:
            kb.add(types.ColorInlineKeyboardButton(
                f"💳 {p['name']}", callback_data=f"deposit:{key}"
            ))
    if not kb.keyboard:
        bot.send_message(message.chat.id, "❌ No payment method available.")
        return
    bot.send_message(message.chat.id, "<b>💳 DEPOSIT</b>\n\nSelect method:", reply_markup=kb)

@bot.callback_query_handler(func=lambda c: c.data.startswith("deposit:"))
def deposit_start(call):
    key = call.data.split(":", 1)[1]
    p = db["settings"]["payments"].get(key)
    if not p or not p["enabled"]:
        return
    bot.send_message(
        call.message.chat.id,
        f"<b>💳 {p['name']}</b>\n\n"
        f"Number / ID: <code>{p['number']}</code>\n\n"
        "Send deposit amount:"
    )
    bot.register_next_step_handler(call.message, deposit_amount, key)

def deposit_amount(message, key):
    try:
        amount = float(message.text.strip())
        if amount <= 0:
            raise ValueError
    except Exception:
        bot.send_message(message.chat.id, "❌ Invalid amount.")
        return
    bot.send_message(message.chat.id, "Send Transaction ID / Reference:")
    bot.register_next_step_handler(message, deposit_reference, key, amount)

def deposit_reference(message, key, amount):
    ref = message.text.strip()
    did = uuid.uuid4().hex[:10]
    method = db["settings"]["payments"][key]["name"]
    db["deposits"][did] = {
        "id": did, "user_id": message.from_user.id, "method": method,
        "amount": amount, "reference": ref, "status": "Pending", "date": now()
    }
    u = get_user(message.from_user.id)
    u["deposits"].append({
        "id": did, "method": method, "amount": amount,
        "reference": ref, "status": "Pending", "date": now()
    })
    save_db(db)

    bot.send_message(message.chat.id, "✅ Deposit request submitted. Status: Pending")
    kb = types.ColorInlineKeyboardMarkup()
    kb.row(
        types.ColorInlineKeyboardButton("✅ Accept", callback_data=f"acceptdep:{did}"),
        types.ColorInlineKeyboardButton("❌ Reject", callback_data=f"rejectdep:{did}")
    )
    bot.send_message(
        ADMIN_ID,
        f"<b>💳 NEW DEPOSIT</b>\n\n"
        f"User ID: <code>{message.from_user.id}</code>\n"
        f"Method: {method}\nAmount: ৳{money(amount)}\n"
        f"Reference: {ref}\nID: {did}",
        reply_markup=kb
    )

@bot.callback_query_handler(func=lambda c: c.data.startswith("acceptdep:"))
def accept_deposit(call):
    if not is_admin(call.from_user.id):
        return
    did = call.data.split(":", 1)[1]
    d = db["deposits"].get(did)
    if not d or d["status"] != "Pending":
        return
    d["status"] = "Accepted"
    u = get_user(d["user_id"])
    u["balance"] += float(d["amount"])
    u["total_deposit"] += float(d["amount"])
    for x in u["deposits"]:
        if x["id"] == did:
            x["status"] = "Accepted"
    save_db(db)
    bot.edit_message_text("✅ Deposit accepted.", call.message.chat.id, call.message.message_id)
    try:
        bot.send_message(d["user_id"], f"✅ Deposit accepted.\nAdded: ৳{money(d['amount'])}\nBalance: ৳{money(u['balance'])}")
    except Exception:
        pass

@bot.callback_query_handler(func=lambda c: c.data.startswith("rejectdep:"))
def reject_deposit(call):
    if not is_admin(call.from_user.id):
        return
    did = call.data.split(":", 1)[1]
    d = db["deposits"].get(did)
    if not d or d["status"] != "Pending":
        return
    d["status"] = "Rejected"
    u = get_user(d["user_id"])
    for x in u["deposits"]:
        if x["id"] == did:
            x["status"] = "Rejected"
    save_db(db)
    bot.edit_message_text("❌ Deposit rejected.", call.message.chat.id, call.message.message_id)
    try:
        bot.send_message(d["user_id"], f"❌ Deposit rejected.\nAmount: ৳{money(d['amount'])}")
    except Exception:
        pass

# =========================
# ADMIN PANEL
# =========================
@bot.message_handler(commands=["admin"])
def admin_command(message):
    if is_admin(message.from_user.id):
        admin_panel(message)
    else:
        bot.reply_to(message, "❌ Admin only.")

@bot.message_handler(func=lambda m: m.text == "⚙️ Admin Panel")
def admin_button(message):
    if is_admin(message.from_user.id):
        admin_panel(message)

def admin_panel(message):
    if not is_admin(message.from_user.id):
        return
    kb = types.ColorInlineKeyboardMarkup()
    kb.add(types.ColorInlineKeyboardButton("➕ Add Proxy Category", callback_data="admin:addcat"))
    kb.row(
        types.ColorInlineKeyboardButton("➕ Add Proxy", callback_data="admin:addproxy"),
        types.ColorInlineKeyboardButton("📂 Categories", callback_data="admin:categories")
    )
    kb.add(types.ColorInlineKeyboardButton("➕ Add Meta (XLSX)", callback_data="admin:addmeta"))
    kb.add(types.ColorInlineKeyboardButton("📱 Meta Products", callback_data="admin:metas"))
    kb.add(types.ColorInlineKeyboardButton("➕ Add Bot Script", callback_data="admin:addscript"))
    kb.add(types.ColorInlineKeyboardButton("🤖 Bot Scripts", callback_data="admin:scripts"))
    kb.add(types.ColorInlineKeyboardButton("💳 Deposit Settings", callback_data="admin:payments"))
    kb.add(types.ColorInlineKeyboardButton("📥 Pending Deposits", callback_data="admin:deposits"))
    kb.add(types.ColorInlineKeyboardButton("💬 Support Settings", callback_data="admin:support"))
    kb.add(types.ColorInlineKeyboardButton("📢 Broadcast Message", callback_data="admin:broadcast"))
    kb.add(types.ColorInlineKeyboardButton("📊 Live Statistics", callback_data="admin:stats"))
    kb.add(types.ColorInlineKeyboardButton("🔒 Force Join (2 Channels)", callback_data="admin:forcejoin"))
    bot.send_message(message.chat.id, "<b>⚙️ ADMIN PANEL</b>", reply_markup=kb)

@bot.callback_query_handler(func=lambda c: c.data.startswith("admin:"))
def admin_actions(call):
    if not is_admin(call.from_user.id):
        return
    action = call.data.split(":", 1)[1]
    if action == "addcat":
        bot.send_message(call.message.chat.id, "Send Proxy Category name:")
        bot.register_next_step_handler(call.message, add_category)
    elif action == "addproxy":
        if not db["categories"]:
            bot.send_message(call.message.chat.id, "❌ Create a category first.")
            return
        kb = types.ColorInlineKeyboardMarkup()
        for cid, c in db["categories"].items():
            kb.add(types.ColorInlineKeyboardButton(c["name"], callback_data=f"addproxycat:{cid}"))
        bot.send_message(call.message.chat.id, "Select category:", reply_markup=kb)
    elif action == "categories":
        list_categories(call)
    elif action == "addmeta":
        bot.send_message(call.message.chat.id, "Send Meta product name:")
        bot.register_next_step_handler(call.message, add_meta_name)
    elif action == "metas":
        list_metas(call)
    elif action == "addscript":
        bot.send_message(call.message.chat.id, "Send Script name:")
        bot.register_next_step_handler(call.message, add_script_name)
    elif action == "scripts":
        list_scripts(call)
    elif action == "payments":
        payment_settings(call)
    elif action == "deposits":
        pending_deposits(call)
    elif action == "support":
        support_settings(call)
    elif action == "forcejoin":
        force_join_settings(call)
    elif action == "broadcast":
        bot.send_message(call.message.chat.id, "📢 Send the message to broadcast to all bot users.\nSend /cancel to stop.")
        bot.register_next_step_handler(call.message, broadcast_message)
    elif action == "stats":
        send_live_statistics(call.message.chat.id)

def broadcast_message(message):
    if not is_admin(message.from_user.id):
        return
    if message.text and message.text.strip() == "/cancel":
        bot.send_message(message.chat.id, "❎ Broadcast cancelled.")
        return
    if not message.text or not message.text.strip():
        bot.send_message(message.chat.id, "❌ Please send a text message. Broadcast cancelled.")
        return
    # Escape HTML so user-provided broadcast text cannot break Telegram formatting.
    broadcast_text = html.escape(message.text.strip())
    user_ids = list(db.get("users", {}).keys())
    sent = 0
    failed = 0
    status = bot.send_message(message.chat.id, f"⏳ Broadcasting to {len(user_ids)} users...")
    for user_id in user_ids:
        try:
            bot.send_message(int(user_id), broadcast_text)
            sent += 1
        except Exception:
            failed += 1
        time.sleep(0.05)  # keep message rate within Telegram limits
    bot.edit_message_text(
        f"✅ <b>Broadcast finished</b>\n\n"
        f"Total users: {len(user_ids)}\n"
        f"Sent: {sent}\nFailed: {failed}",
        message.chat.id,
        status.message_id
    )

def send_live_statistics(chat_id):
    users = list(db.get("users", {}).values())
    total_users = len(users)
    total_deposits = sum(float(u.get("total_deposit", 0) or 0) for u in users)
    total_sales = sum(float(u.get("total_purchase", 0) or 0) for u in users)
    total_referrals = sum(int(u.get("referrals_count", 0) or 0) for u in users)
    pending_deposits = sum(1 for d in db.get("deposits", {}).values() if d.get("status") == "Pending")
    meta_stock = sum(len(p.get("rows", [])) for p in db.get("meta_products", {}).values())
    proxy_stock = sum(int(p.get("stock", len(p.get("inventory", []))) or 0) for p in db.get("proxies", {}).values())
    script_count = len(db.get("scripts", {}))
    text = (
        "<b>📊 LIVE SHOP STATISTICS</b>\n\n"
        f"👥 Total users: <b>{total_users}</b>\n"
        f"🔗 Successful referrals: <b>{total_referrals}</b>\n"
        f"💳 Approved deposits: <b>৳{money(total_deposits)}</b>\n"
        f"🛒 Total sales: <b>৳{money(total_sales)}</b>\n"
        f"⏳ Pending deposits: <b>{pending_deposits}</b>\n\n"
        f"📱 Meta stock: <b>{meta_stock}</b>\n"
        f"🌐 Proxy stock: <b>{proxy_stock}</b>\n"
        f"🤖 Bot scripts: <b>{script_count}</b>"
    )
    bot.send_message(chat_id, text)

# =========================
# ADMIN PROXY
# =========================
def add_category(message):
    name = message.text.strip()
    cid = uuid.uuid4().hex[:8]
    db["categories"][cid] = {"name": name, "created": now()}
    save_db(db)
    bot.send_message(message.chat.id, f"✅ Category added: <b>{name}</b>")

@bot.callback_query_handler(func=lambda c: c.data.startswith("addproxycat:"))
def add_proxy_category(call):
    cid = call.data.split(":", 1)[1]
    bot.send_message(call.message.chat.id, "Proxy Name:")
    bot.register_next_step_handler(call.message, proxy_name, cid)

def proxy_name(message, cid):
    bot.send_message(message.chat.id, "GB/MB:")
    bot.register_next_step_handler(message, proxy_size, cid, message.text.strip())

def proxy_size(message, cid, name):
    bot.send_message(message.chat.id, "Price:")
    bot.register_next_step_handler(message, proxy_price, cid, name, message.text.strip())

def proxy_price(message, cid, name, size):
    try:
        price = float(message.text.strip())
    except Exception:
        bot.send_message(message.chat.id, "❌ Invalid price.")
        return
    bot.send_message(message.chat.id, "Server:")
    bot.register_next_step_handler(message, proxy_server, cid, name, size, price)

def proxy_server(message, cid, name, size, price):
    bot.send_message(message.chat.id, "Port:")
    bot.register_next_step_handler(message, proxy_port, cid, name, size, price, message.text.strip())

def proxy_port(message, cid, name, size, price, server):
    bot.send_message(message.chat.id, "Username:")
    bot.register_next_step_handler(message, proxy_username, cid, name, size, price, server, message.text.strip())

def proxy_username(message, cid, name, size, price, server, port):
    bot.send_message(message.chat.id, "Password:")
    bot.register_next_step_handler(message, proxy_password, cid, name, size, price, server, port, message.text.strip())

def proxy_password(message, cid, name, size, price, server, port, username):
    password = message.text.strip()
    existing = None
    for pid, p in db["proxies"].items():
        if p["category_id"] == cid and p["name"].lower() == name.lower() and str(p["size"]).lower() == str(size).lower() and float(p["price"]) == float(price):
            existing = p
            break
    item = {"server": server, "port": port, "username": username, "password": password}
    if existing:
        existing["inventory"].append(item)
        existing["stock"] = len(existing["inventory"])
        save_db(db)
        bot.send_message(message.chat.id, f"✅ Same proxy found. Stock: <b>{existing['stock']}</b>")
    else:
        pid = uuid.uuid4().hex[:10]
        db["proxies"][pid] = {
            "id": pid, "category_id": cid, "name": name, "size": size,
            "price": float(price), "inventory": [item], "stock": 1, "created": now()
        }
        save_db(db)
        bot.send_message(message.chat.id, f"✅ Proxy added. Stock: <b>1</b>")

# =========================
# ADMIN META
# =========================
def add_meta_name(message):
    name = message.text.strip()
    bot.send_message(message.chat.id, "Price per piece:")
    bot.register_next_step_handler(message, add_meta_price, name)

def add_meta_price(message, name):
    try:
        price = float(message.text.strip())
    except Exception:
        bot.send_message(message.chat.id, "❌ Invalid price.")
        return
    bot.send_message(message.chat.id, "Upload XLSX file:")
    bot.register_next_step_handler(message, receive_meta_xlsx, name, price)

def receive_meta_xlsx(message, name, price):
    if not message.document or not (message.document.file_name or "").lower().endswith(".xlsx"):
        bot.send_message(message.chat.id, "❌ Upload an .xlsx file.")
        return
    try:
        info = bot.get_file(message.document.file_id)
        raw = bot.download_file(info.file_path)
        wb = load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            bot.send_message(message.chat.id, "❌ XLSX is empty.")
            return
        headers = [str(x).strip() if x is not None else "" for x in rows[0]]
        data = []
        for row in rows[1:]:
            if all(x is None for x in row):
                continue
            item = {}
            for i, h in enumerate(headers):
                item[h or f"Column_{i+1}"] = row[i] if i < len(row) else ""
            data.append(item)
        if not data:
            bot.send_message(message.chat.id, "❌ No data rows.")
            return
        existing = None
        for p in db["meta_products"].values():
            if p["name"].lower() == name.lower() and float(p["price"]) == float(price):
                existing = p
                break
        if existing:
            existing["rows"].extend(data)
            stock = len(existing["rows"])
        else:
            mid = uuid.uuid4().hex[:10]
            db["meta_products"][mid] = {
                "id": mid, "name": name, "price": float(price),
                "rows": data, "created": now()
            }
            stock = len(data)
        save_db(db)
        bot.send_message(message.chat.id, f"✅ Meta added/updated.\nName: {name}\nPrice: ৳{money(price)}\nStock: {stock}")
    except Exception as e:
        bot.send_message(message.chat.id, f"❌ XLSX error: <code>{e}</code>")

# =========================
# ADMIN SCRIPT
# =========================
def add_script_name(message):
    name = message.text.strip()
    bot.send_message(message.chat.id, "Price:")
    bot.register_next_step_handler(message, add_script_price, name)

def add_script_price(message, name):
    try:
        price = float(message.text.strip())
    except Exception:
        bot.send_message(message.chat.id, "❌ Invalid price.")
        return
    bot.send_message(message.chat.id, "Upload Script Picture:")
    bot.register_next_step_handler(message, add_script_picture, name, price)

def add_script_picture(message, name, price):
    if not message.photo:
        bot.send_message(message.chat.id, "❌ Send a picture.")
        return
    picture = message.photo[-1].file_id
    bot.send_message(message.chat.id, "Upload Script File:")
    bot.register_next_step_handler(message, add_script_file, name, price, picture)

def add_script_file(message, name, price, picture):
    if not message.document:
        bot.send_message(message.chat.id, "❌ Upload a file.")
        return
    try:
        info = bot.get_file(message.document.file_id)
        raw = bot.download_file(info.file_path)
        sid = uuid.uuid4().hex[:10]
        filename = safe_filename(message.document.file_name or "script_file")
        path = os.path.join(FILES_DIR, f"{sid}_{filename}")
        with open(path, "wb") as f:
            f.write(raw)
        db["scripts"][sid] = {
            "id": sid, "name": name, "price": float(price),
            "picture_file_id": picture, "file_path": path, "created": now()
        }
        save_db(db)
        bot.send_message(message.chat.id, f"✅ Bot Script added: <b>{name}</b>\nPrice: ৳{money(price)}\nStock: Unlimited")
    except Exception as e:
        bot.send_message(message.chat.id, f"❌ File error: <code>{e}</code>")

# =========================
# ADMIN LISTS / SETTINGS
# =========================
def list_categories(call):
    kb = types.ColorInlineKeyboardMarkup()
    for cid, c in db["categories"].items():
        kb.add(types.ColorInlineKeyboardButton(c["name"], callback_data=f"delcat:{cid}"))
    bot.send_message(call.message.chat.id, "📂 Categories — tap to delete:", reply_markup=kb)

def list_metas(call):
    kb = types.ColorInlineKeyboardMarkup()
    for mid, p in db["meta_products"].items():
        kb.add(types.ColorInlineKeyboardButton(
            f"{p['name']} | Stock {len(p['rows'])}",
            callback_data=f"delmeta:{mid}"
        ))
    bot.send_message(call.message.chat.id, "📱 Meta Products:", reply_markup=kb)

def list_scripts(call):
    kb = types.ColorInlineKeyboardMarkup()
    for sid, s in db["scripts"].items():
        kb.add(types.ColorInlineKeyboardButton(
            f"{s['name']} | ৳{money(s['price'])}",
            callback_data=f"delscript:{sid}"
        ))
    bot.send_message(call.message.chat.id, "🤖 Bot Scripts — tap to delete:", reply_markup=kb)

def force_join_settings(call):
    fj = db["settings"].setdefault("force_join", {})
    for key in ("channel1", "channel2"):
        fj.setdefault(key, {"enabled": False, "chat_id": "", "url": ""})
    kb = types.ColorInlineKeyboardMarkup()
    for key, label in (("channel1", "Channel 1"), ("channel2", "Channel 2")):
        ch = fj[key]
        state = "ON 🟢" if ch.get("enabled") else "OFF 🔴"
        kb.add(types.ColorInlineKeyboardButton(f"{label} — {state}", callback_data=f"fjtog:{key}"))
        kb.add(types.ColorInlineKeyboardButton(f"✏️ Set {label}", callback_data=f"fjset:{key}"))
    bot.send_message(
        call.message.chat.id,
        "<b>🔒 Force Join Settings</b>\n\n"
        "You can configure up to 2 required channels.\n"
        "The bot must be able to check membership (normally add the bot as admin in the channel).",
        reply_markup=kb
    )

@bot.callback_query_handler(func=lambda c: c.data.startswith("fjtog:"))
def force_join_toggle(call):
    if not is_admin(call.from_user.id):
        return
    key = call.data.split(":", 1)[1]
    ch = db["settings"]["force_join"][key]
    ch["enabled"] = not ch.get("enabled", False)
    save_db(db)
    bot.answer_callback_query(call.id, "Updated.")
    force_join_settings(call)

@bot.callback_query_handler(func=lambda c: c.data.startswith("fjset:"))
def force_join_set(call):
    if not is_admin(call.from_user.id):
        return
    key = call.data.split(":", 1)[1]
    bot.send_message(
        call.message.chat.id,
        f"Send {key[-1]} Channel chat ID or @username\n\n"
        "Example: @mychannel or -1001234567890"
    )
    bot.register_next_step_handler(call.message, force_join_chat_id, key)

def force_join_chat_id(message, key):
    chat_id = message.text.strip()
    if not chat_id:
        bot.send_message(message.chat.id, "❌ Invalid chat ID/username.")
        return
    db["settings"]["force_join"][key]["chat_id"] = chat_id
    save_db(db)
    bot.send_message(
        message.chat.id,
        "Now send the channel join URL.\nExample: https://t.me/mychannel"
    )
    bot.register_next_step_handler(message, force_join_url, key)

def force_join_url(message, key):
    url = message.text.strip()
    if not (url.startswith("https://t.me/") or url.startswith("http://t.me/") or url.startswith("tg://")):
        bot.send_message(message.chat.id, "❌ Invalid Telegram join URL.")
        return
    db["settings"]["force_join"][key]["url"] = url
    save_db(db)
    bot.send_message(message.chat.id, f"✅ {key} configured. Turn it ON from Force Join Settings.")

def payment_settings(call):
    kb = types.ColorInlineKeyboardMarkup()
    for key, p in db["settings"]["payments"].items():
        kb.add(types.ColorInlineKeyboardButton(
            f"{p['name']} — {'ON 🟢' if p['enabled'] else 'OFF 🔴'}",
            callback_data=f"paytoggle:{key}"
        ))
        kb.add(types.ColorInlineKeyboardButton(
            f"✏️ Set {p['name']} Number/ID",
            callback_data=f"paynumber:{key}"
        ))
    bot.send_message(call.message.chat.id, "💳 Deposit Settings:", reply_markup=kb)

def pending_deposits(call):
    kb = types.ColorInlineKeyboardMarkup()
    for did, d in db["deposits"].items():
        if d["status"] == "Pending":
            kb.add(types.ColorInlineKeyboardButton(
                f"৳{money(d['amount'])} | {d['method']} | {d['user_id']}",
                callback_data=f"viewdep:{did}"
            ))
    if not kb.keyboard:
        bot.send_message(call.message.chat.id, "✅ No pending deposits.")
    else:
        bot.send_message(call.message.chat.id, "📥 Pending Deposits:", reply_markup=kb)

def support_settings(call):
    kb = types.ColorInlineKeyboardMarkup()
    kb.add(types.ColorInlineKeyboardButton("👨‍💻 Set Admin / Dev", callback_data="setsupport:admin"))
    kb.add(types.ColorInlineKeyboardButton("📢 Set Channel", callback_data="setsupport:channel"))
    bot.send_message(call.message.chat.id, "💬 Support Settings:", reply_markup=kb)

@bot.callback_query_handler(func=lambda c: c.data.startswith("paytoggle:"))
def pay_toggle(call):
    if not is_admin(call.from_user.id):
        return
    key = call.data.split(":", 1)[1]
    db["settings"]["payments"][key]["enabled"] = not db["settings"]["payments"][key]["enabled"]
    save_db(db)
    bot.answer_callback_query(call.id, "Updated.")
    payment_settings(call)

@bot.callback_query_handler(func=lambda c: c.data.startswith("paynumber:"))
def pay_number(call):
    if not is_admin(call.from_user.id):
        return
    key = call.data.split(":", 1)[1]
    bot.send_message(call.message.chat.id, f"Send {db['settings']['payments'][key]['name']} Number/ID:")
    bot.register_next_step_handler(call.message, save_pay_number, key)

def save_pay_number(message, key):
    db["settings"]["payments"][key]["number"] = message.text.strip()
    save_db(db)
    bot.send_message(message.chat.id, "✅ Updated.")

@bot.callback_query_handler(func=lambda c: c.data.startswith("setsupport:"))
def set_support(call):
    if not is_admin(call.from_user.id):
        return
    key = call.data.split(":", 1)[1]
    bot.send_message(call.message.chat.id, "Send Telegram URL:")
    bot.register_next_step_handler(call.message, save_support, key)

def save_support(message, key):
    if key == "admin":
        db["settings"]["admin_dev"] = message.text.strip()
    else:
        db["settings"]["channel"] = message.text.strip()
    save_db(db)
    bot.send_message(message.chat.id, "✅ Support setting updated.")

@bot.callback_query_handler(func=lambda c: c.data.startswith("viewdep:"))
def view_dep(call):
    if not is_admin(call.from_user.id):
        return
    did = call.data.split(":", 1)[1]
    d = db["deposits"].get(did)
    if not d:
        return
    kb = types.ColorInlineKeyboardMarkup()
    kb.row(
        types.ColorInlineKeyboardButton("✅ Accept", callback_data=f"acceptdep:{did}"),
        types.ColorInlineKeyboardButton("❌ Reject", callback_data=f"rejectdep:{did}")
    )
    bot.send_message(
        call.message.chat.id,
        f"<b>💳 Deposit</b>\n\nUser: <code>{d['user_id']}</code>\n"
        f"Method: {d['method']}\nAmount: ৳{money(d['amount'])}\n"
        f"Reference: {d['reference']}\nDate: {d['date']}",
        reply_markup=kb
    )

@bot.callback_query_handler(func=lambda c: c.data.startswith("delcat:"))
def del_category(call):
    if not is_admin(call.from_user.id):
        return
    cid = call.data.split(":", 1)[1]
    if any(p["category_id"] == cid for p in db["proxies"].values()):
        bot.answer_callback_query(call.id, "Category contains proxies.")
        return
    db["categories"].pop(cid, None)
    save_db(db)
    bot.answer_callback_query(call.id, "Deleted.")

@bot.callback_query_handler(func=lambda c: c.data.startswith("delmeta:"))
def del_meta(call):
    if not is_admin(call.from_user.id):
        return
    mid = call.data.split(":", 1)[1]
    db["meta_products"].pop(mid, None)
    save_db(db)
    bot.answer_callback_query(call.id, "Deleted.")

@bot.callback_query_handler(func=lambda c: c.data.startswith("delscript:"))
def del_script(call):
    if not is_admin(call.from_user.id):
        return
    sid = call.data.split(":", 1)[1]
    s = db["scripts"].pop(sid, None)
    if s and os.path.exists(s.get("file_path", "")):
        try:
            os.remove(s["file_path"])
        except Exception:
            pass
    save_db(db)
    bot.answer_callback_query(call.id, "Deleted.")

# =========================
# START BOT
# =========================
print("Bot starting...")
while True:
    try:
        bot.infinity_polling(timeout=60, long_polling_timeout=60, skip_pending=True)
    except Exception as e:
        print("Polling error:", e)
        time.sleep(5)  # avoid a tight retry loop during temporary outages
