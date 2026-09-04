# -*- coding: utf-8 -*-
import atexit
from datetime import datetime, timedelta
import json
import logging
import mimetypes
import os
import re
import shutil
import signal
import sqlite3
import struct
import subprocess
import sys
import tempfile
import threading
import time
import zipfile
from flask import Flask
from threading import Thread
import psutil
import requests
import telebot
from telebot import types

# --- Hosting Configuration ---
APP_NAME = "Easy Earning Bux Hosting"

# --- Flask Keep Alive ---
app = Flask("")


@app.route("/")
def home():
    return f"{APP_NAME} - Online"


def run_flask():
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)


def keep_alive():
    t = Thread(target=run_flask)
    t.daemon = True
    t.start()
    print("Flask Keep-Alive server started.")


# --- End Flask Keep Alive ---

# --- Configuration ---
TOKEN = os.getenv("BOT_TOKEN", "")
OWNER_ID = int(os.getenv("OWNER_ID", "0"))
ADMIN_ID = int(os.getenv("ADMIN_ID", str(OWNER_ID)))
YOUR_USERNAME = os.getenv("YOUR_USERNAME", "@shiyam7444")
UPDATE_CHANNEL = os.getenv("UPDATE_CHANNEL", "https://t.me/shiyam744")

# --- Payment / Hosting Configuration ---
PAYMENT_INFO_DEFAULT = "Payment instructions are not configured yet. Please contact the admin."

# Folder setup - using absolute paths
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
UPLOAD_BOTS_DIR = os.path.join(BASE_DIR, "upload_bots")
IROTECH_DIR = os.path.join(BASE_DIR, "inf")
DATABASE_PATH = os.path.join(IROTECH_DIR, "bot_data.db")

# File upload limits
FREE_USER_LIMIT = 0  # Default free limit
SUBSCRIBED_USER_LIMIT = 15
ADMIN_LIMIT = 999
OWNER_LIMIT = float("inf")

# Create necessary directories
os.makedirs(UPLOAD_BOTS_DIR, exist_ok=True)
os.makedirs(IROTECH_DIR, exist_ok=True)

# Initialize bot
if not TOKEN:
    raise RuntimeError("BOT_TOKEN is not configured. Add BOT_TOKEN in Railway Variables.")
if OWNER_ID <= 0:
    raise RuntimeError("OWNER_ID is not configured. Add OWNER_ID in Railway Variables.")
bot = telebot.TeleBot(TOKEN)

# --- Data structures ---
bot_scripts = {}
user_subscriptions = {}
user_files = {}
active_users = set()
admin_ids = {ADMIN_ID, OWNER_ID}
bot_locked = False
user_selected_plan = {}  # Temp state for upload flow

# --- Malware Detection Configuration ---
MALWARE_SIGNATURES = [
    b"MZ",  # Windows executable
    b"\x7fELF",  # Linux executable
    b"\xfe\xed\xfa",  # Mach-O binary
    b"\xce\xfa\xed\xfe",  # Mach-O binary (reverse)
    b"PK",  # ZIP archive
    b"Rar!",  # RAR archive
]

ENCRYPTED_FILE_INDICATORS = [
    b"openssl",
    b"encrypted",
    b"cipher",
    b"AES",
    b"DES",
    b"RSA",
    b"GPG",
    b"PGP",
]

SUSPICIOUS_KEYWORDS = [
    b"ransomware",
    b"trojan",
    b"virus",
    b"malware",
    b"backdoor",
    b"exploit",
    b"payload",
    b"botnet",
    b"keylogger",
    b"rootkit",
]

# --- Logging Setup ---
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger(__name__)

# --- Command Button Layouts ---
COMMAND_BUTTONS_LAYOUT_USER_SPEC = [
    ["Updates Channel"],
    ["Upload File", "Manage Files"],
    ["Hosting Plans", "Bot Stats"],
    ["Server Status", "Support"],
]

ADMIN_COMMAND_BUTTONS_LAYOUT_USER_SPEC = [
    ["Updates Channel"],
    ["Upload File", "Manage Files"],
    ["Hosting Plans", "Admin Panel"],
    ["Server Status", "Bot Stats"],
    ["Support"],
]

# --- Database Setup ---
DB_LOCK = threading.Lock()


def init_db():
    """Initialize the database with required tables"""
    logger.info(f"Initializing database at: {DATABASE_PATH}")
    try:
        conn = sqlite3.connect(DATABASE_PATH, check_same_thread=False)
        c = conn.cursor()
        c.execute(
            """CREATE TABLE IF NOT EXISTS subscriptions
                     (user_id INTEGER PRIMARY KEY, plan_name TEXT, expiry TEXT)"""
        )
        c.execute(
            """CREATE TABLE IF NOT EXISTS user_files
                     (user_id INTEGER, file_name TEXT, file_type TEXT,
                      PRIMARY KEY (user_id, file_name))"""
        )
        c.execute(
            """CREATE TABLE IF NOT EXISTS active_users
                     (user_id INTEGER PRIMARY KEY)"""
        )
        c.execute(
            """CREATE TABLE IF NOT EXISTS admins
                     (user_id INTEGER PRIMARY KEY)"""
        )
        c.execute(
            """CREATE TABLE IF NOT EXISTS plans
                     (plan_id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, file_limit INTEGER, price TEXT, duration INTEGER, buy_link TEXT)"""
        )

        c.execute(
            """CREATE TABLE IF NOT EXISTS payment_settings
                     (setting_key TEXT PRIMARY KEY, setting_value TEXT)"""
        )
        c.execute(
            """CREATE TABLE IF NOT EXISTS deposit_requests
                     (request_id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER,
                      plan_id INTEGER, amount TEXT, reference TEXT, status TEXT DEFAULT 'pending',
                      created_at TEXT, reviewed_at TEXT)"""
        )
        c.execute(
            "INSERT OR IGNORE INTO payment_settings (setting_key, setting_value) VALUES (?, ?)",
            ("payment_info", PAYMENT_INFO_DEFAULT),
        )

        c.execute(
            "INSERT OR IGNORE INTO admins (user_id) VALUES (?)", (OWNER_ID,)
        )
        if ADMIN_ID != OWNER_ID:
            c.execute(
                "INSERT OR IGNORE INTO admins (user_id) VALUES (?)", (ADMIN_ID,)
            )

        conn.commit()
        conn.close()
        logger.info("Database initialized successfully.")
    except Exception as e:
        logger.error(f"❌ Database initialization error: {e}", exc_info=True)


def load_data():
    """Load data from database into memory"""
    logger.info("Loading data from database...")
    try:
        conn = sqlite3.connect(DATABASE_PATH, check_same_thread=False)
        c = conn.cursor()

        c.execute("SELECT user_id, plan_name, expiry FROM subscriptions")
        for row in c.fetchall():
            user_id = row[0]
            plan_name = row[1] if len(row) > 2 else "Premium"
            expiry = row[-1]
            try:
                user_subscriptions[user_id] = {
                    "plan_name": plan_name,
                    "expiry": datetime.fromisoformat(expiry),
                }
            except ValueError:
                logger.warning(
                    f"⚠️ Invalid expiry date format for user {user_id}: {expiry}. Skipping."
                )

        c.execute("SELECT user_id, file_name, file_type FROM user_files")
        for user_id, file_name, file_type in c.fetchall():
            if user_id not in user_files:
                user_files[user_id] = []
            user_files[user_id].append((file_name, file_type))

        c.execute("SELECT user_id FROM active_users")
        active_users.update(user_id for (user_id,) in c.fetchall())

        c.execute("SELECT user_id FROM admins")
        admin_ids.update(user_id for (user_id,) in c.fetchall())

        conn.close()
        logger.info(f"Data loaded successfully.")
    except Exception as e:
        logger.error(f"❌ Error loading data: {e}", exc_info=True)


init_db()
load_data()


# --- Payment & Deposit Helpers ---
def get_payment_info():
    conn = sqlite3.connect(DATABASE_PATH, check_same_thread=False)
    c = conn.cursor()
    c.execute("SELECT setting_value FROM payment_settings WHERE setting_key=?", ("payment_info",))
    row = c.fetchone()
    conn.close()
    return row[0] if row else PAYMENT_INFO_DEFAULT


def set_payment_info(info):
    with DB_LOCK:
        conn = sqlite3.connect(DATABASE_PATH, check_same_thread=False)
        c = conn.cursor()
        c.execute("INSERT OR REPLACE INTO payment_settings (setting_key, setting_value) VALUES (?, ?)", ("payment_info", info.strip()))
        conn.commit()
        conn.close()


def create_deposit_request(user_id, plan_id, amount, reference):
    with DB_LOCK:
        conn = sqlite3.connect(DATABASE_PATH, check_same_thread=False)
        c = conn.cursor()
        c.execute(
            "INSERT INTO deposit_requests (user_id, plan_id, amount, reference, status, created_at) VALUES (?, ?, ?, ?, 'pending', ?)",
            (user_id, plan_id, amount.strip(), reference.strip(), datetime.now().isoformat()),
        )
        request_id = c.lastrowid
        conn.commit()
        conn.close()
        return request_id


def get_pending_deposits():
    conn = sqlite3.connect(DATABASE_PATH, check_same_thread=False)
    c = conn.cursor()
    c.execute(
        "SELECT request_id, user_id, plan_id, amount, reference, created_at FROM deposit_requests WHERE status='pending' ORDER BY request_id DESC"
    )
    rows = c.fetchall()
    conn.close()
    return rows


def get_deposit(request_id):
    conn = sqlite3.connect(DATABASE_PATH, check_same_thread=False)
    c = conn.cursor()
    c.execute(
        "SELECT request_id, user_id, plan_id, amount, reference, status, created_at FROM deposit_requests WHERE request_id=?",
        (request_id,),
    )
    row = c.fetchone()
    conn.close()
    return row


def review_deposit(request_id, status):
    with DB_LOCK:
        conn = sqlite3.connect(DATABASE_PATH, check_same_thread=False)
        c = conn.cursor()
        c.execute(
            "UPDATE deposit_requests SET status=?, reviewed_at=? WHERE request_id=? AND status='pending'",
            (status, datetime.now().isoformat(), request_id),
        )
        changed = c.rowcount > 0
        conn.commit()
        conn.close()
        return changed


# --- Malware Detection Functions ---
def is_suspicious_file(file_content, file_name):
    file_lower = file_name.lower()
    suspicious_extensions = [
        ".exe",
        ".dll",
        ".bat",
        ".cmd",
        ".scr",
        ".com",
        ".pif",
        ".application",
        ".gadget",
        ".msi",
        ".msp",
        ".com",
        ".scr",
        ".hta",
        ".cpl",
        ".msc",
        ".jar",
        ".bin",
        ".deb",
        ".rpm",
        ".apk",
        ".app",
        ".dmg",
        ".iso",
        ".img",
    ]
    if any(file_lower.endswith(ext) for ext in suspicious_extensions):
        return True, f"Suspicious file extension: {file_name}"
    for signature in MALWARE_SIGNATURES:
        if file_content.startswith(signature):
            return True, f"Malware signature detected: {signature}"
    sample_size = min(len(file_content), 4096)
    file_sample = file_content[:sample_size]
    for indicator in ENCRYPTED_FILE_INDICATORS:
        if indicator in file_sample:
            return (
                True,
                f"Encrypted file indicator: {indicator.decode('utf-8', errors='ignore')}",
            )
    sample_text = file_sample.decode("utf-8", errors="ignore").lower()
    for keyword in SUSPICIOUS_KEYWORDS:
        if keyword.decode("utf-8").lower() in sample_text:
            return True, f"Suspicious keyword found: {keyword.decode('utf-8')}"
    return False, "File appears safe"


def scan_file_for_malware(file_content, file_name, user_id):
    if user_id == OWNER_ID:
        return True, "Owner bypassed security check"
    is_suspicious, reason = is_suspicious_file(file_content, file_name)
    if is_suspicious:
        logger.warning(
            f"🚨 Malware detected in {file_name} from user {user_id}: {reason}"
        )
        return False, f"Security violation: {reason}"
    return True, "File passed security check"


# --- Helper Functions ---
def get_user_folder(user_id):
    user_folder = os.path.join(UPLOAD_BOTS_DIR, str(user_id))
    os.makedirs(user_folder, exist_ok=True)
    return user_folder


def get_user_file_limit(user_id):
    if user_id == OWNER_ID:
        return OWNER_LIMIT
    if user_id in admin_ids:
        return ADMIN_LIMIT
    if (
        user_id in user_subscriptions
        and user_subscriptions[user_id]["expiry"] > datetime.now()
    ):
        return SUBSCRIBED_USER_LIMIT
    return FREE_USER_LIMIT


def get_user_file_count(user_id):
    return len(user_files.get(user_id, []))


def is_bot_running(script_owner_id, file_name):
    script_key = f"{script_owner_id}_{file_name}"
    script_info = bot_scripts.get(script_key)
    if script_info and script_info.get("process"):
        try:
            proc = psutil.Process(script_info["process"].pid)
            is_running = (
                proc.is_running() and proc.status() != psutil.STATUS_ZOMBIE
            )
            if not is_running:
                if (
                    "log_file" in script_info
                    and hasattr(script_info["log_file"], "close")
                    and not script_info["log_file"].closed
                ):
                    try:
                        script_info["log_file"].close()
                    except Exception:
                        pass
                if script_key in bot_scripts:
                    del bot_scripts[script_key]
            return is_running
        except psutil.NoSuchProcess:
            if script_key in bot_scripts:
                del bot_scripts[script_key]
            return False
        except Exception:
            return False
    return False


def kill_process_tree(process_info):
    try:
        if (
            "log_file" in process_info
            and hasattr(process_info["log_file"], "close")
            and not process_info["log_file"].closed
        ):
            try:
                process_info["log_file"].close()
            except Exception:
                pass
        process = process_info.get("process")
        if process and hasattr(process, "pid"):
            pid = process.pid
            if pid:
                parent = psutil.Process(pid)
                for child in parent.children(recursive=True):
                    try:
                        child.terminate()
                    except Exception:
                        pass
                try:
                    parent.terminate()
                except Exception:
                    pass
    except Exception as e:
        logger.error(f"❌ Error killing process: {e}")


# --- Module / Package Mapping ---
TELEGRAM_MODULES = {
    "telebot": "pyTelegramBotAPI",
    "telegram": "python-telegram-bot",
    "python_telegram_bot": "python-telegram-bot",
    "aiogram": "aiogram",
    "pyrogram": "pyrogram",
    "telethon": "telethon",
    "bs4": "beautifulsoup4",
    "requests": "requests",
    "pillow": "Pillow",
    "cv2": "opencv-python",
    "flask": "Flask",
    "psutil": "psutil",
}


# --- Automatic & Guided Script Running ---
def monitor_and_guide_error(
    process, log_file_path, script_owner_id, file_name, message_obj_for_reply
):
    """রানিং স্ক্রিপ্ট ব্যাকগ্রাউন্ডে চেক করে কোনো এরর থাকলে ইউজারকে বাটন দিয়ে বুঝিয়ে দেবে"""
    time.sleep(3)
    if process.poll() is not None:
        try:
            with open(log_file_path, "r", encoding="utf-8", errors="ignore") as f:
                log_content = f.read()

            match_py = re.search(
                r"(?:ModuleNotFoundError|ImportError): No module named '(.+?)'",
                log_content,
            )
            match_js = re.search(r"Cannot find module '(.+?)'", log_content)

            missing_module = None
            if match_py:
                missing_module = match_py.group(1).split(".")[0].strip("'\"")
            elif match_js:
                missing_module = match_js.group(1).split("/")[0].strip("'\"")

            if missing_module:
                pkg_name = TELEGRAM_MODULES.get(
                    missing_module.lower(), missing_module
                )
                ext = os.path.splitext(file_name)[1].lower()
                cmd_text = (
                    f"npm install {pkg_name}"
                    if ext == ".js"
                    else f"pip install {pkg_name}"
                )

                error_msg = (
                    f"⚠️ **ফাইল রান হতে সমস্যা হয়েছে!**\n\n"
                    f"📄 **File:** `{file_name}`\n"
                    f"❌ **সমস্যা:** আপনার কোডে `{missing_module}` মডিউলটি মিসিং আছে।\n"
                    f"💻 **প্রয়োজনীয় কমান্ড:** `{cmd_text}`\n\n"
                    f"👇 *নিচের বাটনে প্রেস করে সরাসরি মডিউলটি ইনস্টল করুন:*"
                )

                markup = types.InlineKeyboardMarkup()
                markup.add(
                    types.InlineKeyboardButton(
                        f"📦 Install {pkg_name}",
                        callback_data=f"instmod_{script_owner_id}_{missing_module}_{file_name}",
                    )
                )
                markup.add(
                    types.InlineKeyboardButton(
                        "📄 View Error Logs",
                        callback_data=f"viewlog_{script_owner_id}_{file_name}",
                    )
                )

                bot.reply_to(
                    message_obj_for_reply,
                    error_msg,
                    reply_markup=markup,
                    parse_mode="Markdown",
                )
            else:
                error_msg = (
                    f"⚠️ **আপনার কোডে ভুল (Syntax/Runtime Error) পাওয়া গেছে!**\n\n"
                    f"📄 **File:** `{file_name}`\n"
                    f"সুনির্দিষ্ট এরর জানতে নিচের **View Logs** বাটনে ক্লিক করুন।"
                )
                markup = types.InlineKeyboardMarkup()
                markup.add(
                    types.InlineKeyboardButton(
                        "📄 View Error Logs",
                        callback_data=f"viewlog_{script_owner_id}_{file_name}",
                    )
                )
                bot.reply_to(
                    message_obj_for_reply,
                    error_msg,
                    reply_markup=markup,
                    parse_mode="Markdown",
                )
        except Exception as e:
            logger.error(f"Error checking log file: {e}")


def run_script(
    script_path, script_owner_id, user_folder, file_name, message_obj_for_reply
):
    script_key = f"{script_owner_id}_{file_name}"
    try:
        log_file_path = os.path.join(
            user_folder, f"{os.path.splitext(file_name)[0]}.log"
        )
        log_file = open(log_file_path, "w", encoding="utf-8", errors="ignore")
        process = subprocess.Popen(
            [sys.executable, script_path],
            cwd=user_folder,
            stdout=log_file,
            stderr=log_file,
            stdin=subprocess.PIPE,
        )

        bot_scripts[script_key] = {
            "process": process,
            "log_file": log_file,
            "file_name": file_name,
            "script_owner_id": script_owner_id,
            "start_time": datetime.now(),
            "user_folder": user_folder,
            "type": "py",
            "script_key": script_key,
        }

        bot.reply_to(
            message_obj_for_reply,
            f"🚀 **Python Script Started!**\n📄 File: `{file_name}`\n🆔 PID: `{process.pid}`",
            parse_mode="Markdown",
        )

        threading.Thread(
            target=monitor_and_guide_error,
            args=(
                process,
                log_file_path,
                script_owner_id,
                file_name,
                message_obj_for_reply,
            ),
        ).start()

    except Exception as e:
        bot.reply_to(message_obj_for_reply, f"❌ Error running script: {str(e)}")


def run_js_script(
    script_path, script_owner_id, user_folder, file_name, message_obj_for_reply
):
    script_key = f"{script_owner_id}_{file_name}"
    try:
        log_file_path = os.path.join(
            user_folder, f"{os.path.splitext(file_name)[0]}.log"
        )
        log_file = open(log_file_path, "w", encoding="utf-8", errors="ignore")
        process = subprocess.Popen(
            ["node", script_path],
            cwd=user_folder,
            stdout=log_file,
            stderr=log_file,
            stdin=subprocess.PIPE,
        )

        bot_scripts[script_key] = {
            "process": process,
            "log_file": log_file,
            "file_name": file_name,
            "script_owner_id": script_owner_id,
            "start_time": datetime.now(),
            "user_folder": user_folder,
            "type": "js",
            "script_key": script_key,
        }

        bot.reply_to(
            message_obj_for_reply,
            f"🚀 **JS Script Started!**\n📄 File: `{file_name}`\n🆔 PID: `{process.pid}`",
            parse_mode="Markdown",
        )

        threading.Thread(
            target=monitor_and_guide_error,
            args=(
                process,
                log_file_path,
                script_owner_id,
                file_name,
                message_obj_for_reply,
            ),
        ).start()

    except Exception as e:
        bot.reply_to(
            message_obj_for_reply, f"❌ Error running JS script: {str(e)}"
        )


# --- Database Operations ---
def save_user_file(user_id, file_name, file_type="py"):
    with DB_LOCK:
        conn = sqlite3.connect(DATABASE_PATH, check_same_thread=False)
        c = conn.cursor()
        c.execute(
            "INSERT OR REPLACE INTO user_files (user_id, file_name, file_type) VALUES (?, ?, ?)",
            (user_id, file_name, file_type),
        )
        conn.commit()
        conn.close()
        if user_id not in user_files:
            user_files[user_id] = []
        user_files[user_id] = [
            (fn, ft) for fn, ft in user_files[user_id] if fn != file_name
        ]
        user_files[user_id].append((file_name, file_type))


def remove_user_file_db(user_id, file_name):
    with DB_LOCK:
        conn = sqlite3.connect(DATABASE_PATH, check_same_thread=False)
        c = conn.cursor()
        c.execute(
            "DELETE FROM user_files WHERE user_id = ? AND file_name = ?",
            (user_id, file_name),
        )
        conn.commit()
        conn.close()
        if user_id in user_files:
            user_files[user_id] = [
                f for f in user_files[user_id] if f[0] != file_name
            ]


def add_active_user(user_id):
    active_users.add(user_id)
    with DB_LOCK:
        conn = sqlite3.connect(DATABASE_PATH, check_same_thread=False)
        c = conn.cursor()
        c.execute(
            "INSERT OR IGNORE INTO active_users (user_id) VALUES (?)", (user_id,)
        )
        conn.commit()
        conn.close()


def save_subscription(user_id, plan_name, expiry):
    with DB_LOCK:
        conn = sqlite3.connect(DATABASE_PATH, check_same_thread=False)
        c = conn.cursor()
        c.execute(
            "INSERT OR REPLACE INTO subscriptions (user_id, plan_name, expiry) VALUES (?, ?, ?)",
            (user_id, plan_name, expiry.isoformat()),
        )
        conn.commit()
        conn.close()
        user_subscriptions[user_id] = {"plan_name": plan_name, "expiry": expiry}


def remove_subscription_db(user_id):
    with DB_LOCK:
        conn = sqlite3.connect(DATABASE_PATH, check_same_thread=False)
        c = conn.cursor()
        c.execute("DELETE FROM subscriptions WHERE user_id = ?", (user_id,))
        conn.commit()
        conn.close()
        if user_id in user_subscriptions:
            del user_subscriptions[user_id]


# --- Menu Creation ---
def create_reply_keyboard_main_menu(user_id):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)
    layout_to_use = (
        ADMIN_COMMAND_BUTTONS_LAYOUT_USER_SPEC
        if user_id in admin_ids
        else COMMAND_BUTTONS_LAYOUT_USER_SPEC
    )
    for row in layout_to_use:
        markup.add(*[types.KeyboardButton(text) for text in row])
    return markup


def create_admin_panel_inline():
    markup = types.InlineKeyboardMarkup(row_width=2)
    markup.add(
        types.InlineKeyboardButton("Add Plan", callback_data="add_plan_init"),
        types.InlineKeyboardButton("Manage Plans", callback_data="manage_plans"),
    )
    markup.add(
        types.InlineKeyboardButton("Add Subscription", callback_data="add_subscription"),
        types.InlineKeyboardButton("Remove Subscription", callback_data="remove_subscription"),
    )
    markup.add(
        types.InlineKeyboardButton("Add Admin", callback_data="add_admin"),
        types.InlineKeyboardButton("Remove Admin", callback_data="remove_admin"),
    )
    markup.add(
        types.InlineKeyboardButton("Pending Deposits", callback_data="pending_deposits"),
        types.InlineKeyboardButton("Payment Settings", callback_data="payment_settings"),
    )
    markup.add(
        types.InlineKeyboardButton("Database", callback_data="database_menu"),
        types.InlineKeyboardButton("Broadcast", callback_data="broadcast"),
    )
    markup.add(
        types.InlineKeyboardButton("Lock / Unlock", callback_data="toggle_lock"),
        types.InlineKeyboardButton("Run All Scripts", callback_data="run_all_scripts"),
    )
    markup.add(types.InlineKeyboardButton("Bot Stats", callback_data="stats"))
    return markup


# --- Core User Logic ---
def _logic_send_welcome(message):
    user_id = message.from_user.id
    chat_id = message.chat.id
    user_name = message.from_user.first_name

    if bot_locked and user_id not in admin_ids:
        bot.send_message(chat_id, "⚠️ **Bot is temporarily locked by Admin.**")
        return

    if user_id not in active_users:
        add_active_user(user_id)

    if user_id == OWNER_ID:
        user_status = "👑 **Owner**"
    elif user_id in admin_ids:
        user_status = "🛡️ **Admin**"
    elif (
        user_id in user_subscriptions
        and user_subscriptions[user_id]["expiry"] > datetime.now()
    ):
        sub = user_subscriptions[user_id]
        days_left = (sub["expiry"] - datetime.now()).days
        user_status = f"💎 **{sub.get('plan_name', 'Premium')} Active** ({days_left} Days left)"
    else:
        user_status = "🆓 **No Active Plan**"

    welcome_msg = (
        f"✨ **𝗪𝗲𝗹𝗰𝗼𝗺𝗲, {user_name}!** ✨\n\n"
        f"🆔 **𝗬𝗼𝘂𝗿 𝗜𝗗:** `{user_id}`\n"
        f"🔰 **𝗦𝘁𝗮𝘁𝘂𝘀:** {user_status}\n"
        f"📁 **𝗨𝗽𝗹𝗼𝗮𝗱𝗲𝗱 𝗙𝗶𝗹𝗲𝘀:** `{get_user_file_count(user_id)}` / `{get_user_file_limit(user_id)}`\n\n"
        f"💡 **𝗛𝗼𝘀𝘁 & 𝗥𝘂𝗻 𝘆𝗼𝘂𝗿 𝗣𝘆𝘁𝗵𝗼𝗻 (.𝗽𝘆) & 𝗝𝗦 (.𝗷𝘀) 𝗯𝗼𝘁𝘀 𝟮𝟰/𝟳.**\n"
        f"👇 *Select an option from the menu below:* "
    )
    bot.send_message(
        chat_id,
        welcome_msg,
        reply_markup=create_reply_keyboard_main_menu(user_id),
        parse_mode="Markdown",
    )


def _logic_view_plans(message_or_call):
    chat_id = (
        message_or_call.chat.id
        if isinstance(message_or_call, telebot.types.Message)
        else message_or_call.message.chat.id
    )
    plans = get_all_plans()
    if not plans:
        bot.send_message(chat_id, "No hosting plans are available right now.")
        return

    header = (
        "HOSTING PLANS\n"
        "━━━━━━━━━━━━━━━━━━━━\n"
        "Choose a plan that matches your bot workload.\n"
    )
    bot.send_message(chat_id, header)
    for plan_id, name, limit, price, duration, buy_link in plans:
        card = (
            f"PLAN: {name}\n"
            f"━━━━━━━━━━━━━━━━━━━━\n"
            f"File Limit: {limit}\n"
            f"Duration: {duration} days\n"
            f"Price: {price}\n"
            f"Payment: Manual review\n"
        )
        markup = types.InlineKeyboardMarkup(row_width=1)
        markup.add(types.InlineKeyboardButton(f"Buy {name}", callback_data=f"buy_plan_{plan_id}"))
        if buy_link and str(buy_link).strip():
            markup.add(types.InlineKeyboardButton("Payment / Contact Link", url=buy_link.strip()))
        bot.send_message(chat_id, card, reply_markup=markup)

def _logic_upload_file(message):
    user_id = message.from_user.id
    if bot_locked and user_id not in admin_ids:
        bot.reply_to(message, "⚠️ **Bot is locked by Admin.**")
        return

    has_active_plan = False
    plan_name = "None"

    if user_id in admin_ids or user_id == OWNER_ID:
        has_active_plan = True
        plan_name = "Admin / Owner Unlimited"
    elif user_id in user_subscriptions:
        sub = user_subscriptions[user_id]
        if sub["expiry"] > datetime.now():
            has_active_plan = True
            plan_name = sub.get("plan_name", "Premium Plan")

    if not has_active_plan:
        markup = types.InlineKeyboardMarkup()
        markup.add(
            types.InlineKeyboardButton(
                "💳 View Plans & Buy", callback_data="view_plans_cb"
            )
        )
        bot.reply_to(
            message,
            "❌ **আপনার কোন এক্টিভ প্ল্যান নেই!**\n\n"
            "ফাইল আপলোড করতে হলে প্রথমে একটি প্ল্যান সাবস্ক্রাইব করতে হবে। "
            "নিচের বাটনে ক্লিক করে আমাদের প্ল্যানগুলো দেখুন এবং আপনার পছন্দমতো প্ল্যান কিনুন।",
            reply_markup=markup,
            parse_mode="Markdown",
        )
        return

    markup = types.InlineKeyboardMarkup()
    markup.add(
        types.InlineKeyboardButton(
            f"✅ Continue with {plan_name}", callback_data="confirm_plan_upload"
        )
    )
    bot.reply_to(
        message,
        f"🔰 **𝗔𝗰𝘁𝗶𝘃𝗲 𝗣𝗹𝗮𝗻 𝗗𝗲𝘁𝗲𝗰𝘁𝗲𝗱:** `{plan_name}`\n\n"
        f"ফাইল আপলোড চালু করতে নিচের বাটনে সিলেক্ট করুন:",
        reply_markup=markup,
        parse_mode="Markdown",
    )


def _logic_check_files(message):
    user_id = message.from_user.id
    user_files_list = user_files.get(user_id, [])
    if not user_files_list:
        bot.reply_to(
            message,
            "📂 **Your Uploaded Files:**\n\n*(No files uploaded yet)*",
            parse_mode="Markdown",
        )
        return
    markup = types.InlineKeyboardMarkup(row_width=1)
    for file_name, file_type in sorted(user_files_list):
        is_running = is_bot_running(user_id, file_name)
        status_icon = "🟢 Running" if is_running else "🔴 Stopped"
        btn_text = f"📄 {file_name} ({file_type}) - {status_icon}"
        markup.add(
            types.InlineKeyboardButton(
                btn_text, callback_data=f"file_{user_id}_{file_name}"
            )
        )
    bot.reply_to(
        message,
        "📁 **𝗠𝗮𝗻𝗮𝗴𝗲 𝗬𝗼𝘂𝗿 𝗙𝗶𝗹𝗲𝘀:**",
        reply_markup=markup,
        parse_mode="Markdown",
    )


# --- Document Upload Processing ---
@bot.message_handler(content_types=["document"])
def handle_file_upload_doc(message):
    user_id = message.from_user.id
    chat_id = message.chat.id
    doc = message.document

    if user_id not in admin_ids and user_id != OWNER_ID:
        if (
            user_id not in user_subscriptions
            or user_subscriptions[user_id]["expiry"] <= datetime.now()
        ):
            bot.reply_to(
                message,
                "❌ **আপনার কোন এক্টিভ প্ল্যান নেই! ফাইল আপলোড করতে প্ল্যান ক্রয় করুন।**",
                parse_mode="Markdown",
            )
            return

    file_name = os.path.basename(doc.file_name)
    file_ext = os.path.splitext(file_name)[1].lower()
    if file_ext not in [".py", ".js", ".zip"]:
        bot.reply_to(
            message,
            "⚠️ **Only `.py`, `.js`, and `.zip` files are supported!**",
            parse_mode="Markdown",
        )
        return

    try:
        download_wait_msg = bot.reply_to(
            message,
            f"⏳ **Downloading `{file_name}`...**",
            parse_mode="Markdown",
        )
        file_info_tg_doc = bot.get_file(doc.file_id)
        downloaded_file_content = bot.download_file(file_info_tg_doc.file_path)

        if user_id != OWNER_ID:
            is_safe, reason = scan_file_for_malware(
                downloaded_file_content, file_name, user_id
            )
            if not is_safe:
                bot.edit_message_text(
                    f"🚨 **Security Alert:** {reason}",
                    chat_id,
                    download_wait_msg.message_id,
                    parse_mode="Markdown",
                )
                return

        user_folder = get_user_folder(user_id)
        file_path = os.path.join(user_folder, file_name)
        with open(file_path, "wb") as f:
            f.write(downloaded_file_content)

        bot.edit_message_text(
            f"✅ **File `{file_name}` uploaded successfully!**",
            chat_id,
            download_wait_msg.message_id,
            parse_mode="Markdown",
        )

        if file_ext == ".js":
            save_user_file(user_id, file_name, "js")
            threading.Thread(
                target=run_js_script,
                args=(file_path, user_id, user_folder, file_name, message),
            ).start()
        elif file_ext == ".py":
            save_user_file(user_id, file_name, "py")
            threading.Thread(
                target=run_script,
                args=(file_path, user_id, user_folder, file_name, message),
            ).start()

    except Exception as e:
        bot.reply_to(message, f"❌ **Error:** {str(e)}")


# --- Callback Routing ---
@bot.callback_query_handler(func=lambda call: True)
def handle_callbacks(call):
    user_id = call.from_user.id
    data = call.data

    if data == "view_plans_cb":
        bot.answer_callback_query(call.id)
        _logic_view_plans(call)

    elif data == "confirm_plan_upload":
        bot.answer_callback_query(call.id, "✅ Plan Verified!")
        bot.send_message(
            call.message.chat.id,
            "🚀 **এখন আপনার Python (.py), JS (.js) অথবা ZIP (.zip) ফাইল মেসেজে পাঠান।**",
            parse_mode="Markdown",
        )

    # --- Interactive Module Installer Handler ---
    elif data.startswith("instmod_"):
        _, owner_id, mod_name, fname = data.split("_", 3)
        if user_id != int(owner_id) and user_id not in admin_ids:
            bot.answer_callback_query(
                call.id,
                "❌ আপনি অন্য ইউজারের ফাইল কাস্টমাইজ করতে পারবেন না!",
                show_alert=True,
            )
            return

        bot.answer_callback_query(call.id)
        pkg_name = TELEGRAM_MODULES.get(mod_name.lower(), mod_name)
        ext = os.path.splitext(fname)[1].lower()

        status_msg = bot.send_message(
            call.message.chat.id,
            f"⏳ **`{pkg_name}` মডিউলটি ইনস্টল করা হচ্ছে...**",
            parse_mode="Markdown",
        )

        def do_pip_install():
            if ext == ".js":
                cmd = ["npm", "install", pkg_name]
            else:
                cmd = [sys.executable, "-m", "pip", "install", pkg_name]

            res = subprocess.run(cmd, capture_output=True, text=True)
            if res.returncode == 0:
                bot.edit_message_text(
                    f"✅ **`{pkg_name}` মডিউলটি সফলভাবে ইনস্টল হয়েছে!**\n🚀 ফাইলটি পুনরায় চালু করা হচ্ছে...",
                    call.message.chat.id,
                    status_msg.message_id,
                    parse_mode="Markdown",
                )
                time.sleep(1)
                ufolder = get_user_folder(int(owner_id))
                fpath = os.path.join(ufolder, fname)
                if ext == ".js":
                    run_js_script(
                        fpath, int(owner_id), ufolder, fname, call.message
                    )
                else:
                    run_script(
                        fpath, int(owner_id), ufolder, fname, call.message
                    )
            else:
                bot.edit_message_text(
                    f"❌ **ইনস্টলেশন ব্যর্থ হয়েছে!**\n\n```\n{res.stderr[:300]}\n```",
                    call.message.chat.id,
                    status_msg.message_id,
                    parse_mode="Markdown",
                )

        threading.Thread(target=do_pip_install).start()

    # --- Error Log Viewer Handler ---
    elif data.startswith("viewlog_"):
        _, owner_id, fname = data.split("_", 2)
        ufolder = get_user_folder(int(owner_id))
        log_fpath = os.path.join(
            ufolder, f"{os.path.splitext(fname)[0]}.log"
        )
        if os.path.exists(log_fpath):
            with open(log_fpath, "r", encoding="utf-8", errors="ignore") as f:
                logs = f.read()[-2000:]
            bot.send_message(
                call.message.chat.id,
                f"📜 **Error Log for `{fname}`:**\n\n```\n{logs if logs else 'No logs recorded.'}\n```",
                parse_mode="Markdown",
            )
        else:
            bot.answer_callback_query(
                call.id, "No log file found!", show_alert=True
            )

    # --- Manual Deposit Workflow ---
    elif data.startswith("buy_plan_"):
        plan_id = int(data.split("_")[2])
        plan = get_plan_by_id(plan_id)
        if not plan:
            bot.answer_callback_query(call.id, "Plan not found.", show_alert=True)
            return
        _, name, limit, price, duration, buy_link = plan
        bot.answer_callback_query(call.id)
        msg = (
            f"PLAN: {name}\n"
            f"━━━━━━━━━━━━━━━━━━━━\n"
            f"File Limit: {limit}\n"
            f"Duration: {duration} days\n"
            f"Price: {price}\n\n"
            f"PAYMENT INSTRUCTIONS\n"
            f"━━━━━━━━━━━━━━━━━━━━\n"
            f"{get_payment_info()}\n\n"
            "After payment, submit the amount and transaction/reference ID. "
            "An admin will review it before your plan is activated."
        )
        markup = types.InlineKeyboardMarkup(row_width=1)
        markup.add(types.InlineKeyboardButton("Submit Deposit", callback_data=f"submit_deposit_{plan_id}"))
        markup.add(types.InlineKeyboardButton("Back to Plans", callback_data="view_plans_cb"))
        bot.send_message(call.message.chat.id, msg, reply_markup=markup)

    elif data.startswith("submit_deposit_"):
        plan_id = int(data.split("_")[2])
        if not get_plan_by_id(plan_id):
            bot.answer_callback_query(call.id, "Plan not found.", show_alert=True)
            return
        bot.answer_callback_query(call.id)
        msg = bot.send_message(
            call.message.chat.id,
            "Send deposit details in this format:\n\nAMOUNT | TRANSACTION/REFERENCE ID\n\nExample: 500 BDT | TX123456789\n\nYour deposit will stay pending until an admin approves it.",
        )
        bot.register_next_step_handler(msg, lambda m: process_deposit_submission(m, plan_id))

    # --- Admin Callbacks ---
    elif data == "add_plan_init" and user_id in admin_ids:
        bot.answer_callback_query(call.id)
        msg = bot.send_message(
            call.message.chat.id,
            "Enter Plan Details:\nName | FileLimit | Price | DurationInDays | Payment/ContactLink\n\nExample: Basic | 5 | 500 BDT | 30 | https://t.me/example",
        )
        bot.register_next_step_handler(msg, process_add_plan)

    elif data == "manage_plans" and user_id in admin_ids:
        bot.answer_callback_query(call.id)
        plans = get_all_plans()
        if not plans:
            bot.send_message(call.message.chat.id, "No plans found.")
            return
        markup = types.InlineKeyboardMarkup(row_width=1)
        for p in plans:
            markup.add(types.InlineKeyboardButton(f"Delete {p[1]}", callback_data=f"del_plan_{p[0]}"))
        bot.send_message(call.message.chat.id, "Select a plan to delete:", reply_markup=markup)

    elif data.startswith("del_plan_") and user_id in admin_ids:
        pid = int(data.split("_")[2])
        delete_plan_db(pid)
        bot.answer_callback_query(call.id, "Plan deleted.")
        bot.send_message(call.message.chat.id, "Plan deleted successfully.")

    elif data == "add_subscription" and user_id in admin_ids:
        bot.answer_callback_query(call.id)
        msg = bot.send_message(call.message.chat.id, "Enter: UserID PlanName Days\nExample: 123456789 VIP 30")
        bot.register_next_step_handler(msg, process_add_subscription)

    elif data == "remove_subscription" and user_id in admin_ids:
        bot.answer_callback_query(call.id)
        msg = bot.send_message(call.message.chat.id, "Send the User ID whose subscription should be removed.")
        bot.register_next_step_handler(msg, process_remove_subscription)

    elif data == "add_admin" and user_id == OWNER_ID:
        bot.answer_callback_query(call.id)
        msg = bot.send_message(call.message.chat.id, "Send the Telegram User ID to add as admin.")
        bot.register_next_step_handler(msg, process_add_admin)

    elif data == "remove_admin" and user_id == OWNER_ID:
        bot.answer_callback_query(call.id)
        msg = bot.send_message(call.message.chat.id, "Send the Telegram User ID to remove from admin.")
        bot.register_next_step_handler(msg, process_remove_admin)

    elif data == "payment_settings" and user_id in admin_ids:
        bot.answer_callback_query(call.id)
        msg = bot.send_message(
            call.message.chat.id,
            "Current payment instructions:\n\n" + get_payment_info() + "\n\nSend new payment instructions as one message. You can include bKash, Nagad, Rocket, bank details, or a payment link.",
        )
        bot.register_next_step_handler(msg, process_payment_settings)

    elif data == "pending_deposits" and user_id in admin_ids:
        bot.answer_callback_query(call.id)
        rows = get_pending_deposits()
        if not rows:
            bot.send_message(call.message.chat.id, "No pending deposits.")
            return
        for rid, uid, pid, amount, reference, created_at in rows[:30]:
            plan = get_plan_by_id(pid)
            pname = plan[1] if plan else "Deleted plan"
            text = (
                f"DEPOSIT #{rid}\n"
                f"User: {uid}\n"
                f"Plan: {pname}\n"
                f"Amount: {amount}\n"
                f"Reference: {reference}\n"
                f"Submitted: {created_at[:19]}"
            )
            markup = types.InlineKeyboardMarkup(row_width=2)
            markup.add(
                types.InlineKeyboardButton("Approve", callback_data=f"approve_deposit_{rid}"),
                types.InlineKeyboardButton("Reject", callback_data=f"reject_deposit_{rid}"),
            )
            bot.send_message(call.message.chat.id, text, reply_markup=markup)

    elif data.startswith("approve_deposit_") and user_id in admin_ids:
        rid = int(data.split("_")[2])
        row = get_deposit(rid)
        if not row or row[5] != "pending":
            bot.answer_callback_query(call.id, "Already reviewed or not found.", show_alert=True)
            return
        _, uid, pid, amount, reference, _, _ = row
        plan = get_plan_by_id(pid)
        if not plan:
            bot.answer_callback_query(call.id, "Plan no longer exists.", show_alert=True)
            return
        _, pname, _, _, duration, _ = plan
        current = user_subscriptions.get(uid)
        base = current["expiry"] if current and current["expiry"] > datetime.now() else datetime.now()
        expiry = base + timedelta(days=duration)
        if review_deposit(rid, "approved"):
            save_subscription(uid, pname, expiry)
            bot.answer_callback_query(call.id, "Deposit approved.")
            bot.edit_message_reply_markup(call.message.chat.id, call.message.message_id, reply_markup=None)
            bot.send_message(call.message.chat.id, f"Deposit #{rid} approved. User {uid} is now subscribed to {pname} until {expiry.strftime('%Y-%m-%d %H:%M')}.")
            try:
                bot.send_message(uid, f"Your deposit for {pname} has been approved. Your hosting plan is now active until {expiry.strftime('%Y-%m-%d %H:%M')}.")
            except Exception:
                pass

    elif data.startswith("reject_deposit_") and user_id in admin_ids:
        rid = int(data.split("_")[2])
        row = get_deposit(rid)
        if not row or row[5] != "pending":
            bot.answer_callback_query(call.id, "Already reviewed or not found.", show_alert=True)
            return
        if review_deposit(rid, "rejected"):
            bot.answer_callback_query(call.id, "Deposit rejected.")
            bot.edit_message_reply_markup(call.message.chat.id, call.message.message_id, reply_markup=None)
            bot.send_message(call.message.chat.id, f"Deposit #{rid} rejected.")
            try:
                bot.send_message(row[1], "Your deposit request was rejected. Please contact an admin if you believe this was a mistake.")
            except Exception:
                pass

    elif data == "database_menu" and user_id in admin_ids:
        bot.answer_callback_query(call.id)
        markup = types.InlineKeyboardMarkup(row_width=1)
        markup.add(types.InlineKeyboardButton("Download Database", callback_data="db_download"))
        markup.add(types.InlineKeyboardButton("Upload Database", callback_data="db_upload"))
        markup.add(types.InlineKeyboardButton("Clear All Data", callback_data="db_clear_confirm"))
        bot.send_message(call.message.chat.id, "DATABASE MANAGEMENT\n━━━━━━━━━━━━━━━━━━━━\nBackup, restore, or clear the bot database.", reply_markup=markup)

    elif data == "db_download" and user_id in admin_ids:
        bot.answer_callback_query(call.id)
        try:
            bot.send_document(call.message.chat.id, open(DATABASE_PATH, "rb"), caption="Database backup")
        except Exception as e:
            bot.send_message(call.message.chat.id, f"Database backup failed: {e}")

    elif data == "db_upload" and user_id in admin_ids:
        bot.answer_callback_query(call.id)
        msg = bot.send_message(call.message.chat.id, "Send a valid SQLite database file as a document. The current database will be backed up before replacement.")
        bot.register_next_step_handler(msg, process_database_upload)

    elif data == "db_clear_confirm" and user_id == OWNER_ID:
        bot.answer_callback_query(call.id)
        markup = types.InlineKeyboardMarkup(row_width=2)
        markup.add(types.InlineKeyboardButton("Yes, Clear Everything", callback_data="db_clear_execute"), types.InlineKeyboardButton("Cancel", callback_data="database_menu"))
        bot.send_message(call.message.chat.id, "WARNING: This removes user data, subscriptions, plans, deposit history, active-user data, and uploaded bot files. Admin accounts are preserved. Continue?", reply_markup=markup)

    elif data == "db_clear_execute" and user_id == OWNER_ID:
        bot.answer_callback_query(call.id)
        clear_all_data()
        bot.send_message(call.message.chat.id, "Database and uploaded bot data have been cleared. Admin accounts were preserved.")

    # --- File Management Callbacks ---
    elif data.startswith("file_"):
        _, owner_id, fname = data.split("_", 2)
        is_running = is_bot_running(int(owner_id), fname)
        markup = types.InlineKeyboardMarkup(row_width=2)
        if is_running:
            markup.add(
                types.InlineKeyboardButton(
                    "🛑 Stop", callback_data=f"stop_{owner_id}_{fname}"
                )
            )
        else:
            markup.add(
                types.InlineKeyboardButton(
                    "▶️ Start", callback_data=f"start_{owner_id}_{fname}"
                )
            )
        markup.add(
            types.InlineKeyboardButton(
                "🗑️ Delete", callback_data=f"del_{owner_id}_{fname}"
            )
        )
        bot.send_message(
            call.message.chat.id,
            f"📄 **File:** `{fname}`\n🚦 Status: `{'Running' if is_running else 'Stopped'}`",
            reply_markup=markup,
            parse_mode="Markdown",
        )

    elif data.startswith("start_"):
        _, owner_id, fname = data.split("_", 2)
        owner_id = int(owner_id)
        ufolder = get_user_folder(owner_id)
        fpath = os.path.join(ufolder, fname)
        if not os.path.exists(fpath):
            bot.answer_callback_query(call.id, "File not found.", show_alert=True)
            return
        bot.answer_callback_query(call.id, "Starting...")
        if fname.lower().endswith(".js"):
            run_js_script(fpath, owner_id, ufolder, fname, call.message)
        else:
            run_script(fpath, owner_id, ufolder, fname, call.message)

    elif data.startswith("stop_"):
        _, owner_id, fname = data.split("_", 2)
        skey = f"{owner_id}_{fname}"
        if skey in bot_scripts:
            kill_process_tree(bot_scripts[skey])
            del bot_scripts[skey]
        bot.answer_callback_query(call.id, "Stopped!")
        bot.send_message(
            call.message.chat.id,
            f"🛑 Script `{fname}` stopped.",
            parse_mode="Markdown",
        )

    elif data.startswith("del_"):
        _, owner_id, fname = data.split("_", 2)
        skey = f"{owner_id}_{fname}"
        if skey in bot_scripts:
            kill_process_tree(bot_scripts[skey])
            del bot_scripts[skey]
        remove_user_file_db(int(owner_id), fname)
        ufolder = get_user_folder(int(owner_id))
        fpath = os.path.join(ufolder, fname)
        if os.path.exists(fpath):
            os.remove(fpath)
        bot.answer_callback_query(call.id, "Deleted!")
        bot.send_message(
            call.message.chat.id,
            f"🗑️ File `{fname}` deleted.",
            parse_mode="Markdown",
        )


# --- Step Handlers ---
def process_deposit_submission(message, plan_id):
    try:
        parts = [p.strip() for p in message.text.split("|", 1)]
        if len(parts) != 2 or not parts[0] or not parts[1]:
            raise ValueError("Use AMOUNT | TRANSACTION/REFERENCE ID")
        amount, reference = parts
        if len(reference) < 3:
            raise ValueError("Reference ID is too short")
        request_id = create_deposit_request(message.from_user.id, plan_id, amount, reference)
        plan = get_plan_by_id(plan_id)
        pname = plan[1] if plan else "Unknown"
        bot.reply_to(message, f"Deposit request #{request_id} submitted for {pname}.\nStatus: Pending admin review.")
        for aid in admin_ids:
            try:
                markup = types.InlineKeyboardMarkup(row_width=2)
                markup.add(types.InlineKeyboardButton("Approve", callback_data=f"approve_deposit_{request_id}"), types.InlineKeyboardButton("Reject", callback_data=f"reject_deposit_{request_id}"))
                bot.send_message(aid, f"New deposit request #{request_id}\nUser: {message.from_user.id}\nPlan: {pname}\nAmount: {amount}\nReference: {reference}", reply_markup=markup)
            except Exception:
                pass
    except Exception as e:
        bot.reply_to(message, f"Invalid deposit details: {e}")


def process_payment_settings(message):
    if message.from_user.id not in admin_ids:
        return
    info = message.text.strip()
    if len(info) < 5:
        bot.reply_to(message, "Payment instructions are too short.")
        return
    set_payment_info(info)
    bot.reply_to(message, "Payment instructions updated successfully.")


def process_remove_subscription(message):
    try:
        uid = int(message.text.strip())
        remove_subscription_db(uid)
        bot.reply_to(message, f"Subscription removed for User {uid}.")
    except Exception as e:
        bot.reply_to(message, f"Invalid User ID: {e}")


def process_add_admin(message):
    if message.from_user.id != OWNER_ID:
        return
    try:
        uid = int(message.text.strip())
        with DB_LOCK:
            conn = sqlite3.connect(DATABASE_PATH, check_same_thread=False)
            conn.execute("INSERT OR IGNORE INTO admins (user_id) VALUES (?)", (uid,))
            conn.commit()
            conn.close()
        admin_ids.add(uid)
        bot.reply_to(message, f"User {uid} added as admin.")
    except Exception as e:
        bot.reply_to(message, f"Invalid User ID: {e}")


def process_remove_admin(message):
    if message.from_user.id != OWNER_ID:
        return
    try:
        uid = int(message.text.strip())
        if uid == OWNER_ID:
            bot.reply_to(message, "Owner cannot be removed.")
            return
        with DB_LOCK:
            conn = sqlite3.connect(DATABASE_PATH, check_same_thread=False)
            conn.execute("DELETE FROM admins WHERE user_id=?", (uid,))
            conn.commit()
            conn.close()
        admin_ids.discard(uid)
        bot.reply_to(message, f"User {uid} removed from admin list.")
    except Exception as e:
        bot.reply_to(message, f"Invalid User ID: {e}")


def process_database_upload(message):
    if message.from_user.id not in admin_ids or not message.document:
        return
    name = message.document.file_name.lower()
    if not name.endswith(('.db', '.sqlite', '.sqlite3')):
        bot.reply_to(message, "Only SQLite database files are accepted.")
        return
    try:
        info = bot.get_file(message.document.file_id)
        data = bot.download_file(info.file_path)
        # Validate SQLite header before replacing anything.
        if not data.startswith(b"SQLite format 3\x00"):
            bot.reply_to(message, "Invalid SQLite database file.")
            return
        backup = DATABASE_PATH + ".before_restore"
        shutil.copy2(DATABASE_PATH, backup)
        with open(DATABASE_PATH, "wb") as f:
            f.write(data)
        # Rebuild runtime caches from restored DB.
        user_subscriptions.clear(); user_files.clear(); active_users.clear()
        init_db(); load_data()
        bot.reply_to(message, "Database restored successfully. A pre-restore backup was kept on the server as bot_data.db.before_restore.")
    except Exception as e:
        bot.reply_to(message, f"Database restore failed: {e}")


def clear_all_data():
    # Stop all hosted processes first.
    for key in list(bot_scripts.keys()):
        try:
            kill_process_tree(bot_scripts[key])
        except Exception:
            pass
    bot_scripts.clear()
    with DB_LOCK:
        conn = sqlite3.connect(DATABASE_PATH, check_same_thread=False)
        c = conn.cursor()
        for table in ("subscriptions", "user_files", "active_users", "plans", "deposit_requests"):
            c.execute(f"DELETE FROM {table}")
        c.execute("DELETE FROM payment_settings")
        c.execute("INSERT INTO payment_settings (setting_key, setting_value) VALUES (?, ?)", ("payment_info", PAYMENT_INFO_DEFAULT))
        conn.commit()
        conn.close()
    user_subscriptions.clear(); user_files.clear(); active_users.clear()
    if os.path.isdir(UPLOAD_BOTS_DIR):
        for name in os.listdir(UPLOAD_BOTS_DIR):
            path = os.path.join(UPLOAD_BOTS_DIR, name)
            try:
                if os.path.isdir(path): shutil.rmtree(path)
                else: os.remove(path)
            except Exception as e:
                logger.warning(f"Could not remove {path}: {e}")


def process_add_plan(message):
    try:
        parts = [p.strip() for p in message.text.split("|")]
        name, limit, price, duration, buy_link = (
            parts[0],
            int(parts[1]),
            parts[2],
            int(parts[3]),
            parts[4],
        )
        add_plan_db(name, limit, price, duration, buy_link)
        bot.reply_to(
            message,
            f"✅ **Plan `{name}` added successfully!**",
            parse_mode="Markdown",
        )
    except Exception as e:
        bot.reply_to(message, f"❌ Invalid Format! Error: {e}")


def process_add_subscription(message):
    try:
        parts = message.text.split()
        sub_uid, pname, days = int(parts[0]), parts[1], int(parts[2])
        exp = datetime.now() + timedelta(days=days)
        save_subscription(sub_uid, pname, exp)
        bot.reply_to(
            message,
            f"✅ **Subscription active for User `{sub_uid}` under Plan `{pname}` for {days} days!**",
            parse_mode="Markdown",
        )
    except Exception as e:
        bot.reply_to(message, f"❌ Error: {e}")


# --- Text Handler Mapping ---
BUTTON_MAPPING = {
    "Updates Channel": lambda m: bot.reply_to(m, f"Updates Channel: {UPDATE_CHANNEL}"),
    "Upload File": _logic_upload_file,
    "Manage Files": _logic_check_files,
    "Hosting Plans": _logic_view_plans,
    "Server Status": lambda m: bot.reply_to(m, f"Server Status: Online\nHosted Processes: {len(bot_scripts)}"),
    "Bot Stats": lambda m: bot.reply_to(m, f"Bot Stats\nActive Users: {len(active_users)}\nRunning Bots: {len(bot_scripts)}"),
    "Support": lambda m: bot.reply_to(m, f"Support: {YOUR_USERNAME}"),
    "Admin Panel": lambda m: bot.reply_to(m, "ADMIN CONTROL PANEL", reply_markup=create_admin_panel_inline()),
}


@bot.message_handler(func=lambda m: m.text in BUTTON_MAPPING)
def handle_main_buttons(message):
    BUTTON_MAPPING[message.text](message)


@bot.message_handler(commands=["start"])
def start_cmd(message):
    _logic_send_welcome(message)


# --- Cleanup & Start ---
def cleanup():
    for key in list(bot_scripts.keys()):
        kill_process_tree(bot_scripts[key])


atexit.register(cleanup)

if __name__ == "__main__":
    logger.info(f"Starting {APP_NAME}...")
    keep_alive()
    bot.infinity_polling(timeout=60, long_polling_timeout=30)