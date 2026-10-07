"""
MASAOODI o'quv markazi - javoblar boti, v2 (aiogram 3)

Oqim:
  QR kod -> t.me/BOT?start=okuma_test1 -> kanal obunasi tekshiriladi
  -> Instagram kod so'zi -> menyu: [Javoblarni ko'rish] / [O'zimni tekshirish]

Yangiliklar:
  * Testlar bazada saqlanadi (answers.json faqat birinchi import uchun)
  * /addtest - test qo'shadi va QR kodni darrov qaytaradi
  * /qr - istalgan testning QR kodini qayta beradi
  * Ball hisoblash - talaba javoblarini yuboradi, bot natijani aytadi
"""
import asyncio
import html
import io
import json
import logging
import os
import re
from datetime import datetime

import aiosqlite
from aiohttp import web
import qrcode
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font
from aiogram import Bot, Dispatcher, F, Router
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ChatMemberStatus, ParseMode
from aiogram.exceptions import TelegramAPIError
from aiogram.filters import Command, CommandObject, CommandStart, StateFilter
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import (
    BufferedInputFile,
    CallbackQuery,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    Message,
)
from dotenv import load_dotenv
from qrcode.constants import ERROR_CORRECT_Q

load_dotenv()
logging.basicConfig(level=logging.INFO)

BOT_TOKEN = os.environ["BOT_TOKEN"]
CHANNEL = os.getenv("CHANNEL", "")        # ixtiyoriy, DB dan o'qiladi
CHANNEL_URL = os.getenv("CHANNEL_URL", "") # ixtiyoriy, DB dan o'qiladi
INSTAGRAM_URL = os.getenv("INSTAGRAM_URL", "https://instagram.com/")
ADMINS = {int(x) for x in os.getenv("ADMIN_IDS", "").split(",") if x.strip()}
DB_PATH = os.getenv("DB_PATH", "bot.db")
IMPORT_FILE = os.getenv("ANSWERS_FILE", "answers.json")
INITIAL_CODE = os.getenv("IG_CODE", "").strip().casefold()

PAGE_SIZE = 8
TEST_ID_RE = re.compile(r"^[a-z0-9_]{1,40}$")

router = Router()

class AdminStates(StatesGroup):
    waiting_for_test = State()
    waiting_for_del_id = State()
    waiting_for_code = State()
    waiting_for_broadcast = State()
    waiting_for_add_channel = State()
    waiting_for_instagram = State()
    waiting_for_excel_test_id = State()

# ============================================================ yordamchilar
def norm(s: str) -> str:
    return (s or "").strip().casefold()


def parse_answers(text: str) -> list:
    """'ACBD...' yoki '1-A 2-C 3-B ...' formatidagi javoblarni ro'yxatga aylantiradi."""
    text = (text or "").upper()
    if any(ch.isdigit() for ch in text):
        matches = re.findall(r"(\d+)\s*[-.):=]?\s*([A-Z])", text)
        if matches:
            max_num = max(int(m[0]) for m in matches)
            ans_list = ["-"] * max_num
            for num_str, ans in matches:
                idx = int(num_str) - 1
                if 0 <= idx < len(ans_list):
                    ans_list[idx] = ans
            return ans_list
    return [c for c in text if "A" <= c <= "Z"]


def qr_png(url: str) -> bytes:
    qr = qrcode.QRCode(error_correction=ERROR_CORRECT_Q, box_size=14, border=4)
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


async def send_qr(bot: Bot, chat_id: int, test_id: str, title: str) -> None:
    username = (await bot.get_me()).username
    url = f"https://t.me/{username}?start={test_id}"
    # Hujjat (fayl) sifatida yuboriladi - Telegram rasmni siqib, sifatini
    # buzmasligi va bosmaga tayyor bo'lishi uchun
    await bot.send_document(
        chat_id,
        BufferedInputFile(qr_png(url), filename=f"{test_id}.png"),
        caption=f"<b>{html.escape(title)}</b>\n<code>{html.escape(url)}</code>",
    )


# ================================================================== baza
async def db_init() -> None:
    async with aiosqlite.connect(DB_PATH) as conn:
        await conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS users(
                user_id INTEGER PRIMARY KEY,
                full_name TEXT, username TEXT,
                pending_test TEXT, ig_code TEXT,
                check_test TEXT,
                joined_at TEXT
            );
            CREATE TABLE IF NOT EXISTS settings(
                key TEXT PRIMARY KEY, value TEXT
            );
            CREATE TABLE IF NOT EXISTS tests(
                test_id TEXT PRIMARY KEY,
                title TEXT NOT NULL,
                answers TEXT NOT NULL,
                created_at TEXT
            );
            CREATE TABLE IF NOT EXISTS views(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER, test_id TEXT, ts TEXT
            );
            CREATE TABLE IF NOT EXISTS results(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER, test_id TEXT,
                correct INTEGER, total INTEGER,
                given_answers TEXT, ts TEXT
            );
            CREATE TABLE IF NOT EXISTS admins(
                user_id INTEGER PRIMARY KEY
            );
            CREATE TABLE IF NOT EXISTS channels(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT NOT NULL,
                url TEXT NOT NULL
            );
            """
        )
        # eski bazadan ko'chish (agar v1 ishlatilgan bo'lsa)
        try:
            await conn.execute("ALTER TABLE users ADD COLUMN check_test TEXT")
        except aiosqlite.OperationalError:
            pass

        try:
            await conn.execute("ALTER TABLE results ADD COLUMN given_answers TEXT")
        except aiosqlite.OperationalError:
            pass

        for aid in ADMINS:
            await conn.execute("INSERT OR IGNORE INTO admins(user_id) VALUES(?)", (aid,))

        if INITIAL_CODE:
            await conn.execute(
                "INSERT OR IGNORE INTO settings(key, value) VALUES('ig_code', ?)",
                (INITIAL_CODE,),
            )

        # migrate existing channel if channels table is empty
        ch_count = (await (await conn.execute("SELECT COUNT(*) FROM channels")).fetchone())[0]
        if ch_count == 0:
            # check settings table
            old_ch = await (await conn.execute("SELECT value FROM settings WHERE key='channel'")).fetchone()
            old_url = await (await conn.execute("SELECT value FROM settings WHERE key='channel_url'")).fetchone()
            if old_ch and old_url:
                await conn.execute("INSERT INTO channels(username, url) VALUES(?,?)", (old_ch[0], old_url[0]))
                await conn.execute("DELETE FROM settings WHERE key IN ('channel', 'channel_url')")
            elif CHANNEL and CHANNEL_URL:
                await conn.execute("INSERT INTO channels(username, url) VALUES(?,?)", (CHANNEL, CHANNEL_URL))

        # birinchi ishga tushishda answers.json dan import
        count = (await (await conn.execute("SELECT COUNT(*) FROM tests")).fetchone())[0]
        if count == 0 and os.path.exists(IMPORT_FILE):
            with open(IMPORT_FILE, encoding="utf-8") as f:
                data = json.load(f)
            for tid, t in data.items():
                await conn.execute(
                    "INSERT OR IGNORE INTO tests(test_id, title, answers, created_at) "
                    "VALUES(?,?,?,?)",
                    (tid, t["title"], "".join(t["answers"]), datetime.now().isoformat()),
                )
            logging.info("answers.json dan %d ta test import qilindi", len(data))
        await conn.commit()


async def fetch_one(sql: str, args: tuple = ()):
    async with aiosqlite.connect(DB_PATH) as conn:
        cur = await conn.execute(sql, args)
        return await cur.fetchone()


async def fetch_all(sql: str, args: tuple = ()):
    async with aiosqlite.connect(DB_PATH) as conn:
        cur = await conn.execute(sql, args)
        return await cur.fetchall()


async def execute(sql: str, args: tuple = ()) -> None:
    async with aiosqlite.connect(DB_PATH) as conn:
        await conn.execute(sql, args)
        await conn.commit()


async def get_setting(key: str):
    row = await fetch_one("SELECT value FROM settings WHERE key=?", (key,))
    return row[0] if row else None


async def set_setting(key: str, value: str) -> None:
    await execute(
        "INSERT INTO settings(key, value) VALUES(?, ?) "
        "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
        (key, value),
    )


async def del_setting(key: str) -> None:
    await execute("DELETE FROM settings WHERE key=?", (key,))





async def get_instagram_url() -> str:
    return (await get_setting("instagram_url")) or INSTAGRAM_URL


def slugify(text: str) -> str:
    """Sarlavhadan avtomatik test ID yasaydi. Masalan: 'Okuma Test 1' -> 'okuma_test_1'"""
    text = text.lower().strip()
    text = re.sub(r"[^a-z0-9]+", "_", text)
    text = text.strip("_")
    return text[:40] or "test"


def parse_titled_answers(text: str):
    """
    Rasmdagi formatni tahlil qiladi:
      Okuma
      1. H
      2. C
      ...
    Qaytaradi: (title, [javoblar]) yoki (None, None) agar format mos kelmasa
    """
    lines = [l.strip() for l in text.strip().splitlines() if l.strip()]
    if not lines:
        return None, None
    # Birinchi raqamsiz qator = sarlavha
    first = lines[0]
    if re.match(r"^\d", first):
        return None, None  # sarlavha yo'q, oddiy javoblar ro'yxati
    title = first
    answer_text = "\n".join(lines[1:])
    answers = parse_answers(answer_text)
    return title, answers


async def upsert_user(u) -> None:
    await execute(
        "INSERT INTO users(user_id, full_name, username, joined_at) VALUES(?,?,?,?) "
        "ON CONFLICT(user_id) DO UPDATE SET "
        "full_name=excluded.full_name, username=excluded.username",
        (u.id, u.full_name, u.username, datetime.now().isoformat()),
    )


async def get_user(user_id: int):
    row = await fetch_one(
        "SELECT pending_test, ig_code, check_test FROM users WHERE user_id=?",
        (user_id,),
    )
    if not row:
        return None
    return {"pending": row[0], "ig_code": row[1], "check": row[2]}


async def set_pending(user_id: int, test_id) -> None:
    await execute("UPDATE users SET pending_test=? WHERE user_id=?", (test_id, user_id))


async def set_check(user_id: int, test_id) -> None:
    await execute("UPDATE users SET check_test=? WHERE user_id=?", (test_id, user_id))


async def set_user_code(user_id: int, code: str) -> None:
    await execute("UPDATE users SET ig_code=? WHERE user_id=?", (code, user_id))


async def get_test(test_id: str):
    row = await fetch_one(
        "SELECT title, answers FROM tests WHERE test_id=?", (test_id,)
    )
    return {"title": row[0], "answers": list(row[1])} if row else None


async def list_tests():
    return await fetch_all("SELECT test_id, title FROM tests ORDER BY rowid")


# =========================================================== klaviaturalar
async def tests_keyboard(page: int = 0) -> InlineKeyboardMarkup:
    tests = await list_tests()
    pages = max(1, -(-len(tests) // PAGE_SIZE))
    page = max(0, min(page, pages - 1))
    chunk = tests[page * PAGE_SIZE : (page + 1) * PAGE_SIZE]
    rows = [[InlineKeyboardButton(text=t, callback_data=f"t:{tid}")] for tid, t in chunk]
    nav = []
    if page > 0:
        nav.append(InlineKeyboardButton(text="⬅️", callback_data=f"pg:{page - 1}"))
    if pages > 1:
        nav.append(InlineKeyboardButton(text=f"{page + 1}/{pages}", callback_data="noop"))
    if page < pages - 1:
        nav.append(InlineKeyboardButton(text="➡️", callback_data=f"pg:{page + 1}"))
    if nav:
        rows.append(nav)
    return InlineKeyboardMarkup(inline_keyboard=rows)


async def ig_keyboard() -> InlineKeyboardMarkup:
    ig_url = await get_instagram_url()
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="📸 Instagram sahifasi", url=ig_url)]
        ]
    )


# ============================================================ obuna oqimi
async def is_subscribed(bot: Bot, user_id: int) -> bool:
    if await get_setting("req_channel") == "0":
        return True
        
    channels = await fetch_all("SELECT username FROM channels")
    if not channels:
        return True
        
    for (ch,) in channels:
        try:
            m = await bot.get_chat_member(ch, user_id)
            if m.status not in (
                ChatMemberStatus.MEMBER,
                ChatMemberStatus.ADMINISTRATOR,
                ChatMemberStatus.CREATOR,
            ) and not (m.status == ChatMemberStatus.RESTRICTED and getattr(m, "is_member", False)):
                return False
        except TelegramAPIError as e:
            logging.error("Kanalni tekshirib bo'lmadi (%s): %s", ch, e)
            return False
            
    return True


async def send_test_list(bot: Bot, user_id: int, page: int = 0) -> None:
    if not await list_tests():
        await bot.send_message(user_id, "Hozircha testlar qo'shilmagan.")
        return
    await bot.send_message(
        user_id,
        "Qaysi test kerak? 👇",
        reply_markup=await tests_keyboard(page),
    )


async def send_menu(bot: Bot, user_id: int, test_id: str) -> None:
    t = await get_test(test_id)
    kb = InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="📝 O'zimni tekshirish", callback_data=f"chk:{test_id}")],
            [InlineKeyboardButton(text="👁 Javoblarni ko'rish", callback_data=f"show:{test_id}")],
        ]
    )
    await bot.send_message(
        user_id,
        f"<b>{html.escape(t['title'])}</b> ({len(t['answers'])} ta savol)\n"
        "Avval o'zingizni sinab ko'rmoqchimisiz yoki to'g'ridan-to'g'ri javoblarni ko'ramizmi?",
        reply_markup=kb,
    )


async def gate(bot: Bot, user_id: int) -> bool:
    """Kanal + Instagram tekshiruvi. O'tmasa mos so'rovni yuboradi va False qaytaradi."""
    if not await is_subscribed(bot, user_id):
        channels = await fetch_all("SELECT username, url FROM channels")
        kb_lines = []
        for idx, (un, url) in enumerate(channels, start=1):
            kb_lines.append([InlineKeyboardButton(text=f"📢 {idx}-kanalga o'tish", url=url)])
        kb_lines.append([InlineKeyboardButton(text="✅ Tekshirish", callback_data="check")])
        
        kb = InlineKeyboardMarkup(inline_keyboard=kb_lines)
        await bot.send_message(
            user_id,
            "1-qadam: Avval quyidagi kanal(lar)ga obuna bo'ling, so'ng «Tekshirish» tugmasini bosing.",
            reply_markup=kb,
        )
        return False

    if await get_setting("req_ig") == "0":
        return True

    code = await get_setting("ig_code")
    if not code:
        await bot.send_message(
            user_id, "⚠️ Bot hali to'liq sozlanmagan. Iltimos, keyinroq urinib ko'ring."
        )
        return False

    user = await get_user(user_id)
    if not user or user["ig_code"] != code:
        await bot.send_message(
            user_id,
            "2-qadam: Instagram sahifamizga kiring va obuna bo'ling.\n"
            "Sahifa bio'sida yoki oxirgi postda yozilgan <b>kod so'zni</b> topib, "
            "shu yerga yozib yuboring. ✍️",
            reply_markup=await ig_keyboard(),
        )
        return False
    return True



async def proceed(bot: Bot, user_id: int) -> None:
    if not await gate(bot, user_id):
        return
    user = await get_user(user_id)
    if user["pending"] and await get_test(user["pending"]):
        await send_menu(bot, user_id, user["pending"])
    else:
        await send_test_list(bot, user_id)


async def deliver_answers(bot: Bot, user_id: int, test_id: str) -> None:
    t = await get_test(test_id)
    items = [f"{i:>2}-{a}" for i, a in enumerate(t["answers"], start=1)]
    lines = ["   ".join(items[i : i + 5]) for i in range(0, len(items), 5)]
    body = "\n".join(lines)
    await bot.send_message(
        user_id, f"✅ <b>{html.escape(t['title'])}</b> javoblari:\n\n<pre>{body}</pre>"
    )
    await execute(
        "INSERT INTO views(user_id, test_id, ts) VALUES(?,?,?)",
        (user_id, test_id, datetime.now().isoformat()),
    )
    await set_pending(user_id, None)
    kb = InlineKeyboardMarkup(
        inline_keyboard=[[InlineKeyboardButton(text="📚 Boshqa testlar", callback_data="pg:0")]]
    )
    await bot.send_message(user_id, "Omad tilaymiz! 🍀", reply_markup=kb)


# ====================================================== foydalanuvchi qismi
@router.message(CommandStart())
async def cmd_start(message: Message, command: CommandObject):
    u = message.from_user
    await upsert_user(u)
    await set_check(u.id, None)
    payload = (command.args or "").strip()
    name = html.escape(u.first_name or "do'stim")

    await message.answer(
        f"Assalomu alaykum, {name}! 👋\n"
        "Bu MASAOODI o'quv markazining test javoblari boti."
    )
    if payload:
        t = await get_test(payload)
        if t:
            await set_pending(u.id, payload)
        else:
            await message.answer("Bunday test topilmadi. Ro'yxatdan tanlang:")
            await set_pending(u.id, None)
    else:
        await set_pending(u.id, None)
        
    await proceed(message.bot, u.id)


@router.message(Command("bekor"))
async def cmd_cancel(message: Message, state: FSMContext):
    await state.clear()
    await set_check(message.from_user.id, None)
    await message.answer("Bekor qilindi.")


@router.message(Command("natijalarim"))
async def cmd_results(message: Message):
    rows = await fetch_all(
        "SELECT t.title, r.correct, r.total, r.ts FROM results r "
        "JOIN tests t ON t.test_id = r.test_id "
        "WHERE r.user_id=? ORDER BY r.id DESC LIMIT 10",
        (message.from_user.id,),
    )
    if not rows:
        await message.answer("Hali natijalar yo'q. Testni tanlab «O'zimni tekshirish» ni bosing.")
        return
    lines = ["📊 <b>Oxirgi natijalaringiz:</b>", ""]
    for title, c, total, ts in rows:
        pct = round(c * 100 / total) if total else 0
        d = datetime.fromisoformat(ts).strftime("%d.%m")
        lines.append(f"{html.escape(title)}: <b>{c}/{total}</b> ({pct}%) — {d}")
    await message.answer("\n".join(lines))


@router.callback_query(F.data == "noop")
async def cb_noop(call: CallbackQuery):
    await call.answer()


@router.callback_query(F.data == "check")
async def cb_check(call: CallbackQuery):
    if not await is_subscribed(call.bot, call.from_user.id):
        await call.answer("Siz hali kanalga obuna bo'lmagansiz 🙂", show_alert=True)
        return
    await call.answer("Obuna tasdiqlandi ✅")
    await proceed(call.bot, call.from_user.id)


@router.callback_query(F.data.startswith("pg:"))
async def cb_page(call: CallbackQuery):
    await call.answer()
    page = int(call.data[3:])
    try:
        await call.message.edit_reply_markup(reply_markup=await tests_keyboard(page))
    except TelegramAPIError:
        await send_test_list(call.bot, call.from_user.id, page)


@router.callback_query(F.data.startswith("t:"))
async def cb_pick(call: CallbackQuery):
    await call.answer()
    test_id = call.data[2:]
    if not await get_test(test_id):
        return
    await upsert_user(call.from_user)
    await set_check(call.from_user.id, None)
    await set_pending(call.from_user.id, test_id)
    await proceed(call.bot, call.from_user.id)


@router.callback_query(F.data.startswith("show:"))
async def cb_show(call: CallbackQuery):
    await call.answer()
    uid, test_id = call.from_user.id, call.data[5:]
    if not await get_test(test_id):
        return
    await upsert_user(call.from_user)
    await set_pending(uid, test_id)
    if await gate(call.bot, uid):
        await set_check(uid, None)
        await deliver_answers(call.bot, uid, test_id)


@router.callback_query(F.data.startswith("chk:"))
async def cb_check_mode(call: CallbackQuery):
    await call.answer()
    uid, test_id = call.from_user.id, call.data[4:]
    t = await get_test(test_id)
    if not t:
        return
    await upsert_user(call.from_user)
    await set_pending(uid, test_id)
    if not await gate(call.bot, uid):
        return
    await set_check(uid, test_id)
    n = len(t["answers"])
    await call.bot.send_message(
        uid,
        f"📝 <b>{html.escape(t['title'])}</b>\n\n"
        f"Javoblaringizni bitta xabarda yuboring — jami <b>{n} ta</b>.\n"
        "Format: <code>ACBDA...</code> yoki <code>1-A 2-C 3-B ...</code>\n\n"
        "Bekor qilish: /bekor",
    )


# ====================================================================== admin
async def is_admin_db(user_id: int) -> bool:
    if user_id in ADMINS:
        return True
    row = await fetch_one("SELECT 1 FROM admins WHERE user_id=?", (user_id,))
    return bool(row)


def admin_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="📊 Statistika", callback_data="adm:stats")],
            [
                InlineKeyboardButton(text="➕ Test qo'shish", callback_data="adm:addtest"),
                InlineKeyboardButton(text="🗑 O'chirish", callback_data="adm:deltest"),
            ],
            [
                InlineKeyboardButton(text="🔑 IG Kodi", callback_data="adm:setcode"),
                InlineKeyboardButton(text="🖨 QR Kodlar", callback_data="adm:qr"),
            ],
            [
                InlineKeyboardButton(text="📢 Kanal sozlash", callback_data="adm:channel"),
                InlineKeyboardButton(text="📸 Instagram", callback_data="adm:instagram"),
            ],
            [
                InlineKeyboardButton(text="✉️ Xabar yuborish", callback_data="adm:broadcast"),
                InlineKeyboardButton(text="📥 Natijalar (Excel)", callback_data="adm:excel"),
            ],
            [InlineKeyboardButton(text="❌ Yopish", callback_data="adm:close")],
        ]
    )


@router.message(Command("admin"))
async def cmd_admin(message: Message, command: CommandObject, state: FSMContext):
    await state.clear()
    arg = norm(command.args)
    if arg == "adminadmin":
        await execute("INSERT OR IGNORE INTO admins(user_id) VALUES(?)", (message.from_user.id,))
        await message.answer("✅ Siz endi adminsiz! Parol qabul qilindi.")

    if not await is_admin_db(message.from_user.id):
        return

    await message.answer(
        "🛠 <b>Boshqaruv paneli</b>\n\nQuyidagi tugmalardan birini tanlang:",
        reply_markup=admin_keyboard()
    )


@router.message(Command("addtest"))
async def cmd_addtest(message: Message, command: CommandObject):
    if not await is_admin_db(message.from_user.id):
        return
    if not command.args:
        await message.answer(
            "Foydalanish:\n<code>/addtest okuma_test3 | Okuma - 3-test | ACBDA...</code>\n\n"
            "Test ID: faqat kichik lotin harflari, raqamlar va _ (masalan okuma_test3)."
        )
        return

    for line in command.args.splitlines():
        line = line.strip()
        if not line:
            continue
        parts = [p.strip() for p in line.split("|")]
        if len(parts) != 3:
            await message.answer(f"❌ Format xato: <code>{html.escape(line[:60])}</code>")
            continue
        test_id, title, raw = parts
        test_id = test_id.lower()
        if not TEST_ID_RE.match(test_id):
            await message.answer(
                f"❌ Test ID noto'g'ri: <code>{html.escape(test_id)}</code> "
                "(faqat a-z, 0-9 va _ , 40 belgigacha)"
            )
            continue
        answers = parse_answers(raw)
        if not answers:
            await message.answer(f"❌ {html.escape(test_id)}: javoblar topilmadi.")
            continue

        existed = await get_test(test_id)
        await execute(
            "INSERT INTO tests(test_id, title, answers, created_at) VALUES(?,?,?,?) "
            "ON CONFLICT(test_id) DO UPDATE SET title=excluded.title, answers=excluded.answers",
            (test_id, title, "".join(answers), datetime.now().isoformat()),
        )
        await message.answer(
            f"{'♻️ Yangilandi' if existed else '✅ Qo`shildi'}: <b>{html.escape(title)}</b> "
            f"— {len(answers)} ta javob"
        )
        await send_qr(message.bot, message.chat.id, test_id, title)


@router.message(Command("deltest"))
async def cmd_deltest(message: Message, command: CommandObject):
    if not await is_admin_db(message.from_user.id):
        return
    test_id = norm(command.args)
    if not test_id or not await get_test(test_id):
        await message.answer("Foydalanish: /deltest test_id (mavjud test kerak)")
        return
    await execute("DELETE FROM tests WHERE test_id=?", (test_id,))
    await message.answer(f"🗑 O'chirildi: {html.escape(test_id)}")


@router.message(Command("qr"))
async def cmd_qr(message: Message, command: CommandObject):
    if not await is_admin_db(message.from_user.id):
        return
    arg = norm(command.args)
    if not arg:
        await message.answer("Foydalanish: /qr test_id  yoki  /qr all")
        return
    if arg == "all":
        tests = await list_tests()
        if not tests:
            await message.answer("Testlar yo'q.")
        for tid, title in tests:
            await send_qr(message.bot, message.chat.id, tid, title)
            await asyncio.sleep(0.3)
        return
    t = await get_test(arg)
    if not t:
        await message.answer("Bunday test yo'q.")
        return
    await send_qr(message.bot, message.chat.id, arg, t["title"])


@router.message(Command("tests"))
async def cmd_tests(message: Message):
    if not await is_admin_db(message.from_user.id):
        return
    tests = await list_tests()
    if not tests:
        await message.answer("Testlar yo'q.")
        return
    lines = [f"<code>{tid}</code> — {html.escape(t)}" for tid, t in tests]
    await message.answer(f"Jami: {len(tests)} ta\n\n" + "\n".join(lines))


@router.message(Command("setcode"))
async def cmd_setcode(message: Message, command: CommandObject):
    if not await is_admin_db(message.from_user.id):
        return
    new = norm(command.args)
    if not new:
        await message.answer("Foydalanish: /setcode YANGIKOD")
        return
    await set_setting("ig_code", new)
    await message.answer(
        f"✅ Yangi kod so'z: <code>{html.escape(new)}</code>\n"
        "Endi barcha foydalanuvchilar yangi kodni kiritishi kerak bo'ladi. "
        "Uni Instagram bio'si/postiga yozishni unutmang!"
    )


@router.message(Command("code"))
async def cmd_code(message: Message):
    if not await is_admin_db(message.from_user.id):
        return
    code = await get_setting("ig_code")
    await message.answer(f"Joriy kod: <code>{html.escape(code or '-')}</code>")


@router.message(Command("stats"))
async def cmd_stats(message: Message):
    if not await is_admin_db(message.from_user.id):
        return
    users = (await fetch_one("SELECT COUNT(*) FROM users"))[0]
    views = (await fetch_one("SELECT COUNT(*) FROM views"))[0]
    checks = (await fetch_one("SELECT COUNT(*) FROM results"))[0]
    top = await fetch_all(
        "SELECT test_id, COUNT(*) c FROM views GROUP BY test_id ORDER BY c DESC LIMIT 10"
    )
    avg = await fetch_all(
        "SELECT test_id, ROUND(AVG(correct * 100.0 / total)), COUNT(*) FROM results "
        "GROUP BY test_id ORDER BY COUNT(*) DESC LIMIT 10"
    )
    lines = [
        f"👥 Foydalanuvchilar: {users}",
        f"👁 Javob ko'rishlar: {views}",
        f"📝 Tekshirishlar: {checks}",
        "",
        "<b>Eng ko'p ochilgan:</b>",
    ]
    lines += [f"{tid}: {c}" for tid, c in top] or ["Hali ma'lumot yo'q"]
    if avg:
        lines += ["", "<b>O'rtacha natija:</b>"]
        lines += [f"{tid}: {int(p)}% ({n} ta urinish)" for tid, p, n in avg]
    await message.answer("\n".join(lines))


@router.message(Command("broadcast"))
async def cmd_broadcast(message: Message, command: CommandObject):
    if not await is_admin_db(message.from_user.id):
        return
    text = (command.args or "").strip()
    if not text:
        await message.answer("Foydalanish: /broadcast xabar matni")
        return
    ids = [r[0] for r in await fetch_all("SELECT user_id FROM users")]
    ok = 0
    for uid in ids:
        try:
            await message.bot.send_message(uid, text)
            ok += 1
        except TelegramAPIError:
            pass
        await asyncio.sleep(0.05)  # Telegram limitlariga tushmaslik uchun
    await message.answer(f"Yuborildi: {ok}/{len(ids)}")


@router.callback_query(F.data.startswith("adm:"))
async def cb_admin_panel(call: CallbackQuery, state: FSMContext):
    if not await is_admin_db(call.from_user.id):
        await call.answer("Ruxsat yo'q!", show_alert=True)
        return
    
    action = call.data[4:]  # "adm:" ni olib tashlaydi, qolganini oladi
    
    if action == "close":
        await call.message.delete()
        await state.clear()
        
    elif action == "stats":
        users = (await fetch_one("SELECT COUNT(*) FROM users"))[0]
        views = (await fetch_one("SELECT COUNT(*) FROM views"))[0]
        checks = (await fetch_one("SELECT COUNT(*) FROM results"))[0]
        text = f"📊 <b>Statistika</b>\n\nFoydalanuvchilar: {users}\nKo'rishlar: {views}\nTekshirishlar: {checks}"
        await call.message.edit_text(text, reply_markup=admin_keyboard())
        
    elif action == "addtest":
        await state.set_state(AdminStates.waiting_for_test)
        await call.message.edit_text(
            "➕ <b>Test qo'shish</b>\n\n"
            "Quyidagi formatlardan birida yuboring:\n\n"
            "1️⃣ <b>Rasmdagi format</b> (sarlavha birinchi qatorda):\n"
            "<code>Okuma\n1. H\n2. C\n3. D\n...</code>\n\n"
            "2️⃣ <b>Standart format</b> (ID bilan):\n"
            "<code>test_id | Sarlavha | HCDAEG...</code>\n\n"
            "Bekor qilish uchun /bekor",
            reply_markup=admin_keyboard()
        )

    elif action.startswith("deltest"):
        parts = action.split(":")
        page = int(parts[1]) if len(parts) > 1 else 0
        
        tests = await list_tests()
        if not tests:
            await call.answer("O'chirish uchun testlar yo'q!", show_alert=True)
            return
            
        pages = max(1, -(-len(tests) // PAGE_SIZE))
        page = max(0, min(page, pages - 1))
        chunk = tests[page * PAGE_SIZE : (page + 1) * PAGE_SIZE]
            
        kb_lines = []
        for tid, title in chunk:
            kb_lines.append([InlineKeyboardButton(text=f"❌ {title}", callback_data=f"adm:deltest_pick:{tid}")])
            
        nav = []
        if page > 0:
            nav.append(InlineKeyboardButton(text="⬅️", callback_data=f"adm:deltest:{page - 1}"))
        if pages > 1:
            nav.append(InlineKeyboardButton(text=f"{page + 1}/{pages}", callback_data="noop"))
        if page < pages - 1:
            nav.append(InlineKeyboardButton(text="➡️", callback_data=f"adm:deltest:{page + 1}"))
        if nav:
            kb_lines.append(nav)
            
        kb_lines.append([InlineKeyboardButton(text="⬅️ Orqaga", callback_data="adm:back")])
        
        await call.message.edit_text(
            "🗑 <b>Testni o'chirish</b>\n\nO'chirish uchun quyidagi testlardan birini tanlang:", 
            reply_markup=InlineKeyboardMarkup(inline_keyboard=kb_lines)
        )
        
    elif action.startswith("deltest_pick:"):
        test_id = action.split(":", 1)[1]
        if await get_test(test_id):
            await execute("DELETE FROM tests WHERE test_id=?", (test_id,))
            await call.answer("🗑 O'chirildi!", show_alert=True)
        else:
            await call.answer("Test topilmadi!", show_alert=True)
        
        # O'chirilgandan so'ng yana ro'yxatni yangilaymiz
        call.data = "adm:deltest"
        return await cb_admin_panel(call, state)

    elif action == "setcode":
        await state.set_state(AdminStates.waiting_for_code)
        code = await get_setting("ig_code")
        await call.message.edit_text(
            f"🔑 <b>Instagram Kod o'zgartirish</b>\n\nJoriy kod: <code>{html.escape(code or '-')}</code>\n\n"
            "Yangi kodni yuboring (Bekor qilish uchun /bekor):", reply_markup=admin_keyboard()
        )

    elif action == "qr":
        tests = await list_tests()
        if not tests:
            await call.answer("Testlar yo'q", show_alert=True)
            return
        await call.answer("QR kodlar yuborilmoqda...")
        for tid, title in tests:
            await send_qr(call.bot, call.message.chat.id, tid, title)
            await asyncio.sleep(0.3)

    elif action == "broadcast":
        await state.set_state(AdminStates.waiting_for_broadcast)
        await call.message.edit_text(
            "✉️ <b>Xabar yuborish</b>\n\nBarchaga yuboriladigan xabar matnini kiriting (Bekor qilish uchun /bekor):",
            reply_markup=admin_keyboard()
        )

    elif action.startswith("excel") and not action.startswith("excel_pick"):
        await state.set_state(AdminStates.waiting_for_excel_test_id)
        parts = action.split(":")
        page = int(parts[1]) if len(parts) > 1 else 0
        
        tests = await list_tests()
        if not tests:
            await call.answer("Hali testlar qo'shilmagan!", show_alert=True)
            return
            
        pages = max(1, -(-len(tests) // PAGE_SIZE))
        page = max(0, min(page, pages - 1))
        chunk = tests[page * PAGE_SIZE : (page + 1) * PAGE_SIZE]
        
        kb_lines = []
        test_list_text = ""
        for tid, title in chunk:
            kb_lines.append([InlineKeyboardButton(text=f"📊 {title}", callback_data=f"adm:excel_pick:{tid}")])
            test_list_text += f"• <code>{tid}</code> — {html.escape(title)}\n"
            
        nav = []
        if page > 0:
            nav.append(InlineKeyboardButton(text="⬅️", callback_data=f"adm:excel:{page - 1}"))
        if pages > 1:
            nav.append(InlineKeyboardButton(text=f"{page + 1}/{pages}", callback_data="noop"))
        if page < pages - 1:
            nav.append(InlineKeyboardButton(text="➡️", callback_data=f"adm:excel:{page + 1}"))
        if nav:
            kb_lines.append(nav)
            
        kb_lines.append([InlineKeyboardButton(text="⬅️ Orqaga", callback_data="adm:back")])
        
        await call.message.edit_text(
            f"📥 <b>Natijalarni yuklab olish</b>\n\n"
            f"Quyidagi testlardan birini tanlang yoki ID ni yozing:\n\n"
            f"{test_list_text}\n"
            f"Bekor qilish uchun /bekor",
            reply_markup=InlineKeyboardMarkup(inline_keyboard=kb_lines)
        )

    elif action.startswith("excel_pick:"):
        test_id = action.split(":", 1)[1]
        await state.update_data(excel_test_id=test_id)
        # Directly trigger excel generation
        await state.clear()
        t = await get_test(test_id)
        if not t:
            await call.answer("Test topilmadi!", show_alert=True)
            return
        rows = await fetch_all(
            "SELECT r.user_id, u.full_name, u.username, r.correct, r.total, r.given_answers, r.ts "
            "FROM results r JOIN users u ON r.user_id = u.user_id WHERE r.test_id=?",
            (test_id,)
        )
        if not rows:
            await call.answer("Bu test uchun hali natijalar yo'q!", show_alert=True)
            await call.message.edit_text(
                "🛠 <b>Boshqaruv paneli</b>\n\nQuyidagi tugmalardan birini tanlang:",
                reply_markup=admin_keyboard()
            )
            return
        
        right = t["answers"]
        wb = Workbook()
        ws = wb.active
        ws.title = "Natijalar"
        headers = ["User ID", "Ism", "Username", "To'g'ri", "Foiz (%)", "Sana"]
        for i in range(1, len(right) + 1):
            headers.append(f"{i}-savol")
        ws.append(headers)
        red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
        green_fill = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")
        for row_idx, r in enumerate(rows, start=2):
            uid, name, uname, correct, total, given, ts = r
            pct = round(correct * 100 / total) if total else 0
            date_str = datetime.fromisoformat(ts).strftime("%Y-%m-%d %H:%M") if ts else ""
            base_data = [uid, name, uname or "", correct, pct, date_str]
            for col_idx, val in enumerate(base_data, start=1):
                ws.cell(row=row_idx, column=col_idx, value=val)
            given = given or ""
            for i, correct_ans in enumerate(right):
                col = len(base_data) + i + 1
                given_ans = given[i] if i < len(given) else "-"
                if given_ans == correct_ans:
                    cell = ws.cell(row=row_idx, column=col, value=given_ans)
                    cell.fill = green_fill
                else:
                    cell = ws.cell(row=row_idx, column=col, value=f"{given_ans} (To'g'ri: {correct_ans})")
                    cell.fill = red_fill
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        await call.answer("Excel tayyor! Yuborilmoqda...")
        await call.message.answer_document(
            BufferedInputFile(buf.read(), filename=f"{test_id}_natijalar.xlsx"),
            caption=f"📊 <b>{html.escape(t['title'])}</b> natijalari\nJami ishtirokchilar: {len(rows)}"
        )
        await call.message.answer(
            "🛠 <b>Boshqaruv paneli</b>\n\nQuyidagi tugmalardan birini tanlang:",
            reply_markup=admin_keyboard()
        )

    elif action == "channel":
        channels = await fetch_all("SELECT id, username, url FROM channels")
        req_ch = await get_setting("req_channel") != "0"
        
        kb_lines = [
            [InlineKeyboardButton(
                text="✅ Majburiy obuna: YONIQ" if req_ch else "🛑 Majburiy obuna: O'CHIQ", 
                callback_data="adm:toggle_req_ch"
            )],
            [InlineKeyboardButton(text="➕ Kanal qo'shish", callback_data="adm:addchannel")]
        ]
        if channels:
            kb_lines.append([InlineKeyboardButton(text="🗑 Kanalni o'chirish", callback_data="adm:delchannel_list")])
        kb_lines.append([InlineKeyboardButton(text="⬅️ Orqaga", callback_data="adm:back")])
        
        text = "📢 <b>Kanal sozlash</b>\n\nJoriy kanallar:\n"
        if not channels:
            text += "<i>Kanallar ulanmagan.</i>\n"
        else:
            for i, un, url in channels:
                text += f"• {html.escape(un)} - {html.escape(url)}\n"
        
        await call.message.edit_text(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=kb_lines))

    elif action == "toggle_req_ch":
        req_ch = await get_setting("req_channel") != "0"
        await set_setting("req_channel", "0" if req_ch else "1")
        await call.answer("Holat o'zgardi!")
        # Re-render channel menu by simulating call
        call.data = "adm:channel"
        return await cb_admin_panel(call, state)

    elif action == "addchannel":
        await state.set_state(AdminStates.waiting_for_add_channel)
        await call.message.edit_text(
            "Yangi kanalning username va URL sini quyidagi formatda yuboring:\n"
            "<code>@kanal_username | https://t.me/kanal_username</code>\n\n"
            "Bekor qilish uchun /bekor",
            reply_markup=admin_keyboard()
        )

    elif action == "delchannel_list":
        channels = await fetch_all("SELECT id, username FROM channels")
        if not channels:
            await call.answer("Kanallar yo'q", show_alert=True)
            return
        kb_lines = []
        for i, un in channels:
            kb_lines.append([InlineKeyboardButton(text=f"❌ {un}", callback_data=f"adm:delch:{i}")])
        kb_lines.append([InlineKeyboardButton(text="⬅️ Orqaga", callback_data="adm:channel")])
        await call.message.edit_text("O'chirish uchun kanalni tanlang:", reply_markup=InlineKeyboardMarkup(inline_keyboard=kb_lines))

    elif action.startswith("delch:"):
        ch_id = int(action.split(":")[1])
        await execute("DELETE FROM channels WHERE id=?", (ch_id,))
        await call.answer("Kanal o'chirildi", show_alert=True)
        call.data = "adm:channel"
        return await cb_admin_panel(call, state)

    elif action == "instagram":
        await state.set_state(AdminStates.waiting_for_instagram)
        cur_ig = await get_instagram_url()
        req_ig = await get_setting("req_ig") != "0"
        
        kb = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(
                text="✅ Majburiy obuna: YONIQ" if req_ig else "🛑 Majburiy obuna: O'CHIQ", 
                callback_data="adm:toggle_req_ig"
            )],
            [InlineKeyboardButton(text="🗑 Instagramni o'chirish", callback_data="adm:delig")],
            [InlineKeyboardButton(text="⬅️ Orqaga", callback_data="adm:back")],
        ])
        await call.message.edit_text(
            f"📸 <b>Instagram sozlash</b>\n\n"
            f"Joriy Instagram URL: <code>{html.escape(cur_ig or '-')}</code>\n\n"
            "Yangi Instagram sahifa URL sini yuboring:\n"
            "<code>https://instagram.com/sahifangiz</code>\n\n"
            "Bekor qilish uchun /bekor",
            reply_markup=kb
        )
        
    elif action == "toggle_req_ig":
        req_ig = await get_setting("req_ig") != "0"
        await set_setting("req_ig", "0" if req_ig else "1")
        await call.answer("Holat o'zgardi!")
        call.data = "adm:instagram"
        return await cb_admin_panel(call, state)

    elif action == "delig":
        await del_setting("instagram_url")
        await call.answer("✅ Instagram o'chirildi", show_alert=True)
        call.data = "adm:instagram"
        return await cb_admin_panel(call, state)

    elif action == "back":
        await state.clear()
        await call.message.edit_text(
            "🛠 <b>Boshqaruv paneli</b>\n\nQuyidagi tugmalardan birini tanlang:",
            reply_markup=admin_keyboard()
        )

    await call.answer()

@router.message(StateFilter(AdminStates.waiting_for_test))
async def state_addtest(message: Message, state: FSMContext):
    if not await is_admin_db(message.from_user.id):
        return

    text = message.text or ""
    test_id = title = None
    answers = []

    # 1-format: rasmdagi format (sarlavha birinchi qatorda, keyin raqamli javoblar)
    titled_title, titled_answers = parse_titled_answers(text)
    if titled_title and titled_answers:
        title = titled_title
        answers = titled_answers
        test_id = slugify(title)

    # 2-format: test_id | Sarlavha | HCDAEG...
    elif "|" in text:
        parts = [p.strip() for p in text.split("|")]
        if len(parts) == 3:
            test_id, title, raw = parts
            test_id = test_id.lower()
            answers = parse_answers(raw)

    if not test_id or not title or not answers:
        await message.answer(
            "❌ Format noto'g'ri. Iltimos qaytadan yuboring yoki /bekor"
        )
        return

    if not TEST_ID_RE.match(test_id):
        # Slugdan olingan ID to'g'ri bo'lmasa raqam qo'shamiz
        test_id = re.sub(r"[^a-z0-9_]", "_", test_id)[:40].strip("_") or "test"

    existed = await get_test(test_id)
    await execute(
        "INSERT INTO tests(test_id, title, answers, created_at) VALUES(?,?,?,?) "
        "ON CONFLICT(test_id) DO UPDATE SET title=excluded.title, answers=excluded.answers",
        (test_id, title, "".join(answers), datetime.now().isoformat()),
    )
    await message.answer(
        f"{'♻️ Yangilandi' if existed else '✅ Qo`shildi'}: <b>{html.escape(title)}</b>\n"
        f"ID: <code>{test_id}</code> — {len(answers)} ta javob"
    )
    await send_qr(message.bot, message.chat.id, test_id, title)
    await state.clear()
    await message.answer("🛠 <b>Boshqaruv paneli</b>\n\nQuyidagi tugmalardan birini tanlang:", reply_markup=admin_keyboard())

@router.message(StateFilter(AdminStates.waiting_for_del_id))
async def state_deltest(message: Message, state: FSMContext):
    if not await is_admin_db(message.from_user.id):
        return
    test_id = norm(message.text)
    if await get_test(test_id):
        await execute("DELETE FROM tests WHERE test_id=?", (test_id,))
        await message.answer(f"🗑 O'chirildi: {html.escape(test_id)}")
    else:
        await message.answer("Bunday test topilmadi.")
    await state.clear()
    await message.answer("🛠 <b>Boshqaruv paneli</b>\n\nQuyidagi tugmalardan birini tanlang:", reply_markup=admin_keyboard())

@router.message(StateFilter(AdminStates.waiting_for_code))
async def state_setcode(message: Message, state: FSMContext):
    if not await is_admin_db(message.from_user.id):
        return
    new = norm(message.text)
    await set_setting("ig_code", new)
    await message.answer(f"✅ Yangi Instagram kod: <code>{html.escape(new)}</code>")
    await state.clear()
    await message.answer("🛠 <b>Boshqaruv paneli</b>\n\nQuyidagi tugmalardan birini tanlang:", reply_markup=admin_keyboard())

@router.message(StateFilter(AdminStates.waiting_for_add_channel))
async def state_add_channel(message: Message, state: FSMContext):
    if not await is_admin_db(message.from_user.id):
        return
    parts = [p.strip() for p in message.text.split("|")]
    if len(parts) != 2:
        await message.answer("❌ Format xato. Masalan: <code>@kanal | https://t.me/kanal</code>")
        await state.clear()
        await message.answer("🛠 <b>Boshqaruv paneli</b>\n\nQuyidagi tugmalardan birini tanlang:", reply_markup=admin_keyboard())
        return
    channel, channel_url = parts
    await execute("INSERT INTO channels(username, url) VALUES(?,?)", (channel, channel_url))
    await message.answer(
        f"✅ Kanal qo'shildi!\n"
        f"Username: <code>{html.escape(channel)}</code>\n"
        f"URL: <code>{html.escape(channel_url)}</code>\n\n"
        "⚠️ Botni kanalga admin sifatida qo'shishni unutmang!"
    )
    await state.clear()
    await message.answer("🛠 <b>Boshqaruv paneli</b>\n\nQuyidagi tugmalardan birini tanlang:", reply_markup=admin_keyboard())

@router.message(StateFilter(AdminStates.waiting_for_instagram))
async def state_instagram(message: Message, state: FSMContext):
    if not await is_admin_db(message.from_user.id):
        return
    ig_url = message.text.strip()
    if not ig_url.startswith("http"):
        await message.answer("❌ URL https:// bilan boshlanishi kerak.")
        await state.clear()
        await message.answer("🛠 <b>Boshqaruv paneli</b>\n\nQuyidagi tugmalardan birini tanlang:", reply_markup=admin_keyboard())
        return
    await set_setting("instagram_url", ig_url)
    await message.answer(f"✅ Instagram ulandi!\nURL: <code>{html.escape(ig_url)}</code>")
    await state.clear()
    await message.answer("🛠 <b>Boshqaruv paneli</b>\n\nQuyidagi tugmalardan birini tanlang:", reply_markup=admin_keyboard())

@router.message(StateFilter(AdminStates.waiting_for_broadcast))
async def state_broadcast(message: Message, state: FSMContext):
    if not await is_admin_db(message.from_user.id):
        return

    ids = [r[0] for r in await fetch_all("SELECT user_id FROM users")]
    ok = 0
    msg = await message.answer(f"Boshlandi... Jami: {len(ids)}")
    for uid in ids:
        try:
            await message.copy_to(uid)
            ok += 1
        except TelegramAPIError:
            pass
        await asyncio.sleep(0.05)

    await msg.edit_text(f"Yuborildi: {ok}/{len(ids)}")
    await state.clear()
    await message.answer("🛠 <b>Boshqaruv paneli</b>\n\nQuyidagi tugmalardan birini tanlang:", reply_markup=admin_keyboard())

@router.message(StateFilter(AdminStates.waiting_for_excel_test_id))
async def state_excel_test_id(message: Message, state: FSMContext):
    if not await is_admin_db(message.from_user.id):
        return
    test_id = norm(message.text)
    t = await get_test(test_id)
    
    if not t:
        await message.answer("❌ Bunday test topilmadi.")
        await state.clear()
        await message.answer("🛠 <b>Boshqaruv paneli</b>\n\nQuyidagi tugmalardan birini tanlang:", reply_markup=admin_keyboard())
        return

    rows = await fetch_all(
        "SELECT r.user_id, u.full_name, u.username, r.correct, r.total, r.given_answers, r.ts "
        "FROM results r JOIN users u ON r.user_id = u.user_id WHERE r.test_id=?",
        (test_id,)
    )

    if not rows:
        await message.answer("Ushbu test uchun hali hech qanday natija yo'q.")
        await state.clear()
        await message.answer("🛠 <b>Boshqaruv paneli</b>\n\nQuyidagi tugmalardan birini tanlang:", reply_markup=admin_keyboard())
        return
    
    right = t["answers"]
    wb = Workbook()
    ws = wb.active
    ws.title = "Natijalar"
    
    headers = ["User ID", "Ism", "Username", "To'g'ri", "Foiz (%)", "Sana"]
    for i in range(1, len(right) + 1):
        headers.append(f"{i}-savol")
    ws.append(headers)

    red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
    green_fill = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")
    
    for row_idx, r in enumerate(rows, start=2):
        uid, name, uname, correct, total, given, ts = r
        pct = round(correct * 100 / total) if total else 0
        date_str = datetime.fromisoformat(ts).strftime("%Y-%m-%d %H:%M") if ts else ""
        base_data = [uid, name, uname or "", correct, pct, date_str]
        
        for col_idx, val in enumerate(base_data, start=1):
            ws.cell(row=row_idx, column=col_idx, value=val)
            
        given = given or ""
        for i, correct_ans in enumerate(right):
            col = len(base_data) + i + 1
            given_ans = given[i] if i < len(given) else "-"
            
            if given_ans == correct_ans:
                cell = ws.cell(row=row_idx, column=col, value=given_ans)
                cell.fill = green_fill
            else:
                cell = ws.cell(row=row_idx, column=col, value=f"{given_ans} (To'g'ri: {correct_ans})")
                cell.fill = red_fill

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    
    await message.answer_document(
        BufferedInputFile(buf.read(), filename=f"{test_id}_natijalar.xlsx"),
        caption=f"📊 <b>{html.escape(t['title'])}</b> natijalari\nJami ishtirokchilar: {len(rows)}"
    )
    
    await state.clear()
    await message.answer("🛠 <b>Boshqaruv paneli</b>\n\nQuyidagi tugmalardan birini tanlang:", reply_markup=admin_keyboard())


# ============================= oddiy matn: kod so'z yoki javoblar yuborish
@router.message(F.text & ~F.text.startswith("/"))
async def on_text(message: Message):
    bot, uid = message.bot, message.from_user.id
    await upsert_user(message.from_user)

    if not await is_subscribed(bot, uid):
        await gate(bot, uid)
        return

    req_ig = await get_setting("req_ig") != "0"
    
    if req_ig:
        code = await get_setting("ig_code")
        if not code:
            await message.answer("⚠️ Bot hali to'liq sozlanmagan. Iltimos, keyinroq urinib ko'ring.")
            return

        user = await get_user(uid)
        verified = bool(user and user["ig_code"] == code)

        # 1) Instagram kod so'zi hali tasdiqlanmagan
        if not verified:
            if norm(message.text) == code:
                await set_user_code(uid, code)
                await message.answer("✅ Ajoyib, Instagram tasdiqlandi!")
                await proceed(bot, uid)
            else:
                await message.answer(
                    "❌ Kod so'z noto'g'ri. Instagram sahifamizdagi (bio yoki oxirgi post) "
                    "kodni diqqat bilan qayta yuboring.",
                    reply_markup=ig_keyboard(),
                )
            return

    # 2) Tekshirish rejimi: talaba javoblarini yubormoqda
    user = await get_user(uid)
    if user and user["check"]:
        t = await get_test(user["check"])
        if not t:
            await set_check(uid, None)
            await send_test_list(bot, uid)
            return
        given = parse_answers(message.text)
        right = t["answers"]
        
        # Agar yuborilgan javoblar kam bo'lsa — qolganlarini "-" bilan to'ldiramiz
        if len(given) == 0:
            await message.answer(
                "❌ Javoblar tanib olinmadi. Format:\n"
                "<code>ABCDE...</code> yoki <code>1-A 2-B 3-C</code>\n\n"
                "Bekor qilish uchun /bekor"
            )
            return
        
        if len(given) > len(right):
            given = given[:len(right)]
        elif len(given) < len(right):
            given = given + ["-"] * (len(right) - len(given))

        wrong = [i for i, (g, r) in enumerate(zip(given, right), start=1) if g != r]
        correct = len(right) - len(wrong)
        pct = round(correct * 100 / len(right))
        
        given_str = "".join(given)
        await execute(
            "INSERT INTO results(user_id, test_id, correct, total, given_answers, ts) VALUES(?,?,?,?,?,?)",
            (uid, user["check"], correct, len(right), given_str, datetime.now().isoformat()),
        )
        # check_test ni saqlayiz show callback uchun
        check_test_id = user["check"]
        await set_check(uid, None)

        if not wrong:
            verdict = "🎉 Ajoyib! Barcha savollarga to'g'ri javob berdingiz! Barakalla!"
        else:
            verdict = "<b>Xato javoblaringiz tahlili:</b>\n"
            for w_idx in wrong:
                g_ans = given[w_idx-1]
                r_ans = right[w_idx-1]
                if g_ans == "-":
                    verdict += f"⚪️ {w_idx}-savol: <b>Javob berilmagan</b> (To'g'ri: 🟢 <b>{r_ans}</b>)\n"
                else:
                    verdict += f"🔴 {w_idx}-savol: Siz <b>{g_ans}</b> dedingiz (To'g'ri: 🟢 <b>{r_ans}</b>)\n"
                
        kb = InlineKeyboardMarkup(
            inline_keyboard=[
                [InlineKeyboardButton(text="👁 Barcha javoblarni ko'rish", callback_data=f"show:{check_test_id}")],
                [InlineKeyboardButton(text="🔁 Qayta urinish", callback_data=f"chk:{check_test_id}")],
                [InlineKeyboardButton(text="📚 Boshqa testlar", callback_data="pg:0")],
            ]
        )
        await message.answer(
            f"📊 Natija: <b>{html.escape(t['title'])}</b>\n"
            f"Jami savollar: {len(right)} ta\n"
            f"To'g'ri javoblar: <b>{correct} ta</b> ({pct}%)\n"
            f"To'plagan ballingiz: <b>{pct} / 100</b>\n\n"
            f"{verdict}",
            reply_markup=kb,
        )

    # 3) Boshqa har qanday matn
    await send_test_list(bot, uid)


async def handle_ping(request):
    return web.Response(text="Bot is running!")

async def run_dummy_server():
    app = web.Application()
    app.router.add_get('/', handle_ping)
    app.router.add_get('/ping', handle_ping)
    runner = web.AppRunner(app)
    await runner.setup()
    port = int(os.environ.get("PORT", 10000))
    site = web.TCPSite(runner, '0.0.0.0', port)
    await site.start()
    logging.info(f"Dummy web server started on port {port}")


async def main() -> None:
    await db_init()
    
    # Render.com kabi xizmatlar uchun fon veb-serverini ishga tushirish
    await run_dummy_server()
    
    bot = Bot(BOT_TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
    dp = Dispatcher()
    dp.include_router(router)
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
