import os
from telethon import TelegramClient

# Put your Telegram API ID/HASH here or set environment variables.
API_ID = int(os.getenv("API_ID", "38931809"))
API_HASH = os.getenv("API_HASH", "")
SESSION_NAME = os.getenv("TG_SESSION", "autoforward_user")

if not API_HASH:
    API_HASH = input("Telegram API HASH: ").strip()
phone = input("Telegram phone number (e.g. +8801XXXXXXXXX): ").strip()

client = TelegramClient(SESSION_NAME, API_ID, API_HASH)

async def main():
    await client.start(phone=phone)
    me = await client.get_me()
    print("\nLOGIN SUCCESS")
    print("Name:", getattr(me, "first_name", ""))
    print("User ID:", me.id)
    print("Session file created:", SESSION_NAME + ".session")

with client:
    client.loop.run_until_complete(main())
