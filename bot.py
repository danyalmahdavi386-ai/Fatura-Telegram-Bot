from __future__ import annotations

import logging
import os
import sys
from typing import Optional

from telegram import BotCommand, Update
from telegram.ext import Application, ApplicationBuilder, CommandHandler, ContextTypes

BUILD_VERSION = "telegram-diagnostic-v1"

TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "").strip()
ALLOWED_USER_ID_RAW = os.getenv("ALLOWED_USER_ID", "").strip()
ALLOWED_USER_ID: Optional[int] = (
    int(ALLOWED_USER_ID_RAW) if ALLOWED_USER_ID_RAW else None
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
logger = logging.getLogger("telegram_diagnostic")


def authorized(user_id: Optional[int]) -> bool:
    return ALLOWED_USER_ID is None or user_id == ALLOWED_USER_ID


async def start_command(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
) -> None:
    user_id = update.effective_user.id if update.effective_user else None

    if not authorized(user_id):
        await update.effective_message.reply_text(
            "بات فعال است، اما شناسه شما مجاز نیست.\n"
            f"شناسه شما: {user_id}"
        )
        return

    await update.effective_message.reply_text(
        "✅ بات تشخیصی تلگرام فعال است.\n\n"
        f"نسخه: {BUILD_VERSION}\n"
        f"شناسه شما: {user_id}\n\n"
        "دستور بعدی: /ping"
    )


async def ping_command(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
) -> None:
    user_id = update.effective_user.id if update.effective_user else None
    await update.effective_message.reply_text(
        "🏓 PONG\n"
        f"نسخه: {BUILD_VERSION}\n"
        f"شناسه شما: {user_id}"
    )


async def version_command(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
) -> None:
    await update.effective_message.reply_text(
        f"نسخه فعال: {BUILD_VERSION}"
    )


async def myid_command(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
) -> None:
    user_id = update.effective_user.id if update.effective_user else None
    await update.effective_message.reply_text(
        f"شناسه تلگرام شما: {user_id}"
    )


async def post_init(application: Application) -> None:
    # Polling and webhook cannot be used together.
    await application.bot.delete_webhook(drop_pending_updates=True)

    await application.bot.set_my_commands(
        [
            BotCommand("start", "آزمایش شروع"),
            BotCommand("ping", "آزمایش پاسخ"),
            BotCommand("version", "نمایش نسخه"),
            BotCommand("myid", "نمایش شناسه"),
        ]
    )

    me = await application.bot.get_me()
    logger.info(
        "DIAGNOSTIC_BOT_STARTED | version=%s | bot=@%s | id=%s",
        BUILD_VERSION,
        me.username,
        me.id,
    )


async def error_handler(
    update: object,
    context: ContextTypes.DEFAULT_TYPE,
) -> None:
    logger.error(
        "DIAGNOSTIC_ERROR | type=%s | message=%s",
        type(context.error).__name__ if context.error else "Unknown",
        str(context.error)[:800] if context.error else "No error details",
    )


def main() -> None:
    if not TELEGRAM_BOT_TOKEN:
        raise RuntimeError(
            "TELEGRAM_BOT_TOKEN is missing in Railway Variables."
        )

    if ALLOWED_USER_ID_RAW and not ALLOWED_USER_ID_RAW.isdigit():
        raise RuntimeError(
            "ALLOWED_USER_ID must contain digits only."
        )

    application = (
        ApplicationBuilder()
        .token(TELEGRAM_BOT_TOKEN)
        .post_init(post_init)
        .build()
    )

    application.add_handler(CommandHandler("start", start_command))
    application.add_handler(CommandHandler("ping", ping_command))
    application.add_handler(CommandHandler("version", version_command))
    application.add_handler(CommandHandler("myid", myid_command))
    application.add_error_handler(error_handler)

    logger.info(
        "DIAGNOSTIC_PROCESS_STARTING | version=%s",
        BUILD_VERSION,
    )

    application.run_polling(
        drop_pending_updates=True,
        allowed_updates=Update.ALL_TYPES,
    )


if __name__ == "__main__":
    main()
