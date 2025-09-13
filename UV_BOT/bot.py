from telegram import Update
from telegram.ext import Application, MessageHandler, filters, CallbackContext
from config import BOT_TOKEN
from message_handlers import MessageHandlers


# Создаем экземпляр обработчика сообщений
message_handlers = MessageHandlers()

async def message_handler(update: Update, context: CallbackContext):

    await message_handlers.handle_message(update, context)


def main():

    application = Application.builder().token(BOT_TOKEN).build()
    application.add_handler(MessageHandler(filters.ALL, message_handler))
    application.run_polling()


if __name__ == '__main__':
    main()