# Скопируй этот файл в config.py и заполни своими данными:
# cp config.example.py config.py

# TikTok аккаунт для мониторинга (без @)
USERNAME = "tiktok_username"

# Токен Telegram-бота (получить у @BotFather)
TELEGRAM_TOKEN = "your_telegram_bot_token"

# Твой Telegram Chat ID (узнать у @userinfobot)
CHAT_ID = 123456789

# Прокси (опционально, оставь None если не нужен)
# Формат: "socks5://user:password@host:port"
PROXY = None

# Интервал проверки по умолчанию (в часах)
INTERVAL_HOURS = 5

# Пути к файлам (можно не менять)
EXCEL_FILE = "tiktok_reposts.xlsx"
SEEN_FILE = "seen_ids.json"
DELETED_FILE = "deleted_ids.json"
PROFILE_DIR = "browser_profile"
