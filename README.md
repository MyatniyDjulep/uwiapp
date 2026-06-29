# 🤖 Конвейер «Терминатор» v4.0 — Telegram Web App

## Структура проекта

```
terminator_webapp/
├── bot.py                  ← Основной бот (запускать этот файл)
├── .env                    ← Секреты (создать по образцу .env.example)
├── .env.example            ← Шаблон переменных окружения
├── webapp/
│   └── index.html          ← Telegram Mini App (выложить на хостинг)
├── templates/              ← .docx шаблоны документов
├── static_attachments/     ← Статические вложения к письмам
├── READY_DOCUMENTS/        ← Готовые документы (авто)
└── FINAL_ARCHIVE/          ← Архив всех пакетов (авто)
```

---

## Развёртывание Web App (шаг 1)

Файл `webapp/index.html` должен быть доступен по HTTPS.

### Вариант А — GitHub Pages (бесплатно, быстро)
```
1. Создайте репозиторий на GitHub
2. Положите webapp/index.html как index.html в корень репо
3. Включите GitHub Pages в настройках (Settings → Pages → main branch)
4. URL будет: https://<username>.github.io/<repo>/
```

### Вариант Б — VPS с nginx
```nginx
server {
    listen 443 ssl;
    server_name your-domain.com;

    ssl_certificate     /etc/letsencrypt/live/your-domain.com/fullchain.pem;
    ssl_certificate_key /etc/letsencrypt/live/your-domain.com/privkey.pem;

    location /webapp/ {
        root /var/www/;
        index index.html;
    }
}
```
Положите index.html в `/var/www/webapp/`.

### Вариант В — Cloudflare Pages (бесплатно)
Drag & drop папку webapp/ на dash.cloudflare.com → Pages

---

## Настройка .env (шаг 2)

```env
OPENROUTER_API_KEY=sk-or-v1-...
TELEGRAM_BOT_TOKEN=...
MY_CHAT_ID=322318469

SECOND_BOT_TOKEN=...
SECOND_BOT_CHAT_ID=322318469

SMTP_SERVER=smtp.mail.ru
SMTP_PORT=465
IMAP_SERVER=imap.mail.ru
IMAP_PORT=993
EMAIL_SENDER=maritimenvrsk@mail.ru
EMAIL_PASSWORD=...

PROXY_URL=http://login:pass@host:port

# URL вашего задеплоенного webapp/index.html (ОБЯЗАТЕЛЬНО HTTPS)
WEBAPP_URL=https://your-domain.com/webapp/index.html
```

---

## Регистрация Web App в BotFather (шаг 3)

```
1. /newapp — создать новое Menu Button App
2. Выбрать бота
3. Вставить WEBAPP_URL
4. Или использовать как inline-кнопку (бот делает это автоматически)
```

---

## Установка зависимостей

```bash
pip install pytelegrambotapi openai httpx pymupdf Pillow openpyxl python-docx \
            python-dotenv pystray docx2pdf pythoncom imaplib2
```

> ⚠️ `pythoncom` и `docx2pdf` только для Windows (конвертация DOCX→PDF через MS Word)

---

## Запуск

```bash
python bot.py
```

---

## Как работает Web App

```
Входящее письмо / PDF
        │
        ▼
  ИИ-парсер (Gemini)
        │
        ▼
  Кнопка в боте:
  [⚡ Открыть редактор судна]
        │
        ▼
  Telegram Mini App открывается внутри Telegram
  Пользователь:
   ✅ Проверяет все поля
   ✅ Исправляет при необходимости (тапает на поле → модалка)
   ✅ Выбирает компанию-контрагента
   ✅ Задаёт стартовый номер документов
   ✅ Нажимает "ГЕНЕРИРОВАТЬ"
        │
        ▼
  Данные через sendData() → бот получает web_app_data
        │
        ▼
  generate_all_documents() — генерация DOCX + PDF
        │
        ▼
  ZIP-архив в чат + выбор файлов для email-отправки
```

---

## Что передаёт Web App в бот (JSON)

```json
{
  "action": "generate",
  "vessel_name": "OCEAN PIONEER",
  "ship_type": "BULK CARRIER",
  "imo": "9876543",
  "flag": "PANAMA",
  "owner": "PACIFIC SHIPPING CO.",
  "master_name": "JOHN SMITH",
  "loa": "190.0",
  "breadth": "28.5",
  "draft_fore": "7.2",
  "draft_aft": "7.8",
  "act_number": "27/06/01",
  "diving_date": "27.06.2026",
  "date_start": "26.06.2026",
  "date_end": "30.06.2026",
  "current_date": "27.06.2026",
  "wetted_surface": "3820.5",
  "propeller_area": "60.0",
  "company_key": "almar",
  "start_num": 146
}
```
