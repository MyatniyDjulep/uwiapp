FROM python:3.11-slim

# Установка системных утилит и LibreOffice для конвертации DOCX -> PDF
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    libreoffice-writer \
    fontconfig \
    && rm -rf /var/lib/apt/lists/*

# Копируем Windows шрифты (Times New Roman, Arial, Calibri) для LibreOffice
COPY fonts/ /usr/share/fonts/truetype/msttcorefonts/
RUN fc-cache -f -v

# Рабочая директория
WORKDIR /app

# Установка Python зависимостей
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Копирование исходного кода и шаблонов
COPY . .

# Запуск основного файла (он запустит FastAPI и фоновые потоки бота)
CMD ["python", "webterminator.py"]
