# Weather Screenshot Tool

Инструмент для автоматического создания скриншотов прогноза погоды с сайта Gismeteo и формирования презентации PowerPoint.

## Возможности

- Автоматический сбор прогнозов погоды для Москвы:
  - Текущая погода
  - Прогноз на завтра
  - Прогноз на 3 дня
  - Прогноз на 10 дней
- Создание скриншотов с оптимальными отступами
- Автоматическая генерация презентации PowerPoint

## Требования

- Python 3.7+
- Playwright
- python-pptx
- Pillow

## Установка

1. Клонируйте репозиторий
2. Установите зависимости:
```bash
pip install -r requirements.txt
playwright install chromium
```

## Использование

```bash
python weather_screenshot.py
```

Скрипт создаст:
- Папку `screenshots` с изображениями прогноза погоды
- Файл `weather_report.pptx` с презентацией

## Структура проекта

- `weather_screenshot.py` - основной скрипт
- `requirements.txt` - зависимости проекта
- `screenshots/` - папка для сохранения скриншотов
- `weather_report.pptx` - итоговая презентация 