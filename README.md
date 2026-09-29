# ipo

## Региональные ипотечные таблицы — сборка в Google Colab

[![Открыть в Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/xchekh/ipo/blob/main/%D0%A0%D0%B5%D0%B3%D0%B8%D0%BE%D0%BD%D0%B0%D0%BB%D1%8C%D0%BD%D1%8B%D0%B5%20%D1%82%D0%B0%D0%B1%D0%BB%D0%B8%D1%86%D1%8B%20%E2%80%94%20Colab.ipynb)

Ссылка: https://colab.research.google.com/github/xchekh/ipo/blob/main/%D0%A0%D0%B5%D0%B3%D0%B8%D0%BE%D0%BD%D0%B0%D0%BB%D1%8C%D0%BD%D1%8B%D0%B5%20%D1%82%D0%B0%D0%B1%D0%BB%D0%B8%D1%86%D1%8B%20%E2%80%94%20Colab.ipynb

Ноутбук собирает из сырой выгрузки `Статистические ряды_регионы.xlsx` файл с листами
«кол-во», «объем», «ДДУ» и «Характеристики кредитов» с тикерами как у Дани.

1. Откройте ссылку и выберите **Среда выполнения → Выполнить все**.
2. Загрузите свежую выгрузку `Статистические ряды_регионы.xlsx`.
3. Через 2–3 минуты готовый xlsx скачается сам.

Репозиторий приватный: ссылка открывается только у тех, у кого есть доступ к нему
(в Colab нужно разрешить доступ к приватным репозиториям GitHub). Остальным можно дать
копию ноутбука через Google Drive: в Colab **Файл → Загрузить блокнот**, затем «Поделиться».

Тот же код без Colab: `python build_regional_tables.py "Статистические ряды_регионы.xlsx"` (нужен `pip install openpyxl`).
