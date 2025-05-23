# Website Checker

## Описание
Website Checker — это Python-приложение с графическим интерфейсом (Tkinter), которое парсит сайты из текстового файла через Яндекс.Карты с использованием Selenium. Программа проверяет наличие указанных сайтов в результатах поиска, анализирует их статус и количество отзывов, а затем сохраняет подходящие сайты в файл `Result_yandex_plus.txt`. Поддерживается остановка и возобновление парсинга с сохранением прогресса.

## Основные возможности
- **Загрузка списка сайтов** из текстового файла (`.txt`) через интерфейс.
- **Парсинг через Яндекс.Карты**:
  - Поиск сайтов с использованием Selenium и stealth-режима для обхода защиты.
  - Извлечение данных (URL, статус, отзывы) с помощью BeautifulSoup.
- **Фильтрация сайтов**:
  - Проверка совпадения URL с искомым сайтом.
  - Условия: минимум 2 отзыва, исключение закрытых или неработающих компаний.
- **Сохранение результатов**:
  - Подходящие сайты записываются в `Result_yandex_plus.txt`.
  - Прогресс сохраняется в `cond.json` для возобновления после остановки.
- **Графический интерфейс**:
  - Выбор файла, запуск и остановка парсинга.
  - Отображение логов и счётчика найденных сайтов в реальном времени.
- **Обработка ошибок**: Сохранение скриншота и трассировки в случае сбоя.

## Требования
- Python 3.x
- Библиотеки:
  - `selenium` — для автоматизации браузера.
  - `selenium-stealth` — для обхода защиты от ботов.
  - `beautifulsoup4` — для парсинга HTML.
  - `chromedriver-autoinstaller` — для автоматической установки ChromeDriver.
  - `tkinter` — для графического интерфейса (обычно встроен в Python).
- Установленный браузер Google Chrome.
## 
Проект разработан в рамках коммерческого заказа на фрилансе.
