# telegram-chat-analytics

Инструкция для MacOS

# Получить переписки

У телеграма есть два клиента под MacOS и для того что бы сделать бэкап всех чатов нужен скорее всего не тот, что у вас уже стоит. Поэтому качаем ещё один официальный клиент телеги:

https://desktop.telegram.org/ — для того что бы сделать резервную копию всех переписок.

Settings -> Advanced -> Export Telegram data

Снимаем все галочки кроме Personal chats / Bot chats / Private groups (все сообщения, не только свои).

Формат — JSON

это будет долгооооо

# Запускаем скрипт

git clone ...

python3 -m venv venv
source venv/bin/activate

pip3 install -r requirements.txt

кладём в папку рядом со скриптом переписки result.json

заходим в spreadsheet.py / evolution.py и меняем MY_ID на ваш id

python3 spreadsheet.py
python3 evolution.py

изучаем friends_metrics_2024.xlsx
открываем chat_evolution.html и смотрим в динамике по месяцам
