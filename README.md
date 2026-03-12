@echo off
cd /d "C:\Users\Aizhan.kenesh\Desktop\Macrospravka"

call venv\Scripts\activate.bat

venv\Scripts\python app.py

pause



# ИНСТРУКЦИЯ ПО РАЗВЁРТЫВАНИЮ
# ===========================
# Макросправка — сервис автоматического формирования Excel

## 1. СТРУКТУРА ПАПОК

```
macrospravka/
├── app.py                        ← Запускаемый Flask сервис
├── generator.py                  ← Универсальный генератор Excel
├── sectors_config.py             ← Конфигурация (URL n8n, ID таблиц)
├── requirements.txt              ← Python-зависимости
├── n8n_orchestrator_workflow.json← Импортировать в n8n
├── output/                       ← Папка для готовых Excel файлов
└── excel_templates/              ← Положить сюда 5 шаблонов Excel
    ├── Шаблон_Монетарный сектор.xlsx
    ├── Шаблон_Реальный сектор.xlsx
    ├── Шаблон_Фискальный сектор.xlsx
    ├── Шаблон_Внешний сектор.xlsx
    └── Шаблон_Социальный сектор.xlsx
```

## 2. НАСТРОЙКА

### sectors_config.py — заполнить:
- N8N_BASE_URL = "http://ВАШ_СЕРВЕР_N8N:5678"
- N8N_API_KEY  — взять в n8n: Settings → API → Create API Key
- Для каждого сектора — workflow_id и datatable_id (из URL в n8n)

### Пример получения datatable_id:
  В n8n → Data tables → нажать на таблицу
  URL будет: /projects/.../datatables/ВОТ_ЭТО_И_ЕСТЬ_ID

### Пример получения workflow_id:
  В n8n → открыть воркфлоу
  URL будет: /workflow/ВОТ_ЭТО_И_ЕСТЬ_ID

## 3. УСТАНОВКА И ЗАПУСК

```bash
# Установить зависимости (один раз)
pip install -r requirements.txt

# Запустить сервис
python app.py
```

Сервис запустится на http://localhost:8000
Если нужен другой порт — изменить в последней строке app.py

## 4. НАСТРОЙКА n8n

### Импорт Orchestrator воркфлоу:
1. В n8n: New Workflow → Import from JSON
2. Вставить содержимое n8n_orchestrator_workflow.json
3. Заменить все ID воркфлоу и DataTable на свои
4. В ноде "Send to Python Service" указать URL сервера:
   http://ВАШ_КОМПЬЮТЕР_IP:8000/api/generate

### В каждом суб-воркфлоу (Monetary/Real/...):
- Добавить в конец ноду "Respond to Workflow" (для корректного ожидания)

## 5. ПРАВА ДОСТУПА ДЛЯ ДРУГИХ ПОЛЬЗОВАТЕЛЕЙ

### Вариант А — запустить как сетевой сервис:
```bash
# Запуск с доступом из сети (заменить 0.0.0.0)
python app.py
# Пользователи открывают: http://ВАШ_IP:8000
```

### Вариант Б — создать ярлык/bat-файл для пользователя:
```
start_macrospravka.bat:
  cd C:\macrospravka
  python app.py
  start http://localhost:8000
```

## 6. СЦЕНАРИЙ РАБОТЫ ПОЛЬЗОВАТЕЛЯ

1. Открыть браузер → http://сервер:8000
2. Нажать кнопку "Сформировать макросправку"
3. Наблюдать прогресс по секторам (зелёные галочки)
4. Нажать "⬇ Скачать Excel" для каждого готового сектора

## 7. КАК ДОБАВИТЬ НОВЫЙ СЕКТОР

1. В sectors_config.py — добавить новый элемент в SECTORS
2. Положить шаблон в excel_templates/
3. В n8n Orchestrator добавить ещё один Execute Workflow
4. В ноде "Pack All Data" добавить поле нового сектора
5. В веб-интерфейсе (app.py, HTML_TEMPLATE) добавить карточку

## 8. ОТПРАВКА ПО EMAIL (через n8n)

После ноды "Send to Python Service" добавить в n8n:
- Ноду "Wait" (ждать 60 секунд — время генерации)
- Ноду "Send Email" с вложением файлов из папки output/

## 9. ИНТЕГРАЦИЯ С BITRIX24

В n8n после генерации можно добавить HTTP Request к API Bitrix24:
  POST https://ВАШ_ПОРТАЛ.bitrix24.ru/rest/1/ТОКЕН/disk.folder.uploadfile
  с файлом из output/
