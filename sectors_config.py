# sectors_config.py — Конфигурация всех секторов макросправки
#
# КАК ЗАПОЛНИТЬ:
# 1. N8N_BASE_URL — URL вашего n8n (например http://192.168.1.55:5678)
# 2. N8N_API_KEY  — Settings → API → Create API Key
# 3. Для каждого сектора:
#    - workflow_id  → открыть воркфлоу в n8n, скопировать из URL
#    - datatable_id → открыть Data tables, скопировать из URL таблицы
#    - template     → путь к файлу шаблона Excel
#    - output       → путь к файлу результата

N8N_BASE_URL = "https://n8n.finreg.kz"   # ← заменить на IP вашего n8n
N8N_API_KEY  = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiI4YjkxNWY4OC05ZGQxLTQ2NTUtODI1NS1hOGUyMmM1OTgwMTAiLCJpc3MiOiJuOG4iLCJhdWQiOiJwdWJsaWMtYXBpIiwianRpIjoiNzgxNzEyOGYtMjczMy00NWRhLWE2MTQtMjdiNmVlOTk4ZGMwIiwiaWF0IjoxNzczMjEzMjg0fQ._icxvkecZkowF8yuqmG3UDxn6NdjDEdBYgwgd_uauQ4"       # ← заменить на реальный ключ

SECTORS = {
    "monetary": {
        "name":         "Монетарный сектор",
        "workflow_id":  "nmz45DFpD6Ogf...",           # ← заполнить
        "datatable_id": "zKR6BRvlrBMOL1yC",        # ← заполнить
        "template":     "excel_templates/Шаблон_Монетарный сектор.xlsx",
        "output":       "output/Результат_Монетарный_Сектор.xlsx",
        # header_row, mapping_col, potential_cols — определяются автоматически
    },
    "real": {
        "name":         "Реальный сектор",
        "workflow_id":  "nmz45DFpD6Ogf...",              # ← заполнить
        "datatable_id": "IpqfbaX2s0fjL1Kg",             # уже известен
        "template":     "excel_templates/Шаблон_Реальный сектор.xlsx",
        "output":       "output/Результат_Реальный_Сектор.xlsx",
    },
    "fiscal": {
        "name":         "Фискальный сектор",
        "workflow_id":  "WZRWmlcjTG...",            # ← заполнить
        "datatable_id": "EdlrIGcN8qr4wcpM",          # ← заполнить
        "template":     "excel_templates/Шаблон_Фискальный сектор.xlsx",
        "output":       "output/Результат_Фискальный_Сектор.xlsx",
    },
    "external": {
        "name":         "Внешний сектор",
        "workflow_id":  "6xLd_lohR_sn6Wj...",          # ← заполнить
        "datatable_id": "Uo7YBGwIY5a5Kq2F",        # ← заполнить
        "template":     "excel_templates/Шаблон_Внешний сектор.xlsx",
        "output":       "output/Результат_Внешний_Сектор.xlsx",
    },
    "social": {
        "name":         "Социальный сектор",
        "workflow_id":  "WuDhZ0sjHLQhS...",            # ← заполнить
        "datatable_id": "UpIG2HHDZRnLd8rB",          # ← заполнить
        "template":     "excel_templates/Шаблон_Социальный сектор.xlsx",
        "output":       "output/Результат_Социальный_Сектор.xlsx",
    },
}

# Итоговый файл со всеми 5 секторами как листами
COMBINED_OUTPUT = "output/Результат_Итоговый.xlsx"
