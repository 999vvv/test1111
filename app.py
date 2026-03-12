# app.py
# Flask-сервис: веб-интерфейс для пользователя + API для n8n
#
# Запуск:  python app.py
# Откроется по адресу:  http://localhost:8000
#
# Эндпоинты:
#   GET  /                    — главная страница (веб-интерфейс)
#   POST /api/trigger         — запустить все воркфлоу n8n (кнопка пользователя)
#   POST /api/generate        — принять данные от n8n и сгенерировать Excel
#   GET  /api/status          — текущий статус обработки
#   GET  /api/download/<name> — скачать готовый файл

import os
import json
import threading
import time
import requests
from datetime import datetime
from flask import Flask, request, jsonify, send_file, render_template_string

from generator import process_sector
from sectors_config import SECTORS, N8N_BASE_URL, N8N_API_KEY

app = Flask(__name__)

# ── Глобальное состояние (для простоты; в продакшне — Redis/БД) ─────────────
status_store = {
    "running":    False,
    "started_at": None,
    "finished_at": None,
    "sectors":    {},   # sector_key -> {status, updated, error, output}
    "log":        [],
}


def log(msg):
    entry = f"[{datetime.now().strftime('%H:%M:%S')}] {msg}"
    status_store["log"].append(entry)
    print(entry)


# ── Шаг 1: запрос данных из n8n DataTable через API ──────────────────────────
def fetch_datatable(datatable_id: str) -> list:
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    url = N8N_BASE_URL + "/api/v1/data-tables/" + datatable_id + "/rows"
    headers = {"X-N8N-API-KEY": N8N_API_KEY}
    resp = requests.get(url, headers=headers, timeout=30, verify=False)
    resp.raise_for_status()
    return resp.json().get("data", [])


# ── Шаг 2: запуск воркфлоу через n8n API ─────────────────────────────────────
def trigger_n8n_workflow(workflow_id: str) -> bool:
    """Активирует воркфлоу через n8n API."""
    url = f"{N8N_BASE_URL}/api/v1/workflows/{workflow_id}/activate"
    headers = {"X-N8N-API-KEY": N8N_API_KEY, "Content-Type": "application/json"}
    try:
        # Запускаем через webhook Orchestrator (рекомендуемый способ)
        webhook_url = f"{N8N_BASE_URL}/webhook/macrospravka-orchestrator"
        resp = requests.post(webhook_url, json={"action": "run_all"}, timeout=10, verify=False)
        return resp.status_code in (200, 201, 202)
    except Exception as e:
        log(f"Ошибка запуска воркфлоу {workflow_id}: {e}")
        return False


# ── Основной процесс (в фоновом потоке) ──────────────────────────────────────
def run_all_sectors(json_data_by_sector: dict = None):
    """
    json_data_by_sector: {sector_key: [список строк из DataTable]}
    Если None — данные запрашиваются из n8n DataTable API напрямую.
    """
    status_store["running"]    = True
    status_store["started_at"] = datetime.now().isoformat()
    status_store["finished_at"] = None
    status_store["log"]        = []
    status_store["sectors"]    = {k: {"status": "pending"} for k in SECTORS}

    log("=== Запуск формирования макросправки ===")

    results = {}
    for sector_key, config in SECTORS.items():
        log(f"▶ Обрабатываю: {config['name']}")
        status_store["sectors"][sector_key] = {"status": "processing"}

        try:
            # Получаем данные: либо переданные, либо из n8n API
            if json_data_by_sector and sector_key in json_data_by_sector:
                rows = json_data_by_sector[sector_key]
            else:
                log(f"  Получаю данные из n8n DataTable ({config['datatable_id']})...")
                rows = fetch_datatable(config["datatable_id"])
                log(f"  Получено строк: {len(rows)}")

            result = process_sector(sector_key, config, rows)

            if result["success"]:
                log(f"  ✓ Готово. Обновлено ячеек: {result['updated']} → {result['output']}")
                status_store["sectors"][sector_key] = {
                    "status":  "done",
                    "updated": result["updated"],
                    "output":  os.path.basename(result["output"]),
                    "error":   None,
                }
            else:
                log(f"  ✗ Ошибка: {result['error']}")
                status_store["sectors"][sector_key] = {
                    "status": "error",
                    "error":  result["error"],
                }
        except Exception as e:
            log(f"  ✗ Исключение: {e}")
            status_store["sectors"][sector_key] = {"status": "error", "error": str(e)}

    done_count  = sum(1 for s in status_store["sectors"].values() if s["status"] == "done")
    error_count = sum(1 for s in status_store["sectors"].values() if s["status"] == "error")

    # ── Формируем итоговый файл со всеми 5 секторами ─────────────────────────
    if done_count > 0:
        log("▶ Формирую итоговый файл со всеми секторами...")
        try:
            from generator import build_combined_report
            combined_output = os.path.join("output", "Результат_Итоговый.xlsx")
            sector_results_map = {}
            for sk, sc in status_store["sectors"].items():
                if sc.get("status") == "done":
                    sector_results_map[sk] = {
                        "success": True,
                        "output": os.path.join("output", sc["output"])
                    }
            combined = build_combined_report(sector_results_map, combined_output)
            if combined["success"]:
                log(f"  ✓ Итоговый файл готов: {combined_output} ({combined['sheets']} листов)")
                status_store["combined"] = {
                    "status": "done",
                    "output": os.path.basename(combined_output),
                    "sheets": combined["sheets"]
                }
            else:
                log(f"  ✗ Ошибка итогового файла: {combined['error']}")
                status_store["combined"] = {"status": "error", "error": combined["error"]}
        except Exception as e:
            log(f"  ✗ Исключение при создании итогового файла: {e}")
            status_store["combined"] = {"status": "error", "error": str(e)}

    status_store["running"]     = False
    status_store["finished_at"] = datetime.now().isoformat()
    log(f"=== Завершено: {done_count} успешно, {error_count} с ошибками ===")


# ════════════════════════════════════════════════════════════════════════════
#  МАРШРУТЫ
# ════════════════════════════════════════════════════════════════════════════

@app.route("/")
def index():
    return render_template_string(HTML_TEMPLATE)


@app.route("/api/trigger", methods=["POST"])
def trigger():
    """Вызывается кнопкой в веб-интерфейсе.
    Запускает n8n Orchestrator, который спарсит данные,
    а потом вызовет /api/generate с готовыми данными.
    """
    if status_store["running"]:
        return jsonify({"ok": False, "message": "Процесс уже запущен"}), 409

    # Вариант А: n8n сам вызовет /api/generate после парсинга
    # Просто запускаем Orchestrator webhook
    try:
        webhook_url = f"{N8N_BASE_URL}/webhook/macrospravka-orchestrator"
        resp = requests.post(webhook_url, json={"action": "run_all"}, timeout=10, verify=False)
        if resp.status_code in (200, 201, 202):
            log("n8n Orchestrator запущен. Ожидаю данные...")
            return jsonify({"ok": True, "message": "Воркфлоу запущены. Ожидайте данные от n8n..."})
        else:
            # Запасной вариант: забираем данные сами
            log("n8n webhook недоступен. Запускаю прямой запрос к DataTable...")
            t = threading.Thread(target=run_all_sectors, daemon=True)
            t.start()
            return jsonify({"ok": True, "message": "Запущена прямая генерация из DataTable"})
    except Exception:
        log("n8n недоступен. Запускаю прямой запрос к DataTable...")
        t = threading.Thread(target=run_all_sectors, daemon=True)
        t.start()
        return jsonify({"ok": True, "message": "Запущена прямая генерация из DataTable"})


@app.route("/api/generate", methods=["POST"])
def generate():
    """
    Вызывается из n8n (HTTP Request нода) после завершения всех воркфлоу.
    Тело запроса (JSON):
    {
      "monetary": [...строки DataTable...],
      "real":     [...],
      "fiscal":   [...],
      "external": [...],
      "social":   [...]
    }
    Можно передавать частично — только те секторы, данные которых обновились.
    """
    if status_store["running"]:
        return jsonify({"ok": False, "message": "Уже обрабатывается"}), 409

    body = request.get_json(force=True) or {}
    # Принимаем и по полным именам секторов, и по ключам
    sector_data = {}
    for key in SECTORS:
        if key in body:
            sector_data[key] = body[key]

    if not sector_data:
        return jsonify({"ok": False, "message": "Нет данных секторов в теле запроса"}), 400

    t = threading.Thread(target=run_all_sectors, args=(sector_data,), daemon=True)
    t.start()
    return jsonify({"ok": True, "message": f"Генерация запущена для: {list(sector_data.keys())}"})


@app.route("/api/status")
def get_status():
    return jsonify(status_store)


@app.route("/api/download/<filename>")
def download(filename):
    filepath = os.path.join("output", filename)
    if not os.path.exists(filepath):
        return jsonify({"error": "Файл не найден"}), 404
    return send_file(
        os.path.abspath(filepath),
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ════════════════════════════════════════════════════════════════════════════
#  HTML — веб-интерфейс пользователя (встроен в app.py)
# ════════════════════════════════════════════════════════════════════════════

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Макросправка — Формирование Excel</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }

  body {
    font-family: 'Segoe UI', Arial, sans-serif;
    background: #f3f5f7;
    color: #1f2937;
  }

  .header {
    background: #1a3a5c;
    color: white;
    padding: 20px 32px;
    border-bottom: 1px solid rgba(255,255,255,0.08);
  }

  .header h1 {
    font-size: 22px;
    font-weight: 700;
    margin-bottom: 4px;
  }

  .header span {
    font-size: 13px;
    opacity: 0.82;
  }

  .container {
    max-width: 920px;
    margin: 32px auto;
    padding: 0 16px 32px;
  }

  .card {
    background: #fff;
    border-radius: 14px;
    padding: 24px;
    margin-bottom: 18px;
    box-shadow: 0 4px 18px rgba(15, 23, 42, 0.06);
    border: 1px solid #e7edf3;
  }

  .card h2 {
    font-size: 17px;
    color: #1a3a5c;
    margin-bottom: 18px;
    padding-bottom: 10px;
    border-bottom: 1px solid #e8edf2;
  }

  .status-bar {
    display: flex;
    align-items: center;
    gap: 10px;
    margin-bottom: 18px;
  }

  .status-dot {
    width: 10px;
    height: 10px;
    border-radius: 50%;
    flex: 0 0 auto;
  }

  .dot-idle { background: #9ca3af; }
  .dot-running { background: #f59e0b; animation: pulse 1s infinite; }
  .dot-done { background: #22c55e; }
  .dot-error { background: #ef4444; }

  @keyframes pulse {
    0%,100% { opacity: 1; }
    50% { opacity: 0.45; }
  }

  .info-text {
    font-size: 13px;
    color: #6b7280;
    line-height: 1.5;
  }

  .actions {
    display: flex;
    gap: 12px;
    flex-wrap: wrap;
    align-items: center;
  }

  .btn {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    gap: 8px;
    padding: 12px 22px;
    border: none;
    border-radius: 8px;
    font-size: 14px;
    font-weight: 600;
    cursor: pointer;
    transition: 0.2s ease;
    text-decoration: none;
  }

  .btn-primary {
    background: #1a3a5c;
    color: #fff;
  }

  .btn-primary:hover {
    background: #14304d;
    transform: translateY(-1px);
    box-shadow: 0 6px 16px rgba(26,58,92,0.22);
  }

  .btn-primary:disabled {
    background: #94a3b8;
    cursor: not-allowed;
    transform: none;
    box-shadow: none;
  }

  .btn-secondary {
    background: #eef2f6;
    color: #1a3a5c;
    border: 1px solid #d8e0e8;
  }

  .btn-secondary:hover {
    background: #e6ecf2;
    transform: translateY(-1px);
  }

  .btn-hidden {
    display: none;
  }

  .spinner {
    width: 17px;
    height: 17px;
    border: 3px solid rgba(255,255,255,0.3);
    border-top-color: white;
    border-radius: 50%;
    animation: spin 0.8s linear infinite;
  }

  @keyframes spin {
    to { transform: rotate(360deg); }
  }

  .sector-grid {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(260px, 1fr));
    gap: 14px;
  }

  .sector-card {
    border: 1px solid #e2e8f0;
    border-radius: 12px;
    padding: 16px;
    display: flex;
    justify-content: space-between;
    align-items: center;
    gap: 12px;
    background: #fcfdff;
  }

  .sector-name {
    font-size: 14px;
    font-weight: 600;
    color: #1f2937;
  }

  .badge {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    padding: 4px 10px;
    border-radius: 999px;
    font-size: 12px;
    font-weight: 700;
    white-space: nowrap;
  }

  .badge-pending {
    background: #f3f4f6;
    color: #6b7280;
  }

  .badge-processing {
    background: #fff7df;
    color: #9a6700;
  }

  .badge-done {
    background: #dcfce7;
    color: #166534;
  }

  .badge-error {
    background: #fee2e2;
    color: #991b1b;
  }

  .download-wrap {
    display: flex;
    flex-direction: column;
    align-items: flex-end;
    min-width: 105px;
  }

  .download-link {
    font-size: 12px;
    color: #1a3a5c;
    text-decoration: none;
    font-weight: 600;
    text-align: right;
  }

  .download-link:hover {
    text-decoration: underline;
  }

  .updated-count {
    font-size: 11px;
    color: #6b7280;
    margin-top: 4px;
    text-align: right;
    line-height: 1.3;
  }

  .log-shell {
    border: 1px solid #dbe3eb;
    border-radius: 12px;
    overflow: hidden;
    background: #ffffff;
  }

  .log-toolbar {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 10px 14px;
    background: #f8fafc;
    border-bottom: 1px solid #e5eaf0;
  }

  .log-title {
    font-size: 13px;
    font-weight: 600;
    color: #334155;
  }

  .log-status {
    font-size: 12px;
    color: #64748b;
  }

  .log-box {
    background: #0f172a;
    color: #d1fae5;
    font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
    font-size: 12px;
    line-height: 1.65;
    padding: 16px;
    min-height: 170px;
    max-height: 260px;
    overflow-y: auto;
    white-space: pre-wrap;
    word-break: break-word;
  }

  .log-empty {
    color: #94a3b8;
  }
</style>
</head>
<body>

<div class="header">
  <h1>📊 Макросправка — Автоформирование Excel</h1>
  <span>Автоматический сбор данных и генерация отчётов</span>
</div>

<div class="container">

  <div class="card">
    <h2>Управление</h2>

    <div class="status-bar">
      <div class="status-dot dot-idle" id="statusDot"></div>
      <span class="info-text" id="statusText">Готов к запуску</span>
    </div>

    <div class="actions">
      <button class="btn btn-primary" id="runBtn" onclick="triggerAll()">
        <span>▶ Сформировать макросправку</span>
      </button>

      <button class="btn btn-secondary btn-hidden" id="downloadAllBtn" onclick="downloadAllFiles()">
        <span>⬇ Скачать все файлы</span>
      </button>
    </div>

    <p class="info-text" style="margin-top: 12px;">
      Нажмите кнопку — данные автоматически соберутся из всех источников и сформируются файлы Excel для каждого сектора.
    </p>
  </div>

  <div class="card">
    <h2>Секторы</h2>
    <div class="sector-grid">
      <div class="sector-card">
        <span class="sector-name">Монетарный сектор</span>
        <span class="badge badge-pending" id="s-monetary">Ожидание</span>
      </div>
      <div class="sector-card">
        <span class="sector-name">Реальный сектор</span>
        <span class="badge badge-pending" id="s-real">Ожидание</span>
      </div>
      <div class="sector-card">
        <span class="sector-name">Фискальный сектор</span>
        <span class="badge badge-pending" id="s-fiscal">Ожидание</span>
      </div>
      <div class="sector-card">
        <span class="sector-name">Внешний сектор</span>
        <span class="badge badge-pending" id="s-external">Ожидание</span>
      </div>
      <div class="sector-card">
        <span class="sector-name">Социальный сектор</span>
        <span class="badge badge-pending" id="s-social">Ожидание</span>
      </div>
    </div>
  </div>

  <div class="card">
    <h2>Журнал выполнения</h2>
    <div class="log-shell">
      <div class="log-toolbar">
        <span class="log-title">Системные сообщения</span>
        <span class="log-status" id="logStatus">Ожидание запуска</span>
      </div>
      <div class="log-box" id="logBox"><span class="log-empty">Нажмите кнопку для запуска процесса...</span></div>
    </div>
  </div>

</div>

<script>
const SECTOR_NAMES = {
  monetary: "Монетарный сектор",
  real:     "Реальный сектор",
  fiscal:   "Фискальный сектор",
  external: "Внешний сектор",
  social:   "Социальный сектор",
};

const BADGE_CLASS = {
  pending:    "badge-pending",
  processing: "badge-processing",
  done:       "badge-done",
  error:      "badge-error",
};

const BADGE_TEXT = {
  pending:    "Ожидание",
  processing: "⏳ Обработка...",
  done:       "✓ Готово",
  error:      "✗ Ошибка",
};

let polling = null;
let readyFiles = [];

async function triggerAll() {
  const btn = document.getElementById("runBtn");
  const downloadAllBtn = document.getElementById("downloadAllBtn");
  const logBox = document.getElementById("logBox");
  const logStatus = document.getElementById("logStatus");
  const dot = document.getElementById("statusDot");
  const text = document.getElementById("statusText");

  btn.disabled = true;
  btn.innerHTML = '<div class="spinner"></div><span>Запуск...</span>';

  readyFiles = [];
  downloadAllBtn.classList.add("btn-hidden");

  Object.keys(SECTOR_NAMES).forEach(k => updateSectorBadge(k, "pending", null));

  dot.className = "status-dot dot-running";
  text.textContent = "Выполняется...";
  logStatus.textContent = "Процесс запущен";
  logBox.textContent = "Запуск воркфлоу...";

  try {
    const resp = await fetch("/api/trigger", { method: "POST" });
    if (!resp.ok) throw new Error("Ошибка запуска");
  } catch (e) {
    dot.className = "status-dot dot-error";
    text.textContent = "Ошибка связи с сервером";
    logStatus.textContent = "Ошибка";
    logBox.textContent = "Не удалось запустить процесс: " + e;
    btn.disabled = false;
    btn.innerHTML = "<span>▶ Сформировать макросправку</span>";
    return;
  }

  if (polling) clearInterval(polling);
  polling = setInterval(fetchStatus, 2000);
}

async function fetchStatus() {
  try {
    const resp = await fetch("/api/status");
    if (!resp.ok) return;
    const data = await resp.json();
    updateUI(data);
  } catch (e) {
    console.log("Status error:", e);
  }
}

function updateUI(data) {
  const dot = document.getElementById("statusDot");
  const text = document.getElementById("statusText");
  const btn = document.getElementById("runBtn");
  const downloadAllBtn = document.getElementById("downloadAllBtn");
  const logBox = document.getElementById("logBox");
  const logStatus = document.getElementById("logStatus");

  readyFiles = [];

  if (data.running) {
    dot.className = "status-dot dot-running";
    text.textContent = "Выполняется...";
    logStatus.textContent = "Идёт обработка";
    btn.disabled = true;
    btn.innerHTML = '<div class="spinner"></div><span>Выполняется...</span>';
  } else if (data.finished_at) {
    const hasErrors = Object.values(data.sectors || {}).some(s => s.status === "error");

    if (hasErrors) {
      dot.className = "status-dot dot-error";
      text.textContent = "Завершено с ошибками";
      logStatus.textContent = "Завершено с ошибками";
    } else {
      dot.className = "status-dot dot-done";
      text.textContent = "Успешно завершено";
      logStatus.textContent = "Завершено успешно";
    }

    btn.disabled = false;
    btn.innerHTML = "<span>▶ Сформировать макросправку</span>";

    if (polling) {
      clearInterval(polling);
      polling = null;
    }
  } else {
    dot.className = "status-dot dot-idle";
    text.textContent = "Готов к запуску";
    logStatus.textContent = "Ожидание запуска";
    btn.disabled = false;
    btn.innerHTML = "<span>▶ Сформировать макросправку</span>";
  }

  Object.entries(data.sectors || {}).forEach(([key, info]) => {
    updateSectorBadge(key, info.status, info);
    if (info && info.status === "done" && info.output) {
      readyFiles.push(info.output);
    }
  });

  if (readyFiles.length > 0) {
    downloadAllBtn.classList.remove("btn-hidden");
  } else {
    downloadAllBtn.classList.add("btn-hidden");
  }

  if (data.log && data.log.length) {
    logBox.textContent = data.log.join("\\n");
    logBox.scrollTop = logBox.scrollHeight;
  } else {
    logBox.innerHTML = '<span class="log-empty">Пока нет сообщений...</span>';
  }
}

function updateSectorBadge(key, status, info) {
  const el = document.getElementById("s-" + key);
  if (!el) return;

  el.className = "badge " + (BADGE_CLASS[status] || "badge-pending");
  el.textContent = BADGE_TEXT[status] || status;

  const card = el.closest(".sector-card");
  let wrap = card.querySelector(".download-wrap");

  if (info && info.status === "done" && info.output) {
    if (!wrap) {
      wrap = document.createElement("div");
      wrap.className = "download-wrap";
      card.appendChild(wrap);
    }

    wrap.innerHTML =
      '<a class="download-link" href="/api/download/' + encodeURIComponent(info.output) + '">⬇ Скачать Excel</a>' +
      (info.updated ? '<span class="updated-count">Обновлено ячеек: ' + info.updated + '</span>' : '');
  } else if (wrap) {
    wrap.remove();
  }
}

function downloadAllFiles() {
  if (!readyFiles.length) return;

  readyFiles.forEach((file, index) => {
    setTimeout(() => {
      const a = document.createElement("a");
      a.href = "/api/download/" + encodeURIComponent(file);
      a.download = "";
      document.body.appendChild(a);
      a.click();
      a.remove();
    }, index * 250);
  });
}

fetchStatus();
</script>
</body>
</html>
"""


if __name__ == "__main__":
    os.makedirs("output", exist_ok=True)
    print("=" * 60)
    print("  Макросправка — сервис формирования Excel")
    print("  Открыть интерфейс: http://localhost:8000")
    print("=" * 60)
    app.run(host="0.0.0.0", port=5000, debug=False)


