// Личный кабинет · обработка банковских выписок для 1С
//
// Что делает этот сервер:
//  1. Пускает в кабинет только по паролю (CABINET_PASSWORD из .env)
//  2. Хранит адрес/логин/пароль вашей 1С OData-службы
//  3. Принимает файл выписки (.xlsx), разбирает по строкам (разные банки РК)
//  4. Сверяет контрагента с 1С по БИН/ИИН, при неоднозначности — по названию,
//     подтягивает историю категоризации (статья ДДС, счёт, вид операции,
//     вид задолженности) и договор
//  5. По кнопке «подтвердить» создаёт в 1С ЧЕРНОВИК документа (не проведённый)
//     через OData — createDraftInOnec() внизу файла. Это рабочая реализация,
//     а не заглушка.
//  6. Ведёт историю действий
//
// ХРАНЕНИЕ ДАННЫХ (важно):
//  Если задана переменная окружения DATABASE_URL (строка подключения к Postgres) —
//  операции, история, правила и настройки хранятся в базе и переживают любой
//  передеплой/пересборку. Нужен пакет "pg": выполните `npm install pg`.
//  Без DATABASE_URL всё работает как раньше — через файлы в data/*.json — но
//  на бесплатном плане Render этот диск не постоянный и стирается при каждой
//  пересборке сайта.
//
// СЕССИИ (важно):
//  Вход в кабинет теперь не хранится в памяти процесса, а представляет собой
//  подписанный (HMAC) токен со сроком действия внутри самого токена. Поэтому
//  обновление кода / передеплой на Render больше НЕ разлогинивает — токен,
//  выданный до передеплоя, остаётся валиден, пока не истечёт (12 часов) или
//  пока вы явно не нажмёте «Выйти». Секрет для подписи — переменная окружения
//  SESSION_SECRET (если не задана — используется CABINET_PASSWORD).

require('dotenv').config();
const express = require('express');
const cookieParser = require('cookie-parser');
const multer = require('multer');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

// ---------- модули обработки банковской выписки (lib/) ----------
// Перенесены из этого файла как отдельный первый этап рефакторинга — сама
// бизнес-логика не менялась, только перемещена и (где было продублировано)
// объединена в одну функцию. См. lib/*.js — там же комментарии о том, что
// именно перенесено, а что добавлено новое (dryRun, verifyCreatedDocument).
const { fixKazakhMojibake, normalizeCounterpartyName, toIsoDate } = require('./lib/shared');
const { DEFAULT_RULES, normalizeRuleField, applyRules } = require('./lib/rules');
const { classifyOperation } = require('./lib/classifyOperation');
const bankParser = require('./lib/bankParser');
const { findCounterpartyByBin, searchCounterpartyByText, createCounterpartyInOnec } = require('./lib/resolveCounterparty');
const { resolveSupplierContract, createNoContractInOnec, NO_CONTRACT_NAME } = require('./lib/resolveContract');
const { findHistoricalCategory } = require('./lib/resolveCashFlowArticle');
const { makeFingerprint, checkExistingInOnec, findDuplicateAmongOperations } = require('./lib/duplicateGuard');
const { buildPayload } = require('./lib/buildPayload');
const { writeTo1C } = require('./lib/writeTo1C');
const { dryRun } = require('./lib/dryRun');
const { verifyCreatedDocument } = require('./lib/verifyCreatedDocument');
const { reconcileStatement } = require('./lib/reconcileStatement');

const app = express();
const PORT = process.env.PORT || 3000;
const DATA_DIR = path.join(__dirname, 'data');
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 10 * 1024 * 1024 } });

app.use(express.json());

app.use(cookieParser());

const HTML_PAGE = `<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Разноска · Выписки</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,500;9..144,600&family=Inter:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
  :root{
    --ink:#161A16; --paper:#FAFAF7; --paper-2:#F1F0EA; --line:#DEDCD2; --line-soft:#EAE8DF;
    --green:#2F5D45; --green-soft:#E3ECE4; --green-text:#1F4432;
    --brass:#9C7A3C; --brass-soft:#F3EBDA; --red:#A23B2E; --red-soft:#F5E4E0; --muted:#726F63; --radius:10px;
    font-family:'Inter',sans-serif;
  }
  *{box-sizing:border-box;}
  body{margin:0;background:var(--paper);color:var(--ink);}
  h1,h2,h3,.display{font-family:'Fraunces',serif;font-weight:500;letter-spacing:-0.01em;}
  .mono{font-family:'IBM Plex Mono',monospace;font-variant-numeric:tabular-nums;}

  /* Login */
  #login-screen{min-height:100vh;display:flex;align-items:center;justify-content:center;}
  .login-box{width:340px;border:1px solid var(--line);border-radius:var(--radius);padding:28px;background:#fff;}
  .login-box h1{font-size:20px;margin:0 0 4px;}
  .login-box p{font-size:13px;color:var(--muted);margin:0 0 18px;}
  .login-box input{width:100%;padding:10px 12px;border:1px solid var(--line);border-radius:8px;font-size:14px;margin-bottom:12px;}
  .login-box button{width:100%;padding:10px;border:none;border-radius:8px;background:var(--green);color:#fff;font-weight:500;font-size:14px;cursor:pointer;}
  .login-err{color:var(--red);font-size:12.5px;min-height:16px;margin-top:8px;}

  .app{display:flex;min-height:100vh;}
  .rail{width:236px;flex-shrink:0;border-right:1px solid var(--line);padding:24px 16px;display:flex;flex-direction:column;}
  .brand{display:flex;align-items:center;gap:10px;padding:4px 8px 22px;}
  .brand-mark{width:32px;height:32px;border-radius:8px;background:var(--green);color:#fff;display:flex;align-items:center;justify-content:center;font-family:'Fraunces',serif;font-size:15px;font-weight:600;flex-shrink:0;}
  .brand-name{font-size:14.5px;font-weight:600;line-height:1.25;}
  .brand-sub{font-size:12px;color:var(--muted);}
  .nav{display:flex;flex-direction:column;gap:2px;}
  .nav-item{display:flex;align-items:center;justify-content:space-between;padding:9px 10px;border-radius:8px;font-size:14px;cursor:pointer;color:var(--muted);}
  .nav-item:hover{background:var(--paper-2);color:var(--ink);}
  .nav-item.active{background:var(--green-soft);color:var(--green-text);font-weight:500;}
  .nav-item .count{font-size:11.5px;background:var(--brass);color:#fff;border-radius:20px;padding:1px 7px;font-family:'IBM Plex Mono',monospace;}
  .rail-foot{margin-top:auto;padding-top:16px;border-top:1px solid var(--line-soft);}
  .status-chip{display:flex;align-items:center;gap:7px;font-size:12.5px;color:var(--muted);padding:6px 8px;}
  .dot{width:6px;height:6px;border-radius:50%;background:var(--green);flex-shrink:0;}
  .dot.off{background:var(--red);}
  .logout-link{font-size:12.5px;color:var(--muted);cursor:pointer;padding:6px 8px;display:block;}

  .main{flex:1;min-width:0;}
  .topbar{display:flex;align-items:center;justify-content:space-between;padding:22px 32px;border-bottom:1px solid var(--line);}
  .eyebrow{font-size:11.5px;letter-spacing:0.06em;text-transform:uppercase;color:var(--muted);margin-bottom:4px;}
  .topbar h1{font-size:24px;margin:0;}
  .btn{border:none;border-radius:8px;padding:10px 16px;font-size:13.5px;font-weight:500;cursor:pointer;font-family:'Inter',sans-serif;}
  .btn-primary{background:var(--green);color:#fff;}
  .btn-ghost{background:transparent;border:1px solid var(--line);color:var(--ink);}
  .view{padding:28px 32px;} .hidden{display:none;}
  .stamp-banner{display:flex;align-items:center;gap:12px;border:1px solid var(--brass);background:var(--brass-soft);border-radius:var(--radius);padding:12px 16px;margin-bottom:24px;}
  .stamp{border:1.5px solid var(--brass);color:var(--brass);font-family:'Fraunces',serif;font-size:11px;letter-spacing:0.08em;text-transform:uppercase;padding:3px 9px;border-radius:20px;transform:rotate(-2deg);flex-shrink:0;}
  .stamp-banner p{margin:0;font-size:13px;color:#5c4826;}
  .table-card{border:1px solid var(--line);border-radius:var(--radius);overflow-x:auto;background:#fff;}
  .table-card table{min-width:920px;}
  .purpose-cell{max-width:260px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:var(--muted);cursor:help;}

  .op-list{display:flex;flex-direction:column;gap:8px;}
  .op-row{border:1px solid var(--line);border-radius:10px;background:#fff;overflow:hidden;}
  .op-row-header{display:flex;align-items:center;gap:14px;padding:13px 16px;cursor:pointer;flex-wrap:wrap;}
  .op-row-header:hover{background:var(--paper-2);}
  .op-main{flex:1;min-width:0;}
  .op-name{font-weight:500;font-size:13.5px;}
  .op-meta{font-size:11.5px;color:var(--muted);margin-top:2px;}
  .op-purpose-line{font-size:12px;color:var(--muted);margin-top:3px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
  .op-amount{font-family:'IBM Plex Mono',monospace;font-variant-numeric:tabular-nums;flex-shrink:0;text-align:right;min-width:100px;font-size:13.5px;}
  .op-chevron{flex-shrink:0;color:var(--muted);transition:transform .15s;font-size:12px;}
  .op-row.expanded .op-chevron{transform:rotate(180deg);}
  .op-details{display:none;padding:0 16px 16px;border-top:1px solid var(--line-soft);background:var(--paper-2);}
  .op-row.expanded .op-details{display:block;}
  .op-detail-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin:14px 0;}
  .op-detail-item label{font-size:11px;color:var(--muted);display:block;margin-bottom:2px;text-transform:uppercase;letter-spacing:0.03em;}
  .op-detail-item div{font-size:13px;}
  .op-detail-full{grid-column:1/-1;}
  table{width:100%;border-collapse:collapse;}
  th{text-align:left;font-size:11.5px;text-transform:uppercase;letter-spacing:0.04em;color:var(--muted);font-weight:500;padding:12px 16px;border-bottom:1px solid var(--line);background:var(--paper-2);}
  td{padding:13px 16px;border-bottom:1px solid var(--line-soft);font-size:13.5px;vertical-align:middle;}
  tr:last-child td{border-bottom:none;}
  .amt-in{color:var(--green-text);} .amt-out{color:var(--red);}
  .pill{display:inline-block;font-size:11.5px;padding:3px 9px;border-radius:20px;font-weight:500;}
  .pill-review{background:var(--brass-soft);color:#7A5F2B;}
  .pill-approved{background:var(--green-soft);color:var(--green-text);}
  .pill-new{background:var(--red-soft);color:var(--red);}
  .icon-btn{border:1px solid var(--line);background:#fff;border-radius:6px;padding:6px 10px;font-size:12.5px;cursor:pointer;}
  .icon-btn.confirm{background:var(--green);border-color:var(--green);color:#fff;}
  .empty{text-align:center;padding:60px 20px;color:var(--muted);}
  .settings-card{border:1px solid var(--line);border-radius:var(--radius);padding:20px;background:#fff;margin-bottom:14px;max-width:640px;}
  .settings-card label{font-size:12px;color:var(--muted);display:block;margin-bottom:4px;}
  .settings-card input{width:100%;padding:9px 11px;border:1px solid var(--line);border-radius:8px;font-size:13px;margin-bottom:12px;}
  .limits{border:1px solid var(--line);border-radius:var(--radius);padding:20px;background:var(--paper-2);max-width:640px;}
  .limits ul{list-style:none;margin:0;padding:0;display:flex;flex-direction:column;gap:9px;}
  .limits li{font-size:13px;display:flex;gap:9px;} .limits li:before{content:"✓";color:var(--green);font-weight:600;}
</style>
</head>
<body>

<div id="login-screen">
  <div class="login-box">
    <h1>Разноска</h1>
    <p>Личный кабинет · доступ только по паролю</p>
    <input type="password" id="login-password" placeholder="Пароль">
    <button onclick="doLogin()">Войти</button>
    <div class="login-err" id="login-err"></div>
  </div>
</div>

<div class="app hidden" id="app">
  <div class="rail">
    <div class="brand">
      <div class="brand-mark">Р</div>
      <div><div class="brand-name">Разноска</div><div class="brand-sub">Личный кабинет</div></div>
    </div>
    <div class="nav">
      <div class="nav-item active" data-view="review"><span>Проверка</span><span class="count" id="review-count">0</span></div>
      <div class="nav-item" data-view="history"><span>История</span></div>
      <div class="nav-item" data-view="rules"><span>Правила</span></div>
      <div class="nav-item" data-view="settings"><span>Настройки</span></div>
    </div>
    <div class="rail-foot">
      <div class="status-chip"><span class="dot" id="conn-dot"></span><span id="conn-text">Проверка подключения…</span></div>
      <div class="logout-link" onclick="doLogout()">Выйти</div>
    </div>
  </div>

  <div class="main">

    <div class="view" id="view-review">
      <div class="topbar">
        <div><div class="eyebrow">Банковская выписка</div><h1>Проверка операций</h1></div>
        <div>
          <button class="btn btn-ghost" onclick="clearOperations()">Очистить список</button>
          <button class="btn btn-primary" onclick="document.getElementById('upload-input').click()">+ Загрузить выписку</button>
          <input type="file" id="upload-input" accept=".xlsx,.xls" style="display:none" onchange="doUpload(this.files[0])">
        </div>
      </div>
      <div class="view-body" style="padding:28px 32px 0;">
        <div class="stamp-banner">
          <div class="stamp">Черновик</div>
          <p>Подтверждение создаёт документ в 1С без проведения. Проводите сами после проверки.</p>
        </div>
        <div id="upload-result-banner" class="hidden" style="border:1px solid var(--line);border-radius:var(--radius);padding:12px 16px;margin-bottom:14px;font-size:13px;"></div>
        <div id="review-summary" style="font-size:12.5px;color:var(--muted);margin-bottom:12px;"></div>
        <div id="review-list" class="op-list"><div class="empty">Загрузите файл выписки, чтобы начать</div></div>
      </div>
    </div>

    <div class="view hidden" id="view-history">
      <div class="topbar"><div><div class="eyebrow">Журнал действий</div><h1>История</h1></div></div>
      <div class="view-body" style="padding:28px 32px 0;">
        <div class="table-card"><table>
          <thead><tr><th>Время</th><th>Действие</th><th>Документ</th></tr></thead>
          <tbody id="history-tbody"><tr><td colspan="3" class="empty">Пока пусто</td></tr></tbody>
        </table></div>
      </div>
    </div>

    <div class="view hidden" id="view-rules">
      <div class="topbar"><div><div class="eyebrow">Категоризация</div><h1>Правила</h1></div></div>
      <div class="view-body" style="padding:28px 32px 0;max-width:720px;">
        <p style="font-size:12.5px;color:var(--muted);margin:-6px 0 18px;">Правило вида: если в назначении платежа (или в имени контрагента, или в КНП — коде назначения платежа из выписки) есть такой-то текст — предложить статью ДДС и/или счёт. Статья ДДС обязательна для подтверждения операции. Правила применяются сразу, без перезагрузки выписки.</p>
        <div class="settings-card" style="max-width:none;">
          <div style="display:flex;gap:10px;align-items:flex-end;flex-wrap:wrap;">
            <div style="flex:1;min-width:140px;">
              <label>Проверять поле</label>
              <select id="r-field" style="width:100%;padding:9px 11px;border:1px solid var(--line);border-radius:8px;font-size:13px;">
                <option value="purpose">Назначение платежа</option>
                <option value="counterparty">Имя контрагента</option>
                <option value="knp">КНП (код назначения платежа)</option>
              </select>
            </div>
            <div style="flex:2;min-width:160px;">
              <label>Содержит текст</label>
              <input id="r-contains" placeholder="например: аренда">
            </div>
            <div style="flex:2;min-width:160px;">
              <label>Статья ДДС</label>
              <input id="r-category" placeholder="например: Аренда офиса">
            </div>
            <div style="flex:1;min-width:110px;">
              <label>Счёт</label>
              <input id="r-account" placeholder="например: 3360">
            </div>
            <button class="btn btn-primary" style="margin-bottom:12px;" onclick="addRule()">Добавить</button>
          </div>
          <span id="rules-msg" style="font-size:12.5px;color:var(--muted);"></span>
        </div>

        <div class="settings-card" style="max-width:none;">
          <div class="settings-title" style="margin-bottom:6px;">Массовая загрузка правил текстом</div>
          <p style="font-size:12px;color:var(--muted);margin:0 0 10px;">Одна строка — одно правило, поля через вертикальную черту «|»: <b>текст-условие | статья ДДС | счёт | purpose/counterparty/knp | вид операции | вид задолженности</b> (последние три необязательны). Например:<br><span class="mono" style="font-size:12px;">аренда | Аренда офиса | 3360</span><br><span class="mono" style="font-size:12px;">Смаил | Выдача в подотчет | 1251 | counterparty | Перечисление денежных средств подотчетнику | Оплата поставщикам</span><br><span class="mono" style="font-size:12px;">014 | Оплата за товар | 3310/1710 | knp</span></p>
          <textarea id="r-bulk" rows="5" style="width:100%;padding:10px 12px;border:1px solid var(--line);border-radius:8px;font-size:13px;font-family:'IBM Plex Mono',monospace;" placeholder="аренда | Аренда офиса | 3360&#10;подписка | Подписки и лицензии | 3360&#10;реклама | Реклама и маркетинг | 7210"></textarea>
          <div style="margin-top:10px;">
            <button class="btn btn-primary" onclick="importRules()">Загрузить правила</button>
            <span id="rules-import-msg" style="font-size:12.5px;color:var(--muted);margin-left:10px;"></span>
          </div>
        </div>

        <div class="settings-card" style="max-width:none;">
          <div class="settings-title" style="margin-bottom:6px;">Сделать правила постоянными</div>
          <p style="font-size:12px;color:var(--muted);margin:0 0 10px;">Если DATABASE_URL не настроен, правила (как и подключение к 1С) сбрасываются при каждом обновлении сайта. Нажмите — покажется текст для копирования в Render → Environment → переменная <b>RULES_TEXT</b>. Если DATABASE_URL уже настроен — это не требуется, правила и так постоянны.</p>
          <button class="btn btn-ghost" onclick="exportRules()">Показать текст для Render</button>
          <textarea id="rules-export-box" rows="4" readonly style="display:none;width:100%;margin-top:10px;padding:10px 12px;border:1px solid var(--line);border-radius:8px;font-size:12px;font-family:'IBM Plex Mono',monospace;"></textarea>
        </div>

        <div class="table-card"><table>
          <thead><tr><th>Поле</th><th>Содержит</th><th>Статья ДДС</th><th>Счёт</th><th>Вид операции</th><th>Вид задолж.</th><th></th></tr></thead>
          <tbody id="rules-tbody"><tr><td colspan="7" class="empty">Правил пока нет</td></tr></tbody>
        </table></div>
      </div>
    </div>

    <div class="view hidden" id="view-settings">
      <div class="topbar"><div><div class="eyebrow">Подключение</div><h1>Настройки</h1></div></div>
      <div class="view-body" style="padding:28px 32px 0;">
        <div class="settings-card">
          <label>Адрес OData вашей базы 1С</label>
          <input id="s-url" placeholder="https://1cfresh.kz/a/xxxxx/xxxxxx/odata/standard.odata/">
          <label>Логин</label>
          <input id="s-login" placeholder="логин">
          <label>Пароль (оставьте пустым, если не меняете)</label>
          <input id="s-pass" type="password" placeholder="••••••••">
          <button class="btn btn-ghost" onclick="browseOnec()">Проверить подключение</button>
          <span id="browse-msg" style="font-size:12.5px;color:var(--muted);margin-left:10px;"></span>

          <div id="org-account-pickers" class="hidden" style="margin-top:16px;">
            <label>Организация</label>
            <select id="s-org" style="width:100%;padding:9px 11px;border:1px solid var(--line);border-radius:8px;font-size:13px;margin-bottom:12px;"></select>
            <label>Расчётный счёт</label>
            <select id="s-account" style="width:100%;padding:9px 11px;border:1px solid var(--line);border-radius:8px;font-size:13px;"></select>
          </div>

          <div style="margin-top:16px;">
            <button class="btn btn-primary" onclick="saveSettings()">Сохранить</button>
            <span id="settings-msg" style="font-size:12.5px;color:var(--muted);margin-left:10px;"></span>
          </div>
          <div id="current-org-account" style="font-size:12px;color:var(--muted);margin-top:10px;"></div>
        </div>
        <div id="env-hint" class="hidden" style="border:1px solid var(--brass);background:var(--brass-soft);border-radius:var(--radius);padding:16px;font-size:12.5px;color:#5c4826;margin-bottom:20px;line-height:1.7;"></div>
        <div class="limits">
          <h3 style="margin:0 0 12px;font-size:14.5px;">Ограничения текущего режима</h3>
          <ul>
            <li>Документы создаются только по вашему подтверждению</li>
            <li>Ни один документ не проводится автоматически</li>
            <li>Все действия записываются в историю</li>
          </ul>
        </div>
      </div>
    </div>

  </div>
</div>

<script>
async function api(path, opts={}){
  const res = await fetch(path, {credentials:'include', headers:{'Content-Type':'application/json'}, ...opts});
  if(res.status===401){ showLogin(); throw new Error('Требуется вход'); }
  if(!res.ok){ const e = await res.json().catch(()=>({error:'Ошибка запроса'})); throw new Error(e.error||'Ошибка запроса'); }
  return res.json();
}

function showLogin(){
  document.getElementById('login-screen').classList.remove('hidden');
  document.getElementById('app').classList.add('hidden');
}
function showApp(){
  document.getElementById('login-screen').classList.add('hidden');
  document.getElementById('app').classList.remove('hidden');
  loadAll();
}

async function doLogin(){
  const password = document.getElementById('login-password').value;
  const err = document.getElementById('login-err');
  err.textContent = '';
  try{
    await api('/api/login', {method:'POST', body: JSON.stringify({password})});
    showApp();
  }catch(e){ err.textContent = e.message; }
}
async function doLogout(){ await api('/api/logout', {method:'POST'}); showLogin(); }

async function loadAll(){
  await loadSettings();
  await loadOperations();
  await loadHistory();
  await loadRules();
}

let lastSettings = {};

async function loadSettings(){
  const s = await api('/api/settings');
  lastSettings = s;
  document.getElementById('s-url').value = s.baseUrl || '';
  document.getElementById('s-login').value = s.login || '';
  const dot = document.getElementById('conn-dot');
  const text = document.getElementById('conn-text');
  if(s.baseUrl && s.passwordSet){ dot.classList.remove('off'); text.textContent = 'Подключение настроено'; }
  else { dot.classList.add('off'); text.textContent = 'Подключение не настроено'; }
  const orgAccInfo = document.getElementById('current-org-account');
  if(s.orgName || s.accountName){
    orgAccInfo.textContent = 'Сейчас выбрано: ' + (s.orgName||'—') + ' · ' + (s.accountName||'—');
  } else {
    orgAccInfo.textContent = 'Организация и счёт ещё не выбраны — нажмите «Проверить подключение».';
  }
}

async function browseOnec(){
  const baseUrl = document.getElementById('s-url').value.trim();
  const login = document.getElementById('s-login').value.trim();
  const password = document.getElementById('s-pass').value;
  const msg = document.getElementById('browse-msg');
  msg.textContent = 'Проверяю…';
  try{
    const result = await api('/api/settings/browse', {method:'POST', body: JSON.stringify({baseUrl, login, password})});
    msg.textContent = 'Подключение работает, найдено: организаций — ' + result.organizations.length + ', счетов — ' + result.accounts.length;
    const orgSel = document.getElementById('s-org');
    const accSel = document.getElementById('s-account');
    orgSel.innerHTML = result.organizations.map(o => \`<option value="\${o.key}">\${o.name}</option>\`).join('');
    accSel.innerHTML = result.accounts.map(a => \`<option value="\${a.key}">\${a.name}</option>\`).join('');
    document.getElementById('org-account-pickers').classList.remove('hidden');
  }catch(e){ msg.textContent = 'Ошибка: ' + e.message; }
}

async function saveSettings(){
  const baseUrl = document.getElementById('s-url').value.trim();
  const login = document.getElementById('s-login').value.trim();
  const password = document.getElementById('s-pass').value;
  const msg = document.getElementById('settings-msg');
  const body = {baseUrl, login, password};
  const orgSel = document.getElementById('s-org');
  const accSel = document.getElementById('s-account');
  if(orgSel.options.length){
    body.orgKey = orgSel.value;
    body.orgName = orgSel.options[orgSel.selectedIndex].text;
  }
  if(accSel.options.length){
    body.accountKey = accSel.value;
    body.accountName = accSel.options[accSel.selectedIndex].text;
  }
  try{
    await api('/api/settings', {method:'POST', body: JSON.stringify(body)});
    msg.textContent = 'Сохранено';
    document.getElementById('s-pass').value='';
    loadSettings();
    showEnvVarHint(body);
  }catch(e){ msg.textContent = e.message; }
}

function showEnvVarHint(body){
  const box = document.getElementById('env-hint');
  if(!box) return;
  const lines = [
    'ONEC_BASE_URL = ' + body.baseUrl,
    'ONEC_LOGIN = ' + body.login,
    body.password ? 'ONEC_PASSWORD = (тот пароль, что вы ввели)' : null,
    body.orgKey ? 'ONEC_ORG_KEY = ' + body.orgKey : null,
    body.orgName ? 'ONEC_ORG_NAME = ' + body.orgName : null,
    body.accountKey ? 'ONEC_ACCOUNT_KEY = ' + body.accountKey : null,
    body.accountName ? 'ONEC_ACCOUNT_NAME = ' + body.accountName : null,
  ].filter(Boolean);
  box.innerHTML = '<b>Если DATABASE_URL не настроен — эти переменные нужно один раз добавить в Render (Environment → Add Environment Variable), иначе подключение слетит при следующем обновлении сайта:</b><br><br>' + lines.join('<br>');
  box.classList.remove('hidden');
}

function money(v){ const sign = v>0?'+':''; return sign + v.toLocaleString('ru-RU') + ' \\u20B8'; }

const expandedOps = new Set();

function escapeHtml(s){
  return String(s || '').replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]));
}

async function loadOperations(){
  const ops = await api('/api/operations');
  const list = document.getElementById('review-list');
  const pending = ops.filter(o=>o.status==='review' || o.status==='new_counterparty');
  document.getElementById('review-count').textContent = pending.length;

  const summary = document.getElementById('review-summary');
  if(summary){
    const already = ops.filter(o=>o.status==='already_in_1c').length;
    const done = ops.filter(o=>o.status==='draft_created').length;
    summary.textContent = ops.length
      ? \`Всего в выписке: \${ops.length} · Уже в 1С: \${already} · Черновиков создано: \${done} · К проверке: \${pending.length}\`
      : '';
  }

  if(ops.length===0){ list.innerHTML = '<div class="empty">Загрузите файл выписки, чтобы начать</div>'; return; }

  // Требующие вашего решения — всегда сверху; уже обработанные — вниз,
  // чтобы не приходилось искать проблемные операции среди сотен готовых.
  const priority = (o) => {
    if(o.status==='new_counterparty') return 0;
    if(o.status==='review' && !o.suggestedCategory) return 0;
    if(o.contractStatus==='ambiguous') return 0;
    if(o.status==='review') return 1;
    if(o.status==='already_in_1c') return 2;
    if(o.status==='duplicate') return 2;
    if(o.status==='draft_created') return 3;
    return 1;
  };
  const sorted = [...ops].sort((a,b)=>priority(a)-priority(b));

  list.innerHTML = '';
  sorted.forEach(o=>{
    const done = o.status==='draft_created';
    const isNew = o.status==='new_counterparty';
    const already = o.status==='already_in_1c';
    const isDuplicate = o.status==='duplicate';
    const hasCategory = !!o.suggestedCategory;
    const contractAmbiguous = o.contractStatus === 'ambiguous';
    const needsNoContract = o.contractStatus === 'need_create_no_contract' && !o.contractKey;
    // Бывает, что операция загружена ДО того, как подключение к 1С было
    // настроено — тогда сопоставление контрагента не выполнялось вообще,
    // но статус остался обычным "review". Без реального Ref_Key
    // контрагента документ в 1С создать нельзя — блокируем и предлагаем
    // пересверить.
    const missingCounterparty = !done && !already && !isDuplicate && !isNew && !o.counterpartyKey && o.status !== 'ambiguous_counterparty';
    const ambiguousCounterparty = o.status === 'ambiguous_counterparty';
    let statusHtml, actionHtml;
    if(done){
      statusHtml = '<span class="pill pill-approved">Черновик создан</span>';
      actionHtml = '';
    } else if(already){
      statusHtml = '<span class="pill pill-review" style="background:var(--paper-2);color:var(--muted);">Уже в 1С</span>';
      actionHtml = '';
    } else if(isDuplicate){
      statusHtml = '<span class="pill pill-review" style="background:var(--paper-2);color:var(--muted);">Дубликат</span>';
      actionHtml = '<span style="font-size:11px;color:var(--muted);">похоже, уже разнесено — см. историю</span>';
    } else if(isNew){
      statusHtml = '<span class="pill pill-new">Новый контрагент</span>';
      actionHtml = '<span style="font-size:11px;color:var(--muted);">проверьте ниже ↓</span>';
    } else if(ambiguousCounterparty){
      statusHtml = '<span class="pill pill-new">Несколько контрагентов — выбрать</span>';
      actionHtml = '<span style="font-size:11px;color:var(--muted);">выберите ниже ↓</span>';
    } else if(missingCounterparty){
      statusHtml = '<span class="pill pill-new">Контрагент не сверен</span>';
      actionHtml = '<button class="icon-btn" onclick="event.stopPropagation();rematchOp(\\''+o.id+'\\')">Пересверить с 1С</button>';
    } else if(!hasCategory){
      statusHtml = '<span class="pill pill-new">Нет статьи ДДС</span>';
      actionHtml = '<button class="icon-btn" disabled title="Сначала определите статью ДДС в разделе «Правила»" style="opacity:0.5;cursor:not-allowed;">Подтвердить</button>';
    } else if(contractAmbiguous){
      statusHtml = '<span class="pill pill-new">Неясен договор</span>';
      actionHtml = '<button class="icon-btn" disabled title="У контрагента несколько договоров — нужен ручной выбор" style="opacity:0.5;cursor:not-allowed;">Подтвердить</button>';
    } else {
      statusHtml = '<span class="pill pill-review">На проверке</span>';
      actionHtml = '<button class="icon-btn confirm" onclick="event.stopPropagation();confirmOp(\\''+o.id+'\\')">Подтвердить</button>';
    }
    const contractText = contractAmbiguous
      ? '<span style="color:var(--red)">несколько — выбрать вручную</span>'
      : needsNoContract
        ? '<span style="color:var(--brass);">не найден — <a href="#" onclick="event.preventDefault();event.stopPropagation();createNoContract(\\''+o.id+'\\')">создать «Без договора»</a></span>'
        : (o.contractName || '—');
    const isExpanded = expandedOps.has(o.id);
    const defaultHint = o.isDefaultOnly ? ' <span style="color:var(--brass);font-size:11px;">(по умолчанию — проверьте)</span>' : '';

    const row = document.createElement('div');
    row.className = 'op-row' + (isExpanded ? ' expanded' : '');
    row.innerHTML = \`
      <div class="op-row-header" onclick="toggleOpRow('\${o.id}')">
        <div class="op-main">
          <div class="op-name">\${o.counterparty || '—'}</div>
          <div class="op-meta">\${o.date} · КНП \${o.knp || '—'} · БИН \${o.bin || '—'}</div>
          <div class="op-purpose-line" title="\${o.purpose.replace(/"/g,'&quot;')}">\${o.purpose}</div>
        </div>
        <div class="op-amount \${o.amount>0?'amt-in':'amt-out'}">\${money(o.amount)}</div>
        \${statusHtml}
        <div style="min-width:0;">\${actionHtml}</div>
        <div class="op-chevron">▾</div>
      </div>
      <div class="op-details">
        \${ambiguousCounterparty ? \`
        <div style="background:var(--red-soft);border:1px solid var(--red);border-radius:8px;padding:12px 14px;margin-bottom:14px;">
          <div style="font-size:12.5px;color:var(--red);margin-bottom:8px;">В справочнике нашлось несколько контрагентов с похожим именем — выберите нужного:</div>
          <select id="cp-choice-\${o.id}" style="width:100%;padding:7px 9px;border:1px solid var(--line);border-radius:6px;font-size:13px;margin-bottom:8px;">
            \${(o.counterpartyOptions||[]).map(c => \`<option value="\${c.key}">\${c.name}</option>\`).join('')}
          </select>
          <button class="btn btn-primary" style="padding:7px 12px;font-size:12.5px;" onclick="event.stopPropagation();chooseCounterparty('\${o.id}')">Выбрать этого контрагента</button>
        </div>
        \` : ''}
        \${isNew ? \`
        <div style="background:var(--red-soft);border:1px solid var(--red);border-radius:8px;padding:12px 14px;margin-bottom:14px;">
          <div style="font-size:12.5px;color:var(--red);margin-bottom:10px;">Контрагент не найден автоматически. Прежде чем создавать нового — проверьте, нет ли он уже в 1С под другим написанием:</div>
          <div style="display:flex;gap:8px;margin-bottom:8px;">
            <input id="search-cp-\${o.id}" placeholder="Часть названия для поиска в 1С" style="flex:1;padding:7px 9px;border:1px solid var(--line);border-radius:6px;font-size:13px;" onkeydown="if(event.key==='Enter'){event.preventDefault();searchCounterparty('\${o.id}');}">
            <button class="btn btn-ghost" style="padding:7px 12px;font-size:12.5px;" onclick="event.stopPropagation();searchCounterparty('\${o.id}')">Искать в 1С</button>
          </div>
          <div id="search-cp-results-\${o.id}" style="margin-bottom:10px;"></div>
          <div style="display:flex;align-items:center;gap:10px;border-top:1px solid var(--line);padding-top:10px;">
            <span style="font-size:12px;color:var(--muted);">Проверили — такого контрагента правда нет:</span>
            <button class="icon-btn confirm" onclick="event.stopPropagation();createCounterparty('\${o.id}')">Создать нового</button>
          </div>
        </div>
        \` : ''}
        <div class="op-detail-grid">
          <div class="op-detail-item">
            <label>Статья ДДС\${defaultHint}</label>
            <input class="edit-input" id="cat-\${o.id}" value="\${(o.suggestedCategory||'').replace(/"/g,'&quot;')}" style="width:100%;padding:6px 8px;border:1px solid \${hasCategory?'var(--line)':'var(--red)'};border-radius:6px;font-size:13px;">
          </div>
          <div class="op-detail-item">
            <label>Счёт</label>
            <input class="edit-input mono" id="acc-\${o.id}" value="\${(o.suggestedAccount||'').replace(/"/g,'&quot;')}" style="width:100%;padding:6px 8px;border:1px solid var(--line);border-radius:6px;font-size:13px;">
          </div>
          <div class="op-detail-item">
            <label>Вид операции (только из истории 1С)</label>
            <input class="edit-input" id="opkind-\${o.id}" value="\${(o.suggestedOperationKind||'').replace(/"/g,'&quot;')}" placeholder="например: Перечисление денежных средств подотчетнику" style="width:100%;padding:6px 8px;border:1px solid var(--line);border-radius:6px;font-size:13px;">
          </div>
          <div class="op-detail-item">
            <label>Вид задолженности (только из истории 1С)</label>
            <input class="edit-input" id="debttype-\${o.id}" value="\${(o.suggestedDebtType||'').replace(/"/g,'&quot;')}" placeholder="например: Оплата поставщикам" style="width:100%;padding:6px 8px;border:1px solid var(--line);border-radius:6px;font-size:13px;">
          </div>
          <div class="op-detail-item"><label>Договор</label><div>\${contractText}</div></div>
          <div class="op-detail-item"><label>КНП</label><div class="mono">\${o.knp || '—'}</div></div>
          <div class="op-detail-item op-detail-full"><label>Полное назначение платежа</label><div>\${o.purpose}</div></div>
        </div>
        <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-top:4px;">
          <button class="btn btn-ghost" style="padding:7px 12px;font-size:12.5px;" onclick="event.stopPropagation();saveOverride('\${o.id}', false)">Сохранить правку</button>
          <label style="display:flex;align-items:center;gap:6px;font-size:12.5px;color:var(--muted);cursor:pointer;">
            <input type="checkbox" id="rule-\${o.id}"> запомнить как правило: если
          </label>
          <select id="rule-field-\${o.id}" style="padding:4px 6px;border:1px solid var(--line);border-radius:6px;font-size:12px;">
            <option value="counterparty" selected>контрагент</option>
            <option value="purpose">назначение</option>
            <option value="knp">КНП</option>
          </select>
          <span style="font-size:12.5px;color:var(--muted);">содержит</span>
          <input id="rule-contains-\${o.id}" value="\${(o.counterparty||'').replace(/"/g,'&quot;')}" style="padding:4px 8px;border:1px solid var(--line);border-radius:6px;font-size:12px;width:180px;">
          <span id="override-msg-\${o.id}" style="font-size:12px;color:var(--muted);"></span>
        </div>
      </div>
    \`;
    list.appendChild(row);
  });
}

async function saveOverride(id, silent){
  const category = document.getElementById('cat-'+id).value.trim();
  const account = document.getElementById('acc-'+id).value.trim();
  const operationKind = document.getElementById('opkind-'+id).value.trim();
  const debtType = document.getElementById('debttype-'+id).value.trim();
  const rememberBox = document.getElementById('rule-'+id);
  const remember = rememberBox && rememberBox.checked;
  const msg = document.getElementById('override-msg-'+id);
  const body = { category, account, operationKind, debtType };
  if(remember){
    const fieldSel = document.getElementById('rule-field-'+id);
    const containsInput = document.getElementById('rule-contains-'+id);
    body.saveAsRule = true;
    body.ruleField = fieldSel ? fieldSel.value : 'purpose';
    body.ruleContains = containsInput ? containsInput.value.trim() : category;
  }
  try{
    const result = await api('/api/operations/'+id+'/override', {method:'POST', body: JSON.stringify(body)});
    if(msg) msg.textContent = result.ruleAdded ? 'Сохранено и запомнено как правило' : 'Сохранено';
    loadOperations();
  }catch(e){ if(msg) msg.textContent = e.message; }
}

function toggleOpRow(id){
  if(expandedOps.has(id)) expandedOps.delete(id); else expandedOps.add(id);
  const rows = document.querySelectorAll('.op-row');
  loadOperations.lastToggled = id;
  // просто перерисовываем список — состояние "раскрыто" хранится в expandedOps
  loadOperations();
}

async function confirmOp(id){
  try{
    await api('/api/operations/'+id+'/confirm', {method:'POST'});
    loadOperations(); loadHistory();
  }catch(e){ alert(e.message); }
}

async function createCounterparty(id){
  try{
    await api('/api/operations/'+id+'/create-counterparty', {method:'POST'});
    loadOperations(); loadHistory();
  }catch(e){ alert(e.message); }
}

async function createNoContract(id){
  try{
    await api('/api/operations/'+id+'/create-no-contract', {method:'POST'});
    loadOperations(); loadHistory();
  }catch(e){ alert(e.message); }
}

async function rematchOp(id){
  try{
    await api('/api/operations/'+id+'/rematch', {method:'POST'});
    loadOperations(); loadHistory();
  }catch(e){ alert(e.message); }
}

async function chooseCounterparty(id){
  const select = document.getElementById('cp-choice-'+id);
  if(!select) return;
  const key = select.value;
  const name = select.options[select.selectedIndex].text;
  try{
    await api('/api/operations/'+id+'/choose-counterparty', {method:'POST', body: JSON.stringify({counterpartyKey: key, counterpartyName: name})});
    loadOperations(); loadHistory();
  }catch(e){ alert(e.message); }
}

async function searchCounterparty(id){
  const input = document.getElementById('search-cp-'+id);
  const resultsBox = document.getElementById('search-cp-results-'+id);
  if(!input || !resultsBox) return;
  const query = input.value.trim();
  if(!query){ resultsBox.innerHTML = '<span style="font-size:12px;color:var(--muted);">Введите часть названия</span>'; return; }
  resultsBox.innerHTML = '<span style="font-size:12px;color:var(--muted);">Ищу в 1С…</span>';
  try{
    const result = await api('/api/operations/'+id+'/search-counterparty', {method:'POST', body: JSON.stringify({query})});
    if(!result.options.length){
      resultsBox.innerHTML = '<span style="font-size:12px;color:var(--muted);">Ничего не найдено — похоже, контрагента в 1С действительно ещё нет.</span>';
      return;
    }
    resultsBox.innerHTML = result.options.map(o => \`
      <div style="display:flex;align-items:center;justify-content:space-between;gap:8px;padding:6px 0;border-bottom:1px solid var(--line-soft);font-size:12.5px;">
        <span>\${escapeHtml(o.name)}</span>
        <button class="icon-btn confirm" onclick="event.stopPropagation();pickSearchedCounterparty('\${id}','\${o.key}','\${o.name.replace(/'/g,"\\\\'").replace(/"/g,'&quot;')}')">Это он</button>
      </div>
    \`).join('');
  }catch(e){ resultsBox.innerHTML = '<span style="font-size:12px;color:var(--red);">'+e.message+'</span>'; }
}

async function pickSearchedCounterparty(id, key, name){
  try{
    await api('/api/operations/'+id+'/choose-counterparty', {method:'POST', body: JSON.stringify({counterpartyKey: key, counterpartyName: name})});
    loadOperations(); loadHistory();
  }catch(e){ alert(e.message); }
}

async function loadRules(){
  const rules = await api('/api/rules');
  const tbody = document.getElementById('rules-tbody');
  tbody.innerHTML = '';
  if(rules.length===0){ tbody.innerHTML = '<tr><td colspan="7" class="empty">Правил пока нет</td></tr>'; return; }
  rules.forEach(r=>{
    const tr = document.createElement('tr');
    const fieldLabel = r.field==='counterparty' ? 'Имя контрагента' : (r.field==='knp' ? 'КНП' : 'Назначение платежа');
    tr.innerHTML = \`
      <td>\${fieldLabel}</td>
      <td>\${r.contains}</td>
      <td>\${r.category || '—'}</td>
      <td class="mono">\${r.account || '—'}</td>
      <td style="font-size:12px;">\${r.operationKind || '—'}</td>
      <td style="font-size:12px;">\${r.debtType || '—'}</td>
      <td><button class="icon-btn" onclick="deleteRule('\${r.id}')">Удалить</button></td>
    \`;
    tbody.appendChild(tr);
  });
}

async function addRule(){
  const field = document.getElementById('r-field').value;
  const contains = document.getElementById('r-contains').value.trim();
  const category = document.getElementById('r-category').value.trim();
  const account = document.getElementById('r-account').value.trim();
  const msg = document.getElementById('rules-msg');
  try{
    await api('/api/rules', {method:'POST', body: JSON.stringify({field, contains, category, account})});
    document.getElementById('r-contains').value = '';
    document.getElementById('r-category').value = '';
    document.getElementById('r-account').value = '';
    msg.textContent = 'Правило добавлено';
    loadRules(); loadOperations();
  }catch(e){ msg.textContent = e.message; }
}

async function importRules(){
  const text = document.getElementById('r-bulk').value;
  const msg = document.getElementById('rules-import-msg');
  try{
    const result = await api('/api/rules/import', {method:'POST', body: JSON.stringify({text})});
    msg.textContent = \`Добавлено: \${result.added}\` + (result.skipped ? \`, пропущено: \${result.skipped}\` : '');
    document.getElementById('r-bulk').value = '';
    loadRules(); loadOperations();
  }catch(e){ msg.textContent = e.message; }
}

async function exportRules(){
  const result = await api('/api/rules/export');
  const box = document.getElementById('rules-export-box');
  box.value = result.text;
  box.style.display = 'block';
  box.focus();
  box.select();
}

async function deleteRule(id){
  await api('/api/rules/'+id, {method:'DELETE'});
  loadRules(); loadOperations();
}

async function doUpload(file){
  if(!file) return;
  const fd = new FormData();
  fd.append('file', file);
  const banner = document.getElementById('upload-result-banner');
  if(banner){ banner.classList.remove('hidden'); banner.textContent = 'Загружаю и сверяю с 1С — это может занять минуту...'; banner.style.background='var(--paper-2)'; banner.style.color='var(--muted)'; banner.style.borderColor='var(--line)'; }
  const res = await fetch('/api/upload', {method:'POST', credentials:'include', body: fd});
  if(!res.ok){ const e = await res.json().catch(()=>({})); if(banner) banner.classList.add('hidden'); alert(e.error||'Ошибка загрузки'); return; }
  const result = await res.json().catch(()=>({}));
  if(banner && result.reconciliation){
    const r = result.reconciliation;
    let text = \`Загружено \${r.total} новых операций\`;
    if(result.skippedDuplicates) text += \` (пропущено дублей: \${result.skippedDuplicates})\`;
    if(r.connected){
      text += \`. Сверка с 1С: уже есть в базе — \${r.alreadyIn1c}, новых контрагентов — \${r.newCounterparty}, к проверке — \${r.total - r.alreadyIn1c}.\`;
      banner.style.background='var(--green-soft)'; banner.style.color='var(--green-text)'; banner.style.borderColor='var(--green)';
    } else {
      text += '. Сверка с 1С НЕ выполнена — подключение не настроено (раздел «Настройки»).';
      banner.style.background='var(--red-soft)'; banner.style.color='var(--red)'; banner.style.borderColor='var(--red)';
    }
    banner.textContent = text;
  }
  loadOperations(); loadHistory();
}

async function clearOperations(){
  if(!confirm('Очистить весь список операций в кабинете? Это не затронет данные в 1С — только список здесь, в кабинете.')) return;
  await api('/api/operations/clear', {method:'POST'});
  loadOperations(); loadHistory();
}

async function loadHistory(){
  const h = await api('/api/history');
  const tbody = document.getElementById('history-tbody');
  tbody.innerHTML = '';
  if(h.length===0){ tbody.innerHTML = '<tr><td colspan="3" class="empty">Пока пусто</td></tr>'; return; }
  h.forEach(item=>{
    const tr = document.createElement('tr');
    const t = new Date(item.time).toLocaleString('ru-RU');
    tr.innerHTML = \`<td class="mono" style="color:var(--muted)">\${t}</td><td>\${item.action}</td><td style="color:var(--muted)">\${item.doc}</td>\`;
    tbody.appendChild(tr);
  });
}

document.querySelectorAll('.nav-item').forEach(item=>{
  item.addEventListener('click', ()=>{
    document.querySelectorAll('.nav-item').forEach(i=>i.classList.remove('active'));
    item.classList.add('active');
    document.querySelectorAll('.view').forEach(v=>v.classList.add('hidden'));
    document.getElementById('view-'+item.dataset.view).classList.remove('hidden');
  });
});

// Пытаемся сразу показать кабинет, если сессия уже есть
api('/api/settings').then(showApp).catch(()=>showLogin());
</script>
</body>
</html>
`;

app.get('/', (req, res) => {
  res.type('html').send(HTML_PAGE);
});
// (фронтенд теперь встроен прямо в этот файл — см. HTML_PAGE ниже)

// ---------- хранилище: Postgres (если задан DATABASE_URL), иначе локальные JSON-файлы ----------
// На бесплатном плане Render локальный диск стирается при каждом передеплое. Если вы
// хотите, чтобы операции/история/правила/настройки переживали обновление кода — задайте
// переменную окружения DATABASE_URL (строка подключения к Postgres) в Render →
// Environment и установите пакет pg: `npm install pg`. Без DATABASE_URL сервер работает
// как раньше — через файлы в data/ — и данные по-прежнему будут стираться при пересборке.
let Pool = null;
try {
  ({ Pool } = require('pg'));
} catch (e) {
  Pool = null;
}

const pgPool = (process.env.DATABASE_URL && Pool)
  ? new Pool({
      connectionString: process.env.DATABASE_URL,
      // Большинство бесплатных облачных Postgres (Neon, Supabase, Render Postgres)
      // требуют SSL, но с самоподписанным/неполным доверенным путём сертификатов —
      // поэтому проверку цепочки отключаем, а не сам SSL.
      ssl: process.env.PGSSL === 'false' ? false : { rejectUnauthorized: false },
    })
  : null;

if (process.env.DATABASE_URL && !Pool) {
  console.error(
    'DATABASE_URL задан, но пакет "pg" не установлен. Выполните "npm install pg" и передеплойте — ' +
      'до тех пор данные хранятся в локальных файлах и будут стираться при каждом передеплое.'
  );
}

async function ensureStorage() {
  if (!pgPool) return;
  await pgPool.query(`
    CREATE TABLE IF NOT EXISTS app_storage (
      key TEXT PRIMARY KEY,
      value JSONB NOT NULL,
      updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
    )
  `);
}

function readJsonFile(file, fallback) {
  const p = path.join(DATA_DIR, file);
  if (!fs.existsSync(p)) return fallback;
  return JSON.parse(fs.readFileSync(p, 'utf8'));
}
function writeJsonFile(file, data) {
  fs.writeFileSync(path.join(DATA_DIR, file), JSON.stringify(data, null, 2), 'utf8');
}
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR);

// Единая точка чтения/записи для всего "постоянного" состояния приложения
// (operations, history, rules, settings). За кулисами — либо Postgres, либо файл.
async function getStore(key, fallback) {
  if (pgPool) {
    const { rows } = await pgPool.query('SELECT value FROM app_storage WHERE key = $1', [key]);
    return rows.length ? rows[0].value : fallback;
  }
  return readJsonFile(key + '.json', fallback);
}
async function setStore(key, value) {
  if (pgPool) {
    await pgPool.query(
      `INSERT INTO app_storage (key, value, updated_at) VALUES ($1, $2::jsonb, now())
       ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value, updated_at = now()`,
      [key, JSON.stringify(value)]
    );
    return;
  }
  writeJsonFile(key + '.json', value);
}

// ---------- вход по паролю (сессия — подписанный токен, без хранения на сервере) ----------
// Раньше токен сессии хранился в Map в памяти процесса — при каждом передеплое на
// Render процесс пересоздаётся, Map становится пустой, и все уже вошедшие
// разлогиниваются. Теперь сессия — это подписанный (HMAC-SHA256) токен со сроком
// действия внутри самого токена: сервер ничего не хранит и поэтому переживает
// перезапуск/передеплой. Подделать токен без секрета невозможно.
const SESSION_SECRET = process.env.SESSION_SECRET || process.env.CABINET_PASSWORD || '';
if (!SESSION_SECRET) {
  console.error('Не задан ни SESSION_SECRET, ни CABINET_PASSWORD — подписывать сессии нечем, вход не будет работать.');
}
const SESSION_TTL_MS = 12 * 60 * 60 * 1000; // 12 часов

function signSession(payload) {
  const body = Buffer.from(JSON.stringify(payload)).toString('base64url');
  const sig = crypto.createHmac('sha256', SESSION_SECRET).update(body).digest('base64url');
  return `${body}.${sig}`;
}
function verifySession(token) {
  if (!token || typeof token !== 'string' || !token.includes('.')) return null;
  const [body, sig] = token.split('.');
  if (!body || !sig) return null;
  const expected = crypto.createHmac('sha256', SESSION_SECRET).update(body).digest('base64url');
  const sigBuf = Buffer.from(sig);
  const expBuf = Buffer.from(expected);
  if (sigBuf.length !== expBuf.length || !crypto.timingSafeEqual(sigBuf, expBuf)) return null;
  let payload;
  try {
    payload = JSON.parse(Buffer.from(body, 'base64url').toString('utf8'));
  } catch (e) {
    return null;
  }
  if (!payload.exp || payload.exp < Date.now()) return null;
  return payload;
}
function requireAuth(req, res, next) {
  const session = verifySession(req.cookies.session);
  if (session) { req.session = session; return next(); }
  return res.status(401).json({ error: 'Не авторизовано. Войдите заново.' });
}

app.post('/api/login', (req, res) => {
  const { password } = req.body;
  if (!process.env.CABINET_PASSWORD) {
    return res.status(500).json({ error: 'На сервере не задан CABINET_PASSWORD (см. .env.example)' });
  }
  if (password !== process.env.CABINET_PASSWORD) {
    return res.status(401).json({ error: 'Неверный пароль' });
  }
  const exp = Date.now() + SESSION_TTL_MS;
  const token = signSession({ exp });
  res.cookie('session', token, { httpOnly: true, sameSite: 'lax', maxAge: SESSION_TTL_MS });
  res.json({ ok: true });
});

app.post('/api/logout', requireAuth, (req, res) => {
  res.clearCookie('session');
  res.json({ ok: true });
});

// ---------- настройки подключения к 1С ----------
app.get('/api/settings', requireAuth, async (req, res) => {
  const s = await getStore('settings', {});
  // пароль наружу не отдаём, только признак "задан/не задан"
  res.json({
    baseUrl: s.baseUrl || '',
    login: s.login || '',
    passwordSet: !!s.password,
    orgKey: s.orgKey || '',
    orgName: s.orgName || '',
    accountKey: s.accountKey || '',
    accountName: s.accountName || '',
  });
});

app.post('/api/settings', requireAuth, async (req, res) => {
  const { baseUrl, login, password, orgKey, orgName, accountKey, accountName } = req.body;
  const current = await getStore('settings', {});
  await setStore('settings', {
    baseUrl: baseUrl ?? current.baseUrl,
    login: login ?? current.login,
    password: password || current.password, // пустое поле = пароль не менять
    orgKey: orgKey ?? current.orgKey,
    orgName: orgName ?? current.orgName,
    accountKey: accountKey ?? current.accountKey,
    accountName: accountName ?? current.accountName,
  });
  await addHistory('Обновлены настройки подключения к 1С', '—');
  res.json({ ok: true });
});

// Проверяет подключение к 1С и возвращает список организаций и расчётных
// счетов, чтобы их можно было выбрать из выпадающего списка, не зная GUID.
app.post('/api/settings/browse', requireAuth, async (req, res) => {
  const { baseUrl, login, password } = req.body;
  if (!baseUrl || !login) {
    return res.status(400).json({ error: 'Укажите адрес и логин перед проверкой' });
  }
  const current = await getStore('settings', {});
  const pass = password || current.password;
  const base = baseUrl.replace(/\/+$/, '');
  const auth = Buffer.from(`${login}:${pass}`).toString('base64');

  const attempts = []; // для диагностики — что именно пробовали и что ответила 1С

  async function tryFetchList(catalogName) {
    let response;
    try {
      response = await fetch(`${base}/${catalogName}?$format=json&$select=Ref_Key,Description&$top=100`, {
        headers: { Authorization: `Basic ${auth}`, Accept: 'application/json' },
      });
    } catch (e) {
      attempts.push(`${catalogName}: не удалось подключиться (${e.message})`);
      return null;
    }
    if (!response.ok) {
      const text = await response.text().catch(() => '');
      attempts.push(`${catalogName}: код ${response.status} — ${text.slice(0, 150)}`);
      return null;
    }
    const data = await response.json().catch(() => null);
    if (!data) { attempts.push(`${catalogName}: ответ не в формате JSON`); return null; }
    return (data.value || []).map(x => ({ key: x.Ref_Key, name: x.Description }));
  }

  // Названия справочников отличаются между конфигурациями 1С (РФ/КЗ,
  // Бухгалтерия/ERP/УТ) — пробуем самые частые варианты по очереди.
  async function fetchFirstMatch(catalogNames) {
    for (const name of catalogNames) {
      const result = await tryFetchList(name);
      if (result !== null) return { name, result };
    }
    return null;
  }

  try {
    const orgMatch = await fetchFirstMatch(['Catalog_Организации']);
    if (!orgMatch) {
      return res.status(502).json({ error: 'Не удалось получить справочник организаций. Подробности: ' + attempts.join(' | ') });
    }
    const accMatch = await fetchFirstMatch([
      'Catalog_БанковскиеСчета',
      'Catalog_БанковскиеСчетаОрганизаций',
      'Catalog_РасчетныеСчета',
      'Catalog_РасчётныеСчета',
      'Catalog_СчетаОрганизаций',
    ]);
    if (!accMatch) {
      return res.status(502).json({ error: 'Организации нашлись, но не удалось найти справочник расчётных счетов. Подробности: ' + attempts.join(' | ') });
    }
    res.json({ organizations: orgMatch.result, accounts: accMatch.result, accountCatalogUsed: accMatch.name });
  } catch (e) {
    res.status(502).json({ error: e.message });
  }
});

// ---------- загрузка и разбор выписки (.xlsx) ----------
app.post('/api/upload', requireAuth, upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'Файл не получен' });

  // 1. bankParser: xlsx → единый объект операции (см. lib/bankParser.js),
  // не зависящий от того, какой это банк и как называются его колонки.
  let operations, parseFormat;
  try {
    const parsed = bankParser.parseStatement(req.file.buffer, req.file.originalname);
    operations = parsed.operations;
    parseFormat = parsed.format;
  } catch (e) {
    return res.status(400).json({ error: 'Не получилось прочитать файл: ' + e.message });
  }

  const all = await getStore('operations', []);

  // 2. duplicateGuard: не добавляем операцию повторно, если такая же (по
  // отпечатку — дата+сумма+БИН+название+назначение) уже есть в списке —
  // иначе повторная загрузка того же файла (или пересекающийся период в
  // другом файле) задваивает список.
  operations.forEach(op => { op.fingerprint = makeFingerprint(op); });
  const existingFingerprints = new Set(all.map(o => o.fingerprint || makeFingerprint(o)));
  const newOnly = operations.filter(op => !existingFingerprints.has(op.fingerprint));
  const skippedDuplicates = operations.length - newOnly.length;

  const updated = [...all, ...newOnly];
  await setStore('operations', updated);
  await addHistory(
    `Загружена выписка ${req.file.originalname} (формат разбора: ${parseFormat}) · ${newOnly.length} новых операций` +
      (skippedDuplicates ? ` · ${skippedDuplicates} пропущено как дубли` : ''),
    '—'
  );

  // 3. reconcileStatement: сопоставление контрагента/истории/договора с 1С
  // и проверка "не разнесено ли уже" — только если подключение настроено.
  const settingsForMatch = await getStore('settings', {});
  const reconciliation = await reconcileStatement(newOnly, updated, settingsForMatch);
  if (reconciliation.connected) {
    await setStore('operations', updated);
  }
  await addHistory(
    `Сверка с 1С: из ${reconciliation.total} операций — уже в 1С: ${reconciliation.alreadyIn1c}, новых контрагентов: ${reconciliation.newCounterparty}`,
    '—'
  );

  res.json({ operations: newOnly, skippedDuplicates, reconciliation });
});

// ---------- список операций на проверку ----------
app.get('/api/operations', requireAuth, async (req, res) => {
  const ops = await getStore('operations', []);
  const rules = await getStore('rules', []);
  // Приоритет определения статьи ДДС и счёта:
  //   1. Ваша ручная правка (если поправили в кабинете) — самый высокий приоритет
  //   2. История 1С (как этот контрагент разносился раньше)
  //   3. Ваши правила из раздела «Правила» (включая типовые из регламента и КНП)
  //   4. Базовая логика покупатель/поставщик по регламенту (всегда есть)
  const withCategories = ops.map(op => {
    const { category, account, operationKind, debtType, isDefaultOnly, needsAttention } = classifyOperation(op, rules);
    return {
      ...op,
      suggestedCategory: category,
      suggestedAccount: account,
      suggestedOperationKind: operationKind,
      suggestedDebtType: debtType,
      isDefaultOnly, // категория определена только базовой логикой, стоит перепроверить глазами
      needsAttention,
    };
  });
  res.json(withCategories);
});

// ---------- ручная правка операции перед подтверждением ----------
// Позволяет поправить статью ДДС / счёт / договор прямо в карточке, до
// создания документа в 1С. Если saveAsRule=true — тут же создаётся правило,
// чтобы в следующий раз похожая операция определилась автоматически
// (самообучение на ваших правках).
app.post('/api/operations/:id/override', requireAuth, async (req, res) => {
  const all = await getStore('operations', []);
  const op = all.find(o => o.id === req.params.id);
  if (!op) return res.status(404).json({ error: 'Операция не найдена' });

  const { category, account, operationKind, debtType, contractKey, contractName, saveAsRule, ruleContains, ruleField } = req.body;

  if (category !== undefined) op.manualCategory = String(category || '');
  if (account !== undefined) op.manualAccount = String(account || '');
  if (operationKind !== undefined) op.manualOperationKind = String(operationKind || '');
  if (debtType !== undefined) op.manualDebtType = String(debtType || '');
  if (contractKey !== undefined) {
    op.contractKey = contractKey || null;
    op.contractName = contractName || '';
    op.contractStatus = contractKey ? 'matched' : op.contractStatus;
  }
  await setStore('operations', all);

  let ruleAdded = false;
  if (saveAsRule && ruleContains && (category || account || operationKind || debtType)) {
    const rules = await getStore('rules', []);
    rules.unshift({
      id: crypto.randomUUID(),
      field: normalizeRuleField(ruleField),
      contains: String(ruleContains),
      category: String(category || ''),
      account: String(account || ''),
      operationKind: String(operationKind || ''),
      debtType: String(debtType || ''),
    });
    await setStore('rules', rules);
    ruleAdded = true;
    await addHistory(`Сохранено новое правило по вашей правке: "${ruleContains}" → ${category || ''} ${account ? '(счёт ' + account + ')' : ''}`, '—');
  } else {
    await addHistory(`Ручная правка операции: ${op.counterparty || 'операция'} · статья: ${category || '—'}, счёт: ${account || '—'}`, '—');
  }

  res.json({ ok: true, op, ruleAdded });
});

// ---------- подтверждение операции: создать черновик в 1С ----------
// Вся логика "подтвердить операцию" вынесена в отдельную функцию, чтобы
// ей могли пользоваться и одиночный эндпоинт /confirm, и пакетный
// /batch-create-documents — без копирования одного и того же кода дважды.
// Мутирует op и пишет в историю сама; вызывающий код отвечает только за
// getStore/setStore и формирование HTTP-ответа.
//
// Возвращает:
//   { ok: true,  docNumber, droppedFields }
//   { ok: false, httpStatus, error }
async function confirmOperation(op, all, settings, rules) {
  // ПЕРЕПРОВЕРКА С 1С ПРЯМО ПЕРЕД СОЗДАНИЕМ ДОКУМЕНТА. Данные, с которыми
  // сопоставлялась операция, могли устареть с момента загрузки выписки —
  // между загрузкой и подтверждением могло пройти много времени, кто-то
  // мог создать тот же документ вручную, появиться новые проводки по
  // контрагенту и т.д. Поэтому перед составлением документа не полагаемся
  // на то, что было найдено при загрузке, а анализируем 1С заново.

  // 1. Ещё раз проверяем, не появился ли уже такой документ в 1С (та же
  // дата и сумма) — чтобы не создать дубль, даже если при загрузке
  // выписки его ещё не было.
  try {
    const alreadyExists = await checkExistingInOnec(op, settings);
    if (alreadyExists) {
      op.status = 'already_in_1c';
      await addHistory(`Перепроверка перед подтверждением: в 1С уже есть такой документ — черновик не создаём (${op.counterparty || 'операция'})`, '—');
      return { ok: false, httpStatus: 409, error: 'В 1С уже есть документ с такой же датой и суммой — черновик не создан, чтобы не задвоить. Операция помечена как «Уже в 1С».' };
    }
  } catch (e) {
    // Если сверка технически не удалась (например, база временно недоступна) —
    // не блокируем работу, идём дальше с тем, что уже известно.
  }

  // 2. Заново анализируем контрагента и проводки по нему: если контрагент
  // ещё не был сопоставлен (например, операция загружалась без
  // подключения к 1С) — сопоставляем сейчас; если уже сопоставлен —
  // обновляем историческую категорию/договор по самым свежим проводкам,
  // а не по тому, что было найдено когда-то при загрузке.
  if (!op.counterpartyKey && op.bin) {
    try {
      const found = await findCounterpartyByBin(op.bin, settings, op.counterparty);
      if (found) {
        op.counterpartyKey = found.Ref_Key;
        op.counterpartyMatchedName = found.Description || '';
      } else {
        op.status = 'new_counterparty';
      }
    } catch (e) {
      if (e.ambiguousCounterparty) {
        // Несколько похожих контрагентов — не гадаем, останавливаемся и
        // просим выбрать вручную (та же механика, что при загрузке выписки).
        op.status = 'ambiguous_counterparty';
        op.counterpartyOptions = e.ambiguousCounterparty.map(o => ({ key: o.Ref_Key, name: o.Description }));
      }
      // Иначе — сверка технически не удалась, продолжаем без
      // автосопоставления, ниже сработает обычная проверка "контрагент не
      // сопоставлен".
    }
  }
  if (op.status === 'new_counterparty' || op.status === 'ambiguous_counterparty') {
    return {
      ok: false,
      httpStatus: 400,
      error: 'Контрагент не сопоставлен с 1С — сначала обработайте это в карточке операции (создайте контрагента или выберите из найденных вариантов).',
    };
  }
  if (op.counterpartyKey) {
    try {
      const historical = await findHistoricalCategory(op.counterpartyKey, op.amount, settings, op.purpose);
      if (historical) {
        op.historicalCategory = historical.categoryName || op.historicalCategory || '';
        op.historicalCategoryKey = historical.categoryKey || op.historicalCategoryKey || null;
        op.historicalOperationKind = historical.operationKind || op.historicalOperationKind || '';
        op.historicalDebtType = historical.debtType || op.historicalDebtType || '';
      }
    } catch (e) {
      // Не критично — используем то, что было определено раньше (если было).
    }
    try {
      let contract = await resolveSupplierContract(op.counterpartyKey, op.purpose, settings);
      // Договора нет вообще, и "Без договора" тоже ещё нет — для оплаты
      // поставщику создаём "Без договора" сразу здесь, автоматически, не
      // дожидаясь отдельного клика: findNoContractRecord внутри
      // resolveSupplierContract уже убедился, что его действительно нет.
      if (contract && contract.status === 'need_create_no_contract') {
        try {
          const createdContract = await createNoContractInOnec(op.counterpartyKey, settings);
          contract = { status: 'matched', key: createdContract.Ref_Key, name: createdContract.Description || NO_CONTRACT_NAME };
          await addHistory(`Автоматически создан договор «Без договора» для контрагента: ${op.counterparty || op.counterpartyMatchedName || ''}`, '—');
        } catch (e) {
          // Не критично — просто оставим договор пустым, ниже это попадёт в droppedFields.
        }
      }
      if (contract) {
        op.contractStatus = contract.status;
        if (contract.status === 'matched') {
          op.contractKey = contract.key;
          op.contractName = contract.name;
        } else if (contract.status === 'ambiguous') {
          op.contractOptions = contract.options;
        }
      }
    } catch (e) {
      // Не критично — используем то, что было определено раньше.
    }
  }

  // Статья ДДС обязательна — без неё документ не создаём. Приоритет:
  // ручная правка > история 1С (уже обновлённая на шаге перепроверки выше)
  // > правила > базовая логика (которая всегда что-то определяет, так что
  // пусто может быть, только если что-то совсем не так с данными).
  const { category, account, operationKind, debtType } = classifyOperation(op, rules);
  if (!category) {
    return { ok: false, httpStatus: 400, error: 'Не определена статья ДДС — поправьте вручную в карточке операции, прежде чем подтверждать.' };
  }

  // Проверка дубликатов НА УРОВНЕ НАШЕГО СПИСКА операций (в дополнение к
  // проверке 1С выше в шаге 1): если среди уже загруженных операций есть
  // другая, с таким же отпечатком (дата+сумма+БИН+контрагент+назначение), и
  // по ней документ УЖЕ был создан — значит, эта операция, скорее всего,
  // попала в список второй раз (например, из пересекающегося периода двух
  // выгрузок) и создавать по ней ещё один черновик не нужно.
  op.fingerprint = makeFingerprint(op);
  const duplicateOf = findDuplicateAmongOperations(op, all);
  if (duplicateOf) {
    op.status = 'duplicate';
    await addHistory(`Пропущено как дубликат уже созданного документа (№${duplicateOf.onecDocNumber || '—'}): ${op.counterparty || 'операция'}`, '—');
    return {
      ok: false,
      httpStatus: 409,
      error: 'Операция с такими же датой, суммой, БИН/ИИН, контрагентом и назначением платежа уже разнесена ранее — черновик не создан, чтобы не задвоить.',
    };
  }

  try {
    const result = await createDraftInOnec(op, settings, category, account, operationKind, debtType);
    op.status = 'draft_created';
    op.documentCreated = true;
    op.onecDocNumber = result.docNumber || null;
    // Если какие-то поля (счёт учёта, подотчётник и т.п.) не прижились в вашей
    // конфигурации 1С — черновик всё равно создаётся, но мы честно пишем в
    // историю, что именно не заполнилось, чтобы можно было доглядеть глазами
    // и, при желании, прислать нам точный текст ошибки для донастройки.
    const droppedNote = result.droppedFields && result.droppedFields.length
      ? ` · 1С не приняла поля: ${result.droppedFields.join(', ')} — проверьте их вручную в документе`
      : '';
    await addHistory(`Создан черновик в 1С: ${op.counterparty || 'операция'} · ${op.amount} ₸ · статья: ${category}${droppedNote}`, result.docNumber || '—');
    return { ok: true, docNumber: result.docNumber || null, droppedFields: result.droppedFields || [] };
  } catch (e) {
    await addHistory(`Ошибка при создании черновика: ${e.message}`, '—');
    return { ok: false, httpStatus: 502, error: 'Не удалось создать документ в 1С: ' + e.message };
  }
}

app.post('/api/operations/:id/confirm', requireAuth, async (req, res) => {
  const all = await getStore('operations', []);
  const op = all.find(o => o.id === req.params.id);
  if (!op) return res.status(404).json({ error: 'Операция не найдена' });

  const settings = await getStore('settings', {});
  if (!settings.baseUrl) {
    return res.status(400).json({ error: 'Сначала укажите адрес подключения к 1С в разделе «Настройки»' });
  }
  const rules = await getStore('rules', []);

  const result = await confirmOperation(op, all, settings, rules);
  await setStore('operations', all);

  if (!result.ok) {
    return res.status(result.httpStatus).json({ error: result.error, op });
  }
  res.json({ ok: true, op, droppedFields: result.droppedFields || [] });
});

// ---------- пакетное создание черновиков документов сразу по нескольким операциям ----------
app.post('/api/operations/batch-create-documents', requireAuth, async (req, res) => {
  const { ids } = req.body;
  if (!Array.isArray(ids) || ids.length === 0) {
    return res.status(400).json({ error: 'Передайте массив ids операций для пакетного создания' });
  }

  const settings = await getStore('settings', {});
  if (!settings.baseUrl) {
    return res.status(400).json({ error: 'Сначала укажите адрес подключения к 1С в разделе «Настройки»' });
  }

  const all = await getStore('operations', []);
  const rules = await getStore('rules', []);
  const created = [];
  const errors = [];

  // Обрабатываем ПОСЛЕДОВАТЕЛЬНО, а не все параллельно: во-первых, так
  // проверка дубликатов внутри confirmOperation корректно видит документы,
  // уже созданные другими операциями этого же пакета; во-вторых, не
  // заваливаем 1С множеством одновременных запросов.
  for (const id of ids) {
    const op = all.find(o => o.id === id);
    if (!op) {
      errors.push({ id, error: 'Операция не найдена' });
      continue;
    }
    if (op.documentCreated) {
      errors.push({ id, error: 'Документ по этой операции уже был создан ранее', docNumber: op.onecDocNumber || null });
      continue;
    }
    const result = await confirmOperation(op, all, settings, rules);
    if (result.ok) {
      created.push({ id, docNumber: result.docNumber, droppedFields: result.droppedFields || [] });
    } else {
      errors.push({ id, error: result.error });
    }
  }

  await setStore('operations', all);
  await addHistory(`Пакетное создание документов: ${created.length} создано, ${errors.length} с ошибками из ${ids.length}`, '—');
  res.json({ created, errors });
});

// ---------- выбрать контрагента вручную из неоднозначных вариантов ----------
app.post('/api/operations/:id/choose-counterparty', requireAuth, async (req, res) => {
  const all = await getStore('operations', []);
  const op = all.find(o => o.id === req.params.id);
  if (!op) return res.status(404).json({ error: 'Операция не найдена' });

  const { counterpartyKey, counterpartyName } = req.body;
  if (!counterpartyKey) return res.status(400).json({ error: 'Не выбран контрагент' });

  op.counterpartyKey = counterpartyKey;
  op.counterpartyMatchedName = counterpartyName || '';
  op.status = 'review';
  op.counterpartyOptions = undefined;

  const settings = await getStore('settings', {});
  if (settings.baseUrl) {
    try {
      const historical = await findHistoricalCategory(counterpartyKey, op.amount, settings, op.purpose);
      if (historical) {
        op.historicalCategory = historical.categoryName || '';
        op.historicalCategoryKey = historical.categoryKey || '';
        op.historicalOperationKind = historical.operationKind || '';
        op.historicalDebtType = historical.debtType || '';
      }
      const contract = await resolveSupplierContract(counterpartyKey, op.purpose, settings);
      if (contract) {
        op.contractStatus = contract.status;
        if (contract.status === 'matched') {
          op.contractKey = contract.key;
          op.contractName = contract.name;
        }
      }
    } catch (e) {
      // Не критично — просто без истории/договора для этой операции.
    }
  }

  await setStore('operations', all);
  await addHistory(`Контрагент выбран вручную: ${counterpartyName}`, '—');
  res.json({ ok: true, op });
});

// ---------- пересверить одну операцию с 1С заново ----------
// Нужно, когда операция была загружена до того, как подключение к 1С
// было настроено (или сверка не удалась) — прогоняем ту же логику,
// что и при первой загрузке, но для одной конкретной операции.
app.post('/api/operations/:id/rematch', requireAuth, async (req, res) => {
  const all = await getStore('operations', []);
  const op = all.find(o => o.id === req.params.id);
  if (!op) return res.status(404).json({ error: 'Операция не найдена' });

  const settings = await getStore('settings', {});
  if (!settings.baseUrl) {
    return res.status(400).json({ error: 'Сначала укажите адрес подключения к 1С в разделе «Настройки»' });
  }

  try {
    const alreadyExists = await checkExistingInOnec(op, settings);
    if (alreadyExists) {
      op.status = 'already_in_1c';
      await setStore('operations', all);
      return res.json({ ok: true, op });
    }

    if (op.bin) {
      const found = await findCounterpartyByBin(op.bin, settings, op.counterparty);
      if (found) {
        op.counterpartyKey = found.Ref_Key;
        op.counterpartyMatchedName = found.Description || '';
        op.status = 'review';

        const historical = await findHistoricalCategory(found.Ref_Key, op.amount, settings, op.purpose).catch(() => null);
        if (historical) {
          op.historicalCategory = historical.categoryName || '';
          op.historicalCategoryKey = historical.categoryKey || '';
          op.historicalOperationKind = historical.operationKind || '';
        }

        const contract = await resolveSupplierContract(found.Ref_Key, op.purpose, settings).catch(() => null);
        if (contract) {
          op.contractStatus = contract.status;
          if (contract.status === 'matched') {
            op.contractKey = contract.key;
            op.contractName = contract.name;
          } else if (contract.status === 'ambiguous') {
            op.contractOptions = contract.options;
          }
        }
      } else {
        op.status = 'new_counterparty';
      }
    }
    await setStore('operations', all);
    await addHistory(`Пересверено с 1С: ${op.counterparty || 'операция'}`, '—');
    res.json({ ok: true, op });
  } catch (e) {
    res.status(502).json({ error: 'Не удалось пересверить: ' + e.message });
  }
});

// ---------- создание НОВОГО контрагента (отдельное подтверждение) ----------
// Создание контрагента остаётся строго РУЧНЫМ действием — сервер никогда не
// создаёт нового контрагента сам во время загрузки выписки или подтверждения
// документа. Этот эндпоинт срабатывает только по явному клику на кнопку
// "Создать нового" в карточке операции, и перед созданием ещё раз строго
// проверяет и по нормализованному БИН/ИИН, и по нормализованному названию —
// чтобы не задвоить контрагента, который на самом деле уже есть в 1С.
app.post('/api/operations/:id/create-counterparty', requireAuth, async (req, res) => {
  const all = await getStore('operations', []);
  const op = all.find(o => o.id === req.params.id);
  if (!op) return res.status(404).json({ error: 'Операция не найдена' });
  if (!op.bin && !op.counterparty) {
    return res.status(400).json({ error: 'У операции нет ни БИН/ИИН, ни названия — нечего создавать. Сопоставьте контрагента вручную.' });
  }

  const settings = await getStore('settings', {});
  if (!settings.baseUrl) {
    return res.status(400).json({ error: 'Сначала укажите адрес подключения к 1С в разделе «Настройки»' });
  }

  try {
    // Строгая повторная проверка прямо перед созданием — по нормализованному
    // БИН/ИИН (только цифры) и по нормализованному названию (без ТОО/ЖШС/ИП
    // и кавычек, см. normalizeCounterpartyName) — вдруг контрагент уже есть,
    // просто под слегка другим написанием.
    const existing = await findCounterpartyByBin(op.bin, settings, op.counterparty);
    if (existing) {
      op.counterpartyKey = existing.Ref_Key;
      op.counterpartyMatchedName = existing.Description;
      op.status = 'review'; // контрагент есть — можно подтверждать документ как обычно
      await setStore('operations', all);
      await addHistory(`Контрагент найден в 1С при повторной проверке: ${existing.Description} (БИН ${op.bin || '—'})`, '—');
      return res.json({ ok: true, op, created: false });
    }

    const createdRecord = await createCounterpartyInOnec(op, settings);
    op.counterpartyKey = createdRecord.Ref_Key;
    op.counterpartyMatchedName = createdRecord.Description || op.counterparty || '';
    op.status = 'review';
    await setStore('operations', all);
    await addHistory(`Создан новый контрагент в 1С: ${op.counterpartyMatchedName} (БИН ${op.bin || '—'})`, '—');
    res.json({ ok: true, op, created: true });
  } catch (e) {
    await addHistory(`Ошибка при создании контрагента: ${e.message}`, '—');
    res.status(502).json({ error: 'Не удалось создать контрагента в 1С: ' + e.message });
  }
});

// ---------- создание договора "Без договора" (только по явному действию бухгалтера) ----------
app.post('/api/operations/:id/create-no-contract', requireAuth, async (req, res) => {
  const all = await getStore('operations', []);
  const op = all.find(o => o.id === req.params.id);
  if (!op) return res.status(404).json({ error: 'Операция не найдена' });
  if (!op.counterpartyKey) return res.status(400).json({ error: 'У операции ещё не сопоставлен контрагент' });

  const settings = await getStore('settings', {});
  if (!settings.baseUrl) {
    return res.status(400).json({ error: 'Сначала укажите адрес подключения к 1С в разделе «Настройки»' });
  }

  try {
    // Перепроверяем прямо перед созданием — вдруг договор уже появился
    // (например, кто-то создал его в 1С, пока вы смотрели список).
    const fresh = await resolveSupplierContract(op.counterpartyKey, op.purpose, settings);
    if (fresh.status === 'matched') {
      op.contractStatus = 'matched';
      op.contractKey = fresh.key;
      op.contractName = fresh.name;
      await setStore('operations', all);
      await addHistory(`Договор уже найден в 1С при повторной проверке: ${fresh.name}`, '—');
      return res.json({ ok: true, op });
    }
    if (fresh.status === 'ambiguous') {
      op.contractStatus = 'ambiguous';
      op.contractOptions = fresh.options;
      await setStore('operations', all);
      return res.status(409).json({ error: 'У контрагента уже появилось несколько договоров — выберите нужный вручную.', op });
    }

    const created = await createNoContractInOnec(op.counterpartyKey, settings);
    op.contractStatus = 'matched';
    op.contractKey = created.Ref_Key;
    op.contractName = created.Description || NO_CONTRACT_NAME;
    await setStore('operations', all);
    await addHistory(`Создан договор «Без договора» для контрагента: ${op.counterparty || op.counterpartyMatchedName}`, '—');
    res.json({ ok: true, op });
  } catch (e) {
    await addHistory(`Ошибка при создании договора «Без договора»: ${e.message}`, '—');
    res.status(502).json({ error: 'Не удалось создать договор «Без договора» в 1С: ' + e.message });
  }
});

// ---------- ручной поиск контрагента в 1С (когда автопоиск ничего не нашёл) ----------
// Раньше при статусе "Новый контрагент" было только одно действие — создать
// нового. Это опасно, если контрагент на самом деле уже есть в 1С, просто
// автопоиск его не нашёл (например, из-за другого написания названия) —
// получалось задвоение. Теперь перед созданием можно вручную поискать в
// справочнике 1С и выбрать существующего вместо создания нового (кнопка
// "Это он" в интерфейсе вызывает уже существующий /choose-counterparty).
app.post('/api/operations/:id/search-counterparty', requireAuth, async (req, res) => {
  const all = await getStore('operations', []);
  const op = all.find(o => o.id === req.params.id);
  if (!op) return res.status(404).json({ error: 'Операция не найдена' });

  const { query } = req.body;
  if (!query || query.trim().length < 2) {
    return res.status(400).json({ error: 'Введите минимум 2 символа для поиска' });
  }

  const settings = await getStore('settings', {});
  if (!settings.baseUrl) {
    return res.status(400).json({ error: 'Сначала укажите адрес подключения к 1С в разделе «Настройки»' });
  }

  try {
    const results = await searchCounterpartyByText(query.trim(), settings);
    res.json({ options: results.map(o => ({ key: o.Ref_Key, name: o.Description })) });
  } catch (e) {
    res.status(502).json({ error: 'Не удалось выполнить поиск: ' + e.message });
  }
});

// ---------- правила категоризации (редактируются прямо в кабинете) ----------
app.get('/api/rules', requireAuth, async (req, res) => {
  res.json(await getStore('rules', []));
});

// Отдаёт правила текстом в том же формате, что и массовая загрузка —
// чтобы вставить в Render → Environment → RULES_TEXT и не терять их
// при следующей пересборке сайта (актуально только без DATABASE_URL).
app.get('/api/rules/export', requireAuth, async (req, res) => {
  const rules = await getStore('rules', []);
  const text = rules.map(r => {
    const field = normalizeRuleField(r.field);
    // Если вид операции/задолженности не заданы, а поле — не counterparty/knp,
    // можно смело обрезать хвост, но проще всегда писать все 6 колонок —
    // так формат остаётся однозначным при повторном импорте.
    return `${r.contains} | ${r.category} | ${r.account} | ${field} | ${r.operationKind || ''} | ${r.debtType || ''}`;
  }).join('\n');
  res.json({ text });
});

app.post('/api/rules', requireAuth, async (req, res) => {
  const { field, contains, category, account, operationKind, debtType } = req.body;
  if (!contains || (!category && !account)) {
    return res.status(400).json({ error: 'Укажите условие и хотя бы одно из: категорию или счёт' });
  }
  const rules = await getStore('rules', []);
  rules.push({
    id: crypto.randomUUID(),
    field: normalizeRuleField(field),
    contains: String(contains),
    category: String(category || ''),
    account: String(account || ''),
    operationKind: String(operationKind || ''),
    debtType: String(debtType || ''),
  });
  await setStore('rules', rules);
  await addHistory(`Добавлено правило: "${contains}" → ${category || ''} ${account ? '(счёт ' + account + ')' : ''}`, '—');
  res.json(rules);
});

// Массовый импорт правил текстом — по одному правилу на строку, поля через
// вертикальную черту: текст-условие | категория (статья ДДС) | счёт | (поле: purpose/counterparty/knp, необязательно)
// Пример строки: аренда | Аренда офиса | 3360
app.post('/api/rules/import', requireAuth, async (req, res) => {
  const { text } = req.body;
  if (!text || !text.trim()) {
    return res.status(400).json({ error: 'Вставьте текст с правилами' });
  }
  const rules = await getStore('rules', []);
  const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
  let added = 0;
  let skipped = 0;
  for (const line of lines) {
    const parts = line.split('|').map(p => p.trim());
    if (parts.length < 2 || !parts[0]) { skipped++; continue; }
    // Формат: текст | статья ДДС | счёт | поле(purpose/counterparty/knp) | вид операции | вид задолженности
    // Все поля, кроме первого, необязательны.
    const [contains, category, account, fieldRaw, operationKind, debtType] = parts;
    if (!category && !account && !operationKind && !debtType) { skipped++; continue; }
    rules.push({
      id: crypto.randomUUID(),
      field: normalizeRuleField(fieldRaw),
      contains,
      category: category || '',
      account: account || '',
      operationKind: operationKind || '',
      debtType: debtType || '',
    });
    added++;
  }
  await setStore('rules', rules);
  await addHistory(`Импортировано правил: ${added}${skipped ? `, пропущено строк: ${skipped}` : ''}`, '—');
  res.json({ rules, added, skipped });
});

app.delete('/api/rules/:id', requireAuth, async (req, res) => {
  const rules = await getStore('rules', []);
  const filtered = rules.filter(r => r.id !== req.params.id);
  await setStore('rules', filtered);
  await addHistory('Удалено правило категоризации', '—');
  res.json(filtered);
});

// ---------- очистить список операций (начать заново) ----------
app.post('/api/operations/clear', requireAuth, async (req, res) => {
  await setStore('operations', []);
  await addHistory('Список операций очищен вручную', '—');
  res.json({ ok: true });
});

// ---------- история ----------
app.get('/api/history', requireAuth, async (req, res) => {
  res.json(await getStore('history', []));
});
async function addHistory(action, doc) {
  const h = await getStore('history', []);
  h.unshift({ time: new Date().toISOString(), action, doc });
  await setStore('history', h.slice(0, 500));
}

// createDraftInOnec теперь тонкая обёртка над buildPayload (lib/buildPayload.js)
// + writeTo1C (lib/writeTo1C.js) — раньше сборка payload и его отправка были
// одной функцией. Поведение и сигнатура не изменились, вызывающий код
// (confirmOperation, /batch-create-documents) не нужно трогать.
async function createDraftInOnec(op, settings, resolvedCategory, resolvedAccount, resolvedOperationKind, resolvedDebtType) {
  const { endpoint, payload, missingNoContract } = await buildPayload(
    op, settings, resolvedCategory, resolvedAccount, resolvedOperationKind, resolvedDebtType
  );
  return writeTo1C(endpoint, payload, settings, { missingNoContract });
}

// ---------- начальные данные при первом запуске / восстановление после передеплоя ----------
async function seedDefaultsIfNeeded() {
  // Правила: если ключа ещё нет вообще — создаём типовые по регламенту.
  // Если уже есть — прогоняем разовую миграцию старого формата счёта
  // (было "3360"/"3310", стало "3360/1710"/"3310/1710").
  const existingRules = await getStore('rules', null);
  if (existingRules === null) {
    await setStore('rules', DEFAULT_RULES.map(r => ({ id: crypto.randomUUID(), ...r })));
  } else {
    const upgrades = { '3360': '3360/1710', '3310': '3310/1710' };
    let changed = false;
    for (const r of existingRules) {
      if (upgrades[r.account]) { r.account = upgrades[r.account]; changed = true; }
    }
    if (changed) await setStore('rules', existingRules);
  }

  // Восстановление правил из RULES_TEXT — актуально только в файловом режиме
  // (без DATABASE_URL), потому что только там правила теряются при передеплое.
  if (process.env.RULES_TEXT && !pgPool) {
    const current = await getStore('rules', []);
    const looksLikeOnlyDefaults = current.length <= DEFAULT_RULES.length;
    if (looksLikeOnlyDefaults) {
      const lines = process.env.RULES_TEXT.split('\n').map(l => l.trim()).filter(Boolean);
      const restored = [];
      for (const line of lines) {
        const parts = line.split('|').map(p => p.trim());
        if (parts.length < 2) continue;
        const [contains, category, account, fieldRaw, operationKind, debtType] = parts;
        if (!category && !account && !operationKind && !debtType) continue;
        restored.push({
          id: crypto.randomUUID(),
          field: normalizeRuleField(fieldRaw),
          contains, category: category || '', account: account || '',
          operationKind: operationKind || '', debtType: debtType || '',
        });
      }
      if (restored.length) await setStore('rules', restored);
    }
  }

  // Настройки подключения к 1С: если записи ещё нет — берём из переменных
  // окружения ONEC_*. Если запись есть, но пустая, а в окружении данные
  // есть — тоже восстанавливаем (актуально для файлового режима после
  // передеплоя без DATABASE_URL).
  const envSettings = {
    baseUrl: process.env.ONEC_BASE_URL || '',
    login: process.env.ONEC_LOGIN || '',
    password: process.env.ONEC_PASSWORD || '',
    orgKey: process.env.ONEC_ORG_KEY || '',
    orgName: process.env.ONEC_ORG_NAME || '',
    accountKey: process.env.ONEC_ACCOUNT_KEY || '',
    accountName: process.env.ONEC_ACCOUNT_NAME || '',
  };
  const currentSettings = await getStore('settings', null);
  if (currentSettings === null) {
    await setStore('settings', envSettings.baseUrl ? envSettings : { baseUrl: '', login: '', password: '' });
  } else if (envSettings.baseUrl && !currentSettings.baseUrl) {
    await setStore('settings', envSettings);
  }
}

async function main() {
  await ensureStorage();
  await seedDefaultsIfNeeded();
  app.listen(PORT, () => {
    console.log(`Кабинет запущен: http://localhost:${PORT}`);
    console.log(
      pgPool
        ? 'Хранилище: PostgreSQL (DATABASE_URL задан) — операции/история/правила/настройки переживут передеплой.'
        : 'Хранилище: локальные JSON-файлы — на бесплатном Render они стираются при каждом передеплое. Задайте DATABASE_URL, чтобы это исправить.'
    );
  });
}
main().catch(err => {
  console.error('Не удалось запустить сервер:', err);
  process.exit(1);
});
