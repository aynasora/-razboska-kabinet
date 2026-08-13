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
    if(o.status==='draft_created') return 3;
    return 1;
  };
  const sorted = [...ops].sort((a,b)=>priority(a)-priority(b));

  list.innerHTML = '';
  sorted.forEach(o=>{
    const done = o.status==='draft_created';
    const isNew = o.status==='new_counterparty';
    const already = o.status==='already_in_1c';
    const hasCategory = !!o.suggestedCategory;
    const contractAmbiguous = o.contractStatus === 'ambiguous';
    const needsNoContract = o.contractStatus === 'need_create_no_contract' && !o.contractKey;
    // Бывает, что операция загружена ДО того, как подключение к 1С было
    // настроено — тогда сопоставление контрагента не выполнялось вообще,
    // но статус остался обычным "review". Без реального Ref_Key
    // контрагента документ в 1С создать нельзя — блокируем и предлагаем
    // пересверить.
    const missingCounterparty = !done && !already && !isNew && !o.counterpartyKey && o.status !== 'ambiguous_counterparty';
    const ambiguousCounterparty = o.status === 'ambiguous_counterparty';
    let statusHtml, actionHtml;
    if(done){
      statusHtml = '<span class="pill pill-approved">Черновик создан</span>';
      actionHtml = '';
    } else if(already){
      statusHtml = '<span class="pill pill-review" style="background:var(--paper-2);color:var(--muted);">Уже в 1С</span>';
      actionHtml = '';
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

// Типовые правила по регламенту («Правила ИИ-помощника 1С — Банк / ДДС / расчёты»).
// Создаются один раз, если у вас ещё нет ни одного правила — дальше вы можете
// их редактировать или удалять как обычные правила в разделе «Правила».
const DEFAULT_RULES = [
  { field: 'purpose', contains: 'аренд', category: 'Аренда', account: '3360/1710' },
  { field: 'purpose', contains: 'зарплат', category: 'Заработная плата', account: '3350' },
  { field: 'purpose', contains: 'налог', category: 'Налоги', account: '3130' },
  { field: 'purpose', contains: 'кпн', category: 'Налоги', account: '3130' },
  { field: 'purpose', contains: 'ипн', category: 'Налоги', account: '3120' },
  { field: 'purpose', contains: 'осмс', category: 'Налоги', account: '3150' },
  { field: 'purpose', contains: 'опвр', category: 'Налоги', account: '3150' },
  { field: 'purpose', contains: 'комисси', category: 'Банковские услуги', account: '3310/1710' },
  { field: 'purpose', contains: 'возврат', category: 'Возврат денежных средств', account: '3310/1710' },
];

// Дата у нас хранится как ДД.ММ.ГГГГ (как в банковской выписке), а OData
// в 1С ожидает международный формат ГГГГ-ММ-ДДTчч:мм:сс — переводим перед
// отправкой в 1С, иначе она отвечает «Не удалось разобрать строку как
// значение типа Edm.DateTime».
function toIsoDate(dateStr) {
  const parts = String(dateStr).split('.');
  if (parts.length !== 3) return dateStr; // уже похоже на другой формат — не трогаем
  const [day, month, year] = parts;
  return `${year}-${month.padStart(2, '0')}-${day.padStart(2, '0')}T00:00:00`;
}

// Точные внутренние имена для "Вида операции" — 1С хранит их слитно,
// без пробелов, с заглавной буквы у каждого слова, а не так, как
// показывает на экране. Подтверждённые значения (проверены напрямую
// через OData на вашей базе) — самые надёжные. Для остальных пунктов
// того же списка используется тот же принцип именования как обоснованная
// попытка — если она окажется неверной, при создании документа мы просто
// не отправим это поле, а не провалим создание черновика целиком.
const CONFIRMED_OPERATION_KINDS = {
  'перечисление денежных средств подотчетнику': 'ПеречислениеДенежныхСредствПодотчетнику',
};
function guessOperationKindLiteral(text) {
  const normalized = String(text || '').toLowerCase().trim();
  if (CONFIRMED_OPERATION_KINDS[normalized]) return CONFIRMED_OPERATION_KINDS[normalized];
  if (!text) return null;
  // Обоснованная попытка: каждое слово с заглавной буквы, слитно, без
  // знаков препинания — так называет свои перечисления 1С в этом списке.
  return String(text)
    .split(/[\s.,]+/)
    .filter(Boolean)
    .map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase())
    .join('');
}

function baseDefault(amount) {
  return amount >= 0
    ? { category: 'Реализация работ и услуг', account: '1210/3510' }
    : { category: 'Расчёты с поставщиками и подрядчиками', account: '3310/1710' };
}

// Некоторые латинские и кириллические буквы визуально неотличимы (a/а,
// e/е, o/о, p/р, c/с, x/х, y/у, k/к, h/н, m/м, t/т, b/в и т.д.). Если при
// вводе текста (в правиле или в самой банковской выписке) закралась "не
// та" буква — обычное сравнение строк не находит совпадение, хотя
// глазами всё выглядит одинаково. Эта функция приводит все похожие
// буквы к одному (кириллическому) варианту перед сравнением.
const SCRIPT_NORMALIZE_MAP = {
  a: 'а', e: 'е', o: 'о', p: 'р', c: 'с', x: 'х', y: 'у', k: 'к', h: 'н', m: 'м', t: 'т', b: 'в', i: 'і',
  A: 'А', E: 'Е', O: 'О', P: 'Р', C: 'С', X: 'Х', Y: 'У', K: 'К', H: 'Н', M: 'М', T: 'Т', B: 'В', I: 'І',
};
function normalizeScript(s) {
  return String(s || '').split('').map(ch => SCRIPT_NORMALIZE_MAP[ch] || ch).join('');
}

// Некоторые банки (замечено на выгрузках .xls БЦК / БЦК Бизнес) до сих пор
// используют старую до-юникодную казахскую кодировку для казахских букв: в
// тексте вместо Ә/ә стоит македонская/сербская Ј/ј, а вместо Қ/қ — Ќ/ќ.
// В Excel это выглядит почти неотличимо от нормального написания, но байты
// другие — из-за этого сопоставление контрагента по имени не находит уже
// существующую в 1С запись с правильным написанием и создаёт дубль с
// испорченным именем. Эти четыре символа не встречаются в казахском/русском
// тексте ни в каком легитимном случае, поэтому исправляем их без риска
// испортить что-то другое. Если найдутся другие письма с похожей порчей
// других казахских букв — карту можно будет расширить.
const KAZAKH_MOJIBAKE_MAP = { 'Ј': 'Ә', 'ј': 'ә', 'Ќ': 'Қ', 'ќ': 'қ' };
function fixKazakhMojibake(s) {
  const str = String(s || '');
  let hasMojibake = false;
  for (const ch of str) {
    if (KAZAKH_MOJIBAKE_MAP[ch]) { hasMojibake = true; break; }
  }
  if (!hasMojibake) return str;
  return str.split('').map(ch => KAZAKH_MOJIBAKE_MAP[ch] || ch).join('');
}

// Организационно-правовая форма (ТОО/ИП/АО и т.п.) и кавычки часто пишутся
// по-разному в банковской выписке и в справочнике 1С ("ТОО Ромашка" vs
// "Ромашка" vs "\"Ромашка\" ТОО") — из-за этого поиск по точному названию
// не находит контрагента, который на самом деле уже есть. normalizeCounterpartyName
// убирает эти различия, оставляя только "ядро" названия для сравнения.
const LEGAL_FORM_WORDS = ['тоо', 'жшс', 'ип', 'ао', 'оао', 'зао', 'чуп', 'гу', 'пк', 'кх', 'офз', 'фх', 'нао'];
function normalizeCounterpartyName(name) {
  let s = String(name || '').toLowerCase();
  s = s.replace(/["'«»„“”`]/g, ' ');
  for (const word of LEGAL_FORM_WORDS) {
    s = s.replace(new RegExp(`(^|\\s)${word}(\\s|$)`, 'g'), ' ');
  }
  return s.replace(/\s+/g, ' ').trim();
}

// Правило может проверять назначение платежа, имя контрагента ИЛИ КНП
// (код назначения платежа — стандартный классификатор, который банки РК
// присылают прямо в выписке). normalizeRuleField приводит любое сырое
// значение поля (из формы, bulk-импорта, RULES_TEXT) к одному из трёх
// допустимых значений.
function normalizeRuleField(raw) {
  if (raw === 'counterparty') return 'counterparty';
  if (raw === 'knp') return 'knp';
  return 'purpose';
}
function ruleFieldValue(op, field) {
  if (field === 'counterparty') return op.counterparty;
  if (field === 'knp') return op.knp;
  return op.purpose;
}

// Применяет ваши правила категоризации к операции: смотрит назначение
// платежа, имя контрагента и/или КНП, и если находит совпадение —
// возвращает категорию и счёт. Правила проверяются по порядку, первое
// совпадение побеждает. Если ни одно правило не подошло — возвращает
// базовую логику «покупатель/поставщик» по регламенту, а не пустоту:
// минимум статья по направлению платежа должна быть определена всегда.
function applyRules(op, rules) {
  for (const rule of rules) {
    const raw = ruleFieldValue(op, rule.field) || '';
    const haystack = normalizeScript(raw).toLowerCase();
    const needle = normalizeScript(String(rule.contains || '')).toLowerCase();
    if (needle && haystack.includes(needle)) {
      return {
        category: rule.category || '',
        account: rule.account || '',
        operationKind: rule.operationKind || '',
        debtType: rule.debtType || '',
        source: 'rule',
      };
    }
  }
  const base = baseDefault(op.amount);
  return { ...base, operationKind: '', debtType: '', source: 'default' };
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

  let grid;
  let workbook;
  try {
    workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    // header:1 → читаем как массив строк-массивов, без угадывания заголовков —
    // многие банковские выписки (например БЦК/Kaspi) кладут перед таблицей
    // несколько строк с реквизитами банка, поэтому настоящую шапку таблицы
    // нужно искать, а не считать первой строкой.
    grid = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });
  } catch (e) {
    return res.status(400).json({ error: 'Не получилось прочитать файл: ' + e.message });
  }

  const norm = (s) => String(s || '').toLowerCase().replace(/\s+/g, ' ').trim();

  // Ищем строку-шапку: она содержит одновременно что-то похожее на "дата"
  // и что-то похожее на дебет/кредит/сумму.
  let headerRowIndex = -1;
  let cols = {};
  for (let r = 0; r < Math.min(grid.length, 40); r++) {
    const cells = (grid[r] || []).map(norm);
    const findCol = (test) => cells.findIndex(test);
    const dateCol = findCol(c => c.includes('дата') && !c.includes('валют'));
    const debitCol = findCol(c => c.includes('дебет'));
    const creditCol = findCol(c => c.includes('кредит') && !c.includes('корреспондент'));
    const amountCol = findCol(c => c.includes('сумма') && !c.includes('конверт'));
    if (dateCol !== -1 && (debitCol !== -1 || creditCol !== -1 || amountCol !== -1)) {
      headerRowIndex = r;
      cols = {
        date: dateCol,
        debit: debitCol,
        credit: creditCol,
        amount: amountCol,
        purpose: findCol(c => c.includes('мақсат') || c.includes('назначен') || c.includes('комментарий')),
        counterparty: findCol(c => c.includes('корреспондент') && !c.includes('банк') && !c.includes('бик') && !c.includes('иик')),
        senderBin: findCol(c => (c.includes('бин') || c.includes('иин')) && c.includes('отправ')),
        receiverBin: findCol(c => (c.includes('бин') || c.includes('иин')) && c.includes('получ')),
        binGeneric: findCol(c => c.includes('бин') || c.includes('иин')),
        knp: findCol(c => c.includes('кнп') || c.includes('тмк')),
      };
      break;
    }
  }

  let operations = [];
  const parseAmount = (v) => {
    if (v === '' || v === undefined || v === null) return 0;
    const s = String(v).replace(/\s/g, '').replace(',', '.');
    const n = parseFloat(s);
    return isNaN(n) ? 0 : n;
  };

  if (headerRowIndex !== -1) {
    // Формат с "шапкой банка" + двуязычными заголовками (БЦК, Kaspi Business и похожие)
    for (let r = headerRowIndex + 1; r < grid.length; r++) {
      const row = grid[r];
      if (!row || row.every(c => String(c).trim() === '')) continue;
      const debit = cols.debit !== -1 ? parseAmount(row[cols.debit]) : 0;
      const credit = cols.credit !== -1 ? parseAmount(row[cols.credit]) : 0;
      const single = cols.amount !== -1 ? parseAmount(row[cols.amount]) : 0;
      const amount = credit || (debit ? -debit : single);
      const dateRaw = cols.date !== -1 ? String(row[cols.date] || '') : '';
      if (!dateRaw && !amount) continue;
      const bin = credit
        ? (cols.senderBin !== -1 ? row[cols.senderBin] : row[cols.binGeneric])
        : (cols.receiverBin !== -1 ? row[cols.receiverBin] : row[cols.binGeneric]);
      operations.push({
        id: crypto.randomUUID(),
        date: dateRaw.split(' ')[0] || dateRaw,
        counterparty: fixKazakhMojibake(cols.counterparty !== -1 ? row[cols.counterparty] : ''),
        bin: String(bin || '').trim(),
        purpose: fixKazakhMojibake(cols.purpose !== -1 ? row[cols.purpose] : ''),
        knp: cols.knp !== -1 ? String(row[cols.knp] || '') : '', // код назначения платежа — берём прямо из выписки банка
        amount,
        suggestedCategory: '',
        status: 'review',
        sourceFile: req.file.originalname,
        rowIndex: r,
      });
    }
  } else {
    // Запасной вариант: простые выписки, где заголовки — в самой первой строке
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const objRows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    const pick = (row, names) => {
      for (const n of names) {
        const key = Object.keys(row).find(k => k.trim().toLowerCase() === n);
        if (key) return row[key];
      }
      return '';
    };
    operations = objRows.map((row, i) => {
      const dateRaw = pick(row, ['дата', 'дата операции', 'дата документа']);
      const amountIn = parseFloat(pick(row, ['приход', 'сумма прихода']) || 0) || 0;
      const amountOut = parseFloat(pick(row, ['расход', 'сумма расхода']) || 0) || 0;
      const amountSingle = parseFloat(pick(row, ['сумма']) || 0) || 0;
      const amount = amountIn || (amountOut ? -amountOut : amountSingle);
      return {
        id: crypto.randomUUID(),
        date: String(dateRaw || ''),
        counterparty: fixKazakhMojibake(pick(row, ['контрагент', 'наименование контрагента', 'плательщик/получатель'])),
        bin: String(pick(row, ['бин', 'иин', 'бин/иин']) || ''),
        purpose: fixKazakhMojibake(pick(row, ['назначение платежа', 'назначение', 'комментарий'])),
        knp: String(pick(row, ['кнп', 'код назначения платежа', 'тмк']) || ''),
        amount,
        suggestedCategory: '',
        status: 'review',
        sourceFile: req.file.originalname,
        rowIndex: i,
      };
    }).filter(op => op.date || op.amount);
  }

  const all = await getStore('operations', []);

  // Не добавляем операцию повторно, если такая же (по дате, сумме, БИН и
  // назначению) уже есть в списке — иначе повторная загрузка того же файла
  // задваивает список.
  const dedupKey = (o) => `${o.date}|${o.amount}|${o.bin}|${o.purpose}`;
  const existingKeys = new Set(all.map(dedupKey));
  const newOnly = operations.filter(op => !existingKeys.has(dedupKey(op)));
  const skippedDuplicates = operations.length - newOnly.length;

  const updated = [...all, ...newOnly];
  await setStore('operations', updated);
  await addHistory(
    `Загружена выписка ${req.file.originalname} · ${newOnly.length} новых операций` +
      (skippedDuplicates ? ` · ${skippedDuplicates} пропущено как дубли` : ''),
    '—'
  );

  // Сразу пытаемся сопоставить контрагентов по БИН/ИИН со справочником 1С,
  // и проверить — нет ли такой операции уже среди документов в самой 1С
  // (чтобы не предлагать создавать то, что уже разнесено вручную или
  // другим способом). Всё это — только если подключение уже настроено.
  const settingsForMatch = await getStore('settings', {});
  if (settingsForMatch.baseUrl && settingsForMatch.login) {
    for (const op of newOnly) {
      const target = updated.find(o => o.id === op.id);

      try {
        const alreadyExists = await checkExistingInOnec(op, settingsForMatch);
        if (alreadyExists) {
          target.status = 'already_in_1c';
          continue; // не тратим время на сопоставление контрагента — операция и так пропускается
        }
      } catch (e) {
        // Если сверка не удалась — не блокируем, просто идём дальше как обычно.
      }

      try {
        const found = await findCounterpartyByBin(op.bin, settingsForMatch, op.counterparty);
        if (found) {
          target.counterpartyKey = found.Ref_Key;
          target.counterpartyMatchedName = found.Description || '';

          // Смотрим, как этот контрагент категоризировался раньше в 1С —
          // если найдём паттерн, используем ту же статью ДДС автоматически.
          try {
            const historical = await findHistoricalCategory(found.Ref_Key, op.amount, settingsForMatch, op.purpose);
            if (historical) {
              target.historicalCategory = historical.categoryName || '';
              target.historicalCategoryKey = historical.categoryKey || '';
              target.historicalOperationKind = historical.operationKind || '';
              target.historicalDebtType = historical.debtType || '';
            }
          } catch (e) {
            // Не критично — просто не будет автопредложения по истории для этой операции.
          }

          // Ищем договор контрагента — если он есть в справочнике и найден
          // однозначно, привяжем его к документу автоматически.
          try {
            const contract = await resolveSupplierContract(found.Ref_Key, op.purpose, settingsForMatch);
            target.contractStatus = contract.status;
            if (contract.status === 'matched') {
              target.contractKey = contract.key;
              target.contractName = contract.name;
            } else if (contract.status === 'ambiguous') {
              target.contractOptions = contract.options;
            }
          } catch (e) {
            // Не критично — просто без привязки к договору.
          }
        } else {
          target.status = 'new_counterparty';
        }
      } catch (e) {
        if (e.ambiguousCounterparty) {
          target.status = 'ambiguous_counterparty';
          target.counterpartyOptions = e.ambiguousCounterparty.map(o => ({ key: o.Ref_Key, name: o.Description }));
        }
        // Иначе — проверка не удалась (например, неверный пароль) — не
        // блокируем загрузку, просто оставляем операцию как есть, на проверке.
      }
    }
    await setStore('operations', updated);
  }

  const reconciliation = {
    total: newOnly.length,
    alreadyIn1c: newOnly.filter(o => updated.find(u => u.id === o.id)?.status === 'already_in_1c').length,
    newCounterparty: newOnly.filter(o => updated.find(u => u.id === o.id)?.status === 'new_counterparty').length,
    connected: !!(settingsForMatch.baseUrl && settingsForMatch.login),
  };
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
    const ruleMatch = applyRules(op, rules);
    const category = op.manualCategory || op.historicalCategory || ruleMatch.category;
    const account = op.manualAccount || op.historicalAccount || ruleMatch.account;
    const operationKind = op.manualOperationKind || op.historicalOperationKind || ruleMatch.operationKind;
    const debtType = op.manualDebtType || op.historicalDebtType || ruleMatch.debtType;
    const isDefaultOnly = !op.manualCategory && !op.historicalCategory && ruleMatch.source === 'default';
    return {
      ...op,
      suggestedCategory: category,
      suggestedAccount: account,
      suggestedOperationKind: operationKind,
      suggestedDebtType: debtType,
      isDefaultOnly, // категория определена только базовой логикой, стоит перепроверить глазами
      needsAttention: !category,
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
app.post('/api/operations/:id/confirm', requireAuth, async (req, res) => {
  const all = await getStore('operations', []);
  const op = all.find(o => o.id === req.params.id);
  if (!op) return res.status(404).json({ error: 'Операция не найдена' });

  const settings = await getStore('settings', {});
  if (!settings.baseUrl) {
    return res.status(400).json({ error: 'Сначала укажите адрес подключения к 1С в разделе «Настройки»' });
  }

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
      await setStore('operations', all);
      await addHistory(`Перепроверка перед подтверждением: в 1С уже есть такой документ — черновик не создаём (${op.counterparty || 'операция'})`, '—');
      return res.status(409).json({ error: 'В 1С уже есть документ с такой же датой и суммой — черновик не создан, чтобы не задвоить. Операция помечена как «Уже в 1С».', op });
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
    await setStore('operations', all);
    return res.status(400).json({
      error: 'Контрагент не сопоставлен с 1С — сначала обработайте это в карточке операции (создайте контрагента или выберите из найденных вариантов).',
      op,
    });
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
      const contract = await resolveSupplierContract(op.counterpartyKey, op.purpose, settings);
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
  await setStore('operations', all);

  // Статья ДДС обязательна — без неё документ не создаём. Приоритет:
  // ручная правка > история 1С (уже обновлённая на шаге перепроверки выше)
  // > правила > базовая логика (которая всегда что-то определяет, так что
  // пусто может быть, только если что-то совсем не так с данными).
  const rules = await getStore('rules', []);
  const ruleMatch = applyRules(op, rules);
  const category = op.manualCategory || op.historicalCategory || ruleMatch.category;
  const account = op.manualAccount || op.historicalAccount || ruleMatch.account;
  const operationKind = op.manualOperationKind || op.historicalOperationKind || ruleMatch.operationKind;
  const debtType = op.manualDebtType || op.historicalDebtType || ruleMatch.debtType;
  if (!category) {
    return res.status(400).json({ error: 'Не определена статья ДДС — поправьте вручную в карточке операции, прежде чем подтверждать.' });
  }

  try {
    const result = await createDraftInOnec(op, settings, category, account, operationKind, debtType);
    op.status = 'draft_created';
    op.onecDocNumber = result.docNumber || null;
    await setStore('operations', all);
    // Если какие-то поля (счёт учёта, подотчётник и т.п.) не прижились в вашей
    // конфигурации 1С — черновик всё равно создаётся, но мы честно пишем в
    // историю, что именно не заполнилось, чтобы можно было доглядеть глазами
    // и, при желании, прислать нам точный текст ошибки для донастройки.
    const droppedNote = result.droppedFields && result.droppedFields.length
      ? ` · 1С не приняла поля: ${result.droppedFields.join(', ')} — проверьте их вручную в документе`
      : '';
    await addHistory(`Создан черновик в 1С: ${op.counterparty || 'операция'} · ${op.amount} ₸ · статья: ${category}${droppedNote}`, result.docNumber || '—');
    res.json({ ok: true, op, droppedFields: result.droppedFields || [] });
  } catch (e) {
    await addHistory(`Ошибка при создании черновика: ${e.message}`, '—');
    res.status(502).json({ error: 'Не удалось создать документ в 1С: ' + e.message });
  }
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
app.post('/api/operations/:id/create-counterparty', requireAuth, async (req, res) => {
  const all = await getStore('operations', []);
  const op = all.find(o => o.id === req.params.id);
  if (!op) return res.status(404).json({ error: 'Операция не найдена' });
  if (!op.bin) return res.status(400).json({ error: 'У операции не указан БИН/ИИН — сопоставьте контрагента вручную' });

  const settings = await getStore('settings', {});
  if (!settings.baseUrl) {
    return res.status(400).json({ error: 'Сначала укажите адрес подключения к 1С в разделе «Настройки»' });
  }

  try {
    // На всякий случай проверяем ещё раз прямо перед тем, как отвечать —
    // вдруг контрагент уже появился в 1С (например, кто-то создал его
    // вручную, или его нашли через ручной поиск /search-counterparty).
    const existing = await findCounterpartyByBin(op.bin, settings, op.counterparty);
    if (existing) {
      op.counterpartyKey = existing.Ref_Key;
      op.counterpartyMatchedName = existing.Description;
      op.status = 'review'; // теперь контрагент есть — можно подтверждать документ как обычно
      await setStore('operations', all);
      await addHistory(`Контрагент найден в 1С при повторной проверке: ${existing.Description} (БИН ${op.bin})`, '—');
      return res.json({ ok: true, op });
    }

    // Автоматическое создание контрагента ЗАПРЕЩЕНО по умолчанию — 1С
    // остаётся единственным источником истины для справочника контрагентов.
    // Бухгалтер должен либо найти существующего вручную (/search-counterparty),
    // либо создать контрагента сам в 1С и затем нажать "Пересверить с 1С".
    await addHistory(`Контрагент "${op.counterparty || op.bin}" (БИН ${op.bin}) не найден в 1С — создайте его вручную в справочнике 1С, затем нажмите «Пересверить с 1С»`, '—');
    return res.status(409).json({
      error: 'Автоматическое создание контрагента отключено. Убедитесь через поиск, что его точно нет в 1С, затем создайте контрагента вручную в справочнике 1С и нажмите «Пересверить с 1С».',
      needsManualCreation: true,
      op,
    });
  } catch (e) {
    await addHistory(`Ошибка при проверке контрагента: ${e.message}`, '—');
    res.status(502).json({ error: 'Не удалось проверить контрагента в 1С: ' + e.message });
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

// Ищет контрагента в справочнике 1С по БИН/ИИН через OData.
// Возвращает объект контрагента (с Ref_Key) или null, если не найден.
// Проверяет, нет ли уже среди документов 1С (Платёжное поручение
// входящее/исходящее) операции с такой же датой и суммой — чтобы не
// предлагать создавать то, что уже разнесено (вручную или иначе).
async function checkExistingInOnec(op, settings) {
  const base = settings.baseUrl.replace(/\/+$/, '');
  const auth = Buffer.from(`${settings.login}:${settings.password}`).toString('base64');
  const docType = op.amount >= 0 ? 'Document_ПлатежноеПоручениеВходящее' : 'Document_ПлатежноеПоручениеИсходящее';
  const amount = Math.abs(op.amount);
  // Дата у нас хранится как ДД.ММ.ГГГГ — переводим в ГГГГ-ММ-ДД для фильтра OData
  const parts = String(op.date).split('.');
  if (parts.length !== 3) return false;
  const isoDate = `${parts[2]}-${parts[1]}-${parts[0]}`;
  const filter = encodeURIComponent(
    `Date ge datetime'${isoDate}T00:00:00' and Date le datetime'${isoDate}T23:59:59' and СуммаДокумента eq ${amount}`
  );
  const url = `${base}/${docType}?$format=json&$filter=${filter}&$select=Ref_Key`;

  const response = await fetch(url, { headers: { Authorization: `Basic ${auth}`, Accept: 'application/json' } });
  if (!response.ok) return false; // если сверить не удалось — не блокируем, просто считаем, что не нашли
  const data = await response.json().catch(() => null);
  if (!data) return false;
  return (data.value || []).length > 0;
}

// Пытается найти номер договора прямо в тексте назначения платежа —
// банки/бухгалтеры обычно пишут "по дог. №...", "договор №...", "дог. N..."
function extractContractNumber(purpose) {
  if (!purpose) return null;
  const match = purpose.match(/до?г(?:овор)?\.?\s*(?:№|N|no\.?)?\s*([\w\-\/]{2,})/i);
  return match ? match[1] : null;
}

// Ищет договор(ы) контрагента в 1С. Если найден ровно один — используем
// его автоматически. Если несколько — не гадаем, а помечаем операцию как
// требующую ручного выбора договора. Если в тексте назначения нашёлся
// номер договора — в первую очередь пытаемся сопоставить именно по нему.
const CONTRACT_CATALOG_CANDIDATES = ['Catalog_ДоговорыКонтрагентов', 'Catalog_Договоры'];
const NO_CONTRACT_NAME = 'Без договора';

async function findContractForCounterparty(counterpartyKey, purposeText, settings) {
  if (!counterpartyKey) return { status: 'none' };
  const base = settings.baseUrl.replace(/\/+$/, '');
  const auth = Buffer.from(`${settings.login}:${settings.password}`).toString('base64');

  for (const catalog of CONTRACT_CATALOG_CANDIDATES) {
    const filter = encodeURIComponent(`Владелец_Key eq guid'${counterpartyKey}' or Контрагент_Key eq guid'${counterpartyKey}'`);
    const url = `${base}/${catalog}?$format=json&$filter=${filter}&$select=Ref_Key,Description,Number&$top=20`;
    const response = await fetch(url, { headers: { Authorization: `Basic ${auth}`, Accept: 'application/json' } });
    if (!response.ok) continue; // пробуем следующее возможное название справочника
    const data = await response.json().catch(() => null);
    const list = (data && data.value) || [];
    if (list.length === 0) continue;

    if (list.length === 1) {
      return { status: 'matched', key: list[0].Ref_Key, name: list[0].Description || list[0].Number };
    }

    // Несколько договоров — пробуем сопоставить по номеру, найденному в назначении платежа
    const hint = extractContractNumber(purposeText);
    if (hint) {
      const byNumber = list.find(d =>
        (d.Number && d.Number.includes(hint)) || (d.Description && d.Description.includes(hint))
      );
      if (byNumber) {
        return { status: 'matched', key: byNumber.Ref_Key, name: byNumber.Description || byNumber.Number };
      }
    }
    return { status: 'ambiguous', options: list.map(d => ({ key: d.Ref_Key, name: d.Description || d.Number })) };
  }
  return { status: 'none' }; // у конфигурации либо нет отдельного справочника договоров, либо название другое
}

// Ищет договор конкретного контрагента с названием "Без договора" —
// отдельным точным запросом (не полагаемся только на findContractForCounterparty,
// т.к. если у контрагента больше одного договора, "Без договора" может
// потеряться среди неоднозначных вариантов).
async function findNoContractRecord(counterpartyKey, settings) {
  const base = settings.baseUrl.replace(/\/+$/, '');
  const auth = Buffer.from(`${settings.login}:${settings.password}`).toString('base64');
  for (const catalog of CONTRACT_CATALOG_CANDIDATES) {
    const filter = encodeURIComponent(
      `(Владелец_Key eq guid'${counterpartyKey}' or Контрагент_Key eq guid'${counterpartyKey}') and Description eq '${NO_CONTRACT_NAME}'`
    );
    const url = `${base}/${catalog}?$format=json&$filter=${filter}&$select=Ref_Key,Description&$top=1`;
    const response = await fetch(url, { headers: { Authorization: `Basic ${auth}`, Accept: 'application/json' } });
    if (!response.ok) continue;
    const data = await response.json().catch(() => null);
    const found = data && data.value && data.value[0];
    if (found) return { catalog, key: found.Ref_Key, name: found.Description };
  }
  return null;
}

// Полная логика подбора договора для «Оплата поставщику»:
//   1) если у контрагента найден ровно один договор (или однозначно по
//      номеру в назначении платежа) — используем его;
//   2) если договоров несколько и не определить — статус "ambiguous",
//      выбор за бухгалтером;
//   3) если договоров нет вообще — отдельно ищем именно "Без договора";
//   4) если и его нет — не гадаем и не создаём автоматически, а сообщаем,
//      что нужно создать "Без договора" (кнопкой в кабинете либо вручную в 1С).
async function resolveSupplierContract(counterpartyKey, purposeText, settings) {
  if (!counterpartyKey) return { status: 'none' };
  const primary = await findContractForCounterparty(counterpartyKey, purposeText, settings);
  if (primary.status === 'matched' || primary.status === 'ambiguous') return primary;

  const noContract = await findNoContractRecord(counterpartyKey, settings);
  if (noContract) {
    return { status: 'matched', key: noContract.key, name: noContract.name };
  }
  return { status: 'need_create_no_contract' };
}

// Создаёт в справочнике договоров запись "Без договора" для контрагента.
// Вызывается ТОЛЬКО по явному действию бухгалтера (кнопка в кабинете) —
// никогда автоматически при создании черновика документа.
async function createNoContractInOnec(counterpartyKey, settings) {
  const base = settings.baseUrl.replace(/\/+$/, '');
  const auth = Buffer.from(`${settings.login}:${settings.password}`).toString('base64');

  for (const catalog of CONTRACT_CATALOG_CANDIDATES) {
    const payload = {
      Description: NO_CONTRACT_NAME,
      Владелец_Key: counterpartyKey,
      Контрагент_Key: counterpartyKey,
    };
    const response = await fetch(`${base}/${catalog}?$format=json`, {
      method: 'POST',
      headers: { Authorization: `Basic ${auth}`, 'Content-Type': 'application/json', Accept: 'application/json' },
      body: JSON.stringify(payload),
    });
    if (response.ok) return response.json().catch(() => ({}));
    // Пробуем следующий вариант названия справочника, если этот не подошёл
    // (например, поля Владелец_Key/Контрагент_Key в нём называются иначе).
  }
  throw new Error('Не удалось создать договор "Без договора" — проверьте название справочника договоров в вашей конфигурации 1С');
}

// Ищет статью ДДС в справочнике 1С по точному названию (например,
// «Аренда» или «Расчёты с поставщиками и подрядчиками») и возвращает её
// GUID. Нужно, когда категория пришла из текстового правила или ручной
// правки, а не из уже готового документа в 1С.
async function findCategoryKeyByName(name, settings) {
  const base = settings.baseUrl.replace(/\/+$/, '');
  const auth = Buffer.from(`${settings.login}:${settings.password}`).toString('base64');
  const candidates = ['Catalog_СтатьиДвиженияДенежныхСредств', 'Catalog_СтатьиДДС'];
  const filter = encodeURIComponent(`Description eq '${String(name).replace(/'/g, "''")}'`);
  for (const catalog of candidates) {
    const url = `${base}/${catalog}?$format=json&$filter=${filter}&$select=Ref_Key&$top=1`;
    const response = await fetch(url, { headers: { Authorization: `Basic ${auth}`, Accept: 'application/json' } });
    if (!response.ok) continue;
    const data = await response.json().catch(() => null);
    const item = data && data.value && data.value[0];
    if (item) return item.Ref_Key;
  }
  return null;
}

// "Счёт учёта (БУ)" в табличной части документа — это ссылка на план счетов
// (Chart of Accounts), а не просто текст "1251": нужно найти GUID счёта по
// его коду. Название самого объекта плана счетов в OData отличается между
// конфигурациями 1С — пробуем частые варианты и запоминаем рабочий, как и с
// полем БИН/ИИН. Счёт может храниться у нас как пара "1251/3510"
// (счёт расчётов/счёт авансов) — для табличной части берём основную часть
// до "/".
const CHART_OF_ACCOUNTS_CANDIDATES = ['ChartOfAccounts_Хозрасчетный', 'ChartOfAccounts_Хозрасчетный2', 'ChartOfAccounts_Управленческий'];
let workingChartOfAccounts = null;

async function findAccountKeyByCode(code, settings) {
  if (!code) return null;
  const primaryCode = String(code).split('/')[0].trim();
  if (!primaryCode) return null;
  const base = settings.baseUrl.replace(/\/+$/, '');
  const auth = Buffer.from(`${settings.login}:${settings.password}`).toString('base64');

  async function tryChart(chart) {
    const filter = encodeURIComponent(`Code eq '${primaryCode}'`);
    const url = `${base}/${chart}?$format=json&$filter=${filter}&$select=Ref_Key&$top=1`;
    const response = await fetch(url, { headers: { Authorization: `Basic ${auth}`, Accept: 'application/json' } });
    if (!response.ok) return { exists: false };
    const data = await response.json().catch(() => null);
    if (!data) return { exists: false };
    return { exists: true, found: (data.value || [])[0] || null };
  }

  if (workingChartOfAccounts) {
    const r = await tryChart(workingChartOfAccounts);
    if (r.exists) return r.found ? r.found.Ref_Key : null;
    workingChartOfAccounts = null;
  }
  for (const chart of CHART_OF_ACCOUNTS_CANDIDATES) {
    const r = await tryChart(chart);
    if (r.exists) {
      workingChartOfAccounts = chart;
      return r.found ? r.found.Ref_Key : null;
    }
  }
  return null; // ни один вариант плана счетов не подошёл — не страшно, просто не заполним это поле
}

async function findHistoricalCategory(counterpartyKey, amount, settings, currentPurpose) {
  if (!counterpartyKey) return null;
  const base = settings.baseUrl.replace(/\/+$/, '');
  const auth = Buffer.from(`${settings.login}:${settings.password}`).toString('base64');
  const docType = amount >= 0 ? 'Document_ПлатежноеПоручениеВходящее' : 'Document_ПлатежноеПоручениеИсходящее';

  // Статья ДДС в реальности лежит ВНУТРИ табличной части "РасшифровкаПлатежа",
  // а не на уровне самого документа — поэтому обязательно разворачиваем её
  // через $expand, иначе эти поля просто не приедут в ответе.
  //
  // Смотрим НЕСКОЛЬКО последних документов (не только самый свежий), потому
  // что у одного контрагента может быть несколько РАЗНЫХ типов операций
  // (например, разные виды переводов) — берём не просто "любой заполненный",
  // а тот, чьё назначение платежа больше всего похоже на текущую операцию.
  const filter = encodeURIComponent(`Контрагент_Key eq guid'${counterpartyKey}'`);
  const url = `${base}/${docType}?$format=json&$filter=${filter}&$orderby=Date desc&$top=15&$expand=РасшифровкаПлатежа`;

  const response = await fetch(url, { headers: { Authorization: `Basic ${auth}`, Accept: 'application/json' } });
  if (!response.ok) return null;
  const data = await response.json().catch(() => null);
  const docs = (data && data.value) || [];

  // Похожесть текста назначения платежа: доля общих слов (без учёта
  // регистра), длиной от 3 букв — так короткие союзы/предлоги не мешают.
  function wordSet(text) {
    return new Set(String(text || '').toLowerCase().match(/[а-яёіңғүұқөһa-z]{3,}/g) || []);
  }
  const currentWords = wordSet(currentPurpose);
  function similarity(otherText) {
    if (!currentWords.size) return 0;
    const otherWords = wordSet(otherText);
    if (!otherWords.size) return 0;
    let common = 0;
    for (const w of currentWords) if (otherWords.has(w)) common++;
    return common / currentWords.size;
  }

  // Собираем всех кандидатов, у кого есть хотя бы статья ДДС, и сортируем
  // по похожести назначения платежа (лучшее совпадение — первым). При
  // равной похожести побеждает наличие ещё и заполненного вида операции.
  const candidates = [];
  for (const candidate of docs) {
    const candidateRow = candidate.РасшифровкаПлатежа && candidate.РасшифровкаПлатежа[0];
    const candidateCategoryKey = candidateRow && candidateRow.СтатьяДвиженияДенежныхСредств_Key;
    if (!candidateCategoryKey) continue;
    candidates.push({
      doc: candidate,
      row: candidateRow,
      categoryKey: candidateCategoryKey,
      score: similarity(candidate.НазначениеПлатежа || candidate.Комментарий) + (candidate.ВидОперации ? 0.01 : 0),
    });
  }
  candidates.sort((a, b) => b.score - a.score);
  const best = candidates[0];
  if (!best) return null;
  const { doc, row, categoryKey } = best;
  if (!doc || !categoryKey) return null;

  const debtType = (row && (row.ВидЗадолженности || row.ВидЗадолженности_Key)) || '';

  // Статья хранится как GUID — подтягиваем её человекочитаемое название
  const categoryCatalogs = ['Catalog_СтатьиДвиженияДенежныхСредств', 'Catalog_СтатьиДДС'];
  for (const catalog of categoryCatalogs) {
    const catUrl = `${base}/${catalog}(guid'${categoryKey}')?$format=json&$select=Description`;
    const catResp = await fetch(catUrl, { headers: { Authorization: `Basic ${auth}`, Accept: 'application/json' } });
    if (catResp.ok) {
      const catData = await catResp.json().catch(() => null);
      if (catData && catData.Description) {
        return { categoryKey, categoryName: catData.Description, operationKind: doc.ВидОперации || '', debtType };
      }
    }
  }
  return { categoryKey, categoryName: '', operationKind: doc.ВидОперации || '' };
}

// Название поля с БИН/ИИН отличается между конфигурациями 1С. Вместо
// одного запроса со всеми вариантами сразу (он падает целиком, если хотя
// бы одного поля нет — «Сегмент пути БИН не найден»), пробуем варианты
// по очереди и запоминаем то, которое сработало.
const BIN_FIELD_CANDIDATES = ['ИИН', 'БИН', 'ИНН', 'БИН_ИИН', 'ИИН_БИН', 'ИННЮЛ', 'ИННФЛ', 'КодПоОКПО', 'РегистрационныйНомер'];
let workingBinField = null; // определяется один раз за время работы сервера

async function findCounterpartyByBin(bin, settings, name) {
  const base = settings.baseUrl.replace(/\/+$/, '');
  const auth = Buffer.from(`${settings.login}:${settings.password}`).toString('base64');

  async function tryField(field) {
    const filter = encodeURIComponent(`${field} eq '${bin}'`);
    const url = `${base}/Catalog_Контрагенты?$format=json&$filter=${filter}&$top=1`;
    const response = await fetch(url, { headers: { Authorization: `Basic ${auth}`, Accept: 'application/json' } });
    if (!response.ok) return { fieldExists: false };
    const data = await response.json().catch(() => null);
    if (!data) return { fieldExists: false };
    return { fieldExists: true, found: (data.value || [])[0] || null };
  }

  // Запасной поиск по названию — на случай, если БИН в выписке не
  // считался или не совпал по формату, а контрагент в справочнике на
  // самом деле есть. ВАЖНО: раньше здесь брался первый попавшийся
  // результат ($top=1) — если в справочнике несколько похожих записей,
  // легко было тихо выбрать не того контрагента. Теперь: сначала пробуем
  // точное совпадение по названию; если его нет — смотрим частичные
  // совпадения, и если их больше одного, ЧЕСТНО сообщаем о неоднозначности,
  // а не гадаем.
  async function fetchByFilter(filter, top) {
    const url = `${base}/Catalog_Контрагенты?$format=json&$filter=${encodeURIComponent(filter)}&$top=${top}`;
    const response = await fetch(url, { headers: { Authorization: `Basic ${auth}`, Accept: 'application/json' } });
    if (!response.ok) return null;
    const data = await response.json().catch(() => null);
    return (data && data.value) || [];
  }

  async function tryByName() {
    if (!name || name.trim().length < 3) return null;
    const cleanName = name.trim().replace(/'/g, "''");

    // 1. Точное совпадение по исходному названию — самый надёжный вариант
    const exactList = await fetchByFilter(`Description eq '${cleanName}'`, 2);
    if (exactList && exactList.length === 1) return { match: exactList[0] };
    if (exactList && exactList.length > 1) return { ambiguous: true, options: exactList };

    // 2. Частичное совпадение по исходному названию
    const partialList = await fetchByFilter(`substringof('${cleanName}', Description)`, 10);
    let candidates = partialList || [];

    // 3. Если ничего не нашли — та же попытка, но по "ядру" названия без
    // организационно-правовой формы (ТОО/ИП/АО/…) и кавычек: банковская
    // выписка и справочник 1С часто пишут форму по-разному.
    if (candidates.length === 0) {
      const normalizedTarget = normalizeCounterpartyName(name);
      if (normalizedTarget && normalizedTarget.length >= 3) {
        const wide = await fetchByFilter(`substringof('${normalizedTarget.replace(/'/g, "''")}', Description)`, 15);
        candidates = (wide || []).filter(o => normalizeCounterpartyName(o.Description) === normalizedTarget);
        // Если точных совпадений "ядра" не набралось — берём как есть, лучше
        // честно показать неоднозначность, чем не найти существующего контрагента.
        if (candidates.length === 0) candidates = wide || [];
      }
    }

    if (candidates.length === 0) return null;
    if (candidates.length === 1) return { match: candidates[0] };
    return { ambiguous: true, options: candidates };
  }

  async function resolveByName() {
    const r = await tryByName();
    if (!r) return null;
    if (r.ambiguous) {
      const err = new Error(
        `По имени "${name}" в справочнике нашлось ${r.options.length} разных контрагентов ` +
        `(${r.options.map(o => o.Description).join(', ')}) — не могу понять, какой из них правильный. ` +
        `Сопоставьте контрагента вручную для этой операции.`
      );
      err.ambiguousCounterparty = r.options;
      throw err;
    }
    return r.match;
  }

  // Если рабочее поле уже определено — сразу используем его.
  if (workingBinField && bin) {
    const r = await tryField(workingBinField);
    if (r.fieldExists) {
      if (r.found) return r.found;
      return await resolveByName(); // БИН не совпал — пробуем по названию
    }
    workingBinField = null; // поле перестало работать — определим заново
  }

  if (!bin) return await resolveByName(); // БИН в выписке пуст — сразу по названию

  const errors = [];
  for (const field of BIN_FIELD_CANDIDATES) {
    const r = await tryField(field);
    if (r.fieldExists) {
      workingBinField = field; // запомнили — дальше будет быстро
      if (r.found) return r.found;
      return await resolveByName(); // поле рабочее, но по БИН не нашли — пробуем по названию
    }
    errors.push(field);
  }
  // Ни одно поле с БИН не сработало вообще — прежде чем сдаться, всё
  // равно пробуем по названию.
  return await resolveByName();
}

// Свободный ручной поиск контрагента по части названия — используется, когда
// бухгалтер сам проверяет "а нет ли он уже в 1С", прежде чем соглашаться на
// создание нового. В отличие от tryByName() внутри findCounterpartyByBin (та
// возвращает ошибку при неоднозначности), здесь мы, наоборот, ХОТИМ увидеть
// все подходящие варианты списком, чтобы человек выбрал нужный сам.
async function searchCounterpartyByText(query, settings) {
  const base = settings.baseUrl.replace(/\/+$/, '');
  const auth = Buffer.from(`${settings.login}:${settings.password}`).toString('base64');
  const cleanQuery = query.replace(/'/g, "''");
  const filter = encodeURIComponent(`substringof('${cleanQuery}', Description)`);
  const url = `${base}/Catalog_Контрагенты?$format=json&$filter=${filter}&$select=Ref_Key,Description&$top=20`;
  const response = await fetch(url, { headers: { Authorization: `Basic ${auth}`, Accept: 'application/json' } });
  if (!response.ok) {
    const text = await response.text().catch(() => '');
    throw new Error(`1С ответила ${response.status}: ${text.slice(0, 200)}`);
  }
  const data = await response.json().catch(() => null);
  return (data && data.value) || [];
}

// =====================================================================
// Вызов через OData стандартного интерфейса 1С:Фреш — на основе полей,
// которые реально видны в вашей базе (Date, Number, Организация_Key,
// СчетОрганизации_Key, Контрагент_Key, ВидОперации,
// СтатьяДвиженияДенежныхСредств_Key, СуммаДокумента, НазначениеПлатежа).
//
// settings.baseUrl должен быть вида:
//   https://1cfresh.kz/a/ea170/264256/odata/standard.odata/
// (без имени документа на конце — его добавляем здесь).
//
// ВАЖНО: Организация_Key и СчетОрганизации_Key — это GUID-идентификаторы
// вашей организации и расчётного счёта в справочниках 1С. Их нужно один
// раз узнать (через тот же OData: .../Catalog_Организации?$format=json)
// и вписать в настройки — без них документ не создать.
// =====================================================================
async function createDraftInOnec(op, settings, resolvedCategory, resolvedAccount, resolvedOperationKind, resolvedDebtType) {
  const base = settings.baseUrl.replace(/\/+$/, ''); // убираем лишний / на конце
  const docType = op.amount >= 0 ? 'Document_ПлатежноеПоручениеВходящее' : 'Document_ПлатежноеПоручениеИсходящее';
  const endpoint = `${base}/${docType}`;
  const auth = Buffer.from(`${settings.login}:${settings.password}`).toString('base64');

  const payload = {
    Date: toIsoDate(op.date),
    Posted: false, // черновик — не проводится автоматически
    Организация_Key: settings.orgKey || '',
    СчетОрганизации_Key: settings.accountKey || '',
    Контрагент_Key: op.counterpartyKey || '', // если контрагент не найден — не отправляем, см. ниже
    СуммаДокумента: Math.abs(op.amount),
    НазначениеПлатежа: op.purpose,
    Комментарий: op.purpose, // должно совпадать с назначением платежа, как в вашей 1С
  };

  // Статья ДДС: если вы её НЕ правили руками и в истории 1С уже есть
  // готовый GUID — используем его напрямую. Иначе (ручная правка, правило,
  // или текст без готового GUID) ищем статью по названию в справочнике.
  let categoryKey = null;
  if (!op.manualCategory && op.historicalCategoryKey) {
    categoryKey = op.historicalCategoryKey;
  } else if (resolvedCategory) {
    categoryKey = await findCategoryKeyByName(resolvedCategory, settings);
    // Если не нашли даже по имени — не страшно: текст статьи уже есть в
    // комментарии и в интерфейсе, вы сможете проставить её вручную в 1С.
  }
  if (categoryKey) payload.СтатьяДвиженияДенежныхСредств_Key = categoryKey;

  // "Вид операции" — строгий список в 1С. Приоритет:
  //   1. История (скопировано с уже существующего документа — 100% верно)
  //   2. Подтверждённое точное значение (проверено напрямую через вашу базу)
  //   3. Обоснованная попытка по тому же принципу именования — если 1С её
  //      отклонит, мы автоматически повторим запрос без этого поля (см. ниже).
  let operationKindLiteral = null;
  if (op.historicalOperationKind) {
    operationKindLiteral = op.historicalOperationKind;
  } else if (resolvedOperationKind) {
    operationKindLiteral = guessOperationKindLiteral(resolvedOperationKind);
  } else if (op.amount < 0 && (resolvedCategory === 'Расчёты с поставщиками и подрядчиками' || resolvedCategory === 'Расчеты с поставщиками и подрядчиками')) {
    // Базовый сценарий "Оплата поставщику" (нет ни истории, ни более
    // специфичного правила вроде аренды/зарплаты/налогов) — ставим вид
    // операции по умолчанию. Если точное написание в вашей 1С отличается,
    // это поле безопасно "отвалится" через OPTIONAL_FIELDS ниже.
    operationKindLiteral = guessOperationKindLiteral('Оплата поставщику');
  }
  if (operationKindLiteral) {
    payload.ВидОперации = operationKindLiteral;
  }

  // Договор — привязываем, только если найден однозначно. Если у
  // контрагента несколько договоров и непонятно, какой из них — лучше
  // остановиться и попросить вас выбрать вручную, чем угадать неверно.
  // Если нет ни одного договора, ни "Без договора" — тоже не гадаем и не
  // создаём автоматически: просто оставляем поле пустым и явно сообщаем
  // об этом через droppedFields, чтобы бухгалтер создал "Без договора"
  // кнопкой (/create-no-contract) до или после создания черновика.
  if (op.contractStatus === 'ambiguous') {
    throw new Error('У контрагента несколько договоров, и не удалось определить нужный по назначению платежа — выберите договор вручную перед подтверждением.');
  }
  const missingNoContract = op.contractStatus === 'need_create_no_contract' && !op.contractKey;
  if (op.contractKey) {
    payload.Договор_Key = op.contractKey;
  }

  // ВАЖНО: сумма, статья ДДС и договор в этом документе хранятся не
  // только на уровне самого документа, но и в отдельной табличной части
  // "Расшифровка платежа" (видна как таблица внутри документа в 1С).
  // Без нее табличная часть остаётся пустой строкой, даже если общая
  // сумма наверху документа заполнена правильно.
  //
  // ВАЖНО №2: если GUID-поле (Договор_Key, СтатьяДвиженияДенежныхСредств_Key)
  // неизвестно — его нужно просто НЕ включать в объект, а не отправлять
  // пустую строку. 1С понимает "поля нет" нормально, а вот пустую строку
  // вместо GUID отвергает ошибкой "Не удалось разобрать строку '' как
  // значение типа Edm.Guid".
  const amountAbs = Math.abs(op.amount);
  const lineItem = {
    LineNumber: 1,
    СуммаПлатежа: amountAbs,
    КурсВзаиморасчетов: 1,
    СуммаВзаиморасчетов: amountAbs,
  };
  if (op.contractKey) lineItem.Договор_Key = op.contractKey;
  if (categoryKey) lineItem.СтатьяДвиженияДенежныхСредств_Key = categoryKey;

  // "Вид задолженности" — раньше бралось ТОЛЬКО из истории 1С; если у
  // контрагента ещё нет истории, поле молча оставалось пустым, даже если
  // resolvedDebtType (из правила/ручной правки/подстановки по умолчанию)
  // был известен. Приоритет: история 1С (100% верно) > то, что определили
  // для этой операции.
  const debtTypeValue = op.historicalDebtType || resolvedDebtType || '';
  if (debtTypeValue) lineItem.ВидЗадолженности = debtTypeValue;

  // "Счёт учёта (БУ)" в табличной части — ищем GUID счёта по коду
  // (resolvedAccount, например "1251"). Раньше resolvedAccount вообще не
  // использовался внутри этой функции — счёт нигде не попадал в документ.
  const accountKey = await findAccountKeyByCode(resolvedAccount, settings);
  if (accountKey) lineItem.СчетУчета_Key = accountKey;

  // "Подотчётник" — отдельное поле табличной части именно для операции
  // "Перечисление денежных средств подотчётнику": получатель там обычно
  // указывается той же ссылкой, что и Контрагент_Key в шапке документа.
  // Добавляем это поле, только когда вид операции действительно про
  // подотчётные средства — на остальных документах такого поля нет.
  const isAccountablePersonPayment = /подотчет/i.test(
    String(resolvedOperationKind || '') + ' ' + String(op.historicalOperationKind || '')
  );
  if (isAccountablePersonPayment && op.counterpartyKey) {
    lineItem.Подотчетник_Key = op.counterpartyKey;
  }

  payload.РасшифровкаПлатежа = [lineItem];

  if (!op.counterpartyKey) {
    throw new Error('Контрагент не сопоставлен со справочником 1С — сначала выберите контрагента вручную');
  }

  async function postDocument(body) {
    return fetch(`${endpoint}?$format=json`, {
      method: 'POST',
      headers: {
        Authorization: `Basic ${auth}`,
        'Content-Type': 'application/json',
        Accept: 'application/json',
      },
      body: JSON.stringify(body),
    });
  }

  // Точные названия некоторых полей (Подотчетник_Key, СчетУчета_Key и т.п.)
  // мы не можем угадать со 100% гарантией без доступа к вашей конкретной базе
  // 1С — они могут называться иначе в вашей конфигурации. Поэтому вместо
  // того чтобы один раз попробовать и в случае ошибки просто провалить
  // создание черновика целиком — если 1С отвечает ошибкой, ссылающейся на
  // одно из НЕОБЯЗАТЕЛЬНЫХ полей ниже, мы убираем именно это поле (и из
  // шапки документа, и из строки табличной части) и пробуем снова. Так
  // черновик создаётся почти всегда, а какие поля не прижились — видно в
  // истории действий (see /api/operations/:id/confirm), чтобы можно было
  // подсказать нам точное название поля под вашу конфигурацию.
  const OPTIONAL_FIELDS = [
    'ВидОперации',
    'Договор_Key',
    'СтатьяДвиженияДенежныхСредств_Key',
    'ВидЗадолженности',
    'Подотчетник_Key',
    'СчетУчета_Key',
    'РасшифровкаПлатежа',
  ];
  function fieldIsPresent(payloadObj, fieldName) {
    if (fieldName in payloadObj) return true;
    return Array.isArray(payloadObj.РасшифровкаПлатежа) && payloadObj.РасшифровкаПлатежа.some(row => fieldName in row);
  }
  function stripOptionalField(payloadObj, fieldName) {
    if (fieldName === 'РасшифровкаПлатежа') {
      const { РасшифровкаПлатежа, ...rest } = payloadObj;
      return rest;
    }
    const clone = { ...payloadObj };
    delete clone[fieldName];
    if (Array.isArray(clone.РасшифровкаПлатежа)) {
      clone.РасшифровкаПлатежа = clone.РасшифровкаПлатежа.map(row => {
        if (!(fieldName in row)) return row;
        const rowCopy = { ...row };
        delete rowCopy[fieldName];
        return rowCopy;
      });
    }
    return clone;
  }

  let currentPayload = payload;
  let response = await postDocument(currentPayload);
  const droppedFields = [];
  if (missingNoContract) droppedFields.push('Договор_Key (нет договора и нет «Без договора» — создайте кнопкой «Без договора» или в 1С)');
  let safetyCounter = 0;
  while (!response.ok && safetyCounter < OPTIONAL_FIELDS.length) {
    safetyCounter++;
    const text = await response.text().catch(() => '');
    const badField = OPTIONAL_FIELDS.find(f => fieldIsPresent(currentPayload, f) && text.includes(f));
    if (!badField) {
      throw new Error(`1С ответила ${response.status}: ${text.slice(0, 300)}`);
    }
    droppedFields.push(badField);
    currentPayload = stripOptionalField(currentPayload, badField);
    response = await postDocument(currentPayload);
  }
  if (!response.ok) {
    const text = await response.text().catch(() => '');
    throw new Error(`1С ответила ${response.status}: ${text.slice(0, 300)}`);
  }

  const data = await response.json().catch(() => ({}));
  return { docNumber: data.Number || null, droppedFields };
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
