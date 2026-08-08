// Личный кабинет · обработка банковских выписок для 1С
//
// Что делает этот сервер:
//  1. Пускает в кабинет только по паролю (CABINET_PASSWORD из .env)
//  2. Хранит адрес/логин/пароль вашей 1С REST-службы (в файле data/settings.json)
//  3. Принимает файл выписки (.xlsx), разбирает по строкам
//  4. По кнопке "подтвердить" отправляет операцию в вашу 1С как ЧЕРНОВИК
//     (не проведённый документ) через ваш существующий REST-сервис
//  5. Ведёт историю действий
//
// ВАЖНО: функция createDraftInOnec() в самом низу файла — это ЗАГЛУШКА.
// Её нужно донастроить под ваш конкретный REST-сервис (адрес, формат
// запроса), когда у вас будет эта информация. Пока она просто пишет
// в историю "требуется настройка" и ничего не отправляет в 1С.

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
        <p style="font-size:12.5px;color:var(--muted);margin:-6px 0 18px;">Правило вида: если в назначении платежа (или в имени контрагента) есть такой-то текст — предложить статью ДДС и/или счёт. Статья ДДС обязательна для подтверждения операции. Правила применяются сразу, без перезагрузки выписки.</p>
        <div class="settings-card" style="max-width:none;">
          <div style="display:flex;gap:10px;align-items:flex-end;flex-wrap:wrap;">
            <div style="flex:1;min-width:140px;">
              <label>Проверять поле</label>
              <select id="r-field" style="width:100%;padding:9px 11px;border:1px solid var(--line);border-radius:8px;font-size:13px;">
                <option value="purpose">Назначение платежа</option>
                <option value="counterparty">Имя контрагента</option>
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
          <p style="font-size:12px;color:var(--muted);margin:0 0 10px;">Одна строка — одно правило, поля через вертикальную черту «|»: <b>текст-условие | статья ДДС | счёт</b>. Например:<br><span class="mono" style="font-size:12px;">аренда | Аренда офиса | 3360</span></p>
          <textarea id="r-bulk" rows="5" style="width:100%;padding:10px 12px;border:1px solid var(--line);border-radius:8px;font-size:13px;font-family:'IBM Plex Mono',monospace;" placeholder="аренда | Аренда офиса | 3360&#10;подписка | Подписки и лицензии | 3360&#10;реклама | Реклама и маркетинг | 7210"></textarea>
          <div style="margin-top:10px;">
            <button class="btn btn-primary" onclick="importRules()">Загрузить правила</button>
            <span id="rules-import-msg" style="font-size:12.5px;color:var(--muted);margin-left:10px;"></span>
          </div>
        </div>

        <div class="table-card"><table>
          <thead><tr><th>Поле</th><th>Содержит</th><th>Статья ДДС</th><th>Счёт</th><th></th></tr></thead>
          <tbody id="rules-tbody"><tr><td colspan="5" class="empty">Правил пока нет</td></tr></tbody>
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
  box.innerHTML = '<b>Чтобы это не слетало при каждом обновлении сайта — добавьте эти переменные один раз в Render (Environment → Add Environment Variable):</b><br><br>' + lines.join('<br>');
  box.classList.remove('hidden');
}

function money(v){ const sign = v>0?'+':''; return sign + v.toLocaleString('ru-RU') + ' \\u20B8'; }

const expandedOps = new Set();

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
    let statusHtml, actionHtml;
    if(done){
      statusHtml = '<span class="pill pill-approved">Черновик создан</span>';
      actionHtml = '';
    } else if(already){
      statusHtml = '<span class="pill pill-review" style="background:var(--paper-2);color:var(--muted);">Уже в 1С</span>';
      actionHtml = '';
    } else if(isNew){
      statusHtml = '<span class="pill pill-new">Новый контрагент</span>';
      actionHtml = '<button class="icon-btn" onclick="event.stopPropagation();createCounterparty(\\''+o.id+'\\')">Создать контрагента</button>';
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
    const contractText = contractAmbiguous ? '<span style="color:var(--red)">несколько — выбрать вручную</span>' : (o.contractName || '—');
    const isExpanded = expandedOps.has(o.id);

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
        <div class="op-detail-grid">
          <div class="op-detail-item"><label>Статья ДДС</label><div style="color:\${hasCategory?'var(--ink)':'var(--red)'}">\${o.suggestedCategory || 'не определена'}</div></div>
          <div class="op-detail-item"><label>Счёт</label><div class="mono">\${o.suggestedAccount || '—'}</div></div>
          <div class="op-detail-item"><label>Договор</label><div>\${contractText}</div></div>
          <div class="op-detail-item"><label>КНП</label><div class="mono">\${o.knp || '—'}</div></div>
          <div class="op-detail-item op-detail-full"><label>Полное назначение платежа</label><div>\${o.purpose}</div></div>
        </div>
      </div>
    \`;
    list.appendChild(row);
  });
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

async function loadRules(){
  const rules = await api('/api/rules');
  const tbody = document.getElementById('rules-tbody');
  tbody.innerHTML = '';
  if(rules.length===0){ tbody.innerHTML = '<tr><td colspan="5" class="empty">Правил пока нет</td></tr>'; return; }
  rules.forEach(r=>{
    const tr = document.createElement('tr');
    const fieldLabel = r.field==='counterparty' ? 'Имя контрагента' : 'Назначение платежа';
    tr.innerHTML = \`
      <td>\${fieldLabel}</td>
      <td>\${r.contains}</td>
      <td>\${r.category || '—'}</td>
      <td class="mono">\${r.account || '—'}</td>
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

async function deleteRule(id){
  await api('/api/rules/'+id, {method:'DELETE'});
  loadRules(); loadOperations();
}

async function doUpload(file){
  if(!file) return;
  const fd = new FormData();
  fd.append('file', file);
  const res = await fetch('/api/upload', {method:'POST', credentials:'include', body: fd});
  if(!res.ok){ const e = await res.json().catch(()=>({})); alert(e.error||'Ошибка загрузки'); return; }
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

// ---------- простое файловое хранилище (без базы данных) ----------
function readJson(file, fallback) {
  const p = path.join(DATA_DIR, file);
  if (!fs.existsSync(p)) return fallback;
  return JSON.parse(fs.readFileSync(p, 'utf8'));
}
function writeJson(file, data) {
  fs.writeFileSync(path.join(DATA_DIR, file), JSON.stringify(data, null, 2), 'utf8');
}
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR);
if (!fs.existsSync(path.join(DATA_DIR, 'operations.json'))) writeJson('operations.json', []);
if (!fs.existsSync(path.join(DATA_DIR, 'history.json'))) writeJson('history.json', []);
if (!fs.existsSync(path.join(DATA_DIR, 'rules.json'))) writeJson('rules.json', []);

// ВАЖНО: файлы в data/ стираются при каждой пересборке на бесплатном
// плане Render (диск не постоянный). Чтобы подключение к 1С не терялось
// при каждом обновлении кода, настройки при старте сервера в первую
// очередь берутся из переменных окружения Render (Environment) — они,
// в отличие от файлов, сохраняются между пересборками. Если вы один раз
// сохраните подключение через кабинет, сообщение подскажет, что именно
// добавить в Environment на Render, чтобы это стало постоянным.
const envSettings = {
  baseUrl: process.env.ONEC_BASE_URL || '',
  login: process.env.ONEC_LOGIN || '',
  password: process.env.ONEC_PASSWORD || '',
  orgKey: process.env.ONEC_ORG_KEY || '',
  orgName: process.env.ONEC_ORG_NAME || '',
  accountKey: process.env.ONEC_ACCOUNT_KEY || '',
  accountName: process.env.ONEC_ACCOUNT_NAME || '',
};
if (!fs.existsSync(path.join(DATA_DIR, 'settings.json'))) {
  writeJson('settings.json', envSettings.baseUrl ? envSettings : { baseUrl: '', login: '', password: '' });
} else if (envSettings.baseUrl) {
  // Файл есть (сервер уже работал в эту сессию), но если в нём почему-то
  // пусто, а в переменных окружения есть данные — восстанавливаем из них.
  const current = readJson('settings.json', {});
  if (!current.baseUrl) writeJson('settings.json', envSettings);
}

// Применяет ваши правила категоризации к операции: смотрит назначение
// платежа и/или имя контрагента, и если находит совпадение — возвращает
// категорию и счёт. Правила проверяются по порядку, первое совпадение побеждает.
function applyRules(op, rules) {
  for (const rule of rules) {
    const haystack = (rule.field === 'counterparty' ? op.counterparty : op.purpose) || '';
    if (haystack.toLowerCase().includes(String(rule.contains || '').toLowerCase())) {
      return { category: rule.category || '', account: rule.account || '' };
    }
  }
  return { category: '', account: '' };
}

// ---------- вход по паролю (сессия в памяти сервера) ----------
const sessions = new Map(); // token -> expiry
function requireAuth(req, res, next) {
  const token = req.cookies.session;
  if (token && sessions.has(token) && sessions.get(token) > Date.now()) return next();
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
  const token = crypto.randomBytes(24).toString('hex');
  sessions.set(token, Date.now() + 12 * 60 * 60 * 1000); // 12 часов
  res.cookie('session', token, { httpOnly: true, sameSite: 'lax', maxAge: 12 * 60 * 60 * 1000 });
  res.json({ ok: true });
});

app.post('/api/logout', requireAuth, (req, res) => {
  sessions.delete(req.cookies.session);
  res.clearCookie('session');
  res.json({ ok: true });
});

// ---------- настройки подключения к 1С ----------
app.get('/api/settings', requireAuth, (req, res) => {
  const s = readJson('settings.json', {});
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

app.post('/api/settings', requireAuth, (req, res) => {
  const { baseUrl, login, password, orgKey, orgName, accountKey, accountName } = req.body;
  const current = readJson('settings.json', {});
  writeJson('settings.json', {
    baseUrl: baseUrl ?? current.baseUrl,
    login: login ?? current.login,
    password: password || current.password, // пустое поле = пароль не менять
    orgKey: orgKey ?? current.orgKey,
    orgName: orgName ?? current.orgName,
    accountKey: accountKey ?? current.accountKey,
    accountName: accountName ?? current.accountName,
  });
  addHistory('Обновлены настройки подключения к 1С', '—');
  res.json({ ok: true });
});

// Проверяет подключение к 1С и возвращает список организаций и расчётных
// счетов, чтобы их можно было выбрать из выпадающего списка, не зная GUID.
app.post('/api/settings/browse', requireAuth, async (req, res) => {
  const { baseUrl, login, password } = req.body;
  if (!baseUrl || !login) {
    return res.status(400).json({ error: 'Укажите адрес и логин перед проверкой' });
  }
  const current = readJson('settings.json', {});
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
        counterparty: cols.counterparty !== -1 ? String(row[cols.counterparty] || '') : '',
        bin: String(bin || ''),
        purpose: cols.purpose !== -1 ? String(row[cols.purpose] || '') : '',
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
        counterparty: String(pick(row, ['контрагент', 'наименование контрагента', 'плательщик/получатель']) || ''),
        bin: String(pick(row, ['бин', 'иин', 'бин/иин']) || ''),
        purpose: String(pick(row, ['назначение платежа', 'назначение', 'комментарий']) || ''),
        amount,
        suggestedCategory: '',
        status: 'review',
        sourceFile: req.file.originalname,
        rowIndex: i,
      };
    }).filter(op => op.date || op.amount);
  }

  const all = readJson('operations.json', []);

  // Не добавляем операцию повторно, если такая же (по дате, сумме, БИН и
  // назначению) уже есть в списке — иначе повторная загрузка того же файла
  // задваивает список.
  const dedupKey = (o) => `${o.date}|${o.amount}|${o.bin}|${o.purpose}`;
  const existingKeys = new Set(all.map(dedupKey));
  const newOnly = operations.filter(op => !existingKeys.has(dedupKey(op)));
  const skippedDuplicates = operations.length - newOnly.length;

  const updated = [...all, ...newOnly];
  writeJson('operations.json', updated);
  addHistory(
    `Загружена выписка ${req.file.originalname} · ${newOnly.length} новых операций` +
      (skippedDuplicates ? ` · ${skippedDuplicates} пропущено как дубли` : ''),
    '—'
  );

  // Сразу пытаемся сопоставить контрагентов по БИН/ИИН со справочником 1С,
  // и проверить — нет ли такой операции уже среди документов в самой 1С
  // (чтобы не предлагать создавать то, что уже разнесено вручную или
  // другим способом). Всё это — только если подключение уже настроено.
  const settingsForMatch = readJson('settings.json', {});
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

      if (!op.bin) continue;
      try {
        const found = await findCounterpartyByBin(op.bin, settingsForMatch);
        if (found) {
          target.counterpartyKey = found.Ref_Key;
          target.counterpartyMatchedName = found.Description || '';

          // Смотрим, как этот контрагент категоризировался раньше в 1С —
          // если найдём паттерн, используем ту же статью ДДС автоматически.
          try {
            const historical = await findHistoricalCategory(found.Ref_Key, op.amount, settingsForMatch);
            if (historical) {
              target.historicalCategory = historical.categoryName || '';
              target.historicalCategoryKey = historical.categoryKey || '';
              target.historicalOperationKind = historical.operationKind || '';
            }
          } catch (e) {
            // Не критично — просто не будет автопредложения по истории для этой операции.
          }

          // Ищем договор контрагента — если он есть в справочнике и найден
          // однозначно, привяжем его к документу автоматически.
          try {
            const contract = await findContractForCounterparty(found.Ref_Key, op.purpose, settingsForMatch);
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
        // Если проверка не удалась (например, неверный пароль) — не блокируем
        // загрузку, просто оставляем операцию как есть, на проверке.
      }
    }
    writeJson('operations.json', updated);
  }

  res.json({ operations: newOnly, skippedDuplicates });
});

// ---------- список операций на проверку ----------
app.get('/api/operations', requireAuth, (req, res) => {
  const ops = readJson('operations.json', []);
  const rules = readJson('rules.json', []);
  // Категория/счёт считаются "на лету" по текущим правилам — если вы
  // добавите или измените правило, категории в уже загруженных операциях
  // пересчитаются сразу же, без повторной загрузки файла. Историческая
  // категория из 1С (по прошлым документам контрагента) — приоритетнее.
  const withCategories = ops.map(op => {
    const ruleMatch = applyRules(op, rules);
    const category = op.historicalCategory || ruleMatch.category;
    const account = op.historicalAccount || ruleMatch.account;
    return {
      ...op,
      suggestedCategory: category,
      suggestedAccount: account,
      needsAttention: !category, // ДДС обязательна — если не определена, операцию нельзя тихо подтверждать
    };
  });
  res.json(withCategories);
});

// ---------- подтверждение операции: создать черновик в 1С ----------
app.post('/api/operations/:id/confirm', requireAuth, async (req, res) => {
  const all = readJson('operations.json', []);
  const op = all.find(o => o.id === req.params.id);
  if (!op) return res.status(404).json({ error: 'Операция не найдена' });

  const settings = readJson('settings.json', {});
  if (!settings.baseUrl) {
    return res.status(400).json({ error: 'Сначала укажите адрес подключения к 1С в разделе «Настройки»' });
  }

  // Статья ДДС обязательна — без неё документ не создаём, даже если
  // контрагент найден. Либо она должна прийти из истории 1С, либо из
  // ваших правил категоризации.
  const rules = readJson('rules.json', []);
  const ruleMatch = applyRules(op, rules);
  const category = op.historicalCategory || ruleMatch.category;
  if (!category) {
    return res.status(400).json({ error: 'Не определена статья ДДС — добавьте правило категоризации (раздел «Правила») или дождитесь исторического сопоставления, прежде чем подтверждать.' });
  }

  try {
    const result = await createDraftInOnec(op, settings);
    op.status = 'draft_created';
    op.onecDocNumber = result.docNumber || null;
    writeJson('operations.json', all);
    addHistory(`Создан черновик в 1С: ${op.counterparty || 'операция'} · ${op.amount} ₸ · статья: ${category}`, result.docNumber || '—');
    res.json({ ok: true, op });
  } catch (e) {
    addHistory(`Ошибка при создании черновика: ${e.message}`, '—');
    res.status(502).json({ error: 'Не удалось создать документ в 1С: ' + e.message });
  }
});

// ---------- создание НОВОГО контрагента (отдельное подтверждение) ----------
app.post('/api/operations/:id/create-counterparty', requireAuth, async (req, res) => {
  const all = readJson('operations.json', []);
  const op = all.find(o => o.id === req.params.id);
  if (!op) return res.status(404).json({ error: 'Операция не найдена' });
  if (!op.bin) return res.status(400).json({ error: 'У операции не указан БИН/ИИН — сопоставьте контрагента вручную' });

  const settings = readJson('settings.json', {});
  if (!settings.baseUrl) {
    return res.status(400).json({ error: 'Сначала укажите адрес подключения к 1С в разделе «Настройки»' });
  }

  try {
    // На всякий случай проверяем ещё раз прямо перед созданием — вдруг
    // контрагент уже появился в 1С (например, кто-то создал его вручную).
    const existing = await findCounterpartyByBin(op.bin, settings);
    let key, name;
    if (existing) {
      key = existing.Ref_Key;
      name = existing.Description;
    } else {
      const created = await createCounterpartyInOnec(op, settings);
      key = created.Ref_Key;
      name = created.Description;
    }
    op.counterpartyKey = key;
    op.counterpartyMatchedName = name;
    op.status = 'review'; // теперь контрагент есть — можно подтверждать документ как обычно
    writeJson('operations.json', all);
    addHistory(`Создан контрагент в 1С: ${name || op.counterparty} (БИН ${op.bin})`, '—');
    res.json({ ok: true, op });
  } catch (e) {
    addHistory(`Ошибка при создании контрагента: ${e.message}`, '—');
    res.status(502).json({ error: 'Не удалось создать контрагента в 1С: ' + e.message });
  }
});

// ---------- история ----------
// ---------- правила категоризации (редактируются прямо в кабинете) ----------
app.get('/api/rules', requireAuth, (req, res) => {
  res.json(readJson('rules.json', []));
});

app.post('/api/rules', requireAuth, (req, res) => {
  const { field, contains, category, account } = req.body;
  if (!contains || (!category && !account)) {
    return res.status(400).json({ error: 'Укажите условие и хотя бы одно из: категорию или счёт' });
  }
  const rules = readJson('rules.json', []);
  rules.push({
    id: crypto.randomUUID(),
    field: field === 'counterparty' ? 'counterparty' : 'purpose',
    contains: String(contains),
    category: String(category || ''),
    account: String(account || ''),
  });
  writeJson('rules.json', rules);
  addHistory(`Добавлено правило: "${contains}" → ${category || ''} ${account ? '(счёт ' + account + ')' : ''}`, '—');
  res.json(rules);
});

// Массовый импорт правил текстом — по одному правилу на строку, поля через
// вертикальную черту: текст-условие | категория (статья ДДС) | счёт | (поле: purpose/counterparty, необязательно)
// Пример строки: аренда | Аренда офиса | 3360
app.post('/api/rules/import', requireAuth, (req, res) => {
  const { text } = req.body;
  if (!text || !text.trim()) {
    return res.status(400).json({ error: 'Вставьте текст с правилами' });
  }
  const rules = readJson('rules.json', []);
  const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
  let added = 0;
  let skipped = 0;
  for (const line of lines) {
    const parts = line.split('|').map(p => p.trim());
    if (parts.length < 2 || !parts[0]) { skipped++; continue; }
    const [contains, category, account, fieldRaw] = parts;
    if (!category && !account) { skipped++; continue; }
    rules.push({
      id: crypto.randomUUID(),
      field: fieldRaw === 'counterparty' ? 'counterparty' : 'purpose',
      contains,
      category: category || '',
      account: account || '',
    });
    added++;
  }
  writeJson('rules.json', rules);
  addHistory(`Импортировано правил: ${added}${skipped ? `, пропущено строк: ${skipped}` : ''}`, '—');
  res.json({ rules, added, skipped });
});

app.delete('/api/rules/:id', requireAuth, (req, res) => {
  const rules = readJson('rules.json', []);
  const filtered = rules.filter(r => r.id !== req.params.id);
  writeJson('rules.json', filtered);
  addHistory('Удалено правило категоризации', '—');
  res.json(filtered);
});

// ---------- очистить список операций (начать заново) ----------
app.post('/api/operations/clear', requireAuth, (req, res) => {
  writeJson('operations.json', []);
  addHistory('Список операций очищен вручную', '—');
  res.json({ ok: true });
});

app.get('/api/history', requireAuth, (req, res) => {
  res.json(readJson('history.json', []));
});
function addHistory(action, doc) {
  const h = readJson('history.json', []);
  h.unshift({ time: new Date().toISOString(), action, doc });
  writeJson('history.json', h.slice(0, 500));
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

// Смотрит в 1С, какие документы уже создавались для этого контрагента
// раньше, и берёт статью ДДС из самого свежего такого документа — чтобы
// новая операция была категоризирована так же, как вы (или REST-сервис)
// категоризировали её раньше для этого же контрагента.
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
async function findContractForCounterparty(counterpartyKey, purposeText, settings) {
  if (!counterpartyKey) return { status: 'none' };
  const base = settings.baseUrl.replace(/\/+$/, '');
  const auth = Buffer.from(`${settings.login}:${settings.password}`).toString('base64');
  const candidates = ['Catalog_ДоговорыКонтрагентов', 'Catalog_Договоры'];

  for (const catalog of candidates) {
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

async function findHistoricalCategory(counterpartyKey, amount, settings) {
  if (!counterpartyKey) return null;
  const base = settings.baseUrl.replace(/\/+$/, '');
  const auth = Buffer.from(`${settings.login}:${settings.password}`).toString('base64');
  const docType = amount >= 0 ? 'Document_ПлатежноеПоручениеВходящее' : 'Document_ПлатежноеПоручениеИсходящее';

  const filter = encodeURIComponent(`Контрагент_Key eq guid'${counterpartyKey}'`);
  const url = `${base}/${docType}?$format=json&$filter=${filter}&$orderby=Date desc&$top=1&$select=СтатьяДвиженияДенежныхСредств_Key,ВидОперации`;

  const response = await fetch(url, { headers: { Authorization: `Basic ${auth}`, Accept: 'application/json' } });
  if (!response.ok) return null;
  const data = await response.json().catch(() => null);
  const doc = data && data.value && data.value[0];
  if (!doc || !doc.СтатьяДвиженияДенежныхСредств_Key) return null;

  // Статья хранится как GUID — подтягиваем её человекочитаемое название
  const categoryKey = doc.СтатьяДвиженияДенежныхСредств_Key;
  const categoryCatalogs = ['Catalog_СтатьиДвиженияДенежныхСредств', 'Catalog_СтатьиДДС'];
  for (const catalog of categoryCatalogs) {
    const catUrl = `${base}/${catalog}(guid'${categoryKey}')?$format=json&$select=Description`;
    const catResp = await fetch(catUrl, { headers: { Authorization: `Basic ${auth}`, Accept: 'application/json' } });
    if (catResp.ok) {
      const catData = await catResp.json().catch(() => null);
      if (catData && catData.Description) {
        return { categoryKey, categoryName: catData.Description, operationKind: doc.ВидОперации || '' };
      }
    }
  }
  return { categoryKey, categoryName: '', operationKind: doc.ВидОперации || '' };
}

async function findCounterpartyByBin(bin, settings) {
  const base = settings.baseUrl.replace(/\/+$/, '');
  const auth = Buffer.from(`${settings.login}:${settings.password}`).toString('base64');
  const filter = encodeURIComponent(`БИН eq '${bin}' or ИИН eq '${bin}' or ИННЮЛ eq '${bin}' or ИННФЛ eq '${bin}'`);
  const url = `${base}/Catalog_Контрагенты?$format=json&$filter=${filter}`;

  const response = await fetch(url, {
    headers: { Authorization: `Basic ${auth}`, Accept: 'application/json' },
  });
  if (!response.ok) {
    const text = await response.text().catch(() => '');
    throw new Error(`Поиск контрагента: 1С ответила ${response.status}: ${text.slice(0, 200)}`);
  }
  const data = await response.json().catch(() => ({}));
  const list = data.value || [];
  return list.length > 0 ? list[0] : null;
}

// Создаёт нового контрагента в справочнике 1С по данным из операции выписки.
async function createCounterpartyInOnec(op, settings) {
  const base = settings.baseUrl.replace(/\/+$/, '');
  const auth = Buffer.from(`${settings.login}:${settings.password}`).toString('base64');

  const payload = {
    Description: op.counterparty || op.bin,
    БИН: op.bin,
    Комментарий: 'Создан автоматически · личный кабинет разноски выписок',
  };

  const response = await fetch(`${base}/Catalog_Контрагенты?$format=json`, {
    method: 'POST',
    headers: {
      Authorization: `Basic ${auth}`,
      'Content-Type': 'application/json',
      Accept: 'application/json',
    },
    body: JSON.stringify(payload),
  });
  if (!response.ok) {
    const text = await response.text().catch(() => '');
    throw new Error(`1С ответила ${response.status}: ${text.slice(0, 300)}`);
  }
  return response.json().catch(() => ({}));
}

// =====================================================================
// СЮДА вставлен вызов через OData стандартного интерфейса 1С:Фреш —
// на основе полей, которые реально видны в вашей базе (Date, Number,
// Организация_Key, СчетОрганизации_Key, Контрагент_Key, ВидОперации,
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
async function createDraftInOnec(op, settings) {
  const base = settings.baseUrl.replace(/\/+$/, ''); // убираем лишний / на конце
  const docType = op.amount >= 0 ? 'Document_ПлатежноеПоручениеВходящее' : 'Document_ПлатежноеПоручениеИсходящее';
  const endpoint = `${base}/${docType}`;
  const auth = Buffer.from(`${settings.login}:${settings.password}`).toString('base64');

  const payload = {
    Date: op.date,
    Posted: false, // черновик — не проводится автоматически
    Организация_Key: settings.orgKey || '',
    СчетОрганизации_Key: settings.accountKey || '',
    Контрагент_Key: op.counterpartyKey || '', // если контрагент не найден — не отправляем, см. ниже
    СуммаДокумента: Math.abs(op.amount),
    НазначениеПлатежа: op.purpose,
    Комментарий: op.purpose, // должно совпадать с назначением платежа, как в вашей 1С
  };

  // Если для этого контрагента нашлась статья ДДС / вид операции по прошлым
  // документам — используем их, чтобы категоризация совпадала с тем, как
  // вы (или прежний REST-сервис) разносили такие операции раньше.
  if (op.historicalCategoryKey) {
    payload.СтатьяДвиженияДенежныхСредств_Key = op.historicalCategoryKey;
  }
  if (op.historicalOperationKind) {
    payload.ВидОперации = op.historicalOperationKind;
  }

  // Договор — привязываем, только если найден однозначно. Если у
  // контрагента несколько договоров и непонятно, какой из них — лучше
  // остановиться и попросить вас выбрать вручную, чем угадать неверно.
  if (op.contractStatus === 'ambiguous') {
    throw new Error('У контрагента несколько договоров, и не удалось определить нужный по назначению платежа — выберите договор вручную перед подтверждением.');
  }
  if (op.contractKey) {
    payload.Договор_Key = op.contractKey;
  }

  if (!op.counterpartyKey) {
    throw new Error('Контрагент не сопоставлен со справочником 1С — сначала выберите контрагента вручную');
  }

  const response = await fetch(`${endpoint}?$format=json`, {
    method: 'POST',
    headers: {
      Authorization: `Basic ${auth}`,
      'Content-Type': 'application/json',
      Accept: 'application/json',
    },
    body: JSON.stringify(payload),
  });

  if (!response.ok) {
    const text = await response.text().catch(() => '');
    throw new Error(`1С ответила ${response.status}: ${text.slice(0, 300)}`);
  }

  const data = await response.json().catch(() => ({}));
  return { docNumber: data.Number || null };
}

app.listen(PORT, () => {
  console.log(`Кабинет запущен: http://localhost:${PORT}`);
});
