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
app.use(express.static(path.join(__dirname, 'public')));

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
if (!fs.existsSync(path.join(DATA_DIR, 'settings.json'))) writeJson('settings.json', { baseUrl: '', login: '', password: '' });

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
  res.json({ baseUrl: s.baseUrl || '', login: s.login || '', passwordSet: !!s.password });
});

app.post('/api/settings', requireAuth, (req, res) => {
  const { baseUrl, login, password } = req.body;
  const current = readJson('settings.json', {});
  writeJson('settings.json', {
    baseUrl: baseUrl ?? current.baseUrl,
    login: login ?? current.login,
    password: password || current.password, // пустое поле = пароль не менять
  });
  addHistory('Обновлены настройки подключения к 1С', '—');
  res.json({ ok: true });
});

// ---------- загрузка и разбор выписки (.xlsx) ----------
app.post('/api/upload', requireAuth, upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'Файл не получен' });

  let rows;
  try {
    const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
  } catch (e) {
    return res.status(400).json({ error: 'Не получилось прочитать файл: ' + e.message });
  }

  // Разбор строк выписки. Названия колонок в реальных выписках сильно
  // различаются по банкам — здесь распознаются самые частые варианты.
  // Если ваш банк называет колонки иначе, пришлите пример файла,
  // допишем сопоставление.
  const pick = (row, names) => {
    for (const n of names) {
      const key = Object.keys(row).find(k => k.trim().toLowerCase() === n);
      if (key) return row[key];
    }
    return '';
  };

  const operations = rows.map((row, i) => {
    const dateRaw = pick(row, ['дата', 'дата операции', 'дата документа']);
    const amountIn = parseFloat(pick(row, ['приход', 'сумма прихода', 'дебет']) || 0) || 0;
    const amountOut = parseFloat(pick(row, ['расход', 'сумма расхода', 'кредит']) || 0) || 0;
    const amountSingle = parseFloat(pick(row, ['сумма']) || 0) || 0;
    const amount = amountIn || (amountOut ? -amountOut : amountSingle);
    return {
      id: crypto.randomUUID(),
      date: String(dateRaw || ''),
      counterparty: String(pick(row, ['контрагент', 'наименование контрагента', 'плательщик/получатель']) || ''),
      bin: String(pick(row, ['бин', 'иин', 'бин/иин']) || ''),
      purpose: String(pick(row, ['назначение платежа', 'назначение', 'комментарий']) || ''),
      amount,
      suggestedCategory: '', // сюда позже подключится ИИ-категоризация
      status: 'review', // review | new_counterparty | confirmed | draft_created
      sourceFile: req.file.originalname,
      rowIndex: i,
    };
  }).filter(op => op.date || op.amount);

  const all = readJson('operations.json', []);
  const updated = [...all, ...operations];
  writeJson('operations.json', updated);
  addHistory(`Загружена выписка ${req.file.originalname} · ${operations.length} операций`, '—');

  res.json({ operations });
});

// ---------- список операций на проверку ----------
app.get('/api/operations', requireAuth, (req, res) => {
  res.json(readJson('operations.json', []));
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

  try {
    const result = await createDraftInOnec(op, settings);
    op.status = 'draft_created';
    op.onecDocNumber = result.docNumber || null;
    writeJson('operations.json', all);
    addHistory(`Создан черновик в 1С: ${op.counterparty || 'операция'} · ${op.amount} ₸`, result.docNumber || '—');
    res.json({ ok: true, op });
  } catch (e) {
    addHistory(`Ошибка при создании черновика: ${e.message}`, '—');
    res.status(502).json({ error: 'Не удалось создать документ в 1С: ' + e.message });
  }
});

// ---------- история ----------
app.get('/api/history', requireAuth, (req, res) => {
  res.json(readJson('history.json', []));
});
function addHistory(action, doc) {
  const h = readJson('history.json', []);
  h.unshift({ time: new Date().toISOString(), action, doc });
  writeJson('history.json', h.slice(0, 500));
}

// =====================================================================
// СЮДА нужно вставить настоящий вызов вашего REST-сервиса 1С.
// Пример ниже — общий шаблон (Basic Auth + JSON), его нужно поправить
// под точный адрес и формат, который принимает ваш сервис.
// =====================================================================
async function createDraftInOnec(op, settings) {
  const endpoint = settings.baseUrl; // например: https://ваш-домен/hs/bank/create
  const auth = Buffer.from(`${settings.login}:${settings.password}`).toString('base64');

  const payload = {
    Дата: op.date,
    Проведен: false, // черновик — не проводится автоматически
    Сумма: Math.abs(op.amount),
    ВидОперации: op.amount >= 0 ? 'ПлатежноеПоручениеВходящее' : 'ПлатежноеПоручениеИсходящее',
    Контрагент: op.counterparty,
    БИН: op.bin,
    НазначениеПлатежа: op.purpose,
    Комментарий: 'Черновик создан автоматически · личный кабинет разноски выписок',
  };

  const response = await fetch(endpoint, {
    method: 'POST',
    headers: {
      Authorization: `Basic ${auth}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(payload),
  });

  if (!response.ok) {
    const text = await response.text().catch(() => '');
    throw new Error(`1С ответила ${response.status}: ${text.slice(0, 300)}`);
  }

  const data = await response.json().catch(() => ({}));
  return { docNumber: data.Номер || data.number || null };
}

app.listen(PORT, () => {
  console.log(`Кабинет запущен: http://localhost:${PORT}`);
});
