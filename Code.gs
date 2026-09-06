/***** CONFIG *****/
const VERSION = 'rsvp-game-progressive-v4-2026-09-06';
const SHEET_ID = '1OGcM6ydIZyDnBPxuJ-NFtuEa8jgHFGvR9O45cF83GEU';
const SHEET_NAME = 'RSVP';
const OVERWRITE_BY_NAME = true; // Сохраняет прежнее объединение гостей по ФИО.
const TIMEZONE = 'Europe/Moscow';
const BASE_HEADERS = ['ts_client', 'ts_server', 'firstName', 'lastName', 'willAttend', 'variant'];
const WISH_HEADERS = ['alcohol', 'softDrinks', 'drinkNotes', 'diet', 'foodNotes', 'allergyStatus', 'allergyDetails', 'song', 'notes'];
const META_HEADERS = ['schemaVersion', 'submissionId', 'payloadHash', 'wishesSaved', 'requestHistory'];
const HEADERS = BASE_HEADERS.concat(WISH_HEADERS, META_HEADERS);

/* Установка: заменить старый Code.gs целиком, запустить setupRsvp один раз,
 * затем обновить существующее веб-развёртывание, выбрав новую версию.
 * Первые 6 столбцов сохраняются. Недостающие столбцы добавляются справа.
 * Старые данные, чужие столбцы и формулы в них не перезаписываются.
 * submissionId защищает повторы; ФИО не является проверкой личности.
 * Для однофамильцев с одинаковыми именами нужны отдельные ID приглашений.
 */
function setupRsvp() {
  const lock = LockService.getScriptLock();
  lock.waitLock(8000);
  try {
    _sheet();
    _gameSecret();
    const attempts = _gameSheet('GameAttempts', GAME_ATTEMPT_HEADERS);
    _gameRanking(_gameRows(attempts));
    SpreadsheetApp.flush();
    return {ok: true, version: VERSION};
  } finally { lock.releaseLock(); }
}

function doGet() {
  return _json({ok: true, ping: 'alive', overwrite: OVERWRITE_BY_NAME, version: VERSION, schemaVersion: 2});
}

function doPost(e) {
  let lock;
  let locked = false;
  try {
    const raw = e && e.postData && e.postData.contents;
    if (typeof raw === 'string' && raw.length <= 20000) {
      let input;
      try { input = JSON.parse(raw); } catch (_) { /* _parse supplies the error. */ }
      if (input && input.action) {
        lock = LockService.getScriptLock();
        locked = lock.tryLock(8000);
        if (!locked) return _json({ok:false,error:'BUSY',message:'Повторите через несколько секунд.'});
        return _json(_gameAction(input));
      }
    }
    const data = _parse(e);
    const hash = _hash(JSON.stringify({
      firstName: data.firstName, lastName: data.lastName,
      willAttend: data.willAttend, variant: data.variant, wishes: data.wishes
    })); // ts меняется при повторе и намеренно не участвует в хеше.
    lock = LockService.getScriptLock();
    locked = lock.tryLock(8000);
    if (!locked) return _json({ok: false, error: 'BUSY', message: 'Повторите отправку через несколько секунд.'});

    const {sheet, columns} = _sheet();
    const count = sheet.getLastRow() - 1;
    const rows = count > 0 ? sheet.getRange(2, 1, count, sheet.getLastColumn()).getValues() : [];
    let rowIndex = -1;
    let history = [];
    for (let i = rows.length - 1; i >= 0; i--) {
      const row = rows[i];
      const rowHistory = _history(row[columns.requestHistory - 1]);
      const receipt = data.submissionId && rowHistory.find(item => item.id === data.submissionId);
      if (receipt) {
        if (receipt.hash !== hash) return _json({ok: false, error: 'SUBMISSION_CONFLICT', message: 'Идентификатор уже использован для другого ответа.'});
        if (row[columns.submissionId - 1] !== data.submissionId) {
          return _json({ok: false, error: 'STALE_SUBMISSION', message: 'После этого ответа уже был сохранён новый. Обновите анкету.'});
        }
        return _json({ok: true, wishesSaved: receipt.wishesSaved, duplicate: true, version: VERSION, gameAccess: _gameAccess(data)});
      }
      if (OVERWRITE_BY_NAME && rowIndex === -1 &&
          _norm(row[columns.firstName - 1]) === _norm(data.firstName) &&
          _norm(row[columns.lastName - 1]) === _norm(data.lastName)) {
        rowIndex = i + 2;
        history = rowHistory;
      }
    }
    if (data.submissionId) history.push({id: data.submissionId, hash, wishesSaved: data.wishes !== null});
    const historyJson = JSON.stringify(history);
    if (historyJson.length > 40000) _fail('HISTORY_LIMIT', 'Для этого гостя достигнут предел истории отправок. Свяжитесь с организатором.');

    const targetRow = rowIndex === -1 ? sheet.getLastRow() + 1 : rowIndex;
    if (targetRow > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), targetRow - sheet.getMaxRows());
    const fields = {
      ts_client: data.ts,
      ts_server: Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ssXXX"),
      firstName: data.firstName, lastName: data.lastName,
      willAttend: data.willAttend, variant: data.variant
    };
    if (data.wishes !== null) {
      WISH_HEADERS.forEach(key => { fields[key] = Array.isArray(data.wishes[key]) ? data.wishes[key].join(', ') : data.wishes[key]; });
    }
    // Старый запрос без wishes не стирает ранее заполненную анкету.
    // При отказе пожелания остаются в строке; актуальное участие = false.
    // Сначала сбрасываем квитанцию, затем пишем данные, в конце подтверждение.
    sheet.getRange(targetRow, columns.submissionId).setValue('');
    _writeFields(sheet, targetRow, columns, fields);
    SpreadsheetApp.flush();
    _writeFields(sheet, targetRow, columns, {
      schemaVersion: data.wishes !== null ? 2 : 1,
      submissionId: data.submissionId,
      payloadHash: hash,
      wishesSaved: data.wishes !== null,
      requestHistory: historyJson
    });
    SpreadsheetApp.flush();
    return _json({ok: true, wishesSaved: data.wishes !== null, duplicate: false, version: VERSION, gameAccess: _gameAccess(data)});
  } catch (err) {
    // Не пишем имена, аллергии, тело запроса и внутренние ошибки в публичный ответ.
    if (err.publicCode) return _json({ok: false, error: err.publicCode, message: err.message});
    console.error('RSVP write failed: ' + (err.name || 'Error'));
    return _json({ok: false, error: 'SERVER_ERROR', message: 'Не удалось подтвердить запись. Повторите позже или свяжитесь с организатором.'});
  } finally {
    if (locked) lock.releaseLock();
  }
}

/***** SHEET: вызывать только под ScriptLock *****/
function _sheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  let headers = [];
  if (sheet.getLastRow() > 0) {
    headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
    if (BASE_HEADERS.some((name, i) => headers[i] !== name)) {
      _fail('SHEET_HEADERS', 'Первые шесть заголовков RSVP отличаются от исходной схемы. Проверьте таблицу.');
    }
    for (const name of HEADERS) {
      if (headers.filter(header => header === name).length > 1) _fail('SHEET_HEADERS', 'В RSVP повторяются служебные заголовки.');
    }
  }
  const missing = HEADERS.filter(name => !headers.includes(name));
  if (missing.length) {
    const needed = headers.length + missing.length;
    if (needed > sheet.getMaxColumns()) sheet.insertColumnsAfter(sheet.getMaxColumns(), needed - sheet.getMaxColumns());
    sheet.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
    headers = headers.concat(missing);
  }
  const columns = Object.create(null);
  HEADERS.forEach(name => { columns[name] = headers.indexOf(name) + 1; });
  return {sheet, columns};
}

function _writeFields(sheet, row, columns, fields) {
  const cells = Object.keys(fields).map(name => ({column: columns[name], value: _cell(fields[name])})).sort((a, b) => a.column - b.column);
  // Пакеты смежных столбцов: чужие столбцы между ними остаются нетронутыми.
  let group = [];
  function flush() {
    if (group.length) sheet.getRange(row, group[0].column, 1, group.length).setValues([group.map(cell => cell.value)]);
    group = [];
  }
  cells.forEach(cell => {
    if (group.length && cell.column !== group[group.length - 1].column + 1) flush();
    group.push(cell);
  });
  flush();
}

/***** VALIDATION *****/
function _parse(e) {
  const raw = e && e.postData && e.postData.contents;
  if (typeof raw !== 'string' || !raw.trim()) _fail('EMPTY_BODY', 'Пустой запрос.');
  if (raw.length > 20000) _fail('BODY_TOO_LARGE', 'Слишком длинный ответ.');
  let input;
  try { input = JSON.parse(raw); } catch (_) { _fail('INVALID_JSON', 'Некорректный JSON.'); }
  if (!_object(input)) _fail('INVALID_DATA', 'Ожидается объект.');
  const firstName = _text(input.firstName, 'Имя', 80, true);
  const lastName = _text(input.lastName, 'Фамилия', 80, true);
  if (!_norm(firstName) || !_norm(lastName)) _fail('INVALID_NAME', 'Укажите имя и фамилию.');
  if (typeof input.willAttend !== 'boolean') _fail('INVALID_ATTENDANCE', 'willAttend должен быть true или false.');
  const variant = _text(input.variant, 'Вариант', 40);
  const ts = _text(input.ts, 'Дата', 40);
  if (ts && !Number.isFinite(Date.parse(ts))) _fail('INVALID_DATE', 'Некорректная дата.');
  const submissionId = _text(input.submissionId, 'ID отправки', 100);
  if (submissionId && !/^[a-zA-Z0-9_-]+$/.test(submissionId)) _fail('INVALID_SUBMISSION_ID', 'Некорректный ID отправки.');
  let wishes = null;
  if (Object.prototype.hasOwnProperty.call(input, 'wishes')) {
    if (!_object(input.wishes) || !input.willAttend || input.schemaVersion !== 2 || !submissionId) {
      _fail('INVALID_WISHES', 'Анкета требует schemaVersion=2, submissionId и willAttend=true.');
    }
    const w = input.wishes;
    if (Object.keys(w).some(key => !WISH_HEADERS.includes(key))) _fail('UNKNOWN_FIELD', 'Неизвестное поле анкеты.');
    wishes = {
      alcohol: _choices(w.alcohol, ['Игристое', 'Белое вино', 'Красное вино', 'Крепкие напитки', 'Пиво / сидр', 'Не пью алкоголь']),
      softDrinks: _choices(w.softDrinks, ['Вода без газа', 'Вода с газом', 'Сок / морс', 'Лимонад', 'Чай', 'Кофе']),
      drinkNotes: _text(w.drinkNotes, 'Напитки', 300),
      diet: _choice(w.diet, ['', 'Без ограничений', 'Без мяса, рыбу ем', 'Вегетарианское', 'Веганское', 'Другое']),
      foodNotes: _text(w.foodNotes, 'Питание', 500),
      allergyStatus: _choice(w.allergyStatus, ['', 'Нет известных', 'Есть', 'Обсужу лично']),
      allergyDetails: _text(w.allergyDetails, 'Аллергии', 1200),
      song: _text(w.song, 'Песня', 300),
      notes: _text(w.notes, 'Пожелания', 1500)
    };
    if (wishes.alcohol.includes('Не пью алкоголь') && wishes.alcohol.length > 1) _fail('DRINK_CONFLICT', 'Выберите алкогольные напитки или вариант «Не пью алкоголь».');
    if (wishes.allergyStatus !== 'Есть') wishes.allergyDetails = '';
  } else if (input.schemaVersion !== undefined && input.schemaVersion !== 1) {
    _fail('MISSING_WISHES', 'Не переданы поля анкеты.');
  }
  return {firstName, lastName, willAttend: input.willAttend, variant, ts, submissionId, wishes};
}
function _object(value) { return value !== null && typeof value === 'object' && !Array.isArray(value); }
function _text(value, label, max, required) {
  if (value === undefined || value === null) value = '';
  if (typeof value !== 'string') _fail('INVALID_TEXT', label + ': ожидается текст.');
  value = value.trim();
  if (value.length > max || (required && !value)) _fail('INVALID_TEXT', label + ': укажите от ' + (required ? 1 : 0) + ' до ' + max + ' символов.');
  if (/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/.test(value)) _fail('INVALID_TEXT', label + ': недопустимые символы.');
  return value;
}
function _choices(value, allowed) {
  if (value === undefined) return [];
  if (!Array.isArray(value) || value.length > allowed.length || value.some(item => !allowed.includes(item))) _fail('INVALID_CHOICE', 'Некорректный вариант ответа.');
  return allowed.filter(item => value.includes(item));
}
function _choice(value, allowed) {
  if (value === undefined) return '';
  if (!allowed.includes(value)) _fail('INVALID_CHOICE', 'Некорректный вариант ответа.');
  return value;
}
function _norm(value) {
  return String(value || '').toLowerCase().replace(/\u00A0/g, ' ').replace(/\s+/g, ' ').replace(/ё/g, 'е').replace(/[.\-–—'_]+/g, '').trim();
}
function _history(value) {
  if (!value) return [];
  try {
    const history = JSON.parse(value);
    if (!Array.isArray(history) || history.some(item => !_object(item) || typeof item.id !== 'string' || typeof item.hash !== 'string' || typeof item.wishesSaved !== 'boolean')) throw new Error();
    return history;
  } catch (_) { _fail('INVALID_HISTORY', 'Служебная история RSVP повреждена. Нужна проверка организатором.'); }
}
function _hash(value) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value, Utilities.Charset.UTF_8).map(byte => ('0' + (byte & 255).toString(16)).slice(-2)).join('');
}
function _cell(value) {
  // Апостроф заставляет Sheets принять потенциальную формулу как текст.
  return typeof value === 'string' && /^[\s]*[=+@\-']/.test(value) ? "'" + value : value;
}
function _fail(code, message) { const error = new Error(message); error.publicCode = code; throw error; }
function _json(obj) {
  // TextOutput не имеет setResponseCode: клиент проверяет ok в JSON.
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}


/***** ИГРА: служебные листы создаются в той же таблице *****/
const GAME_VERSION = 'bouquet-lives-v2';
const LEGACY_GAME_VERSION = 'bouquet-lives-v1';
const GAME_LIVES = 5;
const GAME_ATTEMPT_HEADERS = ['ts_server','runId','guestKey','firstName','lastName','score','flowersCaught','ringsCaught','missed','activeSeconds','gameVersion','payloadHash'];
const GAME_RANK_HEADERS = ['Место','Имя','Фамилия','Лучший результат','Попыток','Дата рекорда','Последний результат','guestKey'];

function _gameSecret() {
  const properties = PropertiesService.getScriptProperties();
  let secret = properties.getProperty('WEDDING_GAME_SECRET');
  if (!secret) {
    // Создание вызывается только из setupRsvp или doPost под ScriptLock.
    secret = Utilities.getUuid() + Utilities.getUuid();
    properties.setProperty('WEDDING_GAME_SECRET', secret);
  }
  return secret;
}
function _signGame(data) {
  const encoded = Utilities.base64EncodeWebSafe(JSON.stringify(data), Utilities.Charset.UTF_8).replace(/=+$/, '');
  const signature = Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(encoded, _gameSecret(), Utilities.Charset.UTF_8)).replace(/=+$/, '');
  return encoded + '.' + signature;
}
function _readGameToken(token, type) {
  if (typeof token !== 'string' || token.length > 4000) _fail('GAME_ACCESS', 'Откройте игру после отправки анкеты.');
  const parts = token.split('.');
  if (parts.length !== 2 || !parts.every(part => /^[a-zA-Z0-9_-]+$/.test(part))) _fail('GAME_ACCESS', 'Некорректный доступ к игре.');
  const expected = Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(parts[0], _gameSecret(), Utilities.Charset.UTF_8)).replace(/=+$/, '');
  if (expected.length !== parts[1].length) _fail('GAME_ACCESS', 'Некорректный доступ к игре.');
  let diff = 0;
  for (let i = 0; i < expected.length; i++) diff |= expected.charCodeAt(i) ^ parts[1].charCodeAt(i);
  if (diff) _fail('GAME_ACCESS', 'Некорректный доступ к игре.');
  let data;
  try { data = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0] + '='.repeat((4 - parts[0].length % 4) % 4))).getDataAsString('UTF-8')); }
  catch (_) { _fail('GAME_ACCESS', 'Некорректный доступ к игре.'); }
  if (!data || data.type !== type || !Number.isFinite(data.expires) || data.expires < Date.now()) _fail('GAME_EXPIRED', 'Доступ к игре истёк. Отправьте анкету ещё раз.');
  return data;
}
function _gameAccess(data) {
  if (!data.willAttend || data.wishes === null) return null;
  const guestKey = _hash(_norm(data.lastName) + '\n' + _norm(data.firstName));
  return {firstName:data.firstName, lastName:data.lastName, guestKey,
    token:_signGame({type:'guest', guestKey, firstName:data.firstName, lastName:data.lastName, expires:Date.now()+30*86400000})};
}
function _gameSheet(name, headers) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getMaxColumns() < headers.length) sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length-sheet.getMaxColumns());
  if (!sheet.getLastRow()) sheet.getRange(1,1,1,headers.length).setValues([headers]);
  else {
    const actual = sheet.getRange(1,1,1,headers.length).getValues()[0];
    if (headers.some((header,i) => actual[i] !== header)) _fail('GAME_HEADERS', 'Проверьте заголовки листа '+name+'.');
  }
  return sheet;
}
function _gameRows(sheet) {
  return sheet.getLastRow() > 1 ? sheet.getRange(2,1,sheet.getLastRow()-1,GAME_ATTEMPT_HEADERS.length).getValues() : [];
}
function _gameRanking(rows) {
  const guests = new Map();
  rows.forEach(row => {
    const key = row[2], score = Number(row[5]);
    if (!key || !Number.isInteger(score) || score < 0) _fail('GAME_DATA', 'Повреждена строка GameAttempts.');
    let guest = guests.get(key);
    if (!guest) {
      guest = {key,firstName:row[3],lastName:row[4],best:score,bestAt:row[0],attempts:0,last:score};
      guests.set(key,guest);
    }
    guest.attempts++; guest.last=score;
    if (score > guest.best) {guest.best=score;guest.bestAt=row[0];}
  });
  const sorted = [...guests.values()].sort((a,b) => b.best-a.best || String(a.bestAt).localeCompare(String(b.bestAt)) || a.key.localeCompare(b.key));
  let rank=0, previous=-1;
  sorted.forEach((guest,i) => {if(guest.best!==previous) rank=i+1;guest.rank=rank;previous=guest.best;});
  const sheet = _gameSheet('GameLeaderboard',GAME_RANK_HEADERS);
  const oldRows = Math.max(0,sheet.getLastRow()-1);
  if (sorted.length+1 > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(),sorted.length+1-sheet.getMaxRows());
  if (sorted.length) sheet.getRange(2,1,sorted.length,GAME_RANK_HEADERS.length).setValues(sorted.map(g => [g.rank,_cell(g.firstName),_cell(g.lastName),g.best,g.attempts,g.bestAt,g.last,g.key]));
  if (oldRows > sorted.length) sheet.getRange(sorted.length+2,1,oldRows-sorted.length,GAME_RANK_HEADERS.length).clearContent();
  return sorted;
}
function _gameSpawnInterval(version, score) {
  if (version === LEGACY_GAME_VERSION) return .62;
  if (version !== GAME_VERSION) _fail('GAME_VERSION','Обновите страницу и обработчик игры.');
  const level = Math.min(12, Math.floor(Math.max(0, score) / 5));
  return Math.max(.18, .5 - level * .032);
}
function _gameAction(input) {
  if (input.action === 'startGame') {
    const guest = _readGameToken(input.gameToken,'guest');
    // Прежняя страница не передавала версию; её раунды продолжают работать.
    const gameVersion = input.gameVersion === undefined ? LEGACY_GAME_VERSION : input.gameVersion;
    _gameSpawnInterval(gameVersion, 0);
    const startedAt=Date.now(),runId=Utilities.getUuid();
    return {ok:true,gameVersion,lives:GAME_LIVES,runId,
      runToken:_signGame({type:'round',guestKey:guest.guestKey,firstName:guest.firstName,lastName:guest.lastName,
        startedAt,runId,gameVersion,expires:startedAt+7*86400000})};
  }
  if (input.action !== 'saveGameScore') _fail('UNKNOWN_ACTION','Неизвестное действие.');
  const round = _readGameToken(input.runToken,'round');
  if (input.gameVersion !== round.gameVersion) _fail('GAME_VERSION','Обновите страницу игры.');
  const gameVersion = round.gameVersion;
  _gameSpawnInterval(gameVersion, 0);
  for (const key of ['score','flowersCaught','ringsCaught','missed']) {
    if (!Number.isInteger(input[key]) || input[key] < 0 || input[key] > 100000) _fail('GAME_SCORE','Некорректный результат игры.');
  }
  if (input.missed !== GAME_LIVES || input.score !== input.flowersCaught+3*input.ringsCaught) _fail('GAME_SCORE','Очки не совпадают с результатом раунда.');
  const seconds=input.activeSeconds;
  if (!Number.isFinite(seconds) || seconds < 0 || seconds > 21600 || seconds > (Date.now()-round.startedAt)/1000+3) _fail('GAME_TIME','Некорректная длительность раунда.');
  // Частота растёт только вместе с очками. Частота на итоговом счёте задаёт
  // верхнюю границу количества предметов; прежние раунды сохраняют свой лимит.
  const fastestInterval = _gameSpawnInterval(gameVersion, input.score);
  if (input.flowersCaught+input.ringsCaught+input.missed > Math.floor(seconds/fastestInterval)+1) _fail('GAME_SCORE','Количество предметов не соответствует раунду.');
  const canonical={runId:round.runId,guestKey:round.guestKey,score:input.score,flowersCaught:input.flowersCaught,
    ringsCaught:input.ringsCaught,missed:input.missed,activeSeconds:seconds,gameVersion};
  const hash=_hash(JSON.stringify(canonical));
  const sheet=_gameSheet('GameAttempts',GAME_ATTEMPT_HEADERS);
  let rows=_gameRows(sheet);
  const existing=rows.find(row=>row[1]===round.runId);
  if (existing && existing[11]!==hash) _fail('GAME_CONFLICT','Этот раунд уже записан с другим результатом.');
  if (!existing) {
    const target=sheet.getLastRow()+1;
    if(target>sheet.getMaxRows())sheet.insertRowsAfter(sheet.getMaxRows(),target-sheet.getMaxRows());
    const row=[Utilities.formatDate(new Date(),TIMEZONE,"yyyy-MM-dd'T'HH:mm:ssXXX"),round.runId,round.guestKey,
      _cell(round.firstName),_cell(round.lastName),input.score,input.flowersCaught,input.ringsCaught,input.missed,seconds,gameVersion,hash];
    sheet.getRange(target,1,1,row.length).setValues([row]);
    SpreadsheetApp.flush();
    rows=_gameRows(sheet);
  }
  // При повторе рейтинг тоже пересобирается: восстанавливает сбой после записи попытки.
  const own=_gameRanking(rows).find(guest=>guest.key===round.guestKey);
  SpreadsheetApp.flush();
  return {ok:true,scoreSaved:true,duplicate:!!existing,runId:round.runId,bestScore:own.best,rank:own.rank,attempts:own.attempts};
}
