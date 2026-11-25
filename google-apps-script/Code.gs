/**
 * THERMO LAB — Google Sheets receiver
 * 1) Run setup() once to create the spreadsheet.
 * 2) Deploy as Web App: Execute as Me; Who has access: Anyone.
 * 3) Paste the deployed /exec URL into GOOGLE_SHEETS_WEB_APP_URL in index.html.
 */

const CONFIG = {
  spreadsheetName: 'Теплова лабораторія — результати учнів',
  resultsSheet: 'Результати',
  answersSheet: 'Відповіді',
  experimentsSheet: 'Досліди'
};

function setup() {
  const props = PropertiesService.getScriptProperties();
  let spreadsheetId = props.getProperty('THERMO_LAB_SPREADSHEET_ID');
  let spreadsheet;

  if (spreadsheetId) {
    try {
      spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    } catch (error) {
      spreadsheetId = '';
    }
  }

  if (!spreadsheetId) {
    spreadsheet = SpreadsheetApp.create(CONFIG.spreadsheetName);
    spreadsheetId = spreadsheet.getId();
    props.setProperty('THERMO_LAB_SPREADSHEET_ID', spreadsheetId);
  }

  spreadsheet.setSpreadsheetTimeZone('Europe/Kyiv');
  ensureSheet_(spreadsheet, CONFIG.resultsSheet, [
    'Час надсилання', 'ID звіту', 'Прізвище та ім’я', 'Клас',
    'Лабораторна робота', 'Оцінка', 'Максимум балів', 'Статус',
    'Відповіді з теорії (JSON)', 'Розрахунки учня (JSON)',
    'Висновок', 'Контрольна відповідь', 'Контрольна відповідь правильна',
    'Спроби (JSON)'
  ]);
  ensureSheet_(spreadsheet, CONFIG.answersSheet, [
    'Час надсилання', 'ID звіту', 'Учень', 'Клас', 'Лабораторна',
    'Етап', 'Запитання', 'Відповідь учня', 'Правильна відповідь',
    'Правильно', 'Кількість спроб', 'Пояснення'
  ]);
  ensureSheet_(spreadsheet, CONFIG.experimentsSheet, [
    'Час надсилання', 'ID звіту', 'Учень', 'Клас', 'Лабораторна',
    '№ досліду', 'Тип досліду / речовина', 'Початкові параметри (JSON)',
    'Результати вимірювання (JSON)', 'Розрахунок учня',
    'Розрахунок правильний'
  ]);

  Logger.log('Таблицю створено або вже налаштовано: ' + spreadsheet.getUrl());
  Logger.log('Скопіюй URL таблиці з цього журналу. Не роби таблицю загальнодоступною.');
  return spreadsheet.getUrl();
}

function doGet() {
  return jsonOutput_({ ok: true, service: 'Thermo Lab Sheets receiver', message: 'Endpoint is ready for POST submissions.' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('Порожній запит.');
    }

    const payload = JSON.parse(e.postData.contents);
    validatePayload_(payload);

    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('THERMO_LAB_SPREADSHEET_ID');
    if (!spreadsheetId) {
      throw new Error('Спочатку запусти setup().');
    }

    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const results = spreadsheet.getSheetByName(CONFIG.resultsSheet);
    const answers = spreadsheet.getSheetByName(CONFIG.answersSheet);
    const experiments = spreadsheet.getSheetByName(CONFIG.experimentsSheet);

    // Захист від повторного запису того самого завершення лабораторної.
    const lastRow = results.getLastRow();
    if (lastRow > 1) {
      const duplicate = results.getRange(2, 2, lastRow - 1, 1)
        .createTextFinder(payload.submissionId)
        .matchEntireCell(true)
        .findNext();
      if (duplicate) {
        return jsonOutput_({ ok: true, duplicate: true, message: 'Цей звіт уже збережено.' });
      }
    }

    const student = payload.student;
    const mission = payload.mission;
    const timestamp = safeCell_(payload.submittedAt || new Date().toISOString());
    const labTitle = safeCell_(mission.title);

    results.appendRow([
      timestamp,
      safeCell_(payload.submissionId),
      safeCell_(student.name),
      safeCell_(student.className || ''),
      labTitle,
      numeric_(mission.score),
      numeric_(mission.maxScore || 12),
      safeCell_(mission.status || 'completed'),
      jsonCell_(mission.theoryAnswers || []),
      jsonCell_(mission.calculationAnswers || []),
      safeCell_(mission.conclusion || ''),
      safeCell_(mission.controlAnswer || ''),
      mission.controlCorrect === true ? 'Так' : 'Ні',
      jsonCell_(mission.attempts || {})
    ]);

    (mission.theoryAnswers || []).forEach(function(answer, index) {
      appendAnswer_(answers, timestamp, payload, labTitle, 'Теорія · ' + (index + 1), answer);
    });

    (mission.calculationAnswers || []).forEach(function(answer, index) {
      appendAnswer_(answers, timestamp, payload, labTitle, 'Розрахунок · ' + (index + 1), answer);
    });

    if (mission.conclusion) {
      appendAnswer_(answers, timestamp, payload, labTitle, 'Висновок', {
        question: 'Поясніть результати експерименту власними словами.',
        selectedAnswer: mission.conclusion,
        correctAnswer: '',
        isCorrect: '',
        attempts: ''
      });
    }

    if (mission.controlAnswer) {
      appendAnswer_(answers, timestamp, payload, labTitle, 'Контрольне запитання', {
        question: 'Контрольне запитання наприкінці місії',
        selectedAnswer: mission.controlAnswer,
        correctAnswer: '',
        isCorrect: mission.controlCorrect,
        attempts: mission.controlAttempts || 0
      });
    }

    (mission.experiments || []).forEach(function(experiment, index) {
      const isWater = mission.id === 1;
      const initial = isWater
        ? { coldWaterMass_g: experiment.mc, coldWaterTemp_C: experiment.tc, hotWaterMass_g: experiment.mh, hotWaterTemp_C: experiment.th }
        : { metal: experiment.name, metalMass_g: experiment.mm, metalTemp_C: experiment.tm, waterMass_g: experiment.mw, waterTemp_C: experiment.tw };
      const measured = isWater
        ? { mixtureTemp_C: experiment.t, Qcold_J: experiment.qc, Qhot_J: experiment.qh }
        : { equilibriumTemp_C: experiment.t, calculatedSpecificHeat_J_kgC: experiment.calc, studentAnswer: experiment.studentAnswer, studentCorrect: experiment.studentCorrect, referenceSpecificHeat_J_kgC: experiment.known };

      experiments.appendRow([
        timestamp,
        safeCell_(payload.submissionId),
        safeCell_(student.name),
        safeCell_(student.className || ''),
        labTitle,
        index + 1,
        safeCell_(isWater ? 'Змішування води' : (experiment.name || 'Металевий зразок')),
        jsonCell_(initial),
        jsonCell_(measured),
        isWater ? '' : safeCell_(experiment.studentAnswer == null ? '' : experiment.studentAnswer),
        isWater ? '' : (experiment.studentCorrect === true ? 'Так' : 'Ні')
      ]);
    });

    return jsonOutput_({ ok: true, duplicate: false, message: 'Звіт збережено.' });
  } catch (error) {
    return jsonOutput_({ ok: false, error: String(error && error.message ? error.message : error) });
  } finally {
    lock.releaseLock();
  }
}

function appendAnswer_(sheet, timestamp, payload, labTitle, section, answer) {
  sheet.appendRow([
    timestamp,
    safeCell_(payload.submissionId),
    safeCell_(payload.student.name),
    safeCell_(payload.student.className || ''),
    labTitle,
    safeCell_(section),
    safeCell_(answer.question || ''),
    safeCell_(answer.selectedAnswer == null ? '' : answer.selectedAnswer),
    safeCell_(answer.correctAnswer == null ? '' : answer.correctAnswer),
    answer.isCorrect === true ? 'Так' : (answer.isCorrect === false ? 'Ні' : ''),
    answer.attempts == null ? '' : numeric_(answer.attempts),
    safeCell_(answer.explanation || '')
  ]);
}

function validatePayload_(payload) {
  if (!payload || payload.source !== 'thermo-lab' || payload.schemaVersion !== 1) {
    throw new Error('Непідтримуваний формат даних.');
  }
  if (!payload.submissionId || String(payload.submissionId).length > 150) {
    throw new Error('Некоректний ID звіту.');
  }
  if (!payload.student || !String(payload.student.name || '').trim() || String(payload.student.name).length > 120) {
    throw new Error('Не вказано ім’я учня.');
  }
  if (String(payload.student.className || '').length > 40) {
    throw new Error('Некоректне поле класу.');
  }
  if (!payload.mission || [1, 2].indexOf(Number(payload.mission.id)) === -1) {
    throw new Error('Некоректний номер лабораторної.');
  }
  if (!Array.isArray(payload.mission.experiments) || payload.mission.experiments.length > 3) {
    throw new Error('Некоректний список дослідів.');
  }
}

function ensureSheet_(spreadsheet, name, headers) {
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(name);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#d9eaf7');
    sheet.autoResizeColumns(1, headers.length);
  }
  return sheet;
}

function safeCell_(value) {
  let text = String(value == null ? '' : value).replace(/\u0000/g, '');
  if (text.length > 45000) text = text.slice(0, 45000);
  // Текстові значення не повинні трактуватися як формули Google Sheets.
  if (/^[=+\-@]/.test(text)) text = "'" + text;
  return text;
}

function jsonCell_(value) {
  return safeCell_(JSON.stringify(value));
}

function numeric_(value) {
  const number = Number(value);
  return Number.isFinite(number) ? number : '';
}

function jsonOutput_(value) {
  return ContentService.createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}
