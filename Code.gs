/**
 * SAHS Days 2026 QR Attendance System
 * Bind this Apps Script project to the new Google Spreadsheet:
 * Extensions > Apps Script.
 *
 * IMPORTANT:
 * - Run setupSheets() once before deploying.
 * - Set the spreadsheet timezone to Asia/Manila.
 * - Deploy as a Web App, executing as the owner.
 */

const CONFIG = {
  EVENT_DATES: ['2026-10-19', '2026-10-20'],
  CUTOFF: '06:45',
  TIMEZONE: 'Asia/Manila',
  SHEETS: {
    PARTICIPANTS: 'Participants',
    PARTICIPANT_LOG: 'Participant Attendance Log',
    PARTICIPANT_SUMMARY: 'Participant Summary',
    COMMITTEE: 'Committee',
    COMMITTEE_LOG: 'Committee Attendance Log',
    FOOD_STUB: 'Food Stub Monitoring',
    DASHBOARD: 'Dashboard'
  },
  HEADERS: {
    'Participants': ['ID Number', 'Full Name', 'Program', 'Section', 'Team'],
    'Participant Attendance Log': ['Timestamp', 'Event Date', 'ID Number', 'Full Name', 'Program', 'Section', 'Team', 'Checkpoint', 'Station ID', 'Status'],
    'Participant Summary': ['Event Date', 'ID Number', 'Full Name', 'Program', 'Section', 'Team', 'Morning In', 'Morning Out', 'Afternoon In', 'Afternoon Out'],
    'Committee': ['ID Number', 'Full Name', 'Program', 'Section', 'Committee', 'Position'],
    'Committee Attendance Log': ['Timestamp', 'Event Date', 'ID Number', 'Full Name', 'Program', 'Section', 'Committee', 'Position', 'Checkpoint', 'Station ID', 'Status'],
    'Food Stub Monitoring': ['Event Date', 'ID Number', 'Full Name', 'Program', 'Section', 'Committee', 'Arrival Time', 'Cutoff Time', 'Eligibility', 'Stub Claimed', 'Claim Timestamp', 'Distributed By'],
    'Dashboard': ['Metric', 'Value']
  }
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('SAHS Days 2026 QR Attendance')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open Apps Script from the new attendance spreadsheet (Extensions > Apps Script).');
  ss.setSpreadsheetTimeZone(CONFIG.TIMEZONE);

  Object.keys(CONFIG.HEADERS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const headers = CONFIG.HEADERS[name];
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setWrap(true);
    if (name === CONFIG.SHEETS.PARTICIPANTS || name === CONFIG.SHEETS.COMMITTEE) {
      sh.getRange('A:A').setNumberFormat('@');
    }
  });

  const dashboard = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD);
  dashboard.getRange(2, 1, 8, 2).setValues([
    ['Event dates', CONFIG.EVENT_DATES.join(', ')],
    ['Participant records', '=COUNTA(Participants!A2:A)'],
    ['Committee records', '=COUNTA(Committee!A2:A)'],
    ['Participant scans recorded', '=COUNTIF(\'Participant Attendance Log\'!J2:J,"RECORDED")'],
    ['Committee scans recorded', '=COUNTIF(\'Committee Attendance Log\'!K2:K,"RECORDED")'],
    ['Food stubs marked claimed', '=COUNTIF(\'Food Stub Monitoring\'!J2:J,"YES")'],
    ['Committee eligible by cutoff', '=COUNTIF(\'Food Stub Monitoring\'!I2:I,"ELIGIBLE")'],
    ['Committee late', '=COUNTIF(\'Food Stub Monitoring\'!I2:I,"LATE")']
  ]);
  dashboard.autoResizeColumns(1, 2);
  return 'Setup complete. Import your participant and committee lists under the headers.';
}

function getAppConfig() {
  return {
    eventDates: CONFIG.EVENT_DATES,
    cutoff: CONFIG.CUTOFF,
    checkpoints: ['Morning In', 'Morning Out', 'Afternoon In', 'Afternoon Out'],
    committeeCheckpoints: ['Committee In', 'Committee Out']
  };
}

/**
 * Called by the web app after a QR scan.
 * payload: {id, eventDate, group: 'participant'|'committee',
 *           checkpoint, stationId}
 */
function recordScan(payload) {
  validatePayload_(payload);
  const id = normalizeId_(payload.id);
  const eventDate = String(payload.eventDate);
  const group = payload.group === 'committee' ? 'committee' : 'participant';
  const checkpoint = String(payload.checkpoint);
  const stationId = String(payload.stationId || 'UNASSIGNED').trim().slice(0, 80);
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(15000)) {
    throw new Error('The system is busy processing scans. Please wait a moment and scan again.');
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const masterName = group === 'committee' ? CONFIG.SHEETS.COMMITTEE : CONFIG.SHEETS.PARTICIPANTS;
    const logName = group === 'committee' ? CONFIG.SHEETS.COMMITTEE_LOG : CONFIG.SHEETS.PARTICIPANT_LOG;
    const master = ss.getSheetByName(masterName);
    const log = ss.getSheetByName(logName);
    const person = findPerson_(master, id, group);
    if (!person) {
      return {ok: false, type: 'not_found', message: 'ID not found in the ' + (group === 'committee' ? 'Committee' : 'Participants') + ' list.', id: id};
    }

    const lastRow = log.getLastRow();
    if (lastRow >= 2) {
      const idCol = group === 'committee' ? 3 : 3;
      const dateCol = 2;
      const checkpointCol = group === 'committee' ? 9 : 8;
      const existing = log.getRange(2, 1, lastRow - 1, log.getLastColumn()).getValues();
      const duplicate = existing.some(row =>
        String(row[dateCol - 1]) === eventDate &&
        normalizeId_(row[idCol - 1]) === id &&
        String(row[checkpointCol - 1]) === checkpoint &&
        String(row[row.length - 1]) === 'RECORDED'
      );
      if (duplicate) {
        return {ok: true, duplicate: true, message: 'Already recorded for ' + checkpoint + '.', person: person, timestamp: ''};
      }
    }

    const now = new Date();
    const timestamp = Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
    let row;
    if (group === 'committee') {
      row = [now, eventDate, id, person.name, person.program, person.section, person.committee, person.position, checkpoint, stationId, 'RECORDED'];
    } else {
      row = [now, eventDate, id, person.name, person.program, person.section, person.team, checkpoint, stationId, 'RECORDED'];
    }
    log.appendRow(row);

    let foodStub = null;
    if (group === 'committee' && checkpoint === 'Committee In') {
      foodStub = updateFoodStub_(ss, eventDate, id, person, now);
    }

    return {
      ok: true,
      duplicate: false,
      message: 'Attendance recorded.',
      person: person,
      timestamp: timestamp,
      foodStub: foodStub
    };
  } finally {
    lock.releaseLock();
  }
}

function validatePayload_(p) {
  if (!p || !p.id) throw new Error('No ID number was read. Please scan the QR code again.');
  if (CONFIG.EVENT_DATES.indexOf(String(p.eventDate)) === -1) throw new Error('Select a valid event date.');
  const group = p.group === 'committee' ? 'committee' : 'participant';
  const allowed = group === 'committee'
    ? ['Committee In', 'Committee Out']
    : ['Morning In', 'Morning Out', 'Afternoon In', 'Afternoon Out'];
  if (allowed.indexOf(String(p.checkpoint)) === -1) throw new Error('Select a valid attendance checkpoint.');
}

function normalizeId_(value) {
  return String(value == null ? '' : value).trim().replace(/\s+/g, '');
}

function findPerson_(sheet, id, group) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  // TextFinder avoids reading the entire master list on every scan.
  const cell = sheet.getRange(2, 1, lastRow - 1, 1)
    .createTextFinder(id)
    .matchEntireCell(true)
    .matchCase(true)
    .findNext();
  if (!cell) return null;
  const values = sheet.getRange(cell.getRow(), 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  if (normalizeId_(values[0]) !== id) return null;
  if (group === 'committee') {
    return {
      id: id, name: values[1] || '', program: values[2] || '',
      section: values[3] || '', committee: values[4] || '', position: values[5] || ''
    };
  }
  return {
    id: id, name: values[1] || '', program: values[2] || '',
    section: values[3] || '', team: values[4] || ''
  };
}

/**
 * Refresh the Participant Summary tab from the logs.
 * Run manually after a test or event day; avoid running during peak scanning.
 */
function refreshParticipantSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName(CONFIG.SHEETS.PARTICIPANTS);
  const log = ss.getSheetByName(CONFIG.SHEETS.PARTICIPANT_LOG);
  const summary = ss.getSheetByName(CONFIG.SHEETS.PARTICIPANT_SUMMARY);
  const people = master.getLastRow() >= 2
    ? master.getRange(2, 1, master.getLastRow() - 1, 5).getDisplayValues()
    : [];
  const logs = log.getLastRow() >= 2
    ? log.getRange(2, 1, log.getLastRow() - 1, 10).getValues()
    : [];
  const seen = {};
  logs.forEach(row => {
    const date = String(row[1]);
    const id = normalizeId_(row[2]);
    const checkpoint = String(row[7]);
    if (String(row[9]) === 'RECORDED') seen[date + '|' + id + '|' + checkpoint] = true;
  });
  const output = [];
  CONFIG.EVENT_DATES.forEach(date => {
    people.forEach(person => {
      const id = normalizeId_(person[0]);
      output.push([
        date, id, person[1], person[2], person[3], person[4],
        seen[date + '|' + id + '|Morning In'] ? 'PRESENT' : '',
        seen[date + '|' + id + '|Morning Out'] ? 'PRESENT' : '',
        seen[date + '|' + id + '|Afternoon In'] ? 'PRESENT' : '',
        seen[date + '|' + id + '|Afternoon Out'] ? 'PRESENT' : ''
      ]);
    });
  });
  const existingRows = summary.getLastRow();
  if (existingRows > 1) summary.getRange(2, 1, existingRows - 1, summary.getLastColumn()).clearContent();
  if (output.length) summary.getRange(2, 1, output.length, 10).setValues(output);
  return 'Participant Summary refreshed: ' + output.length + ' participant-day rows.';
}

function updateFoodStub_(ss, eventDate, id, person, now) {
  const sh = ss.getSheetByName(CONFIG.SHEETS.FOOD_STUB);
  const last = sh.getLastRow();
  const cutoffParts = CONFIG.CUTOFF.split(':').map(Number);
  const cutoff = new Date(now);
  cutoff.setHours(cutoffParts[0], cutoffParts[1], 0, 0);
  const eligible = now.getTime() <= cutoff.getTime();
  let targetRow = 0;
  if (last >= 2) {
    const data = sh.getRange(2, 1, last - 1, 12).getDisplayValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === eventDate && normalizeId_(data[i][1]) === id) {
        targetRow = i + 2;
        break;
      }
    }
  }
  const arrivalText = Utilities.formatDate(now, CONFIG.TIMEZONE, 'HH:mm:ss');
  const values = [
    eventDate, id, person.name, person.program, person.section, person.committee,
    arrivalText, CONFIG.CUTOFF, eligible ? 'ELIGIBLE' : 'LATE',
    targetRow ? (sh.getRange(targetRow, 10).getDisplayValue() || 'NO') : 'NO',
    targetRow ? sh.getRange(targetRow, 11).getValue() : '',
    targetRow ? sh.getRange(targetRow, 12).getDisplayValue() : ''
  ];
  if (targetRow) {
    // Preserve the first recorded arrival and its eligibility.
    return {
      eligibility: sh.getRange(targetRow, 9).getDisplayValue(),
      arrivalTime: sh.getRange(targetRow, 7).getDisplayValue(),
      cutoff: CONFIG.CUTOFF
    };
  }
  sh.appendRow(values);
  return {eligibility: values[8], arrivalTime: arrivalText, cutoff: CONFIG.CUTOFF};
}

/** Run from Apps Script editor to inspect whether sheet setup worked. */
function testConfiguration() {
  return {
    spreadsheet: SpreadsheetApp.getActiveSpreadsheet().getName(),
    timezone: SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(),
    config: getAppConfig()
  };
}
