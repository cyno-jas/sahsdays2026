const CONFIG = {
  SPREADSHEET_ID: "1-eka8sQdmjXqASk8mZBOA6VkgEzoxTY6dx_xkmF8M3g",
  TIME_ZONE: "Asia/Manila",
  FOOD_STUB_CUTOFF: "06:45",

  SHEETS: {
    PARTICIPANTS: "Participants",
    PARTICIPANT_LOG: "Participant Attendance Log",
    PARTICIPANT_SUMMARY: "Participant Summary",
    COMMITTEE: "Committee",
    COMMITTEE_LOG: "Committee Attendance Log",
    FOOD_STUB: "Food Stub Monitoring",
    DASHBOARD: "Dashboard"
  },

  PARTICIPANT_CHECKPOINTS: [
    "Morning In",
    "Morning Out",
    "Afternoon In",
    "Afternoon Out"
  ],

  COMMITTEE_CHECKPOINTS: [
    "Committee In",
    "Committee Out"
  ]
};

const HEADERS = {
  "Participants": ["ID", "Full Name", "Program", "Section"],
  "Participant Attendance Log": [
    "Timestamp", "Date", "ID", "Full Name",
    "Program", "Section", "Checkpoint", "Status"
  ],
  "Participant Summary": [
    "ID", "Full Name", "Program", "Section",
    "Morning In", "Morning Out", "Afternoon In", "Afternoon Out"
  ],
  "Committee": ["ID", "Full Name", "Program", "Section"],
  "Committee Attendance Log": [
    "Timestamp", "Date", "ID", "Full Name",
    "Program", "Section", "Checkpoint",
    "Food Stub Eligible", "Status"
  ],
  "Food Stub Monitoring": [
    "Date", "ID", "Full Name", "Program",
    "Section", "First Scan Time", "Eligible", "Stub Issued"
  ],
  "Dashboard": ["Metric", "Count"]
};

/**
 * Run this function once from the Apps Script editor.
 * It creates missing tabs and adds headers to empty tabs.
 */
function setupSheets() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  Object.keys(HEADERS).forEach(function (name) {
    let sheet = ss.getSheetByName(name);

    if (!sheet) {
      sheet = ss.insertSheet(name);
    }

    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, HEADERS[name].length)
        .setValues([HEADERS[name]]);
      sheet.setFrozenRows(1);
    }
  });

  refreshSummary();
  refreshDashboard();

  return "Setup complete.";
}

/**
 * GitHub Pages sends scan requests to this web app.
 * The request body must be JSON with action: "recordScan".
 */
function doPost(e) {
  try {
    const body = e && e.postData && e.postData.contents
      ? JSON.parse(e.postData.contents)
      : {};

    if (body.action !== "recordScan") {
      return jsonResponse({
        success: false,
        message: "Unknown action."
      });
    }

    const result = recordScan(body);

    return jsonResponse(result);
  } catch (error) {
    return jsonResponse({
      success: false,
      message: error.message || String(error)
    });
  }
}

/**
 * Optional browser test endpoint.
 */
function doGet() {
  return jsonResponse({
    success: true,
    message: "SAHS Days 2026 attendance backend is running."
  });
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function recordScan(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const id = normalizeId(payload.id);
    const checkpoint = String(payload.checkpoint || "").trim();

    if (!id) {
      return { success: false, message: "Missing ID." };
    }

    const isCommittee =
      CONFIG.COMMITTEE_CHECKPOINTS.indexOf(checkpoint) !== -1;

    const isParticipant =
      CONFIG.PARTICIPANT_CHECKPOINTS.indexOf(checkpoint) !== -1;

    if (!isCommittee && !isParticipant) {
      return { success: false, message: "Invalid checkpoint." };
    }

    const sourceName = isCommittee
      ? CONFIG.SHEETS.COMMITTEE
      : CONFIG.SHEETS.PARTICIPANTS;

    const logName = isCommittee
      ? CONFIG.SHEETS.COMMITTEE_LOG
      : CONFIG.SHEETS.PARTICIPANT_LOG;

    const sourceSheet = ss.getSheetByName(sourceName);
    const logSheet = ss.getSheetByName(logName);

    if (!sourceSheet || !logSheet) {
      return {
        success: false,
        message: "Required sheet not found. Run setupSheets()."
      };
    }

    const person = findPerson(sourceSheet, id);

    if (!person) {
      return {
        success: false,
        message: "ID not found in " + sourceName + "."
      };
    }

    const now = new Date();
    const date = Utilities.formatDate(
      now, CONFIG.TIME_ZONE, "yyyy-MM-dd"
    );
    const time = Utilities.formatDate(
      now, CONFIG.TIME_ZONE, "HH:mm:ss"
    );

    // Prevent the same ID from being recorded twice
    // at the same checkpoint on the same date.
    if (isDuplicate(logSheet, date, id, checkpoint)) {
      return {
        success: false,
        duplicate: true,
        message: "Already scanned for " + checkpoint + "."
      };
    }

    let eligible = "";

    if (isCommittee && checkpoint === "Committee In") {
      eligible = time <= CONFIG.FOOD_STUB_CUTOFF + ":00"
        ? "Yes"
        : "No";
    }

    if (isCommittee) {
      logSheet.appendRow([
        now, date, person.id, person.name,
        person.program, person.section,
        checkpoint, eligible, "Recorded"
      ]);

      if (checkpoint === "Committee In") {
        updateFoodStub(ss, date, person, time, eligible);
      }
    } else {
      logSheet.appendRow([
        now, date, person.id, person.name,
        person.program, person.section,
        checkpoint, "Recorded"
      ]);
    }

    refreshSummary();
    refreshDashboard();

    return {
      success: true,
      duplicate: false,
      message: "Scan recorded.",
      id: person.id,
      name: person.name,
      checkpoint: checkpoint,
      date: date,
      time: time,
      foodStubEligible: eligible
    };

  } finally {
    lock.releaseLock();
  }
}

function normalizeId(value) {
  return String(value == null ? "" : value).trim();
}

/**
 * Finds a person by ID using the first column.
 * IDs are compared as text to preserve leading zeros.
 */
function findPerson(sheet, id) {
  const values = sheet.getDataRange().getDisplayValues();

  if (values.length < 2) return null;

  for (let r = 1; r < values.length; r++) {
    if (normalizeId(values[r][0]) === id) {
      return {
        id: normalizeId(values[r][0]),
        name: values[r][1] || "",
        program: values[r][2] || "",
        section: values[r][3] || ""
      };
    }
  }

  return null;
}

function isDuplicate(logSheet, date, id, checkpoint) {
  const values = logSheet.getDataRange().getDisplayValues();

  // Log columns: Date = B, ID = C, Checkpoint = G.
  for (let r = 1; r < values.length; r++) {
    if (
      values[r][1] === date &&
      normalizeId(values[r][2]) === id &&
      values[r][6] === checkpoint
    ) {
      return true;
    }
  }

  return false;
}

function updateFoodStub(ss, date, person, time, eligible) {
  const sheet = ss.getSheetByName(CONFIG.SHEETS.FOOD_STUB);
  const values = sheet.getDataRange().getDisplayValues();

  // One food-stub monitoring row per committee ID per date.
  for (let r = 1; r < values.length; r++) {
    if (
      values[r][0] === date &&
      normalizeId(values[r][1]) === person.id
    ) {
      return;
    }
  }

  sheet.appendRow([
    date,
    person.id,
    person.name,
    person.program,
    person.section,
    time,
    eligible,
    "No"
  ]);
}

/**
 * Rebuilds the participant summary using the attendance log.
 * A checkpoint is marked with its first recorded scan time.
 */
function refreshSummary() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const participants = ss.getSheetByName(CONFIG.SHEETS.PARTICIPANTS);
  const log = ss.getSheetByName(CONFIG.SHEETS.PARTICIPANT_LOG);
  const summary = ss.getSheetByName(CONFIG.SHEETS.PARTICIPANT_SUMMARY);

  if (!participants || !log || !summary) return;

  const people = participants.getDataRange().getDisplayValues();
  const logs = log.getDataRange().getValues();

  const checkpoints = CONFIG.PARTICIPANT_CHECKPOINTS;
  const byId = {};

  for (let r = 1; r < logs.length; r++) {
    const row = logs[r];
    const id = normalizeId(row[2]);
    const checkpoint = String(row[6] || "");
    const timestamp = row[0];

    if (!id || checkpoints.indexOf(checkpoint) === -1) continue;

    const key = id + "|" + checkpoint;

    if (!byId[key] && timestamp) {
      byId[key] = Utilities.formatDate(
        new Date(timestamp), CONFIG.TIME_ZONE, "HH:mm:ss"
      );
    }
  }

  const output = [[
    "ID", "Full Name", "Program", "Section",
    "Morning In", "Morning Out", "Afternoon In", "Afternoon Out"
  ]];

  for (let r = 1; r < people.length; r++) {
    const p = people[r];
    const id = normalizeId(p[0]);

    if (!id) continue;

    output.push([
      id, p[1] || "", p[2] || "", p[3] || "",
      byId[id + "|Morning In"] || "",
      byId[id + "|Morning Out"] || "",
      byId[id + "|Afternoon In"] || "",
      byId[id + "|Afternoon Out"] || ""
    ]);
  }

  summary.clearContents();
  summary.getRange(1, 1, output.length, output[0].length)
    .setValues(output);
  summary.setFrozenRows(1);
}

function refreshDashboard() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const participants = ss.getSheetByName(CONFIG.SHEETS.PARTICIPANTS);
  const committee = ss.getSheetByName(CONFIG.SHEETS.COMMITTEE);
  const participantLog = ss.getSheetByName(CONFIG.SHEETS.PARTICIPANT_LOG);
  const committeeLog = ss.getSheetByName(CONFIG.SHEETS.COMMITTEE_LOG);
  const dashboard = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD);

  if (!dashboard) return;

  const countPeople = function (sheet) {
    return sheet ? Math.max(0, sheet.getLastRow() - 1) : 0;
  };

  const rows = [
    ["Metric", "Count"],
    ["Total Participants", countPeople(participants)],
    ["Total Working Committee", countPeople(committee)],
    ["Participant Scan Records", countPeople(participantLog)],
    ["Committee Scan Records", countPeople(committeeLog)],
    ["Last Updated", Utilities.formatDate(
      new Date(), CONFIG.TIME_ZONE, "yyyy-MM-dd HH:mm:ss"
    )]
  ];

  dashboard.clearContents();
  dashboard.getRange(1, 1, rows.length, 2).setValues(rows);
  dashboard.setFrozenRows(1);
}
