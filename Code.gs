
/**
 * SAHS DAYS 2026 — QR ATTENDANCE BACKEND
 * Google Sheets + GitHub Pages JSONP scanner
 */

const CONFIG = {
  SPREADSHEET_ID: "1-eka8sQdmjXqASk8mZBOA6VkgEzoxTY6dx_xkmF8M3g",

  PARTICIPANTS_SHEET: "Participants",
  PARTICIPANT_LOG: "Participant Attendance Log",

  COMMITTEE_SHEET: "Committee",
  COMMITTEE_LOG: "Committee Attendance Log",

  TIMEZONE: "Asia/Manila",
  FOOD_STUB_CUTOFF: "06:45"
};

const PARTICIPANT_LOG_HEADERS = [
  "Timestamp", "Event Date", "ID Number", "Full Name",
  "Program", "Section", "Team", "Checkpoint",
  "Station ID", "Status"
];

const COMMITTEE_LOG_HEADERS = [
  "Timestamp", "Event Date", "ID Number", "Full Name",
  "Program", "Section", "Committee", "Position",
  "Checkpoint", "Food Stub Eligible", "Status"
];

/**
 * Handles requests from GitHub Pages.
 * Supports:
 *   ?action=ping&callback=callbackName
 *   ?action=recordScan&id=123&checkpoint=Morning%20In
 *       &group=Participant&callback=callbackName
 */
function doGet(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  const callback = params.callback || "";

  let result;

  try {
    const action = params.action || "ping";

    if (action === "ping") {
      result = {
        success: true,
        message: "SAHS Days 2026 attendance server is online."
      };
    } else if (action === "recordScan") {
      result = recordAttendance_(params);
    } else {
      result = {
        success: false,
        message: "Unknown action: " + action
      };
    }
  } catch (error) {
    result = {
      success: false,
      message: error && error.message
        ? error.message
        : "An unexpected server error occurred."
    };
  }

  const json = JSON.stringify(result);

  // JSONP allows the GitHub Pages site to receive the response
  // without a browser CORS preflight.
  if (callback && /^[a-zA-Z_$][0-9a-zA-Z_$.[\]]*$/.test(callback)) {
    return ContentService
      .createTextOutput(callback + "(" + json + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function recordAttendance_(params) {
  const id = normalizeId_(params.id || params.idNumber || "");
  const checkpoint = String(params.checkpoint || "").trim();
  const group = normalizeGroup_(params.group || "Participant");

  if (!id) {
    return { success: false, message: "Please provide an ID number." };
  }

  if (!checkpoint) {
    return { success: false, message: "Please select an attendance checkpoint." };
  }

  const isCommittee = group === "Committee";
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  const sourceSheet = ss.getSheetByName(
    isCommittee ? CONFIG.COMMITTEE_SHEET : CONFIG.PARTICIPANTS_SHEET
  );

  const logSheet = ss.getSheetByName(
    isCommittee ? CONFIG.COMMITTEE_LOG : CONFIG.PARTICIPANT_LOG
  );

  if (!sourceSheet) {
    throw new Error(
      "Sheet not found: " +
      (isCommittee ? CONFIG.COMMITTEE_SHEET : CONFIG.PARTICIPANTS_SHEET)
    );
  }

  if (!logSheet) {
    throw new Error(
      "Sheet not found: " +
      (isCommittee ? CONFIG.COMMITTEE_LOG : CONFIG.PARTICIPANT_LOG)
    );
  }

  const person = findPerson_(sourceSheet, id, isCommittee);

  if (!person) {
    return {
      success: false,
      message: "ID number " + id + " was not found in the " +
        (isCommittee ? "Committee" : "Participants") + " sheet."
    };
  }

  const validCheckpoints = isCommittee
    ? ["Committee In", "Committee Out"]
    : ["Morning In", "Morning Out", "Afternoon In", "Afternoon Out"];

  if (!validCheckpoints.includes(checkpoint)) {
    return {
      success: false,
      message: "Invalid checkpoint: " + checkpoint
    };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const now = new Date();
    const timestamp = Utilities.formatDate(
      now, CONFIG.TIMEZONE, "yyyy-MM-dd HH:mm:ss"
    );
    const eventDate = Utilities.formatDate(
      now, CONFIG.TIMEZONE, "yyyy-MM-dd"
    );

    // Check the log for an existing scan for this ID, date and checkpoint.
    const lastRow = logSheet.getLastRow();

    if (lastRow >= 2) {
      const existingRows = logSheet
        .getRange(2, 1, lastRow - 1, isCommittee ? 11 : 10)
        .getDisplayValues();

      const duplicate = existingRows.some(row =>
        normalizeId_(row[2]) === id &&
        String(row[1]).trim() === eventDate &&
        String(row[8]).trim() === checkpoint
      );

      if (duplicate) {
        return {
          success: false,
          duplicate: true,
          message: person.name + " has already been recorded for " +
            checkpoint + " today.",
          person: person
        };
      }
    }

    let row;

    if (isCommittee) {
      const currentTime = Utilities.formatDate(now, CONFIG.TIMEZONE, "HH:mm");
      const eligible = checkpoint === "Committee In"
        ? currentTime <= CONFIG.FOOD_STUB_CUTOFF
        : "";

      row = [
        timestamp, eventDate, person.id, person.name,
        person.program, person.section, person.committee,
        person.position, checkpoint, eligible, "Recorded"
      ];
    } else {
      row = [
        timestamp, eventDate, person.id, person.name,
        person.program, person.section, person.team,
        checkpoint, params.stationId || "", "Recorded"
      ];
    }

    // Write the row in one operation.
    logSheet.appendRow(row);

    return {
      success: true,
      duplicate: false,
      message: "Attendance recorded successfully.",
      timestamp: timestamp,
      eventDate: eventDate,
      checkpoint: checkpoint,
      group: isCommittee ? "Committee" : "Participant",
      foodStubEligible: isCommittee && checkpoint === "Committee In"
        ? (Utilities.formatDate(now, CONFIG.TIMEZONE, "HH:mm") <=
           CONFIG.FOOD_STUB_CUTOFF)
        : null,
      person: person
    };

  } finally {
    lock.releaseLock();
  }
}

/**
 * Finds an ID in the source sheet.
 * Expected Participants columns:
 * A ID Number | B Full Name | C Program | D Section | E Team
 *
 * Expected Committee columns:
 * A ID Number | B Full Name | C Program | D Section
 * E Committee | F Position
 */
function findPerson_(sheet, id, isCommittee) {
  const values = sheet.getDataRange().getDisplayValues();

  if (values.length < 2) return null;

  const headers = values[0].map(h => String(h).trim().toLowerCase());

  const idCol = findHeader_(headers, ["id number", "id", "student id"]);
  const nameCol = findHeader_(headers, ["full name", "name"]);
  const programCol = findHeader_(headers, ["program"]);
  const sectionCol = findHeader_(headers, ["section"]);
  const teamCol = findHeader_(headers, ["team"]);
  const committeeCol = findHeader_(headers, ["committee"]);
  const positionCol = findHeader_(headers, ["position"]);

  if (idCol === -1 || nameCol === -1) {
    throw new Error(
      "The " + sheet.getName() +
      " sheet must have ID Number and Full Name headers in row 1."
    );
  }

  for (let r = 1; r < values.length; r++) {
    const row = values[r];

    if (normalizeId_(row[idCol]) === id) {
      const person = {
        id: normalizeId_(row[idCol]),
        name: row[nameCol] || "",
        program: programCol >= 0 ? row[programCol] : "",
        section: sectionCol >= 0 ? row[sectionCol] : ""
      };

      if (isCommittee) {
        person.committee = committeeCol >= 0 ? row[committeeCol] : "";
        person.position = positionCol >= 0 ? row[positionCol] : "";
      } else {
        person.team = teamCol >= 0 ? row[teamCol] : "";
      }

      return person;
    }
  }

  return null;
}

function findHeader_(headers, candidates) {
  for (let i = 0; i < candidates.length; i++) {
    const index = headers.indexOf(candidates[i]);
    if (index !== -1) return index;
  }
  return -1;
}

function normalizeId_(value) {
  return String(value == null ? "" : value).trim();
}

function normalizeGroup_(value) {
  const group = String(value || "").trim().toLowerCase();

  if (
    group === "committee" ||
    group === "working committee" ||
    group === "workingcommittee"
  ) {
    return "Committee";
  }

  return "Participant";
}

/**
 * Optional: run this from the Apps Script editor to check
 * that the spreadsheet and required sheets are accessible.
 */
function testConnection() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  const requiredSheets = [
    CONFIG.PARTICIPANTS_SHEET,
    CONFIG.PARTICIPANT_LOG,
    CONFIG.COMMITTEE_SHEET,
    CONFIG.COMMITTEE_LOG
  ];

  requiredSheets.forEach(name => {
    if (!ss.getSheetByName(name)) {
      throw new Error("Missing sheet: " + name);
    }
  });

  Logger.log("Connection successful: " + ss.getName());
  return "Connection successful: " + ss.getName();
}
