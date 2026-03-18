// ═══════════════════════════════════════════════════════════════════════════
//  AWS SIS — Google Apps Script v8.0
//  American World School · Student Information System
//
//  WHAT'S NEW IN v8 (over v7):
//  · Lower School Grades & Progress Reports (Pre-K, K, Grades 1–5)
//  · Middle School Grades (Grades 6–8) — letter grades, no HS credits
//  · Lower School Skills Audit — per-student skill mastery tracking
//  · Campus Partner Role — login & data scoped to one campus
//  · Transcript fixes: Subject Area column, Transfer Credit GPA & Points
//  · Skills Mastery, Narratives, Portfolio Notes storage
//  · All 37 sheet tabs created by setupSpreadsheet() (EXISTING DATA SAFE)
//
//  ╔══════════════════════════════════════════════════════════════════════╗
//  ║  UPDATING FROM v7 — EXISTING DATA IS NEVER TOUCHED:                 ║
//  ║  • setupSpreadsheet() ONLY creates tabs that don't exist yet         ║
//  ║  • Headers are only written to brand-new empty sheets                ║
//  ║  • Existing student / grade / user rows are never modified           ║
//  ║  • New columns added to Users sheet header without clearing data     ║
//  ║  • objectsToSheet() always receives the full array from the SIS      ║
//  ║    (localStorage is the source of truth; Sheet = cloud backup)       ║
//  ╚══════════════════════════════════════════════════════════════════════╝
//
//  SETUP (fresh install):
//  1. Create a new Google Sheet
//  2. Extensions → Apps Script → paste this file → Save
//  3. Run setupSpreadsheet() once
//  4. Run createTriggers() once (daily session cleanup)
//  5. Deploy → New deployment → Web app
//     Execute as: Me  |  Who has access: Anyone
//  6. Copy the Web App URL into AWS SIS Settings
//
//  UPDATING FROM v7 (existing sheet):
//  1. Replace old script with this file → Save
//  2. Run upgradeToV8() — adds new tabs, fixes Users header, never touches data
//  3. Re-deploy (New deployment OR edit existing deployment)
//
//  DEFAULT LOGINS (fresh install only):
//  admin           / admin123     (Admin)
//  staff           / staff123     (Staff)
//  coach           / coach123     (Staff)
//  partner_srilanka/ partner123   (Partner — Sri Lanka only)
//  partner_uae     / partner456   (Partner — UAE only)
//  partner_spain   / partner789   (Partner — Spain only)
// ═══════════════════════════════════════════════════════════════════════════

var SESSION_HOURS = 8;
var _SS = null; // Lazy-loaded spreadsheet

function getSpreadsheet() {
  if (!_SS) _SS = SpreadsheetApp.getActiveSpreadsheet();
  return _SS;
}

// ── Sheet tab names ──────────────────────────────────────────────────────────
var SHEET = {
  STUDENTS      : "Students",
  COURSES       : "Courses",
  TRANSFER      : "Transfer",
  ATTENDANCE    : "Attendance",
  INTERVIEWS    : "Interviews",
  FEES          : "Fees",
  COMMS         : "Communications",
  SETTINGS      : "Settings",
  STAFF         : "Staff",
  CATALOG       : "Catalog",
  REMARKS       : "ReportRemarks",
  CALENDAR      : "Calendar",
  LESSONS       : "Lessons",
  UNITS         : "Units",
  EVENTS        : "TPMSEvents",
  PD            : "PD",
  BLOCKS        : "Blocks",
  HEALTH        : "HealthRecords",
  BEHAVIOUR     : "BehaviourLog",
  AT_ASSIGN     : "AT_Assignments",
  AT_SUBS       : "AT_Submissions",
  AT_ASSESS     : "AT_Assessments",
  AT_EP         : "AT_ExactPath",
  AT_NOTES      : "AT_Notes",
  AT_REPORTS    : "AT_Reports",
  PT_ASSIGN     : "PT_Assignments",
  PT_EVALS      : "PT_Evaluations",
  USERS         : "Users",
  SESSIONS      : "Sessions",
  // ── v8 NEW ──────────────────────────────────────────────────────────────
  ELEM_PROGRESS : "ElemProgress",
  MS_GRADES     : "MSGrades",
  ELEM_NARRATIVE: "ElemNarrative",
  ELEM_PORTFOLIO: "ElemPortfolio",
  SKILL_MASTERY : "SkillMastery"
};

// ── Column definitions ───────────────────────────────────────────────────────
var STUDENT_COLS = [
  "id","studentId","firstName","lastName","dob","gender","grade","status",
  "appDate","enrollDate","email","phone","nationality","lang",
  "parent","relation","ecName","ecPhone",
  "address","bloodGroup","allergy","meds","physician","physicianPhone",
  "healthNotes","iep","prevSchool","gpa","cohort",
  "documents","notes","counselorNotes","priority","campus",
  "studentType","yearJoined","gradeJoined","yearGraduated",
  "postSecondary","gradDistinction","alumniNotes"
];
var COURSE_COLS    = ["studentId","id","title","area","type","year","semester","creditsAttempted","creditsEarned","grade","courseStatus","apScore","instructor","section","notes"];
var TRANSFER_COLS  = ["studentId","id","origTitle","sourceSchool","location","accred","origGrade","creditsAwarded","area","type","gradeLevel","status","year","notes"];
var ATT_COLS       = ["attKey","studentId","date","status","note"];
var INT_COLS       = ["id","studentId","date","time","type","interviewer","result","notes","createdAt"];
var FEE_COLS       = ["id","studentId","type","amount","currency","dueDate","paidDate","status","reference","notes"];
var COMM_COLS      = ["id","studentId","date","channel","direction","subject","body","sentBy","tags"];
var STAFF_COLS     = ["id","staffId","firstName","lastName","email","phone","role","department","campus","staffType","joinDate","status","notes","tpmsId"];
var CATALOG_COLS   = ["id","code","title","area","type","credits","description","grade","prerequisites","status"];
var REMARK_COLS    = ["studentId","term","remark","updatedAt"];
var CAL_COLS       = ["id","title","type","date","endDate","time","location","description","campus","createdAt"];
var LESSON_COLS    = ["id","title","subject","grade","date","lessonNum","unitId","status","objectives","materials","activities","assessment","notes","createdAt","teacherId"];
var UNIT_COLS      = ["id","title","subject","grade","year","semester","weeks","standards","description","status","createdAt","teacherId"];
var EVENT_COLS     = ["id","title","type","date","endDate","time","location","description","campus","createdAt"];
var PD_COLS        = ["id","title","type","date","duration","facilitator","attendees","description","outcome","cost","createdAt"];
var BLOCK_COLS     = ["id","name","day","startTime","endTime","subject","teacher","room","cohort","type","color","notes"];
var HEALTH_COLS    = ["studentId","blood","allergy","meds","conditions","vision","vaccines","ec1","ec2","doctor","hospital","iepType","diet","notes","updatedAt"];
var BEH_COLS       = ["id","studentId","type","date","time","location","description","action","followUp","recordedBy","severity","status","parentNotified","notes"];
var AT_ASSIGN_COLS = ["id","title","type","subject","division","maxScore","dateAssigned","dueDate","instructions","status","createdAt"];
var AT_SUBS_COLS   = ["subKey","assignId","studentId","score","status","submittedDate","teacherNote","penaltyApplied","penaltyWaived","penaltyPoints"];
var AT_ASSESS_COLS = ["id","studentId","weekStart","subject","rawScore","maxScore","status","assessmentDate","enteredAt","feedback"];
var AT_EP_COLS     = ["id","studentId","weekStart","subject","targetType","targetValue","actualValue","met","trophies"];
var AT_NOTES_COLS  = ["id","studentId","type","subject","noteText","date","visibility","createdAt"];
var AT_RPT_COLS    = ["id","stuId","type","week","coachNote","generatedAt","generatedBy"];
var PT_ASSIGN_COLS = ["id","sid","mn","mname","ptype","title","brief","status","q","due","deliv","team","cnotes","reflect","wurl","assignedAt","submittedAt","score","mastery","yr"];
var PT_EVAL_COLS   = ["id","aid","sid","ms","cs","ov","mastery","crJson","coJson","comment","date"];
// v8: added "active" column to Users (8 cols total)
var USER_COLS      = ["username","password","role","name","email","campus","createdAt","active","sidPrefix"];
var SESSION_COLS   = ["token","username","role","name","campus","createdAt","expiresAt","sidPrefix"];
// v8 NEW sheets
var ELEM_PROGRESS_COLS  = ["studentId","grade","term","subjectsJson","updatedAt"];
var MS_GRADES_COLS      = ["studentId","grade","term","subjectsJson","updatedAt"];
var ELEM_NARRATIVE_COLS = ["key","studentId","term","general","strengths","growth","goals","parentMsg","updatedAt"];
var ELEM_PORTFOLIO_COLS = ["key","studentId","term","entriesJson","updatedAt"];
var SKILL_MASTERY_COLS  = ["studentId","grade","masteryJson","updatedAt"];


// ════════════════════════════════════════════════════════════════════════════
//  UTILITY HELPERS
// ════════════════════════════════════════════════════════════════════════════
function getSheet(name) {
  return getSpreadsheet().getSheetByName(name);
}

function sheetToObjects(sheetName, cols) {
  var sh = getSheet(sheetName);
  if (!sh) return [];
  var data = sh.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).map(function(row) {
    var obj = {};
    cols.forEach(function(col, i) { obj[col] = (row[i] !== undefined && row[i] !== null) ? row[i] : ""; });
    return obj;
  });
}

// ── Safe bulk-replace: only writes if objects array is non-empty.
// This prevents accidental data wipe if the SIS sends an empty array.
function objectsToSheet(sheetName, cols, objects) {
  if (!objects || objects.length === 0) {
    // Nothing to write — but ensure sheet + header exist
    _ensureSheetHeader(sheetName, cols);
    return;
  }
  var sh = getSheet(sheetName);
  if (!sh) sh = getSpreadsheet().insertSheet(sheetName);
  sh.clearContents();
  var rows = [cols];
  objects.forEach(function(obj) {
    rows.push(cols.map(function(col) {
      var v = obj[col];
      if (v === null || v === undefined) return "";
      if (typeof v === "object") return JSON.stringify(v);
      return v;
    }));
  });
  sh.getRange(1, 1, rows.length, cols.length).setValues(rows);
  styleHeader(sh, cols.length);
}

// ── Force-clear a sheet (explicit wipe — only used when intentionally clearing)
function clearSheet(sheetName, cols) {
  var sh = getSheet(sheetName);
  if (!sh) sh = getSpreadsheet().insertSheet(sheetName);
  sh.clearContents();
  sh.getRange(1, 1, 1, cols.length).setValues([cols]);
  styleHeader(sh, cols.length);
}

// ── Create sheet with header ONLY if it doesn't already exist with data
function _ensureSheetHeader(sheetName, cols) {
  var sh = getSheet(sheetName);
  if (!sh) {
    sh = getSpreadsheet().insertSheet(sheetName);
    sh.getRange(1, 1, 1, cols.length).setValues([cols]);
    styleHeader(sh, cols.length);
    sh.setFrozenRows(1);
  }
  // If sheet exists but is completely empty (no header), write header
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, cols.length).setValues([cols]);
    styleHeader(sh, cols.length);
    sh.setFrozenRows(1);
  }
  return sh;
}

// ── Add a new column header to an existing sheet WITHOUT touching data rows
// Used when upgrading: e.g. adding "active" column to Users sheet
function _addColumnIfMissing(sheetName, cols) {
  var sh = getSheet(sheetName);
  if (!sh) return;
  var lastRow = sh.getLastRow();
  if (lastRow === 0) {
    // Empty sheet — write full header
    sh.getRange(1, 1, 1, cols.length).setValues([cols]);
    styleHeader(sh, cols.length);
    sh.setFrozenRows(1);
    return;
  }
  // Read existing header
  var existingCols = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var changed = false;
  cols.forEach(function(col, i) {
    if (existingCols.indexOf(col) === -1) {
      // Column missing — append it to the header row
      var newColIndex = existingCols.length + 1;
      sh.getRange(1, newColIndex).setValue(col);
      existingCols.push(col);
      changed = true;
    }
  });
  if (changed) styleHeader(sh, sh.getLastColumn());
}

function styleHeader(sh, numCols) {
  if (!sh || numCols === 0) return;
  var hRange = sh.getRange(1, 1, 1, numCols);
  hRange.setBackground("#1A365E").setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(10);
  sh.setFrozenRows(1);
}

function resp(data) {
  return ContentService.createTextOutput(JSON.stringify(
    Object.assign({ status: "ok" }, data)
  )).setMimeType(ContentService.MimeType.JSON);
}

function respErr(msg) {
  return ContentService.createTextOutput(JSON.stringify({ status: "error", message: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonOk(data) { return resp(data); }
function jsonErr(msg)  { return respErr(msg); }


// ════════════════════════════════════════════════════════════════════════════
//  SESSION HELPERS
// ════════════════════════════════════════════════════════════════════════════
function createSession(user) {
  var token = Utilities.getUuid();
  var now   = new Date();
  var exp   = new Date(now.getTime() + SESSION_HOURS * 3600000);
  var sh    = getSheet(SHEET.SESSIONS);
  if (sh) sh.appendRow([
    token, user.username, user.role || "", user.name || "", user.campus || "",
    now.toISOString(), exp.toISOString(), user.sidPrefix || ""
  ]);
  return token;
}

function validateSession(token) {
  if (!token) return null;
  var sh = getSheet(SHEET.SESSIONS);
  if (!sh) return null;
  var data = sh.getDataRange().getValues();
  var now  = new Date();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === token && new Date(data[i][6]) > now) {
      return {
        token     : token,
        username  : String(data[i][1]),
        role      : String(data[i][2]),
        name      : String(data[i][3]),
        campus    : String(data[i][4]),
        sidPrefix : data[i][7] !== undefined ? String(data[i][7]) : ""
      };
    }
  }
  return null;
}

function deleteSession(token) {
  var sh = getSheet(SHEET.SESSIONS);
  if (!sh) return;
  var data = sh.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === token) { sh.deleteRow(i + 1); break; }
  }
}

// ── Campus access: partners are scoped to their assigned campus ──────────────
function campusAllowed(sess, studentCampus) {
  if (!sess) return false;
  if ((sess.role || "").toLowerCase() !== "partner") return true;
  var pc = (sess.campus || "").trim().toLowerCase();
  return !!pc && pc === (studentCampus || "").trim().toLowerCase();
}

function filterByCampus(students, sess) {
  if (!sess || (sess.role || "").toLowerCase() !== "partner") return students;
  var pc = (sess.campus || "").trim().toLowerCase();
  if (!pc) return [];
  return students.filter(function(s) {
    return (s.campus || "").trim().toLowerCase() === pc;
  });
}

function _getAllowedStudentIds(sess) {
  var pc = (sess.campus || "").trim().toLowerCase();
  if (!pc) return [];
  return sheetToObjects(SHEET.STUDENTS, STUDENT_COLS)
    .filter(function(s) { return (s.campus || "").trim().toLowerCase() === pc; })
    .map(function(s) { return String(s.id); });
}


// ════════════════════════════════════════════════════════════════════════════
//  HTTP ROUTER
// ════════════════════════════════════════════════════════════════════════════
function doGet(e) {
  try {
    var params = (e && e.parameter) ? e.parameter : {};
    var action = params.action || "";
    var token  = params.token  || "";

    // Start with direct query params
    var p = { action: action, token: token };

    // 1. Try to parse JSON payload param
    if (params.payload) {
      var raw = params.payload;
      var parsed = null;
      // Try both with and without extra decoding
      var tries = [raw];
      try { tries.push(decodeURIComponent(raw)); } catch(de) {}
      for (var i = 0; i < tries.length; i++) {
        try { parsed = JSON.parse(tries[i]); if (parsed) break; } catch(je) {}
      }
      if (parsed && typeof parsed === "object") {
        Object.keys(parsed).forEach(function(k) { p[k] = parsed[k]; });
      }
    }

    // 2. Direct params override (u/pw sent as explicit query params for reliability)
    if (params.u)    p.username = decodeURIComponent(params.u);
    if (params.pw)   p.password = decodeURIComponent(params.pw);
    if (params.fast) p.fast     = params.fast;

    // 3. Ensure action and token always set
    p.action = p.action || action;
    p.token  = p.token  || token;

    return _handleAction(p);
  } catch(err) {
    return respErr("doGet error: " + err.message);
  }
}

function doPost(e) {
  try {
    var p = JSON.parse(e.postData.contents);
    return _handleAction(p);
  } catch(err) {
    return respErr("doPost error: " + err.message);
  }
}

// ── Shared action handler — called by both doGet and doPost ─────────────────
function _handleAction(p) {
  try {
    var action = p.action;

    // ── Public actions ───────────────────────────────────────────────────────
    if (action === "ping") {
      return resp({ version: "8.0", timestamp: new Date().toISOString() });
    }

    // ── DEBUG — visit ?action=debugLogin in browser to see Users sheet ──────
    // REMOVE THIS after fixing login issues
    if (action === "debugLogin") {
      var sh = getSheet(SHEET.USERS);
      if (!sh) return resp({ error: "Users sheet not found" });
      var data = sh.getDataRange().getValues();
      if (data.length === 0) return resp({ error: "Sheet is empty" });
      var header = data[0];
      var rows = [];
      for (var ri = 1; ri < data.length; ri++) {
        var rowObj = {};
        header.forEach(function(h, i) {
          var val = data[ri][i];
          rowObj[String(h)] = {
            value: String(val),
            type: typeof val,
            length: String(val).length
          };
        });
        rows.push(rowObj);
      }
      return resp({
        header: header,
        headerLower: header.map(function(h){ return String(h).trim().toLowerCase(); }),
        rowCount: data.length - 1,
        rows: rows
      });
    }

    if (action === "login") {
      var inUser = String(p.username || "").trim().toLowerCase();
      var inPass = String(p.password || "").trim();
      if (!inUser || !inPass) return respErr("Username and password are required.");

      var u = null;

      // ── 1. Try PropertiesService cache first (fastest — no sheet read) ──
      try {
        var cacheStr = PropertiesService.getScriptProperties().getProperty("aws_users_cache");
        if (cacheStr) {
          var cacheObj = JSON.parse(cacheStr);
          var cu = cacheObj[inUser];
          if (cu && String(cu.pw || "").trim() === inPass) {
            var activeOk = (!cu.active || cu.active === "true" || cu.active === "TRUE");
            if (activeOk) {
              u = { username: inUser, role: cu.role||"staff", name: cu.name||"",
                    email: cu.email||"", campus: cu.campus||"", active: cu.active||"" };
            }
          }
        }
      } catch(ce) {}

      // ── 2. Fall back to Sheet read if not cached ──────────────────────
      if (!u) {
        var sh = getSheet(SHEET.USERS);
        if (!sh) return respErr("Users sheet not found. Run setupSpreadsheet() first.");
        var data = sh.getDataRange().getValues();
        if (data.length <= 1) return respErr("No user accounts found. Run setupSpreadsheet() first.");

        var header = data[0].map(function(h){ return String(h).trim().toLowerCase(); });
        var ci = {};
        header.forEach(function(h,i){ ci[h]=i; });

        if (ci["password"] === undefined)
          return respErr("Password column missing. Run upgradeToV8() in Apps Script.");

        for (var ri = 1; ri < data.length; ri++) {
          var row = data[ri];
          var rUser = String(row[ci["username"]]||"").trim().toLowerCase();
          var rPass = String(row[ci["password"]]||"").trim();
          var rAct  = ci["active"]!==undefined ? String(row[ci["active"]]||"").trim() : "";
          var aOk   = (rAct===""||rAct.toLowerCase()==="true"||rAct===true);
          if (rUser===inUser && rPass===inPass && aOk) {
            u = {
              username : String(row[ci["username"]]||""),
              role     : String(ci["role"]!==undefined   ? row[ci["role"]]   : "staff"),
              name     : String(ci["name"]!==undefined   ? row[ci["name"]]   : ""),
              email    : String(ci["email"]!==undefined  ? row[ci["email"]]  : ""),
              campus   : String(ci["campus"]!==undefined ? row[ci["campus"]] : ""),
              active   : rAct
            };
            break;
          }
        }

        // Rebuild cache for next time
        try {
          var newCache = {};
          for (var ri2 = 1; ri2 < data.length; ri2++) {
            var r2 = data[ri2];
            var uName = String(r2[ci["username"]]||"").trim().toLowerCase();
            if (uName) newCache[uName] = {
              pw     : String(r2[ci["password"]]||"").trim(),
              role   : String(ci["role"]!==undefined   ? r2[ci["role"]]   : "staff"),
              name   : String(ci["name"]!==undefined   ? r2[ci["name"]]   : ""),
              email  : String(ci["email"]!==undefined  ? r2[ci["email"]]  : ""),
              campus : String(ci["campus"]!==undefined ? r2[ci["campus"]] : ""),
              active : ci["active"]!==undefined ? String(r2[ci["active"]]||"").trim() : ""
            };
          }
          PropertiesService.getScriptProperties().setProperty("aws_users_cache", JSON.stringify(newCache));
        } catch(ce2) {}
      }

      if (!u) return respErr("Invalid credentials or account inactive.");
      var token = createSession(u);
      return resp({
        token : token,
        user  : {
          username  : u.username,
          role      : u.role,
          name      : u.name || u.username,
          campus    : u.campus    || "",
          email     : u.email     || "",
          sidPrefix : u.sidPrefix || ""
        }
      });
    }


    if (action === "logout") {
      deleteSession(p.token);
      return resp({});
    }

    // ── Auth required ────────────────────────────────────────────────────────
    var sess = validateSession(p.token);
    if (!sess) return respErr("Session expired — please log in again.");

    var isAdmin   = ["admin","principal"].indexOf((sess.role||"").toLowerCase()) >= 0;
    var isPartner = (sess.role||"").toLowerCase() === "partner";

    // ── getAll ───────────────────────────────────────────────────────────────
    if (action === "getAll") {
      var allStudents = _loadStudents();
      var students    = filterByCampus(allStudents, sess);
      var settings    = _loadSettings();

      // ── FAST MODE: students + settings only (used by login for speed) ───
      if (p.fast === "1" || p.fast === 1) {
        return resp({
          user     : { username: sess.username, role: sess.role, name: sess.name, campus: sess.campus },
          students : students,
          settings : settings
        });
      }

      // ── FULL MODE: all 20 data tables ────────────────────────────────────
      // Allowed student IDs for partner campus filtering
      var allowedIds  = isPartner ? _getAllowedStudentIds(sess) : null;

      function campusFilter(arr, idField) {
        if (!allowedIds) return arr;
        return arr.filter(function(r) { return allowedIds.indexOf(String(r[idField])) >= 0; });
      }

      // ── Load all data tables ──────────────────────────────────────────────
      var courses    = campusFilter(sheetToObjects(SHEET.COURSES,   COURSE_COLS),   "studentId");
      var transfer   = campusFilter(sheetToObjects(SHEET.TRANSFER,  TRANSFER_COLS), "studentId");
      var attendance = campusFilter(sheetToObjects(SHEET.ATTENDANCE,ATT_COLS),      "studentId");
      var interviews = campusFilter(sheetToObjects(SHEET.INTERVIEWS,INT_COLS),      "studentId");
      var fees       = campusFilter(sheetToObjects(SHEET.FEES,      FEE_COLS),      "studentId");
      var comms      = campusFilter(sheetToObjects(SHEET.COMMS,     COMM_COLS),     "studentId");
      var staff      = sheetToObjects(SHEET.STAFF,   STAFF_COLS);
      var catalog    = sheetToObjects(SHEET.CATALOG,  CATALOG_COLS);
      var remarks    = campusFilter(sheetToObjects(SHEET.REMARKS,   REMARK_COLS),   "studentId");
      var calendar   = sheetToObjects(SHEET.CALENDAR, CAL_COLS);

      // ── TPMS bundle ───────────────────────────────────────────────────────
      var tpms = {
        lessons: sheetToObjects(SHEET.LESSONS, LESSON_COLS),
        units:   sheetToObjects(SHEET.UNITS,   UNIT_COLS),
        events:  sheetToObjects(SHEET.EVENTS,  EVENT_COLS),
        pd:      sheetToObjects(SHEET.PD,      PD_COLS)
      };

      // ── Assignment Tracker ────────────────────────────────────────────────
      var atAssignments  = sheetToObjects(SHEET.AT_ASSIGN,  AT_ASSIGN_COLS);
      var atAssessments  = campusFilter(sheetToObjects(SHEET.AT_ASSESS, AT_ASSESS_COLS), "studentId");
      var atExactPath    = campusFilter(sheetToObjects(SHEET.AT_EP,     AT_EP_COLS),     "studentId");
      var atNotes        = campusFilter(sheetToObjects(SHEET.AT_NOTES,  AT_NOTES_COLS),  "studentId");
      var atReports      = campusFilter(sheetToObjects(SHEET.AT_REPORTS,AT_RPT_COLS),    "stuId");

      // AT Submissions — stored as key-value, filter by allowed studentIds
      var atSubsRaw = sheetToObjects(SHEET.AT_SUBS, AT_SUBS_COLS);
      var atSubs    = allowedIds
        ? atSubsRaw.filter(function(r){ return allowedIds.indexOf(String(r.studentId)) >= 0; })
        : atSubsRaw;

      // ── Health & Behaviour ────────────────────────────────────────────────
      var health    = campusFilter(sheetToObjects(SHEET.HEALTH,    HEALTH_COLS), "studentId");
      var behaviour = campusFilter(sheetToObjects(SHEET.BEHAVIOUR, BEH_COLS),   "studentId");

      // ── Blocks (timetable) ────────────────────────────────────────────────
      var blocks = sheetToObjects(SHEET.BLOCKS, BLOCK_COLS);

      // ── AWSC-27 Project Tracker ───────────────────────────────────────────
      var ptAssignments  = campusFilter(sheetToObjects(SHEET.PT_ASSIGN, PT_ASSIGN_COLS), "sid");
      var ptEvaluations  = campusFilter(sheetToObjects(SHEET.PT_EVALS,  PT_EVAL_COLS),   "sid");

      // ── Lower School Grades & Skills ─────────────────────────────────────
      var elemProgress = sheetToObjects(SHEET.ELEM_PROGRESS,  ELEM_PROGRESS_COLS).map(function(r) {
        try { r.subjects = JSON.parse(r.subjectsJson || "{}"); } catch(e2) { r.subjects = {}; }
        return r;
      });
      var msGrades = sheetToObjects(SHEET.MS_GRADES, MS_GRADES_COLS).map(function(r) {
        try { r.subjects = JSON.parse(r.subjectsJson || "{}"); } catch(e2) { r.subjects = {}; }
        return r;
      });
      var elemNarratives  = sheetToObjects(SHEET.ELEM_NARRATIVE, ELEM_NARRATIVE_COLS);
      var elemPortfolios  = sheetToObjects(SHEET.ELEM_PORTFOLIO, ELEM_PORTFOLIO_COLS).map(function(r) {
        try { r.entries = JSON.parse(r.entriesJson || "[]"); } catch(e2) { r.entries = []; }
        return r;
      });
      var skillMasteryRows = sheetToObjects(SHEET.SKILL_MASTERY, SKILL_MASTERY_COLS);
      var skillMasteryMap  = {};
      skillMasteryRows.forEach(function(r) {
        try { skillMasteryMap[r.studentId] = JSON.parse(r.masteryJson || "{}"); } catch(e2) { skillMasteryMap[r.studentId] = {}; }
      });

      // Apply campus filter to lower school data
      elemProgress    = campusFilter(elemProgress,   "studentId");
      msGrades        = campusFilter(msGrades,        "studentId");
      elemNarratives  = campusFilter(elemNarratives,  "studentId");
      elemPortfolios  = campusFilter(elemPortfolios,  "studentId");

      return resp({
        user           : { username: sess.username, role: sess.role, name: sess.name, campus: sess.campus, sidPrefix: sess.sidPrefix || "" },
        students       : students,
        settings       : settings,
        courses        : courses,
        transfer       : transfer,
        attendance     : attendance,
        interviews     : interviews,
        fees           : fees,
        comms          : comms,
        staff          : staff,
        catalog        : catalog,
        remarks        : remarks,
        calendar       : calendar,
        tpms           : tpms,
        health         : health,
        behaviour      : behaviour,
        blocks         : blocks,
        atAssignments  : atAssignments,
        atSubs         : atSubs,
        atAssessments  : atAssessments,
        atExactPath    : atExactPath,
        atNotes        : atNotes,
        atReports      : atReports,
        ptAssignments  : ptAssignments,
        ptEvaluations  : ptEvaluations,
        elemProgress   : elemProgress,
        msGrades       : msGrades,
        elemNarratives : elemNarratives,
        elemPortfolios : elemPortfolios,
        skillMastery   : skillMasteryMap
      });
    }

    // ════════════════════════════════════════════════════════════════════════
    //  STUDENT CRUD
    // ════════════════════════════════════════════════════════════════════════
    if (action === "add" || action === "create") {
      var s = p.student;
      if (isPartner && !campusAllowed(sess, s.campus)) {
        return respErr("Partners can only add students to their assigned campus.");
      }
      var sh = getSheet(SHEET.STUDENTS);
      if (!sh) return respErr("Students sheet not found.");
      sh.appendRow(STUDENT_COLS.map(function(col) {
        var v = s[col];
        if (col === "documents") return JSON.stringify(Array.isArray(v) ? v : []);
        return (v === undefined || v === null) ? "" : v;
      }));
      return resp({ id: s.id });
    }

    if (action === "update") {
      var s = p.student;
      if (isPartner && !campusAllowed(sess, s.campus)) {
        return respErr("Partners can only edit students from their assigned campus.");
      }
      var sh   = getSheet(SHEET.STUDENTS);
      if (!sh) return respErr("Students sheet not found.");
      var data = sh.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(s.id)) {
          sh.getRange(i + 1, 1, 1, STUDENT_COLS.length).setValues([STUDENT_COLS.map(function(col) {
            var v = s[col];
            if (col === "documents") return JSON.stringify(Array.isArray(v) ? v : []);
            return (v === undefined || v === null) ? "" : v;
          })]);
          break;
        }
      }
      return resp({});
    }

    if (action === "delete") {
      if (isPartner) return respErr("Partners cannot delete student records.");
      var sh   = getSheet(SHEET.STUDENTS);
      if (!sh) return respErr("Students sheet not found.");
      var data = sh.getDataRange().getValues();
      for (var i = data.length - 1; i >= 1; i--) {
        if (String(data[i][0]) === String(p.id)) { sh.deleteRow(i + 1); break; }
      }
      return resp({});
    }

    // ════════════════════════════════════════════════════════════════════════
    //  HS GRADES — Courses & Transfer
    // ════════════════════════════════════════════════════════════════════════
    if (action === "saveCourses") {
      var newCourses = p.courses || [];
      // If a specific studentId is provided, do a MERGE (upsert for that student only)
      // This prevents wiping other students' courses when one student's data is saved
      if (p.studentId) {
        var sid = String(p.studentId);
        var sh = _ensureSheetHeader(SHEET.COURSES, COURSE_COLS);
        var existing = sheetToObjects(SHEET.COURSES, COURSE_COLS);
        // Remove existing rows for this student, add the new ones
        var others = existing.filter(function(c) { return String(c.studentId) !== sid; });
        // Partner campus check
        if (isPartner) {
          var allowed = _getAllowedStudentIds(sess);
          if (allowed.indexOf(sid) < 0) return respErr("Access denied for this student.");
        }
        var merged = others.concat(newCourses);
        objectsToSheet(SHEET.COURSES, COURSE_COLS, merged);
      } else {
        // Full replace (admin bulk sync)
        if (isPartner) {
          var allowed = _getAllowedStudentIds(sess);
          newCourses = newCourses.filter(function(c) { return allowed.indexOf(String(c.studentId)) >= 0; });
        }
        objectsToSheet(SHEET.COURSES, COURSE_COLS, newCourses);
      }
      return resp({});
    }

    if (action === "saveTransfer") {
      var newTransfer = p.transfer || p.transfers || [];
      if (p.studentId) {
        var sid = String(p.studentId);
        if (isPartner) {
          var allowed = _getAllowedStudentIds(sess);
          if (allowed.indexOf(sid) < 0) return respErr("Access denied for this student.");
        }
        var existing = sheetToObjects(SHEET.TRANSFER, TRANSFER_COLS);
        var others = existing.filter(function(t) { return String(t.studentId) !== sid; });
        var merged = others.concat(newTransfer);
        objectsToSheet(SHEET.TRANSFER, TRANSFER_COLS, merged);
      } else {
        if (isPartner) {
          var allowed = _getAllowedStudentIds(sess);
          newTransfer = newTransfer.filter(function(t) { return allowed.indexOf(String(t.studentId)) >= 0; });
        }
        objectsToSheet(SHEET.TRANSFER, TRANSFER_COLS, newTransfer);
      }
      return resp({});
    }

    if (action === "saveECCredits") { return resp({}); }

    // ════════════════════════════════════════════════════════════════════════
    //  LOWER SCHOOL GRADES (Pre-K · K · Grades 1–5)
    // ════════════════════════════════════════════════════════════════════════
    if (action === "saveElemRecords") {
      var records = (p.records || []).map(function(r) {
        return {
          studentId   : r.studentId || "",
          grade       : r.grade || "",
          term        : r.term || "",
          subjectsJson: (typeof r.subjects === "object") ? JSON.stringify(r.subjects) : (r.subjectsJson || "{}"),
          updatedAt   : new Date().toISOString()
        };
      });
      objectsToSheet(SHEET.ELEM_PROGRESS, ELEM_PROGRESS_COLS, records);
      return resp({});
    }

    if (action === "upsertElemRecord") {
      var r   = p.record || {};
      var key = String(r.studentId) + "_" + r.term;
      var sh  = _ensureSheetHeader(SHEET.ELEM_PROGRESS, ELEM_PROGRESS_COLS);
      var row = [
        r.studentId || "", r.grade || "", r.term || "",
        JSON.stringify(r.subjects || {}),
        new Date().toISOString()
      ];
      var data = sh.getDataRange().getValues();
      var found = false;
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]) + "_" + data[i][2] === key) {
          sh.getRange(i+1, 1, 1, ELEM_PROGRESS_COLS.length).setValues([row]);
          found = true; break;
        }
      }
      if (!found) sh.appendRow(row);
      return resp({});
    }

    if (action === "getElemRecords") {
      var records = sheetToObjects(SHEET.ELEM_PROGRESS, ELEM_PROGRESS_COLS).map(function(r) {
        try { r.subjects = JSON.parse(r.subjectsJson || "{}"); } catch(e2) { r.subjects = {}; }
        return r;
      });
      if (isPartner) {
        var allowed = _getAllowedStudentIds(sess);
        records = records.filter(function(r) { return allowed.indexOf(String(r.studentId)) >= 0; });
      }
      return resp({ records: records });
    }

    // ════════════════════════════════════════════════════════════════════════
    //  MIDDLE SCHOOL GRADES (Grades 6–8)
    // ════════════════════════════════════════════════════════════════════════
    if (action === "saveMSRecords") {
      var records = (p.records || []).map(function(r) {
        return {
          studentId   : r.studentId || "",
          grade       : r.grade || "",
          term        : r.term || "",
          subjectsJson: (typeof r.subjects === "object") ? JSON.stringify(r.subjects) : (r.subjectsJson || "{}"),
          updatedAt   : new Date().toISOString()
        };
      });
      objectsToSheet(SHEET.MS_GRADES, MS_GRADES_COLS, records);
      return resp({});
    }

    if (action === "upsertMSRecord") {
      var r   = p.record || {};
      var key = String(r.studentId) + "_" + r.term;
      var sh  = _ensureSheetHeader(SHEET.MS_GRADES, MS_GRADES_COLS);
      var row = [
        r.studentId || "", r.grade || "", r.term || "",
        JSON.stringify(r.subjects || {}),
        new Date().toISOString()
      ];
      var data = sh.getDataRange().getValues();
      var found = false;
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]) + "_" + data[i][2] === key) {
          sh.getRange(i+1, 1, 1, MS_GRADES_COLS.length).setValues([row]);
          found = true; break;
        }
      }
      if (!found) sh.appendRow(row);
      return resp({});
    }

    if (action === "getMSRecords") {
      var records = sheetToObjects(SHEET.MS_GRADES, MS_GRADES_COLS).map(function(r) {
        try { r.subjects = JSON.parse(r.subjectsJson || "{}"); } catch(e2) { r.subjects = {}; }
        return r;
      });
      if (isPartner) {
        var allowed = _getAllowedStudentIds(sess);
        records = records.filter(function(r) { return allowed.indexOf(String(r.studentId)) >= 0; });
      }
      return resp({ records: records });
    }

    // ════════════════════════════════════════════════════════════════════════
    //  ELEM NARRATIVE COMMENTS
    // ════════════════════════════════════════════════════════════════════════
    if (action === "saveElemNarratives") {
      var narrs = (p.narratives || []).map(function(n) {
        return {
          key       : n.key || (String(n.studentId) + "_" + n.term),
          studentId : n.studentId || "",
          term      : n.term || "",
          general   : n.general || "",
          strengths : n.strengths || "",
          growth    : n.growth || "",
          goals     : n.goals || "",
          parentMsg : n.parentMsg || "",
          updatedAt : new Date().toISOString()
        };
      });
      objectsToSheet(SHEET.ELEM_NARRATIVE, ELEM_NARRATIVE_COLS, narrs);
      return resp({});
    }

    if (action === "upsertElemNarrative") {
      var n   = p.narrative || {};
      var key = n.key || (String(n.studentId) + "_" + n.term);
      var sh  = _ensureSheetHeader(SHEET.ELEM_NARRATIVE, ELEM_NARRATIVE_COLS);
      var row = [
        key, n.studentId || "", n.term || "",
        n.general || "", n.strengths || "", n.growth || "",
        n.goals || "", n.parentMsg || "",
        new Date().toISOString()
      ];
      var data  = sh.getDataRange().getValues();
      var found = false;
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]) === key) {
          sh.getRange(i+1, 1, 1, ELEM_NARRATIVE_COLS.length).setValues([row]);
          found = true; break;
        }
      }
      if (!found) sh.appendRow(row);
      return resp({});
    }

    if (action === "getElemNarratives") {
      var narrs = sheetToObjects(SHEET.ELEM_NARRATIVE, ELEM_NARRATIVE_COLS);
      if (isPartner) {
        var allowed = _getAllowedStudentIds(sess);
        narrs = narrs.filter(function(n) { return allowed.indexOf(String(n.studentId)) >= 0; });
      }
      return resp({ narratives: narrs });
    }

    // ════════════════════════════════════════════════════════════════════════
    //  ELEM PORTFOLIO NOTES
    // ════════════════════════════════════════════════════════════════════════
    if (action === "saveElemPortfolios") {
      var ports = (p.portfolios || []).map(function(po) {
        return {
          key        : po.key || ("port_" + String(po.studentId) + "_" + po.term),
          studentId  : po.studentId || "",
          term       : po.term || "",
          entriesJson: (typeof po.entries === "object") ? JSON.stringify(po.entries) : (po.entriesJson || "[]"),
          updatedAt  : new Date().toISOString()
        };
      });
      objectsToSheet(SHEET.ELEM_PORTFOLIO, ELEM_PORTFOLIO_COLS, ports);
      return resp({});
    }

    if (action === "upsertElemPortfolio") {
      var po  = p.portfolio || {};
      var key = po.key || ("port_" + String(po.studentId) + "_" + po.term);
      var sh  = _ensureSheetHeader(SHEET.ELEM_PORTFOLIO, ELEM_PORTFOLIO_COLS);
      var row = [
        key, po.studentId || "", po.term || "",
        JSON.stringify(po.entries || []),
        new Date().toISOString()
      ];
      var data  = sh.getDataRange().getValues();
      var found = false;
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]) === key) {
          sh.getRange(i+1, 1, 1, ELEM_PORTFOLIO_COLS.length).setValues([row]);
          found = true; break;
        }
      }
      if (!found) sh.appendRow(row);
      return resp({});
    }

    // ════════════════════════════════════════════════════════════════════════
    //  SKILLS AUDIT MASTERY
    // ════════════════════════════════════════════════════════════════════════
    if (action === "saveSkillMastery") {
      var masteryMap = p.mastery || {};
      // Only write if we have data
      if (Object.keys(masteryMap).length === 0) return resp({});
      var allStudents = sheetToObjects(SHEET.STUDENTS, STUDENT_COLS);
      var rows = Object.keys(masteryMap).map(function(sid) {
        var stu = allStudents.find(function(s) { return String(s.id) === String(sid); });
        return {
          studentId  : sid,
          grade      : stu ? stu.grade : "",
          masteryJson: JSON.stringify(masteryMap[sid] || {}),
          updatedAt  : new Date().toISOString()
        };
      });
      objectsToSheet(SHEET.SKILL_MASTERY, SKILL_MASTERY_COLS, rows);
      return resp({});
    }

    if (action === "upsertSkillMastery") {
      var sid  = String(p.studentId || "");
      if (!sid) return respErr("studentId required");
      var sh   = _ensureSheetHeader(SHEET.SKILL_MASTERY, SKILL_MASTERY_COLS);
      var allStudents = sheetToObjects(SHEET.STUDENTS, STUDENT_COLS);
      var stu  = allStudents.find(function(s) { return String(s.id) === sid; });
      var row  = [
        sid,
        stu ? stu.grade : "",
        JSON.stringify(p.mastery || {}),
        new Date().toISOString()
      ];
      var data  = sh.getDataRange().getValues();
      var found = false;
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]) === sid) {
          sh.getRange(i+1, 1, 1, SKILL_MASTERY_COLS.length).setValues([row]);
          found = true; break;
        }
      }
      if (!found) sh.appendRow(row);
      return resp({});
    }

    if (action === "getSkillMastery") {
      var rows = sheetToObjects(SHEET.SKILL_MASTERY, SKILL_MASTERY_COLS).map(function(r) {
        try { r.mastery = JSON.parse(r.masteryJson || "{}"); } catch(e2) { r.mastery = {}; }
        return r;
      });
      if (isPartner) {
        var allowed = _getAllowedStudentIds(sess);
        rows = rows.filter(function(r) { return allowed.indexOf(String(r.studentId)) >= 0; });
      }
      var masteryMap = {};
      rows.forEach(function(r) { masteryMap[r.studentId] = r.mastery; });
      return resp({ mastery: masteryMap });
    }

    // ════════════════════════════════════════════════════════════════════════
    //  ATTENDANCE
    // ════════════════════════════════════════════════════════════════════════
    if (action === "saveAttendance") {
      var records = p.records || [];
      var sh      = _ensureSheetHeader(SHEET.ATTENDANCE, ATT_COLS);
      // Merge by attKey — preserves records not in current batch
      var existing = {};
      var allData  = sh.getDataRange().getValues();
      allData.slice(1).forEach(function(row) { if (row[0]) existing[String(row[0])] = row; });
      records.forEach(function(r) {
        existing[String(r.attKey)] = [r.attKey, r.studentId||"", r.date||"", r.status||"", r.note||""];
      });
      var vals = Object.values(existing);
      if (vals.length > 0) {
        var rows = [ATT_COLS].concat(vals);
        sh.clearContents();
        sh.getRange(1, 1, rows.length, ATT_COLS.length).setValues(rows);
        styleHeader(sh, ATT_COLS.length);
      }
      return resp({});
    }

    // ════════════════════════════════════════════════════════════════════════
    //  INTERVIEWS
    // ════════════════════════════════════════════════════════════════════════
    if (action === "saveInterview") {
      var sh = _ensureSheetHeader(SHEET.INTERVIEWS, INT_COLS);
      var r  = p.interview;
      sh.appendRow(INT_COLS.map(function(c) { return r[c] || ""; }));
      return resp({});
    }
    if (action === "deleteInterview") {
      var sh   = getSheet(SHEET.INTERVIEWS);
      if (!sh) return resp({});
      var data = sh.getDataRange().getValues();
      for (var i = data.length - 1; i >= 1; i--) {
        if (String(data[i][0]) === String(p.id)) { sh.deleteRow(i + 1); break; }
      }
      return resp({});
    }

    // ════════════════════════════════════════════════════════════════════════
    //  FEES & COMMUNICATIONS
    // ════════════════════════════════════════════════════════════════════════
    if (action === "saveFees")  { objectsToSheet(SHEET.FEES,  FEE_COLS,  p.fees  || []); return resp({}); }
    if (action === "saveComms") { objectsToSheet(SHEET.COMMS, COMM_COLS, p.comms || []); return resp({}); }
    if (action === "saveFee") {
      var sh = _ensureSheetHeader(SHEET.FEES, FEE_COLS);
      sh.appendRow(FEE_COLS.map(function(c) { return (p.fee||{})[c] || ""; }));
      return resp({});
    }
    if (action === "deleteFee") {
      var sh = getSheet(SHEET.FEES); if (!sh) return resp({});
      var data = sh.getDataRange().getValues();
      for (var i = data.length - 1; i >= 1; i--) {
        if (String(data[i][0]) === String(p.id)) { sh.deleteRow(i + 1); break; }
      }
      return resp({});
    }
    if (action === "saveComm") {
      var sh = _ensureSheetHeader(SHEET.COMMS, COMM_COLS);
      sh.appendRow(COMM_COLS.map(function(c) { return (p.comm||{})[c] || ""; }));
      return resp({});
    }
    if (action === "deleteComm") {
      var sh = getSheet(SHEET.COMMS); if (!sh) return resp({});
      var data = sh.getDataRange().getValues();
      for (var i = data.length - 1; i >= 1; i--) {
        if (String(data[i][0]) === String(p.id)) { sh.deleteRow(i + 1); break; }
      }
      return resp({});
    }

    // ════════════════════════════════════════════════════════════════════════
    //  SETTINGS, STAFF, CATALOG, REMARKS
    // ════════════════════════════════════════════════════════════════════════
    if (action === "saveSettings") {
      if (!isAdmin) return respErr("Admin access required to change settings.");
      var sh   = getSheet(SHEET.SETTINGS);
      if (!sh) sh = getSpreadsheet().insertSheet(SHEET.SETTINGS);
      sh.clearContents();
      sh.appendRow(["key","value"]);
      styleHeader(sh, 2);
      var sett = p.settings || {};
      Object.keys(sett).forEach(function(k) {
        var v = sett[k];
        sh.appendRow([k, typeof v === "object" ? JSON.stringify(v) : v]);
      });
      return resp({});
    }
    if (action === "saveStaff")   { objectsToSheet(SHEET.STAFF,   STAFF_COLS,   p.staff   || []); return resp({}); }
    if (action === "saveCatalog") { objectsToSheet(SHEET.CATALOG,  CATALOG_COLS, p.catalog || []); return resp({}); }
    if (action === "saveRemarks") {
      var existing = sheetToObjects(SHEET.REMARKS, REMARK_COLS);
      if (p.studentId && p.term) {
        // Single upsert: save one student's remark for one term
        var sid = String(p.studentId);
        var term = String(p.term);
        var key = sid + "_" + term;
        var others = existing.filter(function(r) {
          return !(String(r.studentId) === sid && String(r.term) === term);
        });
        var newRow = {
          studentId: sid,
          term: term,
          remark: typeof p.remarks === "string" ? p.remarks : JSON.stringify(p.remarks || ""),
          updatedAt: new Date().toISOString()
        };
        objectsToSheet(SHEET.REMARKS, REMARK_COLS, others.concat([newRow]));
      } else if (p.studentId) {
        // Replace all terms for one student
        var sid = String(p.studentId);
        var others = existing.filter(function(r) { return String(r.studentId) !== sid; });
        objectsToSheet(SHEET.REMARKS, REMARK_COLS, others.concat(p.remarks || []));
      } else {
        // Full replace (bulk sync)
        objectsToSheet(SHEET.REMARKS, REMARK_COLS, p.remarks || []);
      }
      return resp({});
    }

    // ════════════════════════════════════════════════════════════════════════
    //  CALENDAR & TPMS
    // ════════════════════════════════════════════════════════════════════════
    if (action === "saveCalendar")    { objectsToSheet(SHEET.CALENDAR, CAL_COLS,    p.events  || []); return resp({}); }
    if (action === "saveLessons")     { objectsToSheet(SHEET.LESSONS,  LESSON_COLS, p.lessons || []); return resp({}); }
    if (action === "saveUnits")       { objectsToSheet(SHEET.UNITS,    UNIT_COLS,   p.units   || []); return resp({}); }
    if (action === "saveTPMSEvents")  { objectsToSheet(SHEET.EVENTS,   EVENT_COLS,  p.events  || []); return resp({}); }
    if (action === "savePD")          { objectsToSheet(SHEET.PD,       PD_COLS,     p.pd      || []); return resp({}); }
    if (action === "saveBlocks")      { objectsToSheet(SHEET.BLOCKS,   BLOCK_COLS,  p.blocks  || []); return resp({}); }
    if (action === "saveTPMS") {
      if (p.lessons) objectsToSheet(SHEET.LESSONS, LESSON_COLS, p.lessons);
      if (p.units)   objectsToSheet(SHEET.UNITS,   UNIT_COLS,   p.units);
      return resp({});
    }

    // ════════════════════════════════════════════════════════════════════════
    //  HEALTH & BEHAVIOUR
    // ════════════════════════════════════════════════════════════════════════
    if (action === "saveHealth") {
      if (p.studentId) {
        var sid = String(p.studentId);
        var existing = sheetToObjects(SHEET.HEALTH, HEALTH_COLS);
        var others = existing.filter(function(r) { return String(r.studentId) !== sid; });
        objectsToSheet(SHEET.HEALTH, HEALTH_COLS, others.concat(p.health || []));
      } else {
        objectsToSheet(SHEET.HEALTH, HEALTH_COLS, p.health || []);
      }
      return resp({});
    }
    if (action === "saveBehaviour") {
      if (p.studentId) {
        var sid = String(p.studentId);
        var existing = sheetToObjects(SHEET.BEHAVIOUR, BEH_COLS);
        var others = existing.filter(function(r) { return String(r.studentId) !== sid; });
        objectsToSheet(SHEET.BEHAVIOUR, BEH_COLS, others.concat(p.behaviour || []));
      } else {
        objectsToSheet(SHEET.BEHAVIOUR, BEH_COLS, p.behaviour || []);
      }
      return resp({});
    }

    // ════════════════════════════════════════════════════════════════════════
    //  ASSIGNMENT TRACKER
    // ════════════════════════════════════════════════════════════════════════
    if (action === "saveATAssignments") {
      // AT Assignments are school-wide (not per-student) - full replace is fine for admin
      // Partners only save their campus students' submissions, not assignments
      objectsToSheet(SHEET.AT_ASSIGN, AT_ASSIGN_COLS, p.assignments || []);
      return resp({});
    }
    if (action === "saveATSubmissions") {
      var subs = p.submissions || {};
      var newRows = Object.keys(subs).map(function(key) {
        var sub   = subs[key];
        var parts = key.split("_");
        return {
          subKey: key, assignId: parts[0]||"", studentId: parts[1]||"",
          score: sub.score !== undefined ? sub.score : "",
          status: sub.status||"", submittedDate: sub.submittedDate||"",
          teacherNote: sub.teacherNote||"",
          penaltyApplied: sub.penaltyApplied ? "TRUE":"FALSE",
          penaltyWaived:  sub.penaltyWaived  ? "TRUE":"FALSE",
          penaltyPoints:  sub.penaltyPoints || 0
        };
      });
      // If studentId provided, merge (replace only that student's submissions)
      if (p.studentId) {
        var sid = String(p.studentId);
        var existing = sheetToObjects(SHEET.AT_SUBS, AT_SUBS_COLS);
        var others = existing.filter(function(r) { return String(r.studentId) !== sid; });
        objectsToSheet(SHEET.AT_SUBS, AT_SUBS_COLS, others.concat(newRows));
      } else {
        objectsToSheet(SHEET.AT_SUBS, AT_SUBS_COLS, newRows);
      }
      return resp({});
    }
    if (action === "saveATAssessments") { objectsToSheet(SHEET.AT_ASSESS, AT_ASSESS_COLS, p.assessments||[]); return resp({}); }
    if (action === "saveATExactPath") {
      var epRows = (p.exactpath||[]).map(function(r) {
        return Object.assign({}, r, { trophies: JSON.stringify(r.trophies||[]) });
      });
      objectsToSheet(SHEET.AT_EP, AT_EP_COLS, epRows);
      return resp({});
    }
    if (action === "saveATNotes")   { objectsToSheet(SHEET.AT_NOTES,   AT_NOTES_COLS, p.notes  ||[]); return resp({}); }
    if (action === "saveATReports") { objectsToSheet(SHEET.AT_REPORTS, AT_RPT_COLS,   p.reports||[]); return resp({}); }

    // ════════════════════════════════════════════════════════════════════════
    //  AWSC-27 PROJECT TRACKER
    // ════════════════════════════════════════════════════════════════════════
    if (action === "savePTAssignments") {
      var ptRows = (p.assignments||[]).map(function(a) {
        return Object.assign({}, a, {
          team:    JSON.stringify(Array.isArray(a.team) ? a.team : []),
          mastery: a.mastery ? "TRUE":"FALSE"
        });
      });
      // Per-student merge if sid provided
      if (p.sid) {
        var sid = String(p.sid);
        var existing = sheetToObjects(SHEET.PT_ASSIGN, PT_ASSIGN_COLS);
        var others = existing.filter(function(r) { return String(r.sid) !== sid; });
        objectsToSheet(SHEET.PT_ASSIGN, PT_ASSIGN_COLS, others.concat(ptRows));
      } else {
        objectsToSheet(SHEET.PT_ASSIGN, PT_ASSIGN_COLS, ptRows);
      }
      return resp({});
    }
    if (action === "savePTEvaluations") {
      var evRows = (p.evaluations||[]).map(function(ev) {
        return {
          id:ev.id, aid:ev.aid, sid:ev.sid, ms:ev.ms, cs:ev.cs, ov:ev.ov,
          mastery: ev.mastery ? "TRUE":"FALSE",
          crJson:  JSON.stringify(ev.cr||{}),
          coJson:  JSON.stringify(ev.co||{}),
          comment: ev.comment||"", date: ev.date||""
        };
      });
      if (p.sid) {
        var sid = String(p.sid);
        var existing = sheetToObjects(SHEET.PT_EVALS, PT_EVAL_COLS);
        var others = existing.filter(function(r) { return String(r.sid) !== sid; });
        objectsToSheet(SHEET.PT_EVALS, PT_EVAL_COLS, others.concat(evRows));
      } else {
        objectsToSheet(SHEET.PT_EVALS, PT_EVAL_COLS, evRows);
      }
      return resp({});
    }

    // ════════════════════════════════════════════════════════════════════════
    //  USER MANAGEMENT (admin only)
    // ════════════════════════════════════════════════════════════════════════
    if (action === "getUsers") {
      if (!isAdmin) return respErr("Admin access required.");
      var users = sheetToObjects(SHEET.USERS, USER_COLS).map(function(u) {
        return {
          username  : u.username,
          role      : u.role,
          name      : u.name,
          email     : u.email,
          campus    : u.campus,
          active    : (u.active !== "false"),
          sidPrefix : u.sidPrefix || ""
        };
      });
      return resp({ users: users });
    }

    if (action === "saveUser") {
      if (!isAdmin) return respErr("Admin access required.");
      var sh   = getSheet(SHEET.USERS);
      if (!sh) return respErr("Users sheet not found.");
      var u    = p.user;
      var data = sh.getDataRange().getValues();
      // Determine the actual number of columns in the sheet (may be 7 on v7, 8 on v8)
      var shCols = data[0] ? data[0].length : USER_COLS.length;
      var found  = false;
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(u.username)) {
          // Build row — pad to sheet column count for safety
          var row = [
            u.username,
            u.password || data[i][1] || "",
            u.role     || data[i][2] || "staff",
            u.name     || data[i][3] || "",
            u.email    || data[i][4] || "",
            u.campus   || data[i][5] || "",
            data[i][6] || new Date().toISOString(),
            u.active !== undefined ? String(u.active) : (data[i][7] !== undefined ? String(data[i][7]) : "true"),
            u.sidPrefix !== undefined ? String(u.sidPrefix) : (data[i][8] || "")
          ];
          // Only write as many cols as the sheet currently has (backward compat)
          var writeLen = Math.max(shCols, USER_COLS.length);
          sh.getRange(i+1, 1, 1, writeLen).setValues([row.slice(0, writeLen).concat(
            Array(Math.max(0, writeLen - row.length)).fill("")
          )]);
          found = true; break;
        }
      }
      if (!found) {
        sh.appendRow([
          u.username, u.password||"", u.role||"staff",
          u.name||"", u.email||"", u.campus||"",
          new Date().toISOString(),
          u.active !== undefined ? String(u.active) : "true",
          u.sidPrefix||""
        ]);
      }
      // Invalidate user cache so next login reads fresh from sheet
      try { PropertiesService.getScriptProperties().deleteProperty("aws_users_cache"); } catch(ce) {}
      return resp({});
    }

    if (action === "deleteUser") {
      if (!isAdmin) return respErr("Admin access required.");
      var sh   = getSheet(SHEET.USERS);
      if (!sh) return resp({});
      var data = sh.getDataRange().getValues();
      for (var i = data.length - 1; i >= 1; i--) {
        if (String(data[i][0]) === String(p.username)) { sh.deleteRow(i + 1); break; }
      }
      // Invalidate user cache
      try { PropertiesService.getScriptProperties().deleteProperty("aws_users_cache"); } catch(ce) {}
      return resp({});
    }

    return respErr("Unknown action: " + action);

  } catch(err) {
    return respErr("Server error: " + err.message + " | " + (err.stack || ""));
  }
}


// ════════════════════════════════════════════════════════════════════════════
//  PRIVATE HELPERS
// ════════════════════════════════════════════════════════════════════════════
function _loadStudents() {
  return sheetToObjects(SHEET.STUDENTS, STUDENT_COLS).map(function(s) {
    try { s.documents = JSON.parse(s.documents || "[]"); } catch(e) { s.documents = []; }
    s.id = parseInt(s.id) || 0;
    return s;
  });
}

function _loadSettings() {
  var settings = {};
  sheetToObjects(SHEET.SETTINGS, ["key","value"]).forEach(function(r) {
    try { settings[r.key] = JSON.parse(r.value); } catch(e2) { settings[r.key] = r.value; }
  });
  return settings;
}


// ════════════════════════════════════════════════════════════════════════════
//  UPGRADE FROM v7 → v8
//  Run this ONCE after pasting the new script into an existing v7 spreadsheet.
//  ✅ SAFE: never deletes or overwrites any existing data rows.
//  It only: creates missing tabs, adds missing header columns, colours tabs.
// ════════════════════════════════════════════════════════════════════════════
function upgradeToV8() {
  var log = [];

  // ── 1. Create new v8 tabs if they don't exist yet ─────────────────────────
  var newSheets = [
    [SHEET.ELEM_PROGRESS,  ELEM_PROGRESS_COLS,  "#0A6B64"],
    [SHEET.MS_GRADES,      MS_GRADES_COLS,      "#0A6B64"],
    [SHEET.ELEM_NARRATIVE, ELEM_NARRATIVE_COLS, "#6B21A8"],
    [SHEET.ELEM_PORTFOLIO, ELEM_PORTFOLIO_COLS, "#6B21A8"],
    [SHEET.SKILL_MASTERY,  SKILL_MASTERY_COLS,  "#1DBD6A"]
  ];
  newSheets.forEach(function(pair) {
    var name = pair[0], cols = pair[1], colour = pair[2];
    var sh = getSheet(name);
    if (!sh) {
      sh = getSpreadsheet().insertSheet(name);
      sh.getRange(1, 1, 1, cols.length).setValues([cols]);
      styleHeader(sh, cols.length);
      sh.setFrozenRows(1);
      try { sh.setTabColor(colour); } catch(e) {}
      try { sh.setColumnWidth(1, 130); } catch(e) {}
      log.push("Created new tab: " + name);
    } else {
      log.push("Tab already exists (skipped): " + name);
    }
  });

  // ── 2. Fix Users sheet — ensure all required columns exist ────────────────
  var usersSh = getSheet(SHEET.USERS);
  if (usersSh) {
    var usersData = usersSh.getDataRange().getValues();
    if (usersData.length > 0) {
      var usersHeader = usersData[0].map(function(h){ return String(h).trim().toLowerCase(); });
      
      // Add password column if missing (insert at position 2 = column B)
      if (usersHeader.indexOf("password") === -1) {
        // Insert "password" column after "username" (col A)
        usersSh.insertColumnAfter(1);
        usersSh.getRange(1, 2).setValue("password");
        // Set temp password "changeme" for all existing users
        for (var ur = 2; ur <= usersSh.getLastRow(); ur++) {
          usersSh.getRange(ur, 2).setValue("changeme");
        }
        styleHeader(usersSh, usersSh.getLastColumn());
        log.push("Users sheet: added 'password' column — all users set to temp password: changeme");
      } else {
        log.push("Users sheet: 'password' column already exists");
      }
      
      // Add active column if missing
      var updatedHeader = usersSh.getRange(1, 1, 1, usersSh.getLastColumn()).getValues()[0]
        .map(function(h){ return String(h).trim().toLowerCase(); });
      if (updatedHeader.indexOf("active") === -1) {
        var nextCol = usersSh.getLastColumn() + 1;
        usersSh.getRange(1, nextCol).setValue("active");
        styleHeader(usersSh, nextCol);
        log.push("Users sheet: added 'active' column");
      } else {
        log.push("Users sheet: 'active' column already exists");
      }
      // Add sidPrefix column if missing
      var updatedHeader2 = usersSh.getRange(1, 1, 1, usersSh.getLastColumn()).getValues()[0]
        .map(function(h){ return String(h).trim().toLowerCase(); });
      if (updatedHeader2.indexOf("sidprefix") === -1) {
        var sidCol = usersSh.getLastColumn() + 1;
        usersSh.getRange(1, sidCol).setValue("sidPrefix");
        styleHeader(usersSh, sidCol);
        log.push("Users sheet: added 'sidPrefix' column for partner campus ID series");
      } else {
        log.push("Users sheet: 'sidPrefix' column already exists");
      }

    }
  }

  // ── 3. Apply/refresh tab colours to all existing sheets ──────────────────
  var colours = {
    "Students":"#1A365E","Courses":"#0369A1","Transfer":"#0891B2",
    "Attendance":"#059669","HealthRecords":"#7C3AED","BehaviourLog":"#D97706",
    "AT_Assignments":"#D61F31","AT_Submissions":"#B91C1C","AT_Assessments":"#9F1239",
    "AT_ExactPath":"#6D28D9","AT_Notes":"#1D4ED8","AT_Reports":"#0F766E",
    "PT_Assignments":"#B45309","PT_Evaluations":"#92400E",
    "Fees":"#374151","Communications":"#0F2240",
    "ElemProgress":"#0A6B64","MSGrades":"#0A6B64",
    "ElemNarrative":"#6B21A8","ElemPortfolio":"#6B21A8","SkillMastery":"#1DBD6A",
    "Settings":"#6B7280","Users":"#374151","Sessions":"#9CA3AF"
  };
  Object.keys(colours).forEach(function(name) {
    var sh = getSpreadsheet().getSheetByName(name);
    if (sh) { try { sh.setTabColor(colours[name]); } catch(e) {} }
  });
  log.push("Tab colours refreshed");

  // ── 4. Freeze header rows on all sheets ──────────────────────────────────
  getSpreadsheet().getSheets().forEach(function(sh) {
    try { if (sh.getFrozenRows() === 0) sh.setFrozenRows(1); } catch(e) {}
  });
  log.push("Frozen header rows ensured on all sheets");

  SpreadsheetApp.getUi().alert(
    "✅  AWS SIS v8 Upgrade Complete!\n\n" +
    "Summary:\n" + log.map(function(l) { return "  • " + l; }).join("\n") + "\n\n" +
    "⚠️  Your existing data has NOT been modified.\n\n" +
    "Next: Re-deploy → Manage deployments → Edit → New version → Deploy"
  );
}


// ════════════════════════════════════════════════════════════════════════════
//  FRESH INSTALL — setupSpreadsheet()
//  Only run this on a BRAND NEW Google Sheet.
//  If you have existing data, run upgradeToV8() instead.
// ════════════════════════════════════════════════════════════════════════════
function setupSpreadsheet() {
  var allSheets = [
    [SHEET.STUDENTS,       STUDENT_COLS,       "#1A365E"],
    [SHEET.COURSES,        COURSE_COLS,        "#0369A1"],
    [SHEET.TRANSFER,       TRANSFER_COLS,      "#0891B2"],
    [SHEET.ATTENDANCE,     ATT_COLS,           "#059669"],
    [SHEET.INTERVIEWS,     INT_COLS,           "#374151"],
    [SHEET.FEES,           FEE_COLS,           "#374151"],
    [SHEET.COMMS,          COMM_COLS,          "#0F2240"],
    [SHEET.STAFF,          STAFF_COLS,         "#374151"],
    [SHEET.CATALOG,        CATALOG_COLS,       "#374151"],
    [SHEET.REMARKS,        REMARK_COLS,        "#6B7280"],
    [SHEET.CALENDAR,       CAL_COLS,           "#374151"],
    [SHEET.LESSONS,        LESSON_COLS,        "#374151"],
    [SHEET.UNITS,          UNIT_COLS,          "#374151"],
    [SHEET.EVENTS,         EVENT_COLS,         "#374151"],
    [SHEET.PD,             PD_COLS,            "#374151"],
    [SHEET.BLOCKS,         BLOCK_COLS,         "#374151"],
    [SHEET.HEALTH,         HEALTH_COLS,        "#7C3AED"],
    [SHEET.BEHAVIOUR,      BEH_COLS,           "#D97706"],
    [SHEET.AT_ASSIGN,      AT_ASSIGN_COLS,     "#D61F31"],
    [SHEET.AT_SUBS,        AT_SUBS_COLS,       "#B91C1C"],
    [SHEET.AT_ASSESS,      AT_ASSESS_COLS,     "#9F1239"],
    [SHEET.AT_EP,          AT_EP_COLS,         "#6D28D9"],
    [SHEET.AT_NOTES,       AT_NOTES_COLS,      "#1D4ED8"],
    [SHEET.AT_REPORTS,     AT_RPT_COLS,        "#0F766E"],
    [SHEET.PT_ASSIGN,      PT_ASSIGN_COLS,     "#B45309"],
    [SHEET.PT_EVALS,       PT_EVAL_COLS,       "#92400E"],
    [SHEET.ELEM_PROGRESS,  ELEM_PROGRESS_COLS, "#0A6B64"],
    [SHEET.MS_GRADES,      MS_GRADES_COLS,     "#0A6B64"],
    [SHEET.ELEM_NARRATIVE, ELEM_NARRATIVE_COLS,"#6B21A8"],
    [SHEET.ELEM_PORTFOLIO, ELEM_PORTFOLIO_COLS,"#6B21A8"],
    [SHEET.SKILL_MASTERY,  SKILL_MASTERY_COLS, "#1DBD6A"],
    [SHEET.SETTINGS,       ["key","value"],     "#6B7280"],
    [SHEET.USERS,          USER_COLS,          "#374151"],
    [SHEET.SESSIONS,       SESSION_COLS,       "#9CA3AF"]
  ];

  // Rename Sheet1 → Students
  var firstSheet = getSpreadsheet().getSheets()[0];
  if (firstSheet.getName() !== SHEET.STUDENTS) firstSheet.setName(SHEET.STUDENTS);

  allSheets.forEach(function(triple) {
    var name = triple[0], cols = triple[1], colour = triple[2];
    var sh   = getSheet(name);
    if (!sh) sh = getSpreadsheet().insertSheet(name);
    // Only write header if sheet is truly empty (no rows at all)
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, cols.length).setValues([cols]);
    }
    styleHeader(sh, cols.length);
    sh.setFrozenRows(1);
    try { sh.setTabColor(colour); } catch(e) {}
    try { sh.setColumnWidth(1, 130); } catch(e) {}
  });

  // Default Settings (only if Settings sheet is empty)
  var settSheet = getSheet(SHEET.SETTINGS);
  if (settSheet && settSheet.getLastRow() <= 1) {
    var defaults = [
      ["academicYear",  "2024 - 2025"],
      ["schoolName",    "American World School"],
      ["campuses",      JSON.stringify(["Chennai","Sri Lanka","UAE","Spain","Online"])],
      ["capacity",      "Pre-K–5: 120 | 6–8: 80 | 9–12: 100"],
      ["gradeSystem",   "US Standards"],
      ["gradeScale",    "A/B/C/D/F"],
      ["feeCurrency",   JSON.stringify({ base:"INR", rates:{USD:1, INR:83.2} })],
      ["timezone",      "Asia/Kolkata"],
      ["successCoach",  "Ms. Radhika Rupini"],
      ["docs",          JSON.stringify(["Birth Certificate","Immunization Records","Prior School Transcript","Passport / ID Copy","Medical Form","Recommendation Letter","Emergency Contact Form","Technology Agreement"])],
      ["cohorts",       JSON.stringify(["Grade 1A","Grade 1B","Grade 5A","Grade 5B","Grade 9A","Grade 9B","Grade 10A"])],
      ["sidPrefix",     "AWS"],
      ["emailNotif",    "Enabled for status changes and document requests"]
    ];
    defaults.forEach(function(row) { settSheet.appendRow(row); });
  }

  // Default Users (only if Users sheet is empty)
  var usersSheet = getSheet(SHEET.USERS);
  if (usersSheet && usersSheet.getLastRow() <= 1) {
    var now = new Date().toISOString();
    usersSheet.appendRow(["admin",            "admin123",   "admin",   "Administrator",    "admin@aws.edu",          "Chennai",   now, "true", ""]);
    usersSheet.appendRow(["staff",            "staff123",   "staff",   "Staff Member",     "staff@aws.edu",          "Chennai",   now, "true", ""]);
    usersSheet.appendRow(["coach",            "coach123",   "staff",   "Success Coach",    "coach@aws.edu",          "Chennai",   now, "true", ""]);
    usersSheet.appendRow(["partner_srilanka", "partner123", "partner", "Sri Lanka Partner","partner.lk@aws.edu",     "Sri Lanka", now, "true", "SL"]);
    usersSheet.appendRow(["partner_uae",      "partner456", "partner", "UAE Partner",      "partner.uae@aws.edu",    "UAE",       now, "true", "UAE"]);
    usersSheet.appendRow(["partner_spain",    "partner789", "partner", "Spain Partner",    "partner.spain@aws.edu",  "Spain",     now, "true", "SP"]);
  }

  SpreadsheetApp.getUi().alert(
    "✅  AWS SIS v8 Fresh Setup Complete!\n\n" +
    "Created " + allSheets.length + " sheet tabs.\n\n" +
    "Default Logins:\n" +
    "  admin            / admin123\n" +
    "  staff            / staff123\n" +
    "  partner_srilanka / partner123  (Sri Lanka campus only)\n" +
    "  partner_uae      / partner456  (UAE campus only)\n" +
    "  partner_spain    / partner789  (Spain campus only)\n\n" +
    "Next: Deploy → New deployment → Web app\n" +
    "Execute as: Me  |  Access: Anyone\n" +
    "Copy the URL into SIS Settings."
  );
}


// ════════════════════════════════════════════════════════════════════════════
//  UTILITIES
// ════════════════════════════════════════════════════════════════════════════

// ── BUILD USER CACHE — run once after updating passwords ──────────────────────
// Primes the PropertiesService cache so login bypasses sheet reads entirely
function buildUserCache() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  if (!sh) { Logger.log("❌ Users sheet not found"); return; }
  var data = sh.getDataRange().getValues();
  if (data.length <= 1) { Logger.log("❌ No users in sheet"); return; }
  var header = data[0].map(function(h){ return String(h).trim().toLowerCase(); });
  var ci = {};
  header.forEach(function(h,i){ ci[h]=i; });
  var cache = {};
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var uName = String(row[ci["username"]]||"").trim().toLowerCase();
    if (!uName) continue;
    cache[uName] = {
      pw     : String(ci["password"]!==undefined ? row[ci["password"]] : "").trim(),
      role   : String(ci["role"]!==undefined     ? row[ci["role"]]     : "staff"),
      name   : String(ci["name"]!==undefined     ? row[ci["name"]]     : ""),
      email  : String(ci["email"]!==undefined    ? row[ci["email"]]    : ""),
      campus : String(ci["campus"]!==undefined   ? row[ci["campus"]]   : ""),
      active : ci["active"]!==undefined ? String(row[ci["active"]]||"").trim() : ""
    };
    Logger.log("Cached user: " + uName + " / role=" + cache[uName].role);
  }
  PropertiesService.getScriptProperties().setProperty("aws_users_cache", JSON.stringify(cache));
  Logger.log("✅ User cache built with " + Object.keys(cache).length + " users. Login will now be fast.");
  SpreadsheetApp.getUi().alert("✅ User cache built!\n\nCached " + Object.keys(cache).length + " users.\nLogin will now skip the sheet read and respond instantly.\n\nRun this again any time you add or change passwords.");
}

// ── TEST LOGIN — run from editor to diagnose login issues ─────────────────────
// Change username/password below and run. Check Execution Log for result.
function testLogin() {
  // ── Change these to match the credentials you are testing ────────────────
  var testUser = "admin";
  var testPass = "admin123";
  // ─────────────────────────────────────────────────────────────────────────

  var sh = getSheet(SHEET.USERS);
  if (!sh) { Logger.log("❌ Users sheet NOT FOUND"); return; }

  var data = sh.getDataRange().getValues();
  Logger.log("Users sheet rows (including header): " + data.length);

  // Print header
  var header = data[0].map(function(h) { return String(h).trim().toLowerCase(); });
  Logger.log("Header columns: " + JSON.stringify(header));

  // Check for required columns
  var required = ["username", "password", "role", "active"];
  required.forEach(function(col) {
    var idx = header.indexOf(col);
    if (idx === -1) {
      Logger.log("❌ MISSING COLUMN: '" + col + "' — this will cause login to fail!");
    } else {
      Logger.log("✅ Column '" + col + "' found at index: " + idx);
    }
  });
  
  if (header.indexOf("password") === -1) {
    Logger.log("⚠️  FIX: Run upgradeToV8() to automatically add the password column");
    Logger.log("⚠️  OR manually insert a column B called 'password' in the Users sheet");
  }

  // Print each user row
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var uCol = header.indexOf("username");
    var pCol = header.indexOf("password");
    var aCol = header.indexOf("active");
    Logger.log("Row " + i + ": username=[" + String(row[uCol]).trim() + "]"
      + " password=[" + String(row[pCol]).trim() + "]"
      + " active=[" + (aCol >= 0 ? String(row[aCol]).trim() : "N/A") + "]");
  }

  // Simulate exact login logic
  var inUser = testUser.trim().toLowerCase();
  var inPass = testPass.trim();
  var matched = false;

  for (var ri = 1; ri < data.length; ri++) {
    var row = data[ri];
    var colIdx = {};
    header.forEach(function(h, idx) { colIdx[h] = idx; });

    var rowUser   = String(row[colIdx["username"]] || "").trim().toLowerCase();
    var rowPass   = String(row[colIdx["password"]] || "").trim();
    var rowActive = colIdx["active"] !== undefined ? String(row[colIdx["active"]] || "").trim() : "";
    var activeOk  = (rowActive === "" || rowActive === "true" || rowActive === "TRUE");

    Logger.log("Checking row " + ri + ": userMatch=" + (rowUser === inUser)
      + " passMatch=" + (rowPass === inPass) + " activeOk=" + activeOk);

    if (rowUser === inUser && rowPass === inPass && activeOk) {
      Logger.log("✅ LOGIN WOULD SUCCEED for: " + testUser + " (role=" + String(row[colIdx["role"]] || "") + ")");
      matched = true;
      break;
    }
  }

  if (!matched) {
    Logger.log("❌ LOGIN WOULD FAIL — no matching row found for username=[" + testUser + "]");
    Logger.log("Common causes:");
    Logger.log("  1. Username or password has extra spaces");
    Logger.log("  2. Password stored as a number (not text) in Google Sheets");
    Logger.log("  3. Username capitalisation mismatch");
    Logger.log("  4. 'active' column has value other than blank / true");
  }
}

// Clean expired sessions — wire to a daily time-driven trigger
function cleanSessions() {
  var sh = getSheet(SHEET.SESSIONS);
  if (!sh) return;
  var data = sh.getDataRange().getValues();
  var now  = new Date();
  for (var i = data.length - 1; i >= 1; i--) {
    try { if (new Date(data[i][6]) < now) sh.deleteRow(i + 1); } catch(e) {}
  }
}

// Run once to set up daily session cleanup trigger
function createTriggers() {
  // Remove existing cleanSessions triggers to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "cleanSessions") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("cleanSessions").timeBased().everyDays(1).atHour(2).create();
  SpreadsheetApp.getUi().alert("✅ Trigger set: cleanSessions runs daily at 2am.");
}

// Force all users to re-login (after security events)
function resetAllSessions() {
  var sh = getSheet(SHEET.SESSIONS);
  if (!sh) return;
  sh.clearContents();
  sh.getRange(1, 1, 1, SESSION_COLS.length).setValues([SESSION_COLS]);
  styleHeader(sh, SESSION_COLS.length);
  SpreadsheetApp.getUi().alert("✅ All sessions cleared. All users must log in again.");
}

// Export all sheets as CSV files to a Drive folder
function exportAllToCSV() {
  var folder = DriveApp.createFolder("AWS_SIS_Export_" + new Date().toISOString().slice(0,10));
  getSpreadsheet().getSheets().forEach(function(sh) {
    var data = sh.getDataRange().getValues();
    var csv  = data.map(function(row) {
      return row.map(function(cell) {
        var s = String(cell).replace(/"/g, '""');
        return (s.indexOf(",") > -1 || s.indexOf('"') > -1) ? '"'+s+'"' : s;
      }).join(",");
    }).join("\n");
    folder.createFile(sh.getName() + ".csv", csv, MimeType.PLAIN_TEXT);
  });
  SpreadsheetApp.getUi().alert("✅ Export complete — Drive folder: " + folder.getName());
}
