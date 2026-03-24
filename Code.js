// ============================================================
//  OPS TRIP MONITOR — Ormoc Printshoppe
//  BACKEND  (Code.gs)
//  All database operations, CRUD, validation, role checks
// ============================================================

// ============================================================
//  SHEET NAMES
// ============================================================
const OPS_SHEETS = {
  VEHICLES : 'Vehicles',
  TRIPS    : 'Trips',
  SETTINGS : 'Settings',
  RENEWALS : 'Renewal_Alerts',
  ROLES    : 'Role_Permissions',
  AUDIT    : 'Audit_Log',
  DRIVERS  : 'Drivers',
};

// ============================================================
//  COLUMN MAPS  (0-indexed)
// ============================================================

// Vehicles  A-M
const VEH_COL = {
  VEHICLE_ID    : 0,
  PLATE         : 1,
  TYPE          : 2,
  BRAND         : 3,
  BEG_MILEAGE   : 4,
  STATUS        : 5,
  INS_EXPIRY    : 6,
  INS_LINK      : 7,
  LTO_EXPIRY    : 8,
  LTO_LINK      : 9,
  NOTES         : 10,
  CREATED_AT    : 11,
  UPDATED_AT    : 12
};

// Trips  A-AC
const TRIP_COL = {
  TRIP_ID       : 0,
  REQUEST_DATE  : 1,
  REQ_EMP_ID    : 2,
  REQ_NAME      : 3,
  TRIP_TYPE     : 4,
  PURPOSE       : 5,
  RELATED_JO    : 6,
  FROM_LOC      : 7,
  TO_LOC        : 8,
  START_DATE    : 9,
  END_DATE      : 10,
  VEHICLE_ID    : 11,
  PLATE         : 12,
  DRIVER_EMP_ID : 13,
  DRIVER_NAME   : 14,
  STATUS        : 15,
  APPROVED_BY   : 16,
  APPROVAL_DATE : 17,
  REJECT_REASON : 18,
  CANCEL_REASON : 19,
  ACTUAL_START  : 20,
  ACTUAL_END    : 21,
  START_KM      : 22,
  END_KM        : 23,
  DISTANCE      : 24,
  PROOF_LINK    : 25,
  REMARKS       : 26,
  UPDATED_AT    : 27,
  UPDATED_BY    : 28
};

// ============================================================
//  STATUS CONSTANTS
// ============================================================
const TRIP_STATUS = {
  DRAFT     : 'Draft',
  SUBMITTED : 'Submitted',
  APPROVED  : 'Approved',
  REJECTED  : 'Rejected',
  CANCELLED : 'Cancelled',
  COMPLETED : 'Completed'
};

const VEH_STATUS = {
  ACTIVE  : 'Active',
  REPAIR  : 'Under Repair',
  INACTIVE: 'Inactive'
};

// ============================================================
//  WEB APP ENTRY POINT
// ============================================================
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('OPS Trip Monitor — Ormoc Printshoppe')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================
//  BOOTSTRAP — creates all sheets if missing
// ============================================================
function ops_bootstrap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const schemas = [
    {
      name: OPS_SHEETS.VEHICLES,
      headers: ['Vehicle_ID','Plate_Number','Vehicle_Type','Brand_Model',
                'Beginning_Mileage','Status','Insurance_Expiry','Insurance_PDF_Link',
                'LTO_Expiry','LTO_PDF_Link','Notes','Created_At','Updated_At']
    },
    {
      name: OPS_SHEETS.TRIPS,
      headers: ['Trip_ID','Request_Date','Requestor_EmpID','Requestor_Name',
                'Trip_Type','Purpose','Related_JO','From_Location','To_Location',
                'Planned_Start','Planned_End','Vehicle_ID','Plate_Number',
                'Driver_EmpID','Driver_Name','Status','Approved_By','Approval_Date',
                'Rejection_Reason','Cancel_Reason','Actual_Start','Actual_End',
                'Start_Mileage','End_Mileage','Distance_Travelled','GPS_Proof_Link',
                'Remarks','Updated_At','Updated_By']
    },
    {
      name: OPS_SHEETS.SETTINGS,
      headers: ['Setting_Key','Setting_Value']
    },
    {
      name: OPS_SHEETS.ROLES,
      headers: ['Role','Emails','Abilities']
    },
    {
      name: OPS_SHEETS.AUDIT,
      headers: ['DateTime','Action','User','Role','Payload']
    },
    {
      name: 'Drivers',
      headers: ['Driver_ID','Full_Name','Employee_ID','License_ID',
                'License_Expiry','Contact_Number','Status','Notes']
    },
  ];

  schemas.forEach(function(s) {
    if (!ss.getSheetByName(s.name)) {
      const sh = ss.insertSheet(s.name);
      sh.getRange(1, 1, 1, s.headers.length).setValues([s.headers]);
      sh.getRange(1, 1, 1, s.headers.length)
        .setFontWeight('bold').setBackground('#f8fafc');
      sh.setFrozenRows(1);
    }
  });

  // Seed settings defaults
  const set = ss.getSheetByName(OPS_SHEETS.SETTINGS);
  if (set && set.getLastRow() < 2) {
    set.getRange(2, 1, 2, 2).setValues([
      ['renewal_alert_days', '30'],
      ['app_version', '1.0']
    ]);
  }

  return { success: true, message: 'OPS sheets initialized.' };
}

// ============================================================
//  SHEET GETTER — auto-bootstraps if missing
// ============================================================
function ops_sh_(name) {
  let sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) { ops_bootstrap(); sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); }
  return sh;
}

// ============================================================
//  DATE HELPERS
// ============================================================
function ops_fmtDate_(val) {
  if (val === null || val === undefined || val === '') return '';
  try {
    if (val instanceof Date) {
      if (isNaN(val.getTime())) return '';
      var y1 = val.getFullYear();
      if (y1 < 1900 || y1 > 2200) return '';
      return y1 + '-'
        + String(val.getMonth() + 1).padStart(2, '0') + '-'
        + String(val.getDate()).padStart(2, '0');
    }
    if (typeof val === 'number') {
      if (val <= 0) return '';
      var msPerDay  = 86400000;
      var epoch     = new Date(1899, 11, 30).getTime();
      var d         = new Date(epoch + Math.floor(val) * msPerDay);
      if (isNaN(d.getTime())) return '';
      var y2 = d.getFullYear();
      if (y2 < 1900 || y2 > 2200) return '';
      return y2 + '-'
        + String(d.getMonth() + 1).padStart(2, '0') + '-'
        + String(d.getDate()).padStart(2, '0');
    }
    var s = String(val).trim();
    if (!s || s === '0') return '';
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
      var parts = s.split('-');
      var y3    = parseInt(parts[0]);
      if (y3 < 1900 || y3 > 2200) return '';
      return s;
    }
    var parsed = new Date(s);
    if (isNaN(parsed.getTime())) return '';
    var y4 = parsed.getFullYear();
    if (y4 < 1900 || y4 > 2200) return '';
    return y4 + '-'
      + String(parsed.getMonth() + 1).padStart(2, '0') + '-'
      + String(parsed.getDate()).padStart(2, '0');
  } catch(e) {
    Logger.log('ops_fmtDate_ error: ' + e.message + ' | val=' + val);
    return '';
  }
}

function ops_fmtDT_(val) {
  if (val === null || val === undefined || val === '') return '';
  try {
    var d;
    var mo = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    if (val instanceof Date) {
      d = val;
    } else if (typeof val === 'number') {
      if (val <= 0) return '';
      var epoch = new Date(1899, 11, 30).getTime();
      d = new Date(epoch + Math.floor(val) * 86400000);
    } else {
      var s = String(val).trim();
      if (!s || s === '0') return '';
      d = new Date(s);
    }
    if (isNaN(d.getTime())) return '';
    var y = d.getFullYear();
    if (y < 1900 || y > 2200) return '';
    return mo[d.getMonth()] + ' ' + String(d.getDate()).padStart(2, '0')
      + ', ' + y + ' '
      + String(d.getHours()).padStart(2, '0') + ':'
      + String(d.getMinutes()).padStart(2, '0');
  } catch(e) {
    Logger.log('ops_fmtDT_ error: ' + e.message + ' | val=' + val);
    return '';
  }
}

function ops_toISO_(val) {
  if (val === null || val === undefined || val === '') return '';
  try {
    var d;
    if (val instanceof Date) {
      d = val;
    } else if (typeof val === 'number') {
      if (val <= 0) return '';
      var epoch = new Date(1899, 11, 30).getTime();
      d = new Date(epoch + Math.floor(val) * 86400000);
    } else {
      var s = String(val).trim();
      if (!s || s === '0') return '';
      d = new Date(s);
    }
    if (isNaN(d.getTime())) return '';
    var y = d.getFullYear();
    if (y < 1900 || y > 2200) return '';
    return d.toISOString();
  } catch(e) {
    Logger.log('ops_toISO_ error: ' + e.message + ' | val=' + val);
    return '';
  }
}

function ops_now_() {
  return Utilities.formatDate(new Date(), 'Asia/Manila', "yyyy-MM-dd'T'HH:mm:ss");
}

function ops_daysLeft_(dateStr) {
  if (!dateStr) return null;
  try {
    var d    = new Date(dateStr);
    var now  = new Date();
    var diff = Math.round((d - now) / 86400000);
    return isNaN(diff) ? null : diff;
  } catch(e) { return null; }
}

// ============================================================
//  ID GENERATOR
// ============================================================
function ops_genId_(prefix, rows, col) {
  var year = new Date().getFullYear();
  var max  = 0;
  rows.forEach(function(r) {
    var id = String(r[col] || '');
    var m  = id.match(/-(\d{4})$/);
    if (m) { var n = parseInt(m[1]); if (n > max) max = n; }
  });
  return prefix + '-' + year + '-' + String(max + 1).padStart(4, '0');
}

// ============================================================
//  ROLE & PERMISSION
// ============================================================
function ops_getUserInfo_() {
  try {
    const email = Session.getActiveUser().getEmail().toLowerCase();
    const sh    = ops_sh_(OPS_SHEETS.ROLES);
    const lr    = sh.getLastRow();
    if (lr >= 2) {
      const data = sh.getRange(2, 1, lr - 1, 3).getValues();
      for (let i = 0; i < data.length; i++) {
        const role      = String(data[i][0] || '').trim();
        const emails    = String(data[i][1] || '').toLowerCase().split(',').map(function(e) { return e.trim(); });
        const abilities = String(data[i][2] || '').toLowerCase().split(',').map(function(a) { return a.trim(); });
        if (emails.includes(email)) return { email, role, abilities };
      }
    }
    return { email, role: 'No Role', abilities: [] };
  } catch(e) {
    return { email: Session.getActiveUser().getEmail(), role: 'No Role', abilities: [] };
  }
}

function ops_getUserInfoPublic() { return ops_getUserInfo_(); }

function ops_isAdmin_(r)    { return r.toLowerCase().includes('admin'); }
function ops_isApprover_(r) { return r.toLowerCase().includes('approver') || ops_isAdmin_(r); }
function ops_isEncoder_(r)  { return r.toLowerCase().includes('encoder') || r.toLowerCase().includes('operator') || ops_isAdmin_(r); }

function ops_hashPassword_(plaintext) {
  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(plaintext),
    Utilities.Charset.UTF_8
  );
  return bytes.map(function(b) {
    var hex = (b & 0xff).toString(16);
    return hex.length === 1 ? '0' + hex : hex;
  }).join('');
}

function ops_loginUser(email, password) {
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LoginUsers');
    if (!sh) return { success: false, message: 'LoginUsers sheet not found. Please contact admin.' };

    var lr = sh.getLastRow();
    if (lr < 2) return { success: false, message: 'No users registered yet.' };

    // Read 4 columns: email, password, role, driverId
    var data       = sh.getRange(2, 1, lr - 1, 4).getValues();
    var inputEmail = String(email    || '').trim().toLowerCase();
    var inputPw    = String(password || '').trim();
    var inputHash  = ops_hashPassword_(inputPw);

    for (var i = 0; i < data.length; i++) {
      var rowEmail = String(data[i][0] || '').trim().toLowerCase();
      var rowPw    = String(data[i][1] || '').trim();
      var rowRole  = String(data[i][2] || '').trim();
      var rowDrvId = String(data[i][3] || '').trim();

      if (rowEmail !== inputEmail) continue;

      var isHashed   = /^[0-9a-f]{64}$/.test(rowPw);
      var passwordOk = isHashed
        ? (rowPw === inputHash)
        : (rowPw === inputPw);

      if (passwordOk) {
        if (!isHashed) {
          try {
            sh.getRange(i + 2, 2).setValue(inputHash);
            SpreadsheetApp.flush();
            Logger.log('Password upgraded to hash for: ' + rowEmail);
          } catch(upgradeErr) {
            Logger.log('Hash upgrade failed (non-critical): ' + upgradeErr.message);
          }
        }
        return {
          success  : true,
          email    : rowEmail,
          role     : rowRole  || 'No Role',
          driverId : rowDrvId || '',
          message  : 'Login successful.'
        };
      } else {
        return { success: false, message: 'Incorrect password. Please try again.' };
      }
    }

    return { success: false, message: 'Email not found. Please check your email or contact admin.' };
  } catch(e) {
    return { success: false, message: 'Login error: ' + e.message };
  }
}

// ============================================================
//  SESSION VERIFY
// ============================================================
function ops_verifySession(email) {
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LoginUsers');
    if (!sh) return { success: false, message: 'LoginUsers sheet not found.' };

    var lr = sh.getLastRow();
    if (lr < 2) return { success: false, message: 'No users registered.' };

    var data       = sh.getRange(2, 1, lr - 1, 4).getValues();
    var inputEmail = String(email || '').trim().toLowerCase();

    for (var i = 0; i < data.length; i++) {
      var rowEmail = String(data[i][0] || '').trim().toLowerCase();
      var rowRole  = String(data[i][2] || '').trim();
      var rowDrvId = String(data[i][3] || '').trim();
      if (rowEmail === inputEmail) {
        return {
          success  : true,
          email    : rowEmail,
          role     : rowRole || 'No Role',
          driverId : rowDrvId
        };
      }
    }

    return { success: false, message: 'Session invalid. Please log in again.' };
  } catch(e) {
    return { success: false, message: 'Verify error: ' + e.message };
  }
}

// ============================================================
//  AUDIT LOG
// ============================================================
function ops_audit_(action, payload) {
  try {
    const sh   = ops_sh_(OPS_SHEETS.AUDIT);
    const user = ops_getUserInfo_();
    sh.insertRowBefore(2);
    sh.getRange(2, 1, 1, 5).setValues([[
      ops_now_(), action, user.email, user.role, JSON.stringify(payload)
    ]]);
  } catch(e) { Logger.log('audit error: ' + e.message); }
}

// ============================================================
//  SETTINGS
// ============================================================
function ops_getSettings_() {
  try {
    const sh = ops_sh_(OPS_SHEETS.SETTINGS);
    const lr = sh.getLastRow();
    if (lr < 2) return { renewal_alert_days: '30' };
    const rows = sh.getRange(2, 1, lr - 1, 2).getValues();
    const out  = {};
    rows.forEach(function(r) { if (r[0]) out[String(r[0]).trim()] = String(r[1]).trim(); });
    return out;
  } catch(e) { return { renewal_alert_days: '30' }; }
}

// ============================================================
//  DATABASE LINK HELPER
// ============================================================
function ops_getDBId_(label) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const linkSheet = ss.getSheetByName('DatabaseLink');
  if (!linkSheet) throw new Error("Sheet 'DatabaseLink' not found.");
  const lastRow = linkSheet.getLastRow();
  if (lastRow < 2) throw new Error("'DatabaseLink' sheet is empty.");
  const labels = linkSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const urls   = linkSheet.getRange(2, 2, lastRow - 1, 1).getValues().flat();
  const idx    = labels.indexOf(label);
  if (idx === -1) throw new Error('Label "' + label + '" not found in DatabaseLink sheet.');
  const url = urls[idx];
  if (!url || url.toString().trim() === '') throw new Error('URL for "' + label + '" is empty.');
  const match = url.toString().match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  if (!match || !match[1]) throw new Error('Could not extract Spreadsheet ID from URL for "' + label + '".');
  return match[1];
}

// ============================================================
//  JO LIST
// ============================================================
function ops_getJOList() {
  try {
    let joDbId;
    try {
      joDbId = ops_getDBId_('JODatabase');
    } catch(e) {
      return { success: false, message: 'DatabaseLink error: ' + e.message };
    }

    let extSS;
    try {
      extSS = SpreadsheetApp.openById(joDbId);
    } catch(e) {
      return { success: false, message: 'Cannot open JODatabase (ID: ' + joDbId + '). Check sharing permissions: ' + e.message };
    }

    const joSh = extSS.getSheetByName('Line-up JOs');
    if (!joSh) {
      const shNames = extSS.getSheets().map(function(s) { return s.getName(); });
      return { success: false, message: '"Line-up JOs" not found. Available sheets: ' + shNames.join(', ') };
    }

    const lr = joSh.getLastRow();
    if (lr < 2) return { success: true, data: [] };

    const data = joSh.getRange(2, 1, lr - 1, 12).getValues();
    const list = [];
    data.forEach(function(r) {
      const joNumber = String(r[11] || '').trim();
      const jobDesc  = String(r[8]  || '').trim();
      if (joNumber) list.push({ joNumber: joNumber, jobDesc: jobDesc });
    });

    return { success: true, data: list };
  } catch(e) {
    return { success: false, message: 'ops_getJOList error: ' + e.message };
  }
}

// ============================================================
//  EMPLOYEE LISTS
// ============================================================
function ops_getEmployeeList() {
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('EmployeName');
    if (!sh) return { success: false, message: 'EmployeName sheet not found.' };
    var lr = sh.getLastRow();
    if (lr < 2) return { success: true, data: [] };
    var data = sh.getRange(2, 1, lr - 1, 3).getValues();
    var list = [];
    data.forEach(function(r) {
      var empId   = String(r[0] || '').trim();
      var team    = String(r[1] || '').trim();
      var empName = String(r[2] || '').trim();
      if (empName) list.push({ empId: empId, empName: empName, team: team });
    });
    return { success: true, data: list };
  } catch(e) {
    return { success: false, message: 'ops_getEmployeeList error: ' + e.message };
  }
}

function ops_getEmployeeList_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('EmployeName');
    if (!sh) return [];
    const lr = sh.getLastRow();
    if (lr < 2) return [];
    return sh.getRange(2, 1, lr - 1, 3).getValues()
      .filter(function(r) { return r[2] && String(r[2]).trim(); })
      .map(function(r) {
        return {
          empCode : String(r[0] || '').trim(),
          team    : String(r[1] || '').trim(),
          name    : String(r[2] || '').trim()
        };
      });
  } catch(e) {
    Logger.log('ops_getEmployeeList_ error: ' + e.message);
    return [];
  }
}

// ============================================================
//  COMBINED INIT DATA
// ============================================================
function getDashboardInitData() {
  try {
    const user     = ops_getUserInfo_();
    const trips    = ops_getAllTrips_();
    const vehicles = ops_getAllVehicles_();
    const settings = ops_getSettings_();
    const alertDays= parseInt(settings.renewal_alert_days) || 30;

    const stats = {
      total    : trips.length,
      pending  : trips.filter(function(t) { return t.status === TRIP_STATUS.SUBMITTED; }).length,
      approved : trips.filter(function(t) { return t.status === TRIP_STATUS.APPROVED; }).length,
      completed: trips.filter(function(t) { return t.status === TRIP_STATUS.COMPLETED; }).length,
      vehicles : vehicles.filter(function(v) { return v.status === VEH_STATUS.ACTIVE; }).length
    };

    const recent = trips.slice(-5).reverse();
    const alerts = ops_buildRenewalAlerts_(vehicles, alertDays);

    return { success: true, user, stats, recent, alerts, settings };
  } catch(e) { return { success: false, message: e.message }; }
}

function getTripsInitData() {
  try {
    const user      = ops_getUserInfo_();
    const trips     = ops_getAllTrips_();
    const vehicles  = ops_getAllVehicles_();
    const drivers   = ops_getAllDrivers_();
    const employees = ops_getEmployeeList_();
    const joResult  = ops_getJOList();
    const joList    = joResult.success ? joResult.data : [];
    const joError   = joResult.success ? null : joResult.message;
    const tripTypes = ops_getTripTypes_().map(function(r) { return { value: r.value }; });
    return { success: true, user, trips, vehicles, drivers, employees, joList, joError, tripTypes };
  } catch(e) { return { success: false, message: e.message }; }
}

function getVehiclesInitData() {
  try {
    const user     = ops_getUserInfo_();
    const vehicles = ops_getAllVehicles_();
    return { success: true, user, vehicles };
  } catch(e) { return { success: false, message: e.message }; }
}

function getApprovalInitData() {
  try {
    const user  = ops_getUserInfo_();
    const trips = ops_getAllTrips_().filter(function(t) {
      return t.status === TRIP_STATUS.SUBMITTED;
    });
    return { success: true, user, trips };
  } catch(e) { return { success: false, message: e.message }; }
}

function getCompletionInitData() {
  try {
    const user  = ops_getUserInfo_();
    const trips = ops_getAllTrips_().filter(function(t) {
      return t.status === TRIP_STATUS.APPROVED;
    });
    return { success: true, user, trips };
  } catch(e) { return { success: false, message: e.message }; }
}

function getRenewalsInitData() {
  try {
    const user     = ops_getUserInfo_();
    const vehicles = ops_getAllVehicles_();
    const settings = ops_getSettings_();
    const alertDays= parseInt(settings.renewal_alert_days) || 30;
    const alerts   = ops_buildRenewalAlerts_(vehicles, alertDays);
    return { success: true, user, alerts };
  } catch(e) { return { success: false, message: e.message }; }
}

function getReportsInitData() {
  try {
    const user     = ops_getUserInfo_();
    const trips    = ops_getAllTrips_();
    const vehicles = ops_getAllVehicles_();
    const report   = ops_buildReports_(trips, vehicles);
    return {
      success     : true,
      user,
      report,
      rawTrips    : trips,
      rawVehicles : vehicles
    };
  } catch(e) { return { success: false, message: e.message }; }
}

function getDriversInitData() {
  try {
    const user    = ops_getUserInfo_();
    const drivers = ops_getAllDrivers_();
    return { success: true, user, drivers };
  } catch(e) { return { success: false, message: e.message }; }
}

// ============================================================
//  DRIVER DASHBOARD DATA
//  ✅ FIXED: Uses Driver_ID as the single source of truth.
//  Triple fallback: Driver_ID → email → name (for legacy trips)
// ============================================================

function getDriverDashboardData(sessionEmail) {
  try {
    var email = String(sessionEmail || '').toLowerCase().trim();
    if (!email) return { success: false, message: 'Session expired. Please log in again.' };
    var user = { email: email, role: 'driver', abilities: [] };

    var driverId   = '';
    var driverName = '';

    // ── Step 1: Look up Driver_ID from LoginUsers col 4 by email ──
    try {
      var ss      = SpreadsheetApp.getActiveSpreadsheet();
      var loginSh = ss.getSheetByName('LoginUsers');
      if (loginSh && loginSh.getLastRow() >= 2) {
        var loginData = loginSh.getRange(2, 1, loginSh.getLastRow() - 1, 4).getValues();
        for (var i = 0; i < loginData.length; i++) {
          var rowEmail = String(loginData[i][0] || '').trim().toLowerCase();
          var rowRole  = String(loginData[i][2] || '').trim().toLowerCase();
          var rowDrvId = String(loginData[i][3] || '').trim();
          if (rowEmail === email && rowRole === 'driver' && rowDrvId) {
            driverId = rowDrvId;
            break;
          }
        }
      }
    } catch(e) {
      Logger.log('LoginUsers lookup error: ' + e.message);
    }

    // ── Step 2: Get driver name from Drivers sheet using Driver_ID ──
    if (driverId) {
      try {
        var driverSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Drivers');
        if (driverSh && driverSh.getLastRow() >= 2) {
          var driverData = driverSh.getRange(2, 1, driverSh.getLastRow() - 1, 8).getValues();
          for (var j = 0; j < driverData.length; j++) {
            if (String(driverData[j][0] || '').trim() === driverId) {
              driverName = String(driverData[j][1] || '').trim();
              break;
            }
          }
        }
      } catch(e) {
        Logger.log('Drivers sheet lookup error: ' + e.message);
      }
    }

    Logger.log('Driver login: email=[' + email + '] driverId=[' + driverId + '] name=[' + driverName + ']');

    var allTrips = ops_getAllTrips_();
    var myTrips  = [];

    // ── Step 3a: PRIMARY — match by Driver_ID in DRIVER_EMP_ID column ──
    // Works for all trips created after this fix is deployed.
    if (driverId) {
      myTrips = allTrips.filter(function(t) {
        return String(t.driverEmpId || '').trim() === driverId;
      });
      Logger.log('Matched by driverId: ' + myTrips.length);
    }

    // ── Step 3b: FALLBACK — match by email in DRIVER_EMP_ID column ──
    // Catches trips saved during the period when email was stored there instead.
    if (myTrips.length === 0 && email) {
      var emailMatches = allTrips.filter(function(t) {
        return String(t.driverEmpId || '').trim().toLowerCase() === email;
      });
      if (emailMatches.length > 0) {
        Logger.log('Matched by email fallback: ' + emailMatches.length);
        myTrips = emailMatches;
      }
    }

    // ── Step 3c: FALLBACK — match by driver name string ──
    // Catches very old trips that only stored the name.
    if (myTrips.length === 0 && driverName) {
      var norm = function(s) { return String(s || '').toLowerCase().replace(/\s+/g, ' ').trim(); };
      var n    = norm(driverName);
      var nameMatches = allTrips.filter(function(t) { return norm(t.driverName) === n; });
      if (nameMatches.length > 0) {
        Logger.log('Matched by name fallback: ' + nameMatches.length);
        myTrips = nameMatches;
      }
    }

    Logger.log('Total matched trips: ' + myTrips.length);

    return {
      success    : true,
      trips      : myTrips,
      driverName : driverName,
      driverId   : driverId,
      driverEmail: email,
      user       : user
    };

  } catch(e) {
    Logger.log('getDriverDashboardData error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ============================================================
//  VEHICLES — CRUD
// ============================================================
function ops_getAllVehicles_() {
  const sh = ops_sh_(OPS_SHEETS.VEHICLES);
  const lr = sh.getLastRow();
  if (lr < 2) return [];
  return sh.getRange(2, 1, lr - 1, 13).getValues()
    .filter(function(r) { return r[VEH_COL.VEHICLE_ID] && String(r[VEH_COL.VEHICLE_ID]).trim(); })
    .map(function(r, i) {
      return {
        rowIndex  : i + 2,
        vehicleId : String(r[VEH_COL.VEHICLE_ID]).trim(),
        plate     : String(r[VEH_COL.PLATE]).trim(),
        type      : String(r[VEH_COL.TYPE]).trim(),
        brand     : String(r[VEH_COL.BRAND] || '').trim(),
        begMileage: parseFloat(r[VEH_COL.BEG_MILEAGE]) || 0,
        status    : String(r[VEH_COL.STATUS] || VEH_STATUS.ACTIVE).trim(),
        insExpiry : ops_fmtDate_(r[VEH_COL.INS_EXPIRY]),
        insLink   : String(r[VEH_COL.INS_LINK] || '').trim(),
        ltoExpiry : ops_fmtDate_(r[VEH_COL.LTO_EXPIRY]),
        ltoLink   : String(r[VEH_COL.LTO_LINK] || '').trim(),
        notes     : String(r[VEH_COL.NOTES] || '').trim(),
        createdAt : ops_fmtDT_(r[VEH_COL.CREATED_AT]),
        updatedAt : ops_fmtDT_(r[VEH_COL.UPDATED_AT])
      };
    });
}

function ops_addVehicle(payload) {
  if (!payload.plate || !payload.type)
    return { success: false, message: 'Plate Number and Vehicle Type are required.' };

  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch(e) {
    return { success: false, message: 'Server busy. Please try again in a moment.' };
  }

  try {
    var user = ops_getUserInfo_();
    if (!ops_isAdmin_(user.role) && !ops_isEncoder_(user.role))
      return { success: false, message: 'Access denied.' };

    var vehicles = ops_getAllVehicles_();
    if (vehicles.some(function(v) {
      return v.plate.toLowerCase() === payload.plate.trim().toLowerCase();
    })) return { success: false, message: 'Plate number "' + payload.plate + '" already exists.' };

    var sh  = ops_sh_(OPS_SHEETS.VEHICLES);
    var id  = ops_genId_('V', vehicles.map(function(v) { return [v.vehicleId]; }), 0);
    var now = ops_now_();

    sh.getRange(sh.getLastRow() + 1, 1, 1, 13).setValues([[
      id,
      payload.plate.trim().toUpperCase(),
      payload.type,
      payload.brand     || '',
      parseFloat(payload.begMileage) || 0,
      payload.status    || VEH_STATUS.ACTIVE,
      payload.insExpiry || '',
      payload.insLink   || '',
      payload.ltoExpiry || '',
      payload.ltoLink   || '',
      payload.notes     || '',
      now, now
    ]]);

    SpreadsheetApp.flush();
    ops_audit_('OPS_ADD_VEHICLE', { vehicleId: id, plate: payload.plate, by: user.email });
    return { success: true, message: 'Vehicle ' + id + ' added.', vehicleId: id };

  } catch(e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

function ops_updateVehicle(payload) {
  try {
    const user = ops_getUserInfo_();
    if (!ops_isAdmin_(user.role) && !ops_isEncoder_(user.role))
      return { success: false, message: 'Access denied.' };
    if (!payload.vehicleId) return { success: false, message: 'Vehicle ID required.' };

    const sh   = ops_sh_(OPS_SHEETS.VEHICLES);
    const lr   = sh.getLastRow();
    const data = sh.getRange(2, 1, lr - 1, 13).getValues();
    let rowIdx = -1;
    data.forEach(function(r, i) {
      if (String(r[VEH_COL.VEHICLE_ID]).trim() === payload.vehicleId) rowIdx = i + 2;
    });
    if (rowIdx === -1) return { success: false, message: 'Vehicle not found.' };

    sh.getRange(rowIdx, 1, 1, 13).setValues([[
      payload.vehicleId,
      payload.plate.trim().toUpperCase(),
      payload.type,
      payload.brand    || '',
      parseFloat(payload.begMileage) || 0,
      payload.status   || VEH_STATUS.ACTIVE,
      payload.insExpiry|| '',
      payload.insLink  || '',
      payload.ltoExpiry|| '',
      payload.ltoLink  || '',
      payload.notes    || '',
      data[rowIdx - 2][VEH_COL.CREATED_AT],
      ops_now_()
    ]]);

    ops_audit_('OPS_UPDATE_VEHICLE', { vehicleId: payload.vehicleId, by: user.email });
    return { success: true, message: 'Vehicle ' + payload.vehicleId + ' updated.' };
  } catch(e) { return { success: false, message: e.message }; }
}

function ops_deleteVehicle(vehicleId) {
  try {
    const user = ops_getUserInfo_();
    if (!ops_isAdmin_(user.role))
      return { success: false, message: 'Admin access required to delete vehicles.' };
    if (!vehicleId) return { success: false, message: 'Vehicle ID required.' };

    const sh = ops_sh_(OPS_SHEETS.VEHICLES);
    const lr = sh.getLastRow();
    if (lr < 2) return { success: false, message: 'No vehicles found.' };

    const data = sh.getRange(2, 1, lr - 1, 1).getValues();
    let rowIdx = -1;
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === vehicleId) { rowIdx = i + 2; break; }
    }
    if (rowIdx === -1) return { success: false, message: 'Vehicle ' + vehicleId + ' not found.' };

    sh.deleteRow(rowIdx);
    ops_audit_('OPS_DELETE_VEHICLE', { vehicleId, by: user.email });
    return { success: true, message: 'Vehicle ' + vehicleId + ' permanently deleted.' };
  } catch(e) { return { success: false, message: e.message }; }
}

// ============================================================
//  TRIPS — CRUD
// ============================================================
function ops_getAllTrips_() {
  const sh = ops_sh_(OPS_SHEETS.TRIPS);
  const lr = sh.getLastRow();
  if (lr < 2) return [];
  return sh.getRange(2, 1, lr - 1, 29).getValues()
    .filter(function(r) { return r[TRIP_COL.TRIP_ID] && String(r[TRIP_COL.TRIP_ID]).trim(); })
    .map(function(r, i) {
      return {
        rowIndex       : i + 2,
        tripId         : String(r[TRIP_COL.TRIP_ID]).trim(),
        requestDate    : ops_fmtDT_(r[TRIP_COL.REQUEST_DATE]),
        requestDateISO : ops_toISO_(r[TRIP_COL.REQUEST_DATE]),
        reqEmpId       : String(r[TRIP_COL.REQ_EMP_ID]   || '').trim(),
        reqName        : String(r[TRIP_COL.REQ_NAME]      || '').trim(),
        tripType       : String(r[TRIP_COL.TRIP_TYPE]     || '').trim(),
        purpose        : String(r[TRIP_COL.PURPOSE]       || '').trim(),
        relatedJo      : String(r[TRIP_COL.RELATED_JO]    || '').trim(),
        fromLoc        : String(r[TRIP_COL.FROM_LOC]      || '').trim(),
        toLoc          : String(r[TRIP_COL.TO_LOC]        || '').trim(),
        startDate      : ops_fmtDT_(r[TRIP_COL.START_DATE]),
        endDate        : ops_fmtDT_(r[TRIP_COL.END_DATE]),
        startDateISO   : ops_toISO_(r[TRIP_COL.START_DATE]),
        endDateISO     : ops_toISO_(r[TRIP_COL.END_DATE]),
        vehicleId      : String(r[TRIP_COL.VEHICLE_ID]    || '').trim(),
        plate          : String(r[TRIP_COL.PLATE]         || '').trim(),
        driverEmpId    : String(r[TRIP_COL.DRIVER_EMP_ID] || '').trim(),
        driverName     : String(r[TRIP_COL.DRIVER_NAME]   || '').trim(),
        status         : String(r[TRIP_COL.STATUS]        || TRIP_STATUS.DRAFT).trim(),
        approvedBy     : String(r[TRIP_COL.APPROVED_BY]   || '').trim(),
        approvalDate   : ops_fmtDT_(r[TRIP_COL.APPROVAL_DATE]),
        rejectReason   : String(r[TRIP_COL.REJECT_REASON] || '').trim(),
        cancelReason   : String(r[TRIP_COL.CANCEL_REASON] || '').trim(),
        actualStart    : ops_fmtDT_(r[TRIP_COL.ACTUAL_START]),
        actualEnd      : ops_fmtDT_(r[TRIP_COL.ACTUAL_END]),
        startKm        : parseFloat(r[TRIP_COL.START_KM])  || 0,
        endKm          : parseFloat(r[TRIP_COL.END_KM])    || 0,
        distance       : parseFloat(r[TRIP_COL.DISTANCE])  || 0,
        proofLink      : String(r[TRIP_COL.PROOF_LINK]    || '').trim(),
        remarks        : String(r[TRIP_COL.REMARKS]       || '').trim(),
        updatedAt      : ops_fmtDT_(r[TRIP_COL.UPDATED_AT]),
        updatedBy      : String(r[TRIP_COL.UPDATED_BY]    || '').trim()
      };
    });
}

function ops_saveTrip(payload) {
  if (!payload.reqName)   return { success: false, message: 'Requestor Name required.' };
  if (!payload.tripType)  return { success: false, message: 'Trip Type required.' };
  if (!payload.purpose)   return { success: false, message: 'Purpose required.' };
  if (!payload.fromLoc)   return { success: false, message: 'From Location required.' };
  if (!payload.toLoc)     return { success: false, message: 'To Location required.' };
  if (!payload.startDate) return { success: false, message: 'Planned Start required.' };
  if (!payload.endDate)   return { success: false, message: 'Planned End required.' };

  if (payload.status === TRIP_STATUS.SUBMITTED) {
    if (!payload.vehicleId)  return { success: false, message: 'Vehicle required to submit.' };
    if (!payload.driverName) return { success: false, message: 'Driver required to submit.' };
  }

  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch(e) {
    return { success: false, message: 'Server busy. Please try again in a moment.' };
  }

  try {
    var user  = ops_getUserInfo_();
    var trips = ops_getAllTrips_();
    var sh    = ops_sh_(OPS_SHEETS.TRIPS);
    var id    = ops_genId_('T', trips.map(function(t) { return [t.tripId]; }), 0);
    var now   = ops_now_();

    sh.getRange(sh.getLastRow() + 1, 1, 1, 29).setValues([[
      id, now,
      payload.reqEmpId    || '',
      payload.reqName     || '',
      payload.tripType    || '',
      payload.purpose     || '',
      payload.relatedJo   || '',
      payload.fromLoc     || '',
      payload.toLoc       || '',
      payload.startDate   || '',
      payload.endDate     || '',
      payload.vehicleId   || '',
      payload.plate       || '',
      payload.driverEmpId || '',
      payload.driverName  || '',
      payload.status      || TRIP_STATUS.DRAFT,
      '', '', '', '', '', '', 0, 0, 0,
      payload.proofLink   || '',
      payload.remarks     || '',
      now, user.email
    ]]);

    SpreadsheetApp.flush();
    ops_audit_('OPS_SAVE_TRIP', { tripId: id, status: payload.status, by: user.email });
    return { success: true, message: 'Trip ' + id + ' saved as ' + payload.status + '.', tripId: id };

  } catch(e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

function ops_approveTrip(tripId) {
  try {
    const user = ops_getUserInfo_();
    if (!ops_isApprover_(user.role))
      return { success: false, message: 'Approver access required.' };

    const row = ops_getTripRow_(tripId);
    if (!row) return { success: false, message: 'Trip not found.' };
    if (row.data[TRIP_COL.STATUS] !== TRIP_STATUS.SUBMITTED)
      return { success: false, message: 'Trip must be Submitted to approve.' };

    const sh  = ops_sh_(OPS_SHEETS.TRIPS);
    var now = ops_now_();
    sh.getRange(row.idx, TRIP_COL.STATUS        + 1).setValue(TRIP_STATUS.APPROVED);
    sh.getRange(row.idx, TRIP_COL.APPROVED_BY   + 1).setValue(user.email);
    sh.getRange(row.idx, TRIP_COL.APPROVAL_DATE + 1).setValue(now);
    sh.getRange(row.idx, TRIP_COL.UPDATED_AT    + 1).setValue(now);
    sh.getRange(row.idx, TRIP_COL.UPDATED_BY    + 1).setValue(user.email);

    ops_audit_('OPS_APPROVE_TRIP', { tripId, by: user.email });
    return { success: true, message: 'Trip ' + tripId + ' approved.' };
  } catch(e) { return { success: false, message: e.message }; }
}

function ops_rejectTrip(tripId, reason) {
  try {
    const user = ops_getUserInfo_();
    if (!ops_isApprover_(user.role))
      return { success: false, message: 'Approver access required.' };
    if (!reason || !reason.trim())
      return { success: false, message: 'Rejection reason required.' };

    const row = ops_getTripRow_(tripId);
    if (!row) return { success: false, message: 'Trip not found.' };

    const sh  = ops_sh_(OPS_SHEETS.TRIPS);
    var now = ops_now_();
    sh.getRange(row.idx, TRIP_COL.STATUS        + 1).setValue(TRIP_STATUS.REJECTED);
    sh.getRange(row.idx, TRIP_COL.REJECT_REASON + 1).setValue(reason.trim());
    sh.getRange(row.idx, TRIP_COL.APPROVED_BY   + 1).setValue(user.email);
    sh.getRange(row.idx, TRIP_COL.APPROVAL_DATE + 1).setValue(now);
    sh.getRange(row.idx, TRIP_COL.UPDATED_AT    + 1).setValue(now);
    sh.getRange(row.idx, TRIP_COL.UPDATED_BY    + 1).setValue(user.email);

    ops_audit_('OPS_REJECT_TRIP', { tripId, reason, by: user.email });
    return { success: true, message: 'Trip ' + tripId + ' rejected.' };
  } catch(e) { return { success: false, message: e.message }; }
}

function ops_cancelTrip(tripId, reason) {
  try {
    const user = ops_getUserInfo_();
    if (!reason || !reason.trim())
      return { success: false, message: 'Cancel reason required.' };

    const row = ops_getTripRow_(tripId);
    if (!row) return { success: false, message: 'Trip not found.' };

    const allowed = [TRIP_STATUS.DRAFT, TRIP_STATUS.SUBMITTED, TRIP_STATUS.APPROVED];
    if (!allowed.includes(row.data[TRIP_COL.STATUS]))
      return { success: false, message: 'Cannot cancel a ' + row.data[TRIP_COL.STATUS] + ' trip.' };

    const sh  = ops_sh_(OPS_SHEETS.TRIPS);
    var now = ops_now_();
    sh.getRange(row.idx, TRIP_COL.STATUS        + 1).setValue(TRIP_STATUS.CANCELLED);
    sh.getRange(row.idx, TRIP_COL.CANCEL_REASON + 1).setValue(reason.trim());
    sh.getRange(row.idx, TRIP_COL.UPDATED_AT    + 1).setValue(now);
    sh.getRange(row.idx, TRIP_COL.UPDATED_BY    + 1).setValue(user.email);

    ops_audit_('OPS_CANCEL_TRIP', { tripId, reason, by: user.email });
    return { success: true, message: 'Trip ' + tripId + ' cancelled.' };
  } catch(e) { return { success: false, message: e.message }; }
}

function ops_completeTrip(payload) {
  if (!payload.tripId)      return { success: false, message: 'Trip ID required.' };
  if (!payload.actualStart) return { success: false, message: 'Actual Start required.' };
  if (!payload.actualEnd)   return { success: false, message: 'Actual End required.' };

  var startKm = parseFloat(payload.startKm) || 0;
  var endKm   = parseFloat(payload.endKm)   || 0;
  if (isNaN(startKm) || isNaN(endKm))
    return { success: false, message: 'Mileage must be a number.' };
  if (endKm < startKm)
    return { success: false, message: 'End mileage cannot be less than start mileage.' };

  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch(e) {
    return { success: false, message: 'Server busy. Please try again in a moment.' };
  }

  try {
    var user = ops_getUserInfo_();
    if (!ops_isEncoder_(user.role))
      return { success: false, message: 'Encoder access required.' };

    var row = ops_getTripRow_(payload.tripId);
    if (!row) return { success: false, message: 'Trip not found.' };

    var currentStatus = row.data[TRIP_COL.STATUS];
    if (currentStatus === TRIP_STATUS.COMPLETED)
      return { success: false, message: 'Trip ' + payload.tripId + ' is already completed.' };
    if (currentStatus !== TRIP_STATUS.APPROVED)
      return { success: false, message: 'Trip must be Approved before completing. Current status: ' + currentStatus };

    var distance = endKm - startKm;
    var sh       = ops_sh_(OPS_SHEETS.TRIPS);
    var now      = ops_now_();

    sh.getRange(row.idx, TRIP_COL.STATUS       + 1).setValue(TRIP_STATUS.COMPLETED);
    sh.getRange(row.idx, TRIP_COL.ACTUAL_START + 1).setValue(payload.actualStart);
    sh.getRange(row.idx, TRIP_COL.ACTUAL_END   + 1).setValue(payload.actualEnd);
    sh.getRange(row.idx, TRIP_COL.START_KM     + 1).setValue(startKm);
    sh.getRange(row.idx, TRIP_COL.END_KM       + 1).setValue(endKm);
    sh.getRange(row.idx, TRIP_COL.DISTANCE     + 1).setValue(distance);
    sh.getRange(row.idx, TRIP_COL.PROOF_LINK   + 1).setValue(payload.proofLink || '');
    sh.getRange(row.idx, TRIP_COL.REMARKS      + 1).setValue(payload.remarks   || '');
    sh.getRange(row.idx, TRIP_COL.UPDATED_AT   + 1).setValue(now);
    sh.getRange(row.idx, TRIP_COL.UPDATED_BY   + 1).setValue(user.email);

    SpreadsheetApp.flush();
    ops_audit_('OPS_COMPLETE_TRIP', { tripId: payload.tripId, distance, by: user.email, via: 'ops' });
    return { success: true, message: 'Trip ' + payload.tripId + ' marked Completed. Distance: ' + distance + ' km.' };

  } catch(e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
//  DRIVER COMPLETE TRIP
//  ✅ FIXED: Triple-fallback ownership check matches dashboard
// ============================================================
function ops_driverCompleteTrip(payload) {
  if (!payload.tripId)      return { success: false, message: 'Trip ID required.' };
  if (!payload.actualStart) return { success: false, message: 'Actual Start required.' };
  if (!payload.actualEnd)   return { success: false, message: 'Actual End required.' };

  var startKm = parseFloat(payload.startKm) || 0;
  var endKm   = parseFloat(payload.endKm)   || 0;
  if (isNaN(startKm) || isNaN(endKm))
    return { success: false, message: 'Mileage must be a number.' };
  if (endKm < startKm)
    return { success: false, message: 'End mileage cannot be less than start mileage.' };

  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch(e) {
    return { success: false, message: 'Server busy. Please try again in a moment.' };
  }

  try {
    var role = String(payload.sessionEmail ? 'driver' : '').toLowerCase();
    var user = { email: String(payload.sessionEmail || '').toLowerCase().trim(), role: 'driver' };

    if (!user.email)
      return { success: false, message: 'Session expired. Please log in again.' };

    var row = ops_getTripRow_(payload.tripId);
    if (!row) return { success: false, message: 'Trip not found.' };

    var currentStatus = row.data[TRIP_COL.STATUS];
    if (currentStatus === TRIP_STATUS.COMPLETED)
      return { success: false, message: 'Trip ' + payload.tripId + ' is already completed.' };
    if (currentStatus !== TRIP_STATUS.APPROVED)
      return { success: false, message: 'Trip must be Approved before completing. Current status: ' + currentStatus };

    // ── Resolve this driver's Driver_ID and name from LoginUsers + Drivers ──
    var myEmail    = user.email.toLowerCase().trim();
    var myDriverId = '';
    var myDriverName = '';

    var ss      = SpreadsheetApp.getActiveSpreadsheet();
    var loginSh = ss.getSheetByName('LoginUsers');
    if (loginSh && loginSh.getLastRow() >= 2) {
      var loginData = loginSh.getRange(2, 1, loginSh.getLastRow() - 1, 4).getValues();
      for (var i = 0; i < loginData.length; i++) {
        if (String(loginData[i][0] || '').trim().toLowerCase() === myEmail) {
          myDriverId = String(loginData[i][3] || '').trim();
          break;
        }
      }
    }

    if (myDriverId) {
      var driverSh = ss.getSheetByName('Drivers');
      if (driverSh && driverSh.getLastRow() >= 2) {
        var driverData = driverSh.getRange(2, 1, driverSh.getLastRow() - 1, 2).getValues();
        for (var k = 0; k < driverData.length; k++) {
          if (String(driverData[k][0] || '').trim() === myDriverId) {
            myDriverName = String(driverData[k][1] || '').trim();
            break;
          }
        }
      }
    }

    var tripDriverEmpId = String(row.data[TRIP_COL.DRIVER_EMP_ID] || '').trim();
    var tripDriverName  = String(row.data[TRIP_COL.DRIVER_NAME]   || '').trim();
    var norm = function(s) { return String(s || '').toLowerCase().replace(/\s+/g, ' ').trim(); };

    // Triple-fallback ownership check (mirrors getDriverDashboardData)
    var ownsTrip =
      (myDriverId   && tripDriverEmpId === myDriverId) ||
      (myEmail      && norm(tripDriverEmpId) === myEmail) ||
      (myDriverName && norm(tripDriverName) === norm(myDriverName));

    if (!ownsTrip) {
      Logger.log('Ownership failed: myDriverId=[' + myDriverId + '] tripEmpId=[' + tripDriverEmpId + '] myEmail=[' + myEmail + ']');
      return { success: false, message: 'Access denied. This trip is not assigned to you.' };
    }

    var distance = endKm - startKm;
    var sh       = ops_sh_(OPS_SHEETS.TRIPS);
    var now      = ops_now_();

    sh.getRange(row.idx, TRIP_COL.STATUS       + 1).setValue(TRIP_STATUS.COMPLETED);
    sh.getRange(row.idx, TRIP_COL.ACTUAL_START + 1).setValue(payload.actualStart);
    sh.getRange(row.idx, TRIP_COL.ACTUAL_END   + 1).setValue(payload.actualEnd);
    sh.getRange(row.idx, TRIP_COL.START_KM     + 1).setValue(startKm);
    sh.getRange(row.idx, TRIP_COL.END_KM       + 1).setValue(endKm);
    sh.getRange(row.idx, TRIP_COL.DISTANCE     + 1).setValue(distance);
    sh.getRange(row.idx, TRIP_COL.REMARKS      + 1).setValue(payload.remarks || '');
    sh.getRange(row.idx, TRIP_COL.UPDATED_AT   + 1).setValue(now);
    sh.getRange(row.idx, TRIP_COL.UPDATED_BY   + 1).setValue(user.email);

    SpreadsheetApp.flush();
    ops_audit_('DRIVER_COMPLETE_TRIP', { tripId: payload.tripId, distance, by: user.email, driverId: myDriverId, via: 'driver' });
    return { success: true, message: 'Trip ' + payload.tripId + ' completed! Distance: ' + distance + ' km.' };

  } catch(e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

function ops_getTripRow_(tripId) {
  const sh = ops_sh_(OPS_SHEETS.TRIPS);
  const lr = sh.getLastRow();
  if (lr < 2) return null;
  const data = sh.getRange(2, 1, lr - 1, 29).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][TRIP_COL.TRIP_ID]).trim() === tripId)
      return { idx: i + 2, data: data[i] };
  }
  return null;
}

// ============================================================
//  RENEWAL ALERTS
// ============================================================
function ops_buildRenewalAlerts_(vehicles, alertDays) {
  const rows = [];
  vehicles.forEach(function(v) {
    var insDays = v.insExpiry ? ops_daysLeft_(v.insExpiry) : null;
    rows.push({
      vehicleId  : v.vehicleId,
      plate      : v.plate,
      docType    : 'Insurance',
      expiry     : v.insExpiry || '—',
      daysLeft   : insDays,
      alertStatus: !v.insExpiry      ? 'No Date'
                 : insDays === null  ? 'No Date'
                 : insDays < 0      ? 'Expired'
                 : insDays <= alertDays ? 'Due in ' + insDays + ' days'
                 : 'OK'
    });
    var ltoDays = v.ltoExpiry ? ops_daysLeft_(v.ltoExpiry) : null;
    rows.push({
      vehicleId  : v.vehicleId,
      plate      : v.plate,
      docType    : 'LTO',
      expiry     : v.ltoExpiry || '—',
      daysLeft   : ltoDays,
      alertStatus: !v.ltoExpiry      ? 'No Date'
                 : ltoDays === null  ? 'No Date'
                 : ltoDays < 0      ? 'Expired'
                 : ltoDays <= alertDays ? 'Due in ' + ltoDays + ' days'
                 : 'OK'
    });
  });
  rows.sort(function(a, b) {
    var sa = a.daysLeft === null ? 9999 : a.daysLeft;
    var sb = b.daysLeft === null ? 9999 : b.daysLeft;
    return sa - sb;
  });
  return rows;
}

// ============================================================
//  REPORTS
// ============================================================
function ops_buildReports_(trips, vehicles) {
  const completed = trips.filter(function(t) { return t.status === TRIP_STATUS.COMPLETED; });

  const byVehicle = {};
  completed.forEach(function(t) {
    if (!byVehicle[t.plate]) byVehicle[t.plate] = { count: 0, km: 0 };
    byVehicle[t.plate].count++;
    byVehicle[t.plate].km += (parseFloat(t.distance) || 0);
  });

  const byDriver = {};
  completed.forEach(function(t) {
    const key = t.driverName || 'Unknown';
    if (!byDriver[key]) byDriver[key] = { count: 0, km: 0 };
    byDriver[key].count++;
    byDriver[key].km += (parseFloat(t.distance) || 0);
  });

  const byType = {};
  trips.forEach(function(t) {
    byType[t.tripType] = (byType[t.tripType] || 0) + 1;
  });

  const mileageSummary = vehicles.map(function(v) {
    const recorded = (byVehicle[v.plate] || {}).km || 0;
    return { plate: v.plate, brand: v.brand, begMileage: v.begMileage, recordedKm: recorded };
  });

  return { byVehicle, byDriver, byType, mileageSummary };
}

// ============================================================
//  DRIVERS — CRUD
// ============================================================
const DRIVER_COL = {
  DRIVER_ID      : 0,
  NAME           : 1,
  EMP_ID         : 2,
  LICENSE_ID     : 3,
  LICENSE_EXPIRY : 4,
  CONTACT        : 5,
  STATUS         : 6,
  NOTES          : 7
};

function ops_getAllDrivers_() {
  const sh = ops_sh_(OPS_SHEETS.DRIVERS);
  const lr = sh.getLastRow();
  if (lr < 2) return [];

  var loginMap = {};
  try {
    var loginSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LoginUsers');
    if (loginSh && loginSh.getLastRow() >= 2) {
      loginSh.getRange(2, 1, loginSh.getLastRow() - 1, 4).getValues().forEach(function(r) {
        if (String(r[2]).trim().toLowerCase() === 'driver') {
          loginMap[String(r[3]).trim()] = String(r[0]).trim(); // driverId → email
        }
      });
    }
  } catch(e) {}

  return sh.getRange(2, 1, lr - 1, 8).getValues()
    .filter(function(r) { return r[DRIVER_COL.DRIVER_ID] && String(r[DRIVER_COL.DRIVER_ID]).trim(); })
    .map(function(r) {
      var dId        = String(r[DRIVER_COL.DRIVER_ID] || '').trim();
      var loginEmail = loginMap[dId] || '';
      return {
        driverId      : String(r[DRIVER_COL.DRIVER_ID]).trim(),
        name          : String(r[DRIVER_COL.NAME]           || '').trim(),
        empId         : String(r[DRIVER_COL.EMP_ID]         || '').trim(),
        licenseId     : String(r[DRIVER_COL.LICENSE_ID]     || '').trim(),
        licenseExpiry : ops_fmtDate_(r[DRIVER_COL.LICENSE_EXPIRY]),
        contact       : String(r[DRIVER_COL.CONTACT]        || '').trim(),
        status        : String(r[DRIVER_COL.STATUS]         || 'Active').trim(),
        notes         : String(r[DRIVER_COL.NOTES]          || '').trim(),
        loginEmail    : loginEmail
      };
    });
}

function ops_addDriver(payload) {
  if (!payload.name)      return { success: false, message: 'Full Name required.' };
  if (!payload.licenseId) return { success: false, message: 'Driver License ID required.' };
  if (!payload.email)     return { success: false, message: 'Email required para sa driver account.' };
  if (!payload.password)  return { success: false, message: 'Password required para sa driver account.' };

  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch(e) {
    return { success: false, message: 'Server busy. Please try again in a moment.' };
  }

  try {
    var user = ops_getUserInfo_();
    if (!ops_isAdmin_(user.role) && !ops_isEncoder_(user.role))
      return { success: false, message: 'Access denied.' };

    var drivers = ops_getAllDrivers_();
    if (drivers.some(function(d) {
      return d.licenseId.toLowerCase() === payload.licenseId.trim().toLowerCase();
    })) return { success: false, message: 'License ID "' + payload.licenseId + '" already exists.' };

    var loginSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LoginUsers');
    if (!loginSh) {
      loginSh = SpreadsheetApp.getActiveSpreadsheet().insertSheet('LoginUsers');
      loginSh.getRange(1, 1, 1, 4).setValues([['Email', 'Password', 'Role', 'Driver_ID']]);
      loginSh.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#f8fafc');
      loginSh.setFrozenRows(1);
    }

    if (loginSh.getLastRow() >= 2) {
      var loginData = loginSh.getRange(2, 1, loginSh.getLastRow() - 1, 1).getValues();
      var emailExists = loginData.some(function(r) {
        return String(r[0]).trim().toLowerCase() === payload.email.trim().toLowerCase();
      });
      if (emailExists) return { success: false, message: 'Email "' + payload.email + '" already exists sa LoginUsers.' };
    }

    const sh  = ops_sh_(OPS_SHEETS.DRIVERS);
    const id  = ops_genId_('D', drivers.map(function(d) { return [d.driverId]; }), 0);

    sh.getRange(sh.getLastRow() + 1, 1, 1, 8).setValues([[
      id,
      payload.name.trim(),
      payload.empId         || '',
      payload.licenseId.trim().toUpperCase(),
      payload.licenseExpiry || '',
      payload.contact       || '',
      payload.status        || 'Active',
      payload.notes         || ''
    ]]);

    loginSh.getRange(loginSh.getLastRow() + 1, 1, 1, 4).setValues([[
      payload.email.trim().toLowerCase(),
      ops_hashPassword_(payload.password),
      'driver',
      id
    ]]);

    SpreadsheetApp.flush();
    ops_audit_('OPS_ADD_DRIVER', { driverId: id, name: payload.name, email: payload.email, by: user.email });
    return {
      success  : true,
      message  : 'Driver ' + id + ' added. Login account created for ' + payload.email + '.',
      driverId : id
    };

  } catch(e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

function ops_updateDriver(payload) {
  try {
    const user = ops_getUserInfo_();
    if (!ops_isAdmin_(user.role) && !ops_isEncoder_(user.role))
      return { success: false, message: 'Access denied.' };
    if (!payload.driverId)  return { success: false, message: 'Driver ID required.' };
    if (!payload.name)      return { success: false, message: 'Full Name required.' };
    if (!payload.licenseId) return { success: false, message: 'Driver License ID required.' };

    const sh   = ops_sh_(OPS_SHEETS.DRIVERS);
    const lr   = sh.getLastRow();
    if (lr < 2) return { success: false, message: 'No drivers found.' };
    const data = sh.getRange(2, 1, lr - 1, 10).getValues();
    let rowIdx = -1;
    data.forEach(function(r, i) {
      if (String(r[DRIVER_COL.DRIVER_ID]).trim() === payload.driverId) rowIdx = i + 2;
    });
    if (rowIdx === -1) return { success: false, message: 'Driver not found.' };

    sh.getRange(rowIdx, 1, 1, 8).setValues([[
      payload.driverId,
      payload.name.trim(),
      payload.empId         || '',
      payload.licenseId.trim().toUpperCase(),
      payload.licenseExpiry || '',
      payload.contact       || '',
      payload.status        || 'Active',
      payload.notes         || ''
    ]]);

    if (payload.email || payload.password) {
      var loginSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LoginUsers');
      if (loginSh && loginSh.getLastRow() >= 2) {
        var loginData = loginSh.getRange(2, 1, loginSh.getLastRow() - 1, 3).getValues();
        var found = false;
        for (var i = 0; i < loginData.length; i++) {
          var rowEmail = String(loginData[i][0] || '').toLowerCase();
          var rowRole  = String(loginData[i][2] || '').toLowerCase();
          if (rowRole === 'driver' && payload.email && rowEmail === payload.email.toLowerCase()) {
            found = true;
            if (payload.password) loginSh.getRange(i + 2, 2).setValue(ops_hashPassword_(payload.password));
            break;
          }
        }
        if (!found && payload.email && payload.password) {
          loginSh.getRange(loginSh.getLastRow() + 1, 1, 1, 4).setValues([[
            payload.email.trim().toLowerCase(),
            ops_hashPassword_(payload.password),
            'driver',
            payload.driverId
          ]]);
        }
      }
    }

    ops_audit_('OPS_UPDATE_DRIVER', { driverId: payload.driverId, by: user.email });
    return { success: true, message: 'Driver ' + payload.driverId + ' updated.' };
  } catch(e) { return { success: false, message: e.message }; }
}

function ops_deleteDriver(driverId) {
  try {
    const user = ops_getUserInfo_();
    if (!ops_isAdmin_(user.role))
      return { success: false, message: 'Admin access required to delete drivers.' };
    if (!driverId) return { success: false, message: 'Driver ID required.' };
    const sh = ops_sh_(OPS_SHEETS.DRIVERS);
    const lr = sh.getLastRow();
    if (lr < 2) return { success: false, message: 'No drivers found.' };
    const data = sh.getRange(2, 1, lr - 1, 1).getValues();
    let rowIdx = -1;
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === driverId) { rowIdx = i + 2; break; }
    }
    if (rowIdx === -1) return { success: false, message: 'Driver ' + driverId + ' not found.' };
    sh.deleteRow(rowIdx);
    ops_audit_('OPS_DELETE_DRIVER', { driverId, by: user.email });
    return { success: true, message: 'Driver ' + driverId + ' permanently deleted.' };
  } catch(e) { return { success: false, message: e.message }; }
}

// ============================================================
//  TRIP TYPES — CRUD
// ============================================================
function ops_getTripTypes_() {
  const sh = ops_sh_(OPS_SHEETS.SETTINGS);
  const lr = sh.getLastRow();
  if (lr < 2) return [];
  return sh.getRange(2, 1, lr - 1, 2).getValues()
    .map(function(r, i) { return { row: i + 2, key: String(r[0]).trim(), value: String(r[1]).trim() }; })
    .filter(function(r) { return r.key === 'trip_type' && r.value; });
}

function getTripTypesInitData() {
  try {
    const user  = ops_getUserInfo_();
    const types = ops_getTripTypes_().map(function(r) { return { row: r.row, value: r.value }; });
    return { success: true, user, types };
  } catch(e) { return { success: false, message: e.message }; }
}

function ops_addTripType(value) {
  try {
    const user = ops_getUserInfo_();
    if (!ops_isAdmin_(user.role) && !ops_isEncoder_(user.role))
      return { success: false, message: 'Access denied.' };
    value = String(value || '').trim();
    if (!value) return { success: false, message: 'Trip Type value required.' };

    const existing = ops_getTripTypes_();
    if (existing.some(function(r) { return r.value.toLowerCase() === value.toLowerCase(); }))
      return { success: false, message: '"' + value + '" already exists.' };

    const sh = ops_sh_(OPS_SHEETS.SETTINGS);
    sh.getRange(sh.getLastRow() + 1, 1, 1, 2).setValues([['trip_type', value]]);
    ops_audit_('ADD_TRIP_TYPE', { value, by: user.email });
    return { success: true, message: 'Trip Type "' + value + '" added.' };
  } catch(e) { return { success: false, message: e.message }; }
}

function ops_updateTripType(row, newValue) {
  try {
    const user = ops_getUserInfo_();
    if (!ops_isAdmin_(user.role) && !ops_isEncoder_(user.role))
      return { success: false, message: 'Access denied.' };
    newValue = String(newValue || '').trim();
    if (!newValue) return { success: false, message: 'Value required.' };
    row = parseInt(row);
    if (!row || row < 2) return { success: false, message: 'Invalid row.' };
    const sh = ops_sh_(OPS_SHEETS.SETTINGS);
    sh.getRange(row, 2).setValue(newValue);
    ops_audit_('UPDATE_TRIP_TYPE', { row, newValue, by: user.email });
    return { success: true, message: 'Trip Type updated.' };
  } catch(e) { return { success: false, message: e.message }; }
}

function ops_deleteTripType(row) {
  try {
    const user = ops_getUserInfo_();
    if (!ops_isAdmin_(user.role) && !ops_isEncoder_(user.role))
      return { success: false, message: 'Access denied.' };
    row = parseInt(row);
    if (!row || row < 2) return { success: false, message: 'Invalid row.' };

    const sh  = ops_sh_(OPS_SHEETS.SETTINGS);
    const key = String(sh.getRange(row, 1).getValue()).trim();
    if (key !== 'trip_type') return { success: false, message: 'Row is not a Trip Type.' };

    sh.deleteRow(row);
    ops_audit_('DELETE_TRIP_TYPE', { row, by: user.email });
    return { success: true, message: 'Trip Type deleted.' };
  } catch(e) { return { success: false, message: e.message }; }
}

// ============================================================
//  ONE-TIME MIGRATION — run once from Apps Script Editor
//  Fixes any existing trips where email was stored in DRIVER_EMP_ID
// ============================================================
function ops_fixExistingTripDriverIds() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var tripSh  = ss.getSheetByName('Trips');
  var loginSh = ss.getSheetByName('LoginUsers');

  if (!tripSh || !loginSh) { Logger.log('ERROR: Trips or LoginUsers sheet not found.'); return; }

  // Build email → driverId map
  var emailToId = {};
  if (loginSh.getLastRow() >= 2) {
    loginSh.getRange(2, 1, loginSh.getLastRow() - 1, 4).getValues().forEach(function(r) {
      var email    = String(r[0] || '').trim().toLowerCase();
      var role     = String(r[2] || '').trim().toLowerCase();
      var driverId = String(r[3] || '').trim();
      if (role === 'driver' && email && driverId) emailToId[email] = driverId;
    });
  }
  Logger.log('Email→ID map: ' + JSON.stringify(emailToId));

  if (tripSh.getLastRow() < 2) { Logger.log('No trips found.'); return; }

  var tripData = tripSh.getRange(2, 1, tripSh.getLastRow() - 1, 15).getValues();
  var fixed = 0, skipped = 0;

  tripData.forEach(function(row, i) {
    var currentEmpId = String(row[13] || '').trim(); // DRIVER_EMP_ID = col 14 (index 13)
    if (currentEmpId.includes('@')) {
      var mappedId = emailToId[currentEmpId.toLowerCase()];
      if (mappedId) {
        tripSh.getRange(i + 2, 14).setValue(mappedId);
        Logger.log('Fixed row ' + (i + 2) + ': ' + currentEmpId + ' → ' + mappedId);
        fixed++;
      } else {
        Logger.log('WARNING row ' + (i + 2) + ': "' + currentEmpId + '" not in LoginUsers — skipped.');
        skipped++;
      }
    } else {
      skipped++;
    }
  });

  SpreadsheetApp.flush();
  Logger.log('Done. Fixed: ' + fixed + ', Skipped/OK: ' + skipped);
}

// ============================================================
//  UTILITY — fix LoginUsers header if missing col 4
// ============================================================
function ops_fixLoginUsersHeader() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('LoginUsers');
  if (!sh) { Logger.log('LoginUsers not found.'); return; }
  sh.getRange(1, 1, 1, 4).setValues([['Email', 'Password', 'Role', 'Driver_ID']]);
  sh.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#f8fafc');
  sh.setFrozenRows(1);
  Logger.log('LoginUsers header fixed.');
}

function ops_fixSpreadsheetTimezone() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone('Asia/Manila');
  Logger.log('Spreadsheet timezone updated to: ' + ss.getSpreadsheetTimeZone());
}

function ops_migratePasswordsToHash() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var sh  = ss.getSheetByName('LoginUsers');
  if (!sh) { Logger.log('LoginUsers sheet not found.'); return; }
  var lr = sh.getLastRow();
  if (lr < 2) { Logger.log('No users to migrate.'); return; }
  var data     = sh.getRange(2, 1, lr - 1, 2).getValues();
  var migrated = 0, skipped = 0;
  for (var i = 0; i < data.length; i++) {
    var rowPw = String(data[i][1] || '').trim();
    if (/^[0-9a-f]{64}$/.test(rowPw)) { skipped++; continue; }
    sh.getRange(i + 2, 2).setValue(ops_hashPassword_(rowPw));
    migrated++;
  }
  SpreadsheetApp.flush();
  Logger.log('Migration complete. Migrated: ' + migrated + ', Already hashed: ' + skipped);
}

function ops_resetPassword(targetEmail, newPassword) {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var sh  = ss.getSheetByName('LoginUsers');
  if (!sh) { Logger.log('LoginUsers sheet not found.'); return; }
  var lr   = sh.getLastRow();
  if (lr < 2) { Logger.log('No users found.'); return; }
  var data  = sh.getRange(2, 1, lr - 1, 1).getValues();
  var email = String(targetEmail || '').trim().toLowerCase();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0] || '').trim().toLowerCase() === email) {
      sh.getRange(i + 2, 2).setValue(ops_hashPassword_(newPassword));
      SpreadsheetApp.flush();
      Logger.log('Password reset successful for: ' + email);
      return;
    }
  }
  Logger.log('Email not found: ' + email);
}

function seedTripTypes() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  var types = ['Owner Errand','Supplier Delivery - Ormoc','Supplier Delivery - Outside Ormoc','Signage Installation','Other'];
  types.forEach(function(t) { sh.appendRow(['trip_type', t]); });
}