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
  if (!val) return '';
  try {
    var s = String(val).trim();
    if (!s || s === '0' || s === '') return '';

    // Already YYYY-MM-DD string — return as-is (from HTML date input)
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;

    // GAS Date object — use local getters to avoid UTC shift
    var d = (val instanceof Date) ? val : new Date(s);
    if (isNaN(d.getTime())) return '';
    var year = d.getFullYear();
    if (year < 1900 || year > 2200) return '';
    return year + '-'
      + String(d.getMonth() + 1).padStart(2, '0') + '-'
      + String(d.getDate()).padStart(2, '0');
  } catch(e) { return ''; }
}

function ops_fmtDT_(val) {
  if (!val) return '';
  try {
    const d = (val instanceof Date) ? val : new Date(val);
    if (isNaN(d.getTime())) return '';
    const mo = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    return mo[d.getMonth()] + ' ' + String(d.getDate()).padStart(2,'0')
      + ', ' + d.getFullYear() + ' '
      + String(d.getHours()).padStart(2,'0') + ':'
      + String(d.getMinutes()).padStart(2,'0');
  } catch(e) { return ''; }
}

function ops_daysLeft_(dateStr) {
  if (!dateStr || dateStr === '') return null;
  try {
    // dateStr is YYYY-MM-DD from ops_fmtDate_
    var parts = String(dateStr).split('-');
    if (parts.length !== 3) return null;
    var expiry = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
    var today  = new Date();
    today.setHours(0, 0, 0, 0);
    return Math.ceil((expiry - today) / 86400000);
  } catch(e) { return null; }
}

// DEBUG HELPER — run this in GAS editor to check raw sheet values
function ops_debugVehicles() {
  var sh = ops_sh_(OPS_SHEETS.VEHICLES);
  var lr = sh.getLastRow();
  if (lr < 2) { Logger.log("No data"); return; }
  var data = sh.getRange(2, 1, lr - 1, 13).getValues();
  data.forEach(function(r) {
    Logger.log(JSON.stringify({
      id: r[0], plate: r[1],
      insRaw: String(r[6]), insType: typeof r[6], insFmt: ops_fmtDate_(r[6]),
      ltoRaw: String(r[8]), ltoType: typeof r[8], ltoFmt: ops_fmtDate_(r[8])
    }));
  });
}

// ============================================================
//  ID GENERATOR
// ============================================================
function ops_genId_(prefix, rows, col) {
  const year = new Date().getFullYear();
  let max = 0;
  rows.forEach(function(r) {
    const id = String(r[col] || '');
    const m  = id.match(/-(\d{4})$/);
    if (m) { const n = parseInt(m[1]); if (n > max) max = n; }
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
        const role     = String(data[i][0] || '').trim();
        const emails   = String(data[i][1] || '').toLowerCase().split(',').map(function(e) { return e.trim(); });
        const abilities= String(data[i][2] || '').toLowerCase().split(',').map(function(a) { return a.trim(); });
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



// LOGIN — verifies email + password against LoginUsers sheet
function ops_loginUser(email, password) {
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LoginUsers');
    if (!sh) return { success: false, message: 'LoginUsers sheet not found. Please contact admin.' };

    var lr = sh.getLastRow();
    if (lr < 2) return { success: false, message: 'No users registered yet.' };

    // Read columns: A = Email, B = Password, C = Role
    var data = sh.getRange(2, 1, lr - 1, 3).getValues();

    var inputEmail = String(email || '').trim().toLowerCase();
    var inputPw    = String(password || '').trim();

    for (var i = 0; i < data.length; i++) {
      var rowEmail = String(data[i][0] || '').trim().toLowerCase();
      var rowPw    = String(data[i][1] || '').trim();
      var rowRole  = String(data[i][2] || '').trim();

      if (rowEmail === inputEmail) {
        if (rowPw === inputPw) {
          return {
            success : true,
            email   : rowEmail,
            role    : rowRole || 'No Role',
            message : 'Login successful.'
          };
        } else {
          return { success: false, message: 'Incorrect password. Please try again.' };
        }
      }
    }

    return { success: false, message: 'Email not found. Please check your email or contact admin.' };
  } catch(e) {
    return { success: false, message: 'Login error: ' + e.message };
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
      new Date(), action, user.email, user.role, JSON.stringify(payload)
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
//  JO DATABASE — fetch from external linked spreadsheet
// ============================================================
// ============================================================
//  DATABASE LINK HELPER — extracts Spreadsheet ID from DatabaseLink sheet
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

function ops_getJOList() {
  try {
    // 1. Get JODatabase spreadsheet ID from DatabaseLink sheet
    let joDbId;
    try {
      joDbId = ops_getDBId_('JODatabase');
    } catch(e) {
      return { success: false, message: 'DatabaseLink error: ' + e.message };
    }

    // 2. Open external spreadsheet by ID — more reliable than openByUrl
    let extSS;
    try {
      extSS = SpreadsheetApp.openById(joDbId);
    } catch(e) {
      return { success: false, message: 'Cannot open JODatabase (ID: ' + joDbId + '). Check sharing permissions: ' + e.message };
    }

    // 3. Get Line-up JOs sheet
    const joSh = extSS.getSheetByName('Line-up JOs');
    if (!joSh) {
      const shNames = extSS.getSheets().map(function(s) { return s.getName(); });
      return { success: false, message: '"Line-up JOs" not found. Available sheets: ' + shNames.join(', ') };
    }

    // 4. Read Column I (index 8) = Job Description, Column L (index 11) = JO Number
    const lr = joSh.getLastRow();
    if (lr < 2) return { success: true, data: [] };

    const data = joSh.getRange(2, 1, lr - 1, 12).getValues(); // cols A–L

    const list = [];
    data.forEach(function(r) {
      const joNumber = String(r[11] || '').trim(); // Column L
      const jobDesc  = String(r[8]  || '').trim(); // Column I
      if (joNumber) list.push({ joNumber: joNumber, jobDesc: jobDesc });
    });

    return { success: true, data: list };
  } catch(e) {
    return { success: false, message: 'ops_getJOList error: ' + e.message };
  }
}

function ops_getEmployeeList() {
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('EmployeName');
    if (!sh) return { success: false, message: 'EmployeName sheet not found.' };
    var lr = sh.getLastRow();
    if (lr < 2) return { success: true, data: [] };
    var data = sh.getRange(2, 1, lr - 1, 3).getValues();
    var list = [];
    data.forEach(function(r) {
      var empId   = String(r[0] || '').trim(); // Column A: Employee Code
      var team    = String(r[1] || '').trim(); // Column B: Team
      var empName = String(r[2] || '').trim(); // Column C: Name of Employee
      if (empName) list.push({ empId: empId, empName: empName, team: team });
    });
    return { success: true, data: list };
  } catch(e) {
    return { success: false, message: 'ops_getEmployeeList error: ' + e.message };
  }
}


// ============================================================
//  COMBINED INIT DATA (one call per tab)
// ============================================================

// Dashboard init — stats + recent trips + renewal alerts
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

// Trips tab init — trips + vehicles (for dropdown)
function getTripsInitData() {
  try {
    var user      = ops_getUserInfo_();
    var trips     = ops_getAllTrips_();
    var vehicles  = ops_getAllVehicles_();
    var joResult  = ops_getJOList();
    var empResult = ops_getEmployeeList();
    var joList    = joResult.success  ? joResult.data  : [];
    var joError   = joResult.success  ? null            : joResult.message;
    var empList   = empResult.success ? empResult.data : [];
    return { success: true, user: user, trips: trips, vehicles: vehicles, joList: joList, joError: joError, empList: empList };
  } catch(e) { return { success: false, message: e.message }; }
}

// ============================================================
//  DRIVER DASHBOARD — I-PASTE SA Code.js
//  (i-add after getTripsInitData function)
// ============================================================

// Returns trips assigned to the logged-in driver
function getDriverDashboardData() {
  try {
    var user     = ops_getUserInfo_();
    var email    = user.email.toLowerCase().trim();
    var allTrips = ops_getAllTrips_();

    var driverName = '';
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();

      // ✅ STEP 1: Find the driver name by matching email in LoginUsers
      //    then use the ROW INDEX to find corresponding driver in Drivers sheet
      var loginSh = ss.getSheetByName('LoginUsers');
      var driverSh = ss.getSheetByName('Drivers');

      if (loginSh && driverSh && loginSh.getLastRow() >= 2 && driverSh.getLastRow() >= 2) {
        var loginData  = loginSh.getRange(2, 1, loginSh.getLastRow() - 1, 3).getValues();
        var driverData = driverSh.getRange(2, 1, driverSh.getLastRow() - 1, 8).getValues();

        // ✅ Get only driver-role logins in order
        var driverLogins = [];
        for (var i = 0; i < loginData.length; i++) {
          var rowRole = String(loginData[i][2] || '').trim().toLowerCase();
          if (rowRole === 'driver') {
            driverLogins.push({
              email : String(loginData[i][0] || '').trim().toLowerCase(),
              index : driverLogins.length  // nth driver login = nth driver in Drivers sheet
            });
          }
        }

        // ✅ Find which nth driver login matches this email
        var matchedIndex = -1;
        for (var j = 0; j < driverLogins.length; j++) {
          if (driverLogins[j].email === email) {
            matchedIndex = driverLogins[j].index;
            break;
          }
        }

        // ✅ Get the driver name from Drivers sheet using matched index
        if (matchedIndex >= 0 && driverData[matchedIndex]) {
          driverName = String(driverData[matchedIndex][1] || '').trim(); // col B = Full_Name
        }

        // ✅ FALLBACK: if index match fails, try matching by empId = email directly
        if (!driverName) {
          for (var k = 0; k < driverData.length; k++) {
            var empId = String(driverData[k][2] || '').trim().toLowerCase();
            if (empId === email) {
              driverName = String(driverData[k][1] || '').trim();
              break;
            }
          }
        }
      }
    } catch(e) {
      Logger.log('Driver name lookup error: ' + e.message);
    }

    // ✅ Filter trips by driverName (case-insensitive trim match)
    var myTrips = [];
    if (driverName) {
      myTrips = allTrips.filter(function(t) {
        return (t.driverName || '').trim().toLowerCase() === driverName.toLowerCase();
      });
    }

    // ✅ Debug log — check sa GAS Execution Logs kung mag-fail pa
    Logger.log('Driver lookup: email=' + email + ' | found name=[' + driverName + '] | trips=' + myTrips.length);

    // ✅ EXTRA DEBUG — i-log tanan nga driver names sa Trips sheet para makita ang exact spelling
    var allDriverNames = allTrips.map(function(t) { return '[' + t.driverName + ']'; }).join(', ');
    Logger.log('All driverNames in Trips sheet: ' + allDriverNames);

    return {
      success    : true,
      trips      : myTrips,
      driverName : driverName,
      driverEmail: email,
      user       : user
    };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// Driver completes a trip — same as ops_completeTrip but allows driver role
function ops_driverCompleteTrip(payload) {
  try {
    var user = ops_getUserInfo_();

    // Allow driver role in addition to encoder/admin
    var role = (user.role || '').toLowerCase();
    if (!ops_isEncoder_(user.role) && role !== 'driver') {
      return { success: false, message: 'Access denied. Driver or Encoder role required.' };
    }

    if (!payload.tripId)      return { success: false, message: 'Trip ID required.' };
    if (!payload.actualStart) return { success: false, message: 'Actual Start required.' };
    if (!payload.actualEnd)   return { success: false, message: 'Actual End required.' };

    var startKm = parseFloat(payload.startKm) || 0;
    var endKm   = parseFloat(payload.endKm)   || 0;
    if (endKm < startKm) return { success: false, message: 'End mileage cannot be less than start mileage.' };

    var row = ops_getTripRow_(payload.tripId);
    if (!row) return { success: false, message: 'Trip not found.' };
    if (row.data[TRIP_COL.STATUS] !== TRIP_STATUS.APPROVED) {
      return { success: false, message: 'Trip must be Approved before completing.' };
    }

    var distance = endKm - startKm;
    var sh  = ops_sh_(OPS_SHEETS.TRIPS);
    var now = new Date();

    sh.getRange(row.idx, TRIP_COL.STATUS       + 1).setValue(TRIP_STATUS.COMPLETED);
    sh.getRange(row.idx, TRIP_COL.ACTUAL_START  + 1).setValue(payload.actualStart);
    sh.getRange(row.idx, TRIP_COL.ACTUAL_END    + 1).setValue(payload.actualEnd);
    sh.getRange(row.idx, TRIP_COL.START_KM      + 1).setValue(startKm);
    sh.getRange(row.idx, TRIP_COL.END_KM        + 1).setValue(endKm);
    sh.getRange(row.idx, TRIP_COL.DISTANCE      + 1).setValue(distance);
    sh.getRange(row.idx, TRIP_COL.REMARKS       + 1).setValue(payload.remarks || '');
    sh.getRange(row.idx, TRIP_COL.UPDATED_AT    + 1).setValue(now);
    sh.getRange(row.idx, TRIP_COL.UPDATED_BY    + 1).setValue(user.email);

    ops_audit_('DRIVER_COMPLETE_TRIP', { tripId: payload.tripId, distance: distance, by: user.email });
    return {
      success  : true,
      message  : 'Trip ' + payload.tripId + ' completed! Distance: ' + distance + ' km.'
    };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// Vehicles tab init
function getVehiclesInitData() {
  try {
    const user     = ops_getUserInfo_();
    const vehicles = ops_getAllVehicles_();
    return { success: true, user, vehicles };
  } catch(e) { return { success: false, message: e.message }; }
}

// Approval tab init — submitted trips only
function getApprovalInitData() {
  try {
    const user  = ops_getUserInfo_();
    const trips = ops_getAllTrips_().filter(function(t) {
      return t.status === TRIP_STATUS.SUBMITTED;
    });
    return { success: true, user, trips };
  } catch(e) { return { success: false, message: e.message }; }
}

// Completion tab init — approved trips only
function getCompletionInitData() {
  try {
    const user  = ops_getUserInfo_();
    const trips = ops_getAllTrips_().filter(function(t) {
      return t.status === TRIP_STATUS.APPROVED;
    });
    return { success: true, user, trips };
  } catch(e) { return { success: false, message: e.message }; }
}

// Renewals tab init
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

// Reports tab init
function getReportsInitData() {
  try {
    const user     = ops_getUserInfo_();
    const trips    = ops_getAllTrips_();
    const vehicles = ops_getAllVehicles_();
    const report   = ops_buildReports_(trips, vehicles);
    return { success: true, user, report };
  } catch(e) { return { success: false, message: e.message }; }
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
  try {
    const user = ops_getUserInfo_();
    if (!ops_isAdmin_(user.role) && !ops_isEncoder_(user.role))
      return { success: false, message: 'Access denied.' };
    if (!payload.plate || !payload.type)
      return { success: false, message: 'Plate Number and Vehicle Type are required.' };

    const vehicles = ops_getAllVehicles_();
    if (vehicles.some(function(v) { return v.plate.toLowerCase() === payload.plate.trim().toLowerCase(); }))
      return { success: false, message: 'Plate number "' + payload.plate + '" already exists.' };

    const sh  = ops_sh_(OPS_SHEETS.VEHICLES);
    const id  = ops_genId_('V', vehicles.map(function(v) { return [v.vehicleId]; }), 0);
    const now = new Date();

    sh.getRange(sh.getLastRow() + 1, 1, 1, 13).setValues([[
      id,
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
      now, now
    ]]);

    ops_audit_('OPS_ADD_VEHICLE', { vehicleId: id, plate: payload.plate, by: user.email });
    return { success: true, message: 'Vehicle ' + id + ' added.', vehicleId: id };
  } catch(e) { return { success: false, message: e.message }; }
}

function ops_updateVehicle(payload) {
  try {
    const user = ops_getUserInfo_();
    if (!ops_isAdmin_(user.role) && !ops_isEncoder_(user.role))
      return { success: false, message: 'Access denied.' };
    if (!payload.vehicleId) return { success: false, message: 'Vehicle ID required.' };

    const sh  = ops_sh_(OPS_SHEETS.VEHICLES);
    const lr  = sh.getLastRow();
    const data= sh.getRange(2, 1, lr - 1, 13).getValues();
    let rowIdx= -1;
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
      new Date()
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
        rowIndex     : i + 2,
        tripId       : String(r[TRIP_COL.TRIP_ID]).trim(),
        requestDate  : ops_fmtDT_(r[TRIP_COL.REQUEST_DATE]),
        reqEmpId     : String(r[TRIP_COL.REQ_EMP_ID]  || '').trim(),
        reqName      : String(r[TRIP_COL.REQ_NAME]     || '').trim(),
        tripType     : String(r[TRIP_COL.TRIP_TYPE]    || '').trim(),
        purpose      : String(r[TRIP_COL.PURPOSE]      || '').trim(),
        relatedJo    : String(r[TRIP_COL.RELATED_JO]   || '').trim(),
        fromLoc      : String(r[TRIP_COL.FROM_LOC]     || '').trim(),
        toLoc        : String(r[TRIP_COL.TO_LOC]       || '').trim(),
        startDate    : ops_fmtDT_(r[TRIP_COL.START_DATE]),
        endDate      : ops_fmtDT_(r[TRIP_COL.END_DATE]),
        vehicleId    : String(r[TRIP_COL.VEHICLE_ID]   || '').trim(),
        plate        : String(r[TRIP_COL.PLATE]        || '').trim(),
        driverEmpId  : String(r[TRIP_COL.DRIVER_EMP_ID]|| '').trim(),
        driverName   : String(r[TRIP_COL.DRIVER_NAME]  || '').trim(),
        status       : String(r[TRIP_COL.STATUS]       || TRIP_STATUS.DRAFT).trim(),
        approvedBy   : String(r[TRIP_COL.APPROVED_BY]  || '').trim(),
        approvalDate : ops_fmtDT_(r[TRIP_COL.APPROVAL_DATE]),
        rejectReason : String(r[TRIP_COL.REJECT_REASON]|| '').trim(),
        cancelReason : String(r[TRIP_COL.CANCEL_REASON]|| '').trim(),
        actualStart  : ops_fmtDT_(r[TRIP_COL.ACTUAL_START]),
        actualEnd    : ops_fmtDT_(r[TRIP_COL.ACTUAL_END]),
        startKm      : parseFloat(r[TRIP_COL.START_KM]) || 0,
        endKm        : parseFloat(r[TRIP_COL.END_KM])   || 0,
        distance     : parseFloat(r[TRIP_COL.DISTANCE]) || 0,
        proofLink    : String(r[TRIP_COL.PROOF_LINK]   || '').trim(),
        remarks      : String(r[TRIP_COL.REMARKS]      || '').trim(),
        updatedAt    : ops_fmtDT_(r[TRIP_COL.UPDATED_AT]),
        updatedBy    : String(r[TRIP_COL.UPDATED_BY]   || '').trim()
      };
    });
}

function ops_saveTrip(payload) {
  try {
    const user = ops_getUserInfo_();
    // Validate required fields
    if (!payload.reqName)   return { success: false, message: 'Requestor Name required.' };
    if (!payload.tripType)  return { success: false, message: 'Trip Type required.' };
    if (!payload.purpose)   return { success: false, message: 'Purpose required.' };
    if (!payload.fromLoc)   return { success: false, message: 'From Location required.' };
    if (!payload.toLoc)     return { success: false, message: 'To Location required.' };
    if (!payload.startDate) return { success: false, message: 'Planned Start required.' };
    if (!payload.endDate)   return { success: false, message: 'Planned End required.' };

    // Extra validation if submitting
    if (payload.status === TRIP_STATUS.SUBMITTED) {
      if (!payload.vehicleId) return { success: false, message: 'Vehicle required to submit.' };
      if (!payload.driverName)return { success: false, message: 'Driver required to submit.' };
    }

    const trips = ops_getAllTrips_();
    const sh    = ops_sh_(OPS_SHEETS.TRIPS);
    const id    = ops_genId_('T', trips.map(function(t) { return [t.tripId]; }), 0);
    const now   = new Date();

    sh.getRange(sh.getLastRow() + 1, 1, 1, 29).setValues([[
      id, now,
      payload.reqEmpId   || '',
      payload.reqName    || '',
      payload.tripType   || '',
      payload.purpose    || '',
      payload.relatedJo  || '',
      payload.fromLoc    || '',
      payload.toLoc      || '',
      payload.startDate  || '',
      payload.endDate    || '',
      payload.vehicleId  || '',
      payload.plate      || '',
      payload.driverEmpId|| '',
      payload.driverName || '',
      payload.status     || TRIP_STATUS.DRAFT,
      '', '', '', '', '', '', 0, 0, 0,
      payload.proofLink  || '',
      payload.remarks    || '',
      now, user.email
    ]]);

    ops_audit_('OPS_SAVE_TRIP', { tripId: id, status: payload.status, by: user.email });
    return { success: true, message: 'Trip ' + id + ' saved as ' + payload.status + '.', tripId: id };
  } catch(e) { return { success: false, message: e.message }; }
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
    const now = new Date();
    sh.getRange(row.idx, TRIP_COL.STATUS       + 1).setValue(TRIP_STATUS.APPROVED);
    sh.getRange(row.idx, TRIP_COL.APPROVED_BY  + 1).setValue(user.email);
    sh.getRange(row.idx, TRIP_COL.APPROVAL_DATE+ 1).setValue(now);
    sh.getRange(row.idx, TRIP_COL.UPDATED_AT   + 1).setValue(now);
    sh.getRange(row.idx, TRIP_COL.UPDATED_BY   + 1).setValue(user.email);

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
    const now = new Date();
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
    const now = new Date();
    sh.getRange(row.idx, TRIP_COL.STATUS        + 1).setValue(TRIP_STATUS.CANCELLED);
    sh.getRange(row.idx, TRIP_COL.CANCEL_REASON + 1).setValue(reason.trim());
    sh.getRange(row.idx, TRIP_COL.UPDATED_AT    + 1).setValue(now);
    sh.getRange(row.idx, TRIP_COL.UPDATED_BY    + 1).setValue(user.email);

    ops_audit_('OPS_CANCEL_TRIP', { tripId, reason, by: user.email });
    return { success: true, message: 'Trip ' + tripId + ' cancelled.' };
  } catch(e) { return { success: false, message: e.message }; }
}

function ops_completeTrip(payload) {
  try {
    const user = ops_getUserInfo_();
    if (!ops_isEncoder_(user.role))
      return { success: false, message: 'Encoder access required.' };
    if (!payload.tripId)     return { success: false, message: 'Trip ID required.' };
    if (!payload.actualStart)return { success: false, message: 'Actual Start required.' };
    if (!payload.actualEnd)  return { success: false, message: 'Actual End required.' };

    const startKm = parseFloat(payload.startKm) || 0;
    const endKm   = parseFloat(payload.endKm)   || 0;
    if (isNaN(startKm) || isNaN(endKm)) return { success: false, message: 'Mileage must be a number.' };
    if (endKm < startKm) return { success: false, message: 'End mileage cannot be less than start mileage.' };

    const row = ops_getTripRow_(payload.tripId);
    if (!row) return { success: false, message: 'Trip not found.' };
    if (row.data[TRIP_COL.STATUS] !== TRIP_STATUS.APPROVED)
      return { success: false, message: 'Trip must be Approved before completing.' };

    const distance = endKm - startKm;
    const sh  = ops_sh_(OPS_SHEETS.TRIPS);
    const now = new Date();

    sh.getRange(row.idx, TRIP_COL.STATUS      + 1).setValue(TRIP_STATUS.COMPLETED);
    sh.getRange(row.idx, TRIP_COL.ACTUAL_START + 1).setValue(payload.actualStart);
    sh.getRange(row.idx, TRIP_COL.ACTUAL_END  + 1).setValue(payload.actualEnd);
    sh.getRange(row.idx, TRIP_COL.START_KM    + 1).setValue(startKm);
    sh.getRange(row.idx, TRIP_COL.END_KM      + 1).setValue(endKm);
    sh.getRange(row.idx, TRIP_COL.DISTANCE    + 1).setValue(distance);
    sh.getRange(row.idx, TRIP_COL.PROOF_LINK  + 1).setValue(payload.proofLink  || '');
    sh.getRange(row.idx, TRIP_COL.REMARKS     + 1).setValue(payload.remarks    || '');
    sh.getRange(row.idx, TRIP_COL.UPDATED_AT  + 1).setValue(now);
    sh.getRange(row.idx, TRIP_COL.UPDATED_BY  + 1).setValue(user.email);

    ops_audit_('OPS_COMPLETE_TRIP', { tripId: payload.tripId, distance, by: user.email });
    return { success: true, message: 'Trip ' + payload.tripId + ' marked Completed. Distance: ' + distance + ' km.' };
  } catch(e) { return { success: false, message: e.message }; }
}

// ── Trip row finder helper ──
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
    // Always add Insurance row — even if no expiry date set
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
    // Always add LTO row
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
  // Sort: Expired first, then Due Soon, then OK, then No Date
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

  // By vehicle
  const byVehicle = {};
  completed.forEach(function(t) {
    if (!byVehicle[t.plate]) byVehicle[t.plate] = { count: 0, km: 0 };
    byVehicle[t.plate].count++;
    byVehicle[t.plate].km += (parseFloat(t.distance) || 0);
  });

  // By driver
  const byDriver = {};
  completed.forEach(function(t) {
    const key = t.driverName || 'Unknown';
    if (!byDriver[key]) byDriver[key] = { count: 0, km: 0 };
    byDriver[key].count++;
    byDriver[key].km += (parseFloat(t.distance) || 0);
  });

  // By type — all trips
  const byType = {};
  trips.forEach(function(t) {
    byType[t.tripType] = (byType[t.tripType] || 0) + 1;
  });

  // Mileage summary
  const mileageSummary = vehicles.map(function(v) {
    const recorded = (byVehicle[v.plate] || {}).km || 0;
    return { plate: v.plate, brand: v.brand, begMileage: v.begMileage, recordedKm: recorded };
  });

  return { byVehicle, byDriver, byType, mileageSummary };
}

// ============================================================
//  DRIVERS — CRUD
//  I-paste kini sa PINAKA-UBOS sa Code.js
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

  // ✅ Load login emails from LoginUsers (role = driver)
  var loginMap = {};
  try {
    var loginSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LoginUsers');
    if (loginSh && loginSh.getLastRow() >= 2) {
      loginSh.getRange(2, 1, loginSh.getLastRow() - 1, 3).getValues().forEach(function(r) {
        if (String(r[2]).trim().toLowerCase() === 'driver') {
          loginMap[String(r[0]).trim().toLowerCase()] = String(r[0]).trim();
        }
      });
    }
  } catch(e) {}

  return sh.getRange(2, 1, lr - 1, 8).getValues()
    .filter(function(r) { return r[DRIVER_COL.DRIVER_ID] && String(r[DRIVER_COL.DRIVER_ID]).trim(); })
    .map(function(r) {
      var driverName = String(r[DRIVER_COL.NAME] || '').trim().toLowerCase();
      // Match login email by name (first word match)
      var loginEmail = '';
      Object.keys(loginMap).forEach(function(em) {
        if (em.indexOf(driverName.split(' ')[0]) > -1 || driverName.indexOf(em.split('@')[0]) > -1) {
          loginEmail = loginMap[em];
        }
      });
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
function getDriversInitData() {
  try {
    const user    = ops_getUserInfo_();
    const drivers = ops_getAllDrivers_();
    return { success: true, user, drivers };
  } catch(e) { return { success: false, message: e.message }; }
}

function ops_addDriver(payload) {
  try {
    const user = ops_getUserInfo_();
    if (!ops_isAdmin_(user.role) && !ops_isEncoder_(user.role))
      return { success: false, message: 'Access denied.' };
    if (!payload.name)      return { success: false, message: 'Full Name required.' };
    if (!payload.licenseId) return { success: false, message: 'Driver License ID required.' };
    if (!payload.email)     return { success: false, message: 'Email required para sa driver account.' };
    if (!payload.password)  return { success: false, message: 'Password required para sa driver account.' };

    const drivers = ops_getAllDrivers_();
    if (drivers.some(function(d) {
      return d.licenseId.toLowerCase() === payload.licenseId.trim().toLowerCase();
    })) return { success: false, message: 'License ID "' + payload.licenseId + '" already exists.' };

    var loginSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LoginUsers');
    if (!loginSh) {
      loginSh = SpreadsheetApp.getActiveSpreadsheet().insertSheet('LoginUsers');
      loginSh.getRange(1, 1, 1, 3).setValues([['Email', 'Password', 'Role']]);
      loginSh.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#f8fafc');
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
    const now = new Date();

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

    loginSh.getRange(loginSh.getLastRow() + 1, 1, 1, 3).setValues([[
      payload.email.trim().toLowerCase(),
      payload.password,
      'driver'
    ]]);

    ops_audit_('OPS_ADD_DRIVER', { driverId: id, name: payload.name, email: payload.email, by: user.email });
    return {
      success  : true,
      message  : 'Driver ' + id + ' added. Login account created for ' + payload.email + '.',
      driverId : id
    };
  } catch(e) { return { success: false, message: e.message }; }
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

    // ✅ Update LoginUsers if email or password provided
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
            if (payload.password) loginSh.getRange(i + 2, 2).setValue(payload.password);
            break;
          }
        }
        // If not found, insert new login row
        if (!found && payload.email && payload.password) {
          loginSh.getRange(loginSh.getLastRow() + 1, 1, 1, 3).setValues([[
            payload.email.trim().toLowerCase(),
            payload.password,
            'Driver'
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
//  PASTE THESE 2 THINGS IN YOUR Code.js
// ============================================================

// ── PART 1: New function — paste anywhere (e.g. after ops_getSettings_) ──
// Reads EmployeeName sheet: Col A = Employee Code, Col B = Team, Col C = Name of Employee

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
    Logger.log('ops_getEmployeList_ error: ' + e.message);
    return [];
  }
}


// ── PART 2: Replace your existing getTripsInitData() with this ──
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

// ============================================================
//  TRIP TYPES — CRUD (Settings sheet, key = 'trip_type')
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
    const types = ops_getTripTypes_().map(function(r) {
      return { row: r.row, value: r.value };
    });
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

    const existing = ops_getTripTypes_().filter(function(r) { return r.row !== row; });
    if (existing.some(function(r) { return r.value.toLowerCase() === newValue.toLowerCase(); }))
      return { success: false, message: '"' + newValue + '" already exists.' };

    const sh = ops_sh_(OPS_SHEETS.SETTINGS);
    sh.getRange(row, 2).setValue(newValue);
    ops_audit_('UPDATE_TRIP_TYPE', { row, newValue, by: user.email });
    return { success: true, message: 'Trip Type updated to "' + newValue + '".' };
  } catch(e) { return { success: false, message: e.message }; }
}

function ops_deleteTripType(row) {
  try {
    const user = ops_getUserInfo_();
    if (!ops_isAdmin_(user.role))
      return { success: false, message: 'Admin access required to delete Trip Types.' };
    row = parseInt(row);
    if (!row || row < 2) return { success: false, message: 'Invalid row.' };

    const sh = ops_sh_(OPS_SHEETS.SETTINGS);
    // Validate it's actually a trip_type row
    const key = String(sh.getRange(row, 1).getValue()).trim();
    if (key !== 'trip_type') return { success: false, message: 'Row is not a Trip Type.' };

    sh.deleteRow(row);
    ops_audit_('DELETE_TRIP_TYPE', { row, by: user.email });
    return { success: true, message: 'Trip Type deleted.' };
  } catch(e) { return { success: false, message: e.message }; }
}

function seedTripTypes() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  var types = [
    'Owner Errand',
    'Supplier Delivery - Ormoc',
    'Supplier Delivery - Outside Ormoc',
    'Signage Installation',
    'Other'
  ];
  types.forEach(function(t) { sh.appendRow(['trip_type', t]); });
}