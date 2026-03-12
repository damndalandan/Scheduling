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
  AUDIT    : 'Audit_Log'
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
    }
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
    const d = (val instanceof Date) ? val : new Date(val);
    if (isNaN(d.getTime())) return '';
    return d.getFullYear() + '-'
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
  if (!dateStr) return null;
  try {
    const diff = Math.ceil((new Date(dateStr) - new Date()) / 86400000);
    return diff;
  } catch(e) { return null; }
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
    const user     = ops_getUserInfo_();
    const trips    = ops_getAllTrips_();
    const vehicles = ops_getAllVehicles_();
    return { success: true, user, trips, vehicles };
  } catch(e) { return { success: false, message: e.message }; }
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
    if (v.insExpiry) {
      const days = ops_daysLeft_(v.insExpiry);
      rows.push({
        vehicleId: v.vehicleId, plate: v.plate,
        docType: 'Insurance', expiry: v.insExpiry,
        daysLeft: days,
        alertStatus: days === null ? 'No Date' : days < 0 ? 'Expired' : days <= alertDays ? 'Due in ' + days + ' days' : 'OK'
      });
    }
    if (v.ltoExpiry) {
      const days = ops_daysLeft_(v.ltoExpiry);
      rows.push({
        vehicleId: v.vehicleId, plate: v.plate,
        docType: 'LTO', expiry: v.ltoExpiry,
        daysLeft: days,
        alertStatus: days === null ? 'No Date' : days < 0 ? 'Expired' : days <= alertDays ? 'Due in ' + days + ' days' : 'OK'
      });
    }
  });
  rows.sort(function(a, b) { return (a.daysLeft === null ? 9999 : a.daysLeft) - (b.daysLeft === null ? 9999 : b.daysLeft); });
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