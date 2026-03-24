// ============================================================
//  OPS TRIP MONITOR — DATABASE SETUP & REPAIR SCRIPT
//  Ormoc Printshoppe
//
//  HOW TO USE:
//  1. Open your Google Spreadsheet
//  2. Go to Extensions → Apps Script
//  3. Paste this ENTIRE file into a new script file
//  4. Click Save, then Run → setupDatabase()
//  5. After it finishes, also run → repairExistingData()
//     to fix any Driver_EmpID mismatches in existing trips
//
//  SAFE TO RUN MULTIPLE TIMES — it skips sheets that already
//  exist and only adds missing columns/data.
// ============================================================


// ============================================================
//  MAIN ENTRY POINT — Run this first
// ============================================================
function setupDatabase() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var log = [];

  log.push('=== OPS Trip Monitor — Database Setup ===');
  log.push('Started: ' + new Date().toLocaleString());
  log.push('');

  _setupVehicles(ss, log);
  _setupTrips(ss, log);
  _setupDrivers(ss, log);
  _setupLoginUsers(ss, log);
  _setupSettings(ss, log);
  _setupRolePermissions(ss, log);
  _setupAuditLog(ss, log);
  _setupDatabaseLink(ss, log);
  _setupEmployeeName(ss, log);

  log.push('');
  log.push('=== Setup Complete ===');
  log.push('All sheets are ready. Now run repairExistingData() to fix any data issues.');

  Logger.log(log.join('\n'));
  SpreadsheetApp.getUi().alert('✅ Database Setup Complete!\n\n' + log.join('\n') + '\n\nNow also run repairExistingData() to fix existing trip data.');
}


// ============================================================
//  REPAIR — Run this after setupDatabase()
//  Fixes Driver_EmpID in Trips and Driver_ID in LoginUsers
// ============================================================
function repairExistingData() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var log = [];

  log.push('=== Repair Existing Data ===');
  log.push('Started: ' + new Date().toLocaleString());
  log.push('');

  _repairLoginUsersDriverId(ss, log);
  _repairTripDriverEmpId(ss, log);

  log.push('');
  log.push('=== Repair Complete ===');

  Logger.log(log.join('\n'));
  SpreadsheetApp.getUi().alert('✅ Repair Complete!\n\n' + log.join('\n'));
}


// ============================================================
//  SHEET: Vehicles
// ============================================================
function _setupVehicles(ss, log) {
  var name    = 'Vehicles';
  var headers = [
    'Vehicle_ID', 'Plate_Number', 'Vehicle_Type', 'Brand_Model',
    'Beginning_Mileage', 'Status', 'Insurance_Expiry', 'Insurance_PDF_Link',
    'LTO_Expiry', 'LTO_PDF_Link', 'Notes', 'Created_At', 'Updated_At'
  ];
  var widths = [110, 110, 100, 130, 120, 80, 130, 160, 110, 160, 150, 160, 160];

  var sh = _ensureSheet(ss, name, headers, widths, log);

  // Column validations
  _addDropdown(sh, 2, 6, 1000, ['Active', 'Under Repair', 'Inactive']);
  log.push('  ✅ Status dropdown added (Active / Under Repair / Inactive)');
}


// ============================================================
//  SHEET: Trips
// ============================================================
function _setupTrips(ss, log) {
  var name    = 'Trips';
  var headers = [
    'Trip_ID', 'Request_Date', 'Requestor_EmpID', 'Requestor_Name',
    'Trip_Type', 'Purpose', 'Related_JO', 'From_Location', 'To_Location',
    'Planned_Start', 'Planned_End', 'Vehicle_ID', 'Plate_Number',
    'Driver_EmpID', 'Driver_Name', 'Status',
    'Approved_By', 'Approval_Date', 'Rejection_Reason', 'Cancel_Reason',
    'Actual_Start', 'Actual_End', 'Start_Mileage', 'End_Mileage',
    'Distance_Travelled', 'GPS_Proof_Link', 'Remarks', 'Updated_At', 'Updated_By'
  ];
  var widths = [
    110, 160, 120, 140,
    160, 200, 110, 140, 140,
    140, 140, 110, 110,
    110, 140, 90,
    180, 160, 160, 160,
    140, 140, 110, 110,
    120, 160, 160, 160, 180
  ];

  var sh = _ensureSheet(ss, name, headers, widths, log);

  // Status dropdown
  _addDropdown(sh, 2, 16, 1000, ['Draft', 'Submitted', 'Approved', 'Rejected', 'Cancelled', 'Completed']);
  log.push('  ✅ Status dropdown added');
}


// ============================================================
//  SHEET: Drivers
//  IMPORTANT: Column D = Driver_ID (used to link to LoginUsers)
// ============================================================
function _setupDrivers(ss, log) {
  var name    = 'Drivers';
  var headers = [
    'Driver_ID', 'Full_Name', 'Employee_ID', 'License_ID',
    'License_Expiry', 'Contact_Number', 'Status', 'Notes'
  ];
  var widths = [110, 160, 100, 120, 120, 130, 80, 200];

  var sh = _ensureSheet(ss, name, headers, widths, log);

  _addDropdown(sh, 2, 7, 1000, ['Active', 'Inactive']);
  log.push('  ✅ Status dropdown added (Active / Inactive)');
  log.push('  ⚠️  Driver_ID here must match col D of LoginUsers for trip matching to work.');
}


// ============================================================
//  SHEET: LoginUsers
//  Col A = Email, B = Password (SHA-256 hash), C = Role, D = Driver_ID
// ============================================================
function _setupLoginUsers(ss, log) {
  var name    = 'LoginUsers';
  var headers = ['Email', 'Password', 'Role', 'Driver_ID'];
  var widths  = [200, 300, 80, 110];

  var sh = ss.getSheetByName(name);
  var isNew = !sh;

  sh = _ensureSheet(ss, name, headers, widths, log);

  // Role dropdown
  _addDropdown(sh, 2, 3, 1000, ['admin', 'driver', 'encoder', 'approver']);
  log.push('  ✅ Role dropdown added');
  log.push('  ℹ️  Password column stores SHA-256 hashes. Use ops_hashPassword_() to generate.');
  log.push('  ⚠️  Driver_ID (col D) MUST match Driver_ID in Drivers sheet for the driver to see their trips.');

  // Add a note to the header cell
  try {
    var noteCell = sh.getRange('D1');
    noteCell.setNote('This must match the Driver_ID in the Drivers sheet.\nExample: D-2026-0001\nLeave blank for non-driver roles.');
  } catch(e) {}
}


// ============================================================
//  SHEET: Settings
// ============================================================
function _setupSettings(ss, log) {
  var name    = 'Settings';
  var headers = ['Setting_Key', 'Setting_Value'];
  var widths  = [160, 300];

  var sh    = ss.getSheetByName(name);
  var isNew = !sh;
  sh        = _ensureSheet(ss, name, headers, widths, log);

  if (isNew || sh.getLastRow() < 2) {
    sh.getRange(2, 1, 7, 2).setValues([
      ['renewal_alert_days',              '30'],
      ['app_version',                      '1.0'],
      ['trip_type',                        'Owner Errand'],
      ['trip_type',                        'Supplier Delivery - Ormoc'],
      ['trip_type',                        'Supplier Delivery - Outside Ormoc'],
      ['trip_type',                        'Signage Installation'],
      ['trip_type',                        'Other']
    ]);
    log.push('  ✅ Default settings and trip types seeded');
  } else {
    log.push('  ℹ️  Settings already has data — skipped seeding');
  }
}


// ============================================================
//  SHEET: Role_Permissions
// ============================================================
function _setupRolePermissions(ss, log) {
  var name    = 'Role_Permissions';
  var headers = ['Role', 'Emails', 'Abilities'];
  var widths  = [100, 400, 200];

  var sh    = ss.getSheetByName(name);
  var isNew = !sh;
  sh        = _ensureSheet(ss, name, headers, widths, log);

  if (isNew || sh.getLastRow() < 2) {
    // Get the active user email as default admin
    var adminEmail = Session.getActiveUser().getEmail() || 'admin@yourdomain.com';
    sh.getRange(2, 1, 1, 3).setValues([[
      'admin',
      adminEmail,
      'admin,encoder,approver'
    ]]);
    log.push('  ✅ Admin role seeded with: ' + adminEmail);
    log.push('  ℹ️  Add more roles here: encoder, approver, driver (drivers use LoginUsers, not this sheet)');
  } else {
    log.push('  ℹ️  Role_Permissions already has data — skipped seeding');
  }
}


// ============================================================
//  SHEET: Audit_Log
// ============================================================
function _setupAuditLog(ss, log) {
  var name    = 'Audit_Log';
  var headers = ['DateTime', 'Action', 'User', 'Role', 'Payload'];
  var widths  = [160, 160, 200, 80, 400];
  _ensureSheet(ss, name, headers, widths, log);
}


// ============================================================
//  SHEET: DatabaseLink
// ============================================================
function _setupDatabaseLink(ss, log) {
  var name    = 'DatabaseLink';
  var headers = ['Name', 'Link'];
  var widths  = [130, 500];

  var sh    = ss.getSheetByName(name);
  var isNew = !sh;
  sh        = _ensureSheet(ss, name, headers, widths, log);

  if (isNew || sh.getLastRow() < 2) {
    sh.getRange(2, 1, 1, 2).setValues([
      ['JODatabase', 'PASTE_YOUR_JO_SPREADSHEET_URL_HERE']
    ]);
    log.push('  ⚠️  Paste your JO Database Google Sheets URL in cell B2 of DatabaseLink sheet!');
  } else {
    log.push('  ℹ️  DatabaseLink already has data — skipped seeding');
  }
}


// ============================================================
//  SHEET: EmployeName  (note: intentional spelling from original)
// ============================================================
function _setupEmployeeName(ss, log) {
  var name    = 'EmployeName';
  var headers = ['Employee Code', 'Team', 'NAME OF EMPLOYEE'];
  var widths  = [120, 130, 200];
  _ensureSheet(ss, name, headers, widths, log);
  log.push('  ℹ️  Add your employee list here for the requestor autocomplete dropdown.');
}


// ============================================================
//  REPAIR: Fix LoginUsers — ensure every driver row has Driver_ID
// ============================================================
function _repairLoginUsersDriverId(ss, log) {
  log.push('--- Repairing LoginUsers Driver_ID column ---');

  var loginSh  = ss.getSheetByName('LoginUsers');
  var driverSh = ss.getSheetByName('Drivers');

  if (!loginSh)  { log.push('  ❌ LoginUsers sheet not found'); return; }
  if (!driverSh) { log.push('  ❌ Drivers sheet not found');    return; }

  // Fix header if col D is blank or wrong
  var d1 = loginSh.getRange(1, 4).getValue();
  if (!d1 || String(d1).trim() === '') {
    loginSh.getRange(1, 4).setValue('Driver_ID');
    loginSh.getRange(1, 4).setFontWeight('bold').setBackground('#f8fafc');
    log.push('  ✅ Fixed LoginUsers header — added Driver_ID to column D');
  }

  var loginLr = loginSh.getLastRow();
  if (loginLr < 2) { log.push('  ℹ️  No users in LoginUsers yet'); return; }

  var loginData  = loginSh.getRange(2, 1, loginLr - 1, 4).getValues();

  // Build name → driverId lookup from Drivers sheet
  var driverLr   = driverSh.getLastRow();
  var driverData = driverLr >= 2 ? driverSh.getRange(2, 1, driverLr - 1, 8).getValues() : [];

  // Map: normalised name → driverId
  var nameToId = {};
  // Map: email prefix → driverId (kiyo@gmail.com → prefix "kiyo")
  var prefixToId = {};
  driverData.forEach(function(r) {
    var did  = String(r[0] || '').trim();
    var name = String(r[1] || '').trim().toLowerCase();
    if (did && name) {
      nameToId[name] = did;
      // Also index each word
      name.split(/\s+/).forEach(function(w) {
        if (w.length > 2) prefixToId[w] = did;
      });
    }
  });

  var fixed = 0;
  for (var i = 0; i < loginData.length; i++) {
    var email    = String(loginData[i][0] || '').trim().toLowerCase();
    var role     = String(loginData[i][2] || '').trim().toLowerCase();
    var existing = String(loginData[i][3] || '').trim();

    if (role !== 'driver') continue;
    if (existing && existing !== 'undefined' && existing !== 'nan') {
      log.push('  ✅ ' + email + ' already has Driver_ID: ' + existing);
      continue;
    }

    // Try to match by email prefix against driver names
    var prefix = email.split('@')[0].toLowerCase();
    var matched = nameToId[prefix] || prefixToId[prefix] || '';

    // If no match by prefix, try each word of prefix
    if (!matched) {
      prefix.split(/[._\-]/).forEach(function(part) {
        if (!matched && part.length > 2 && prefixToId[part]) matched = prefixToId[part];
      });
    }

    if (matched) {
      loginSh.getRange(i + 2, 4).setValue(matched);
      log.push('  ✅ ' + email + ' → assigned Driver_ID: ' + matched);
      fixed++;
    } else {
      log.push('  ⚠️  ' + email + ' — could not auto-match to a driver. Set Driver_ID manually in LoginUsers col D.');
    }
  }

  log.push('  Fixed ' + fixed + ' driver rows in LoginUsers');
}


// ============================================================
//  REPAIR: Fix Trips — normalize Driver_EmpID to Driver_ID format
// ============================================================
function _repairTripDriverEmpId(ss, log) {
  log.push('--- Repairing Trips Driver_EmpID column ---');

  var tripSh   = ss.getSheetByName('Trips');
  var loginSh  = ss.getSheetByName('LoginUsers');
  var driverSh = ss.getSheetByName('Drivers');

  if (!tripSh)   { log.push('  ❌ Trips sheet not found');     return; }
  if (!loginSh)  { log.push('  ❌ LoginUsers sheet not found'); return; }
  if (!driverSh) { log.push('  ❌ Drivers sheet not found');    return; }

  var tripLr = tripSh.getLastRow();
  if (tripLr < 2) { log.push('  ℹ️  No trips yet'); return; }

  // Build lookup maps from LoginUsers: email → driverId
  var loginLr   = loginSh.getLastRow();
  var loginData = loginLr >= 2 ? loginSh.getRange(2, 1, loginLr - 1, 4).getValues() : [];
  var emailToId = {};
  loginData.forEach(function(r) {
    var email = String(r[0] || '').trim().toLowerCase();
    var role  = String(r[2] || '').trim().toLowerCase();
    var did   = String(r[3] || '').trim();
    if (role === 'driver' && email && did && did !== 'nan') {
      emailToId[email] = did;
    }
  });

  // Build lookup: empId (numeric string) → driverId from Drivers sheet
  var driverLr   = driverSh.getLastRow();
  var driverData = driverLr >= 2 ? driverSh.getRange(2, 1, driverLr - 1, 8).getValues() : [];
  var empIdToId  = {};
  driverData.forEach(function(r) {
    var did   = String(r[0] || '').trim();
    var empId = String(r[2] || '').trim();
    if (did && empId) empIdToId[empId] = did;
  });

  var tripData = tripSh.getRange(2, 1, tripLr - 1, 15).getValues();
  var fixed = 0, skipped = 0, already = 0;

  for (var i = 0; i < tripData.length; i++) {
    var tripId = String(tripData[i][0] || '').trim();
    if (!tripId) continue;

    // Driver_EmpID is column 14, index 13
    var raw = String(tripData[i][13] || '').trim();
    if (!raw || raw === 'undefined' || raw === 'null' || raw === 'nan') {
      skipped++;
      continue;
    }

    // Already correct format D-YYYY-NNNN
    if (/^D-\d{4}-\d{4}$/.test(raw)) {
      log.push('  ✅ ' + tripId + ': Driver_EmpID already correct (' + raw + ')');
      already++;
      continue;
    }

    var newId = '';

    // Case 1: it's an email
    if (raw.indexOf('@') !== -1) {
      newId = emailToId[raw.toLowerCase()] || '';
      if (newId) log.push('  ✅ ' + tripId + ': email "' + raw + '" → ' + newId);
    }

    // Case 2: it's a numeric empId (e.g. "101", "101.0", "Emp-101")
    if (!newId) {
      var numOnly = raw.replace(/[^0-9]/g, '');
      if (numOnly) newId = empIdToId[numOnly] || '';
      if (newId) log.push('  ✅ ' + tripId + ': empId "' + raw + '" → ' + newId);
    }

    if (newId) {
      tripSh.getRange(i + 2, 14).setValue(newId);
      fixed++;
    } else {
      log.push('  ⚠️  ' + tripId + ': "' + raw + '" — no match found. Fix manually.');
      skipped++;
    }
  }

  SpreadsheetApp.flush();
  log.push('  Fixed: ' + fixed + ' | Already correct: ' + already + ' | Could not fix: ' + skipped);
}


// ============================================================
//  HELPER: Create sheet if missing, add headers + formatting
// ============================================================
function _ensureSheet(ss, name, headers, widths, log) {
  var sh = ss.getSheetByName(name);

  if (!sh) {
    sh = ss.insertSheet(name);
    log.push('  ✅ Created sheet: ' + name);
  } else {
    log.push('  ℹ️  Sheet already exists: ' + name + ' — checking headers...');
  }

  // Always ensure header row is correct
  var headerRange = sh.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold')
             .setBackground('#1e293b')
             .setFontColor('#f8fafc')
             .setFontSize(10)
             .setWrap(false);

  sh.setFrozenRows(1);

  // Set column widths
  if (widths && widths.length) {
    for (var i = 0; i < widths.length; i++) {
      sh.setColumnWidth(i + 1, widths[i]);
    }
  }

  // Alternate row banding
  try {
    var bandings = sh.getBandings();
    if (!bandings.length) {
      sh.getRange(1, 1, 1000, headers.length).applyRowBanding(
        SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false
      );
    }
  } catch(e) {}

  return sh;
}


// ============================================================
//  HELPER: Add data validation dropdown to a column range
// ============================================================
function _addDropdown(sh, startRow, col, numRows, values) {
  try {
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(values, true)
      .setAllowInvalid(true)
      .build();
    sh.getRange(startRow, col, numRows, 1).setDataValidation(rule);
  } catch(e) {
    // Non-critical — skip if it fails
  }
}


// ============================================================
//  UTILITY: Hash a password (SHA-256) for pasting into LoginUsers
//  Usage: Run this from Apps Script editor, check Logs for hash
// ============================================================
function hashPassword() {
  var plaintext = 'yourpasswordhere'; // ← change this before running
  var bytes  = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(plaintext),
    Utilities.Charset.UTF_8
  );
  var hash = bytes.map(function(b) {
    var hex = (b & 0xff).toString(16);
    return hex.length === 1 ? '0' + hex : hex;
  }).join('');
  Logger.log('Password: ' + plaintext);
  Logger.log('Hash:     ' + hash);
  Logger.log('Paste this hash into column B of LoginUsers.');
}


// ============================================================
//  UTILITY: Print a full diagnostic of what each driver sees
//  Run this to verify everything is connected correctly
// ============================================================
function diagnosDriverMatching() {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var loginSh  = ss.getSheetByName('LoginUsers');
  var driverSh = ss.getSheetByName('Drivers');
  var tripSh   = ss.getSheetByName('Trips');
  var log      = [];

  log.push('=== Driver Trip Matching Diagnostic ===\n');

  if (!loginSh || !driverSh || !tripSh) {
    log.push('Missing one or more sheets. Run setupDatabase() first.');
    Logger.log(log.join('\n'));
    return;
  }

  var loginLr   = loginSh.getLastRow();
  var loginData = loginLr >= 2 ? loginSh.getRange(2, 1, loginLr - 1, 4).getValues() : [];

  var tripLr   = tripSh.getLastRow();
  var tripData = tripLr >= 2 ? tripSh.getRange(2, 1, tripLr - 1, 16).getValues() : [];

  loginData.forEach(function(lr) {
    var email = String(lr[0] || '').trim().toLowerCase();
    var role  = String(lr[2] || '').trim().toLowerCase();
    var did   = String(lr[3] || '').trim();

    if (role !== 'driver') return;

    var matched = tripData.filter(function(tr) {
      return String(tr[13] || '').trim() === did;
    });

    var approved = matched.filter(function(tr) {
      return String(tr[15] || '').trim() === 'Approved';
    });

    log.push('Driver: ' + email);
    log.push('  Driver_ID in LoginUsers: ' + (did || '❌ EMPTY — driver will see NO trips'));
    log.push('  All matched trips: '   + (matched.map(function(t) { return t[0]; }).join(', ') || 'none'));
    log.push('  Approved trips: '      + (approved.map(function(t) { return t[0]; }).join(', ') || 'none'));
    log.push('');
  });

  Logger.log(log.join('\n'));
  SpreadsheetApp.getUi().alert(log.join('\n'));
}