/**
 * IBRA Address Book — Google Sheet → Firebase Realtime Database sync
 * ====================================================================
 * Paste into the Apps Script project bound to (or with access to) the
 * member data spreadsheet. Doesn't need to be the same project that powers
 * the existing doGet() endpoints - this only ever writes, using its own
 * service-account auth, so it can live in a brand new standalone Apps
 * Script project if that's easier.
 *
 * ---- ONE-TIME SETUP ----
 *
 * 1. Add the official "OAuth2 for Apps Script" library to this project:
 *      Editor (left sidebar) > Libraries > "+" > paste this Script ID:
 *      1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF
 *      Choose the latest version. Identifier should default to "OAuth2".
 *
 * 2. Create a Google Cloud service account for your Firebase project:
 *      Firebase Console > Project settings > Service accounts >
 *      "Generate new private key" — this downloads a JSON file.
 *    Keep that JSON file private. Never paste its contents into chat,
 *    a public repo, or anywhere outside Script Properties below.
 *
 * 3. Grant that service account database access:
 *      Google Cloud Console > IAM & Admin > IAM > find the service account
 *      (its email looks like ...@YOUR-PROJECT.iam.gserviceaccount.com) >
 *      Edit > Add role > "Firebase Realtime Database Admin".
 *
 * 4. In this Apps Script project: Project Settings (gear icon) > Script
 *    Properties > Add script property, and add THREE properties using
 *    values from the downloaded JSON file:
 *      FIREBASE_DB_URL        -> e.g. https://YOUR-PROJECT-default-rtdb.firebaseio.com
 *      FIREBASE_CLIENT_EMAIL  -> the "client_email" field from the JSON
 *      FIREBASE_PRIVATE_KEY   -> the "private_key" field from the JSON,
 *                                 pasted exactly as-is (including the
 *                                 literal \n sequences in the string)
 *
 * 5. Run `bootstrapFullSync` once manually from the Apps Script editor to
 *    seed Firebase with every row currently in the sheet (the incremental
 *    triggers below only react to FUTURE edits - they won't retroactively
 *    push what's already there). The first run will ask you to authorize
 *    the script - approve it. Check Firebase Console > Realtime Database
 *    afterwards to confirm data appeared under "ibra/addbook".
 *
 * 6. Run `installIbraSyncTriggers` once manually. This installs the
 *    triggers that keep Firebase in sync automatically from then on - you
 *    only run this once, ever (re-running it just replaces the old
 *    triggers, which is harmless).
 *
 * 7. Set Firebase Realtime Database rules (Firebase Console > Realtime
 *    Database > Rules) to:
 *      {
 *        "rules": {
 *          "ibra": {
 *            ".read": true,
 *            ".write": false
 *          }
 *        }
 *      }
 *    Public read (same as today), no public write - only this script's
 *    service account can write, and it bypasses these rules entirely.
 *
 * ---- DATA SHAPE ----
 * Each member is stored at ibra/addbook/m<MembershipNo>, e.g.
 * ibra/addbook/m1242 — a prefixed key, not a bare number, so Realtime
 * Database doesn't silently coerce the whole "addbook" node into a giant
 * sparse array (it does this automatically when every child key looks like
 * a plain integer). Each record is a plain object with named fields - see
 * FIELD_NAMES below. This list's order must exactly match the sheet's
 * column order (A, B, C, ...) AND the field list index.html expects in
 * fetchAddBookData(). If a column is ever added or reordered in the sheet,
 * update FIELD_NAMES here to match.
 */

const SHEET_ID = '17Tf6EBg-FlyXNhl3INAgp6TvJ5lKyY0Ov-wnD-RN-Vc';
const TAB_NAME = 'All Addresses';
const FIREBASE_PATH_DATA = 'ibra/addbook';

const FIELD_NAMES = ['id','name','add1','add2','add3','district','place','pincode','fixed','mobile1','mobile2','state','email','srno','sbac','doj','dor','spname','area','blank1','dob','fampen','pensionbr','scanned','newempid','dod','photoid','blank3','blank4'];

function rowToRecord_(rowValues) {
  const record = {};
  FIELD_NAMES.forEach((field, i) => { record[field] = rowValues[i]; });
  return record;
}

function firebaseKeyForId_(id) {
  return 'm' + id; // prefix avoids the numeric-key -> array coercion gotcha
}

function isBlankOrInvalidId_(id) {
  return id === '' || id === null || id === undefined || isNaN(Number(id));
}

/**
 * One-time (or as-needed) full bootstrap: pushes every current row as one
 * write. Run this once before the incremental triggers take over, and any
 * time you want to force a complete resync (e.g. after a bulk edit done
 * with triggers temporarily disabled).
 */
function bootstrapFullSync() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TAB_NAME);
  if (!sheet) throw new Error(`Tab "${TAB_NAME}" not found in spreadsheet ${SHEET_ID}`);

  const rows = sheet.getDataRange().getValues();
  const out = {};
  let skipped = 0;
  for (let r = 1; r < rows.length; r++) { // skip header row
    const id = rows[r][0];
    if (isBlankOrInvalidId_(id)) { skipped++; continue; }
    out[firebaseKeyForId_(id)] = rowToRecord_(rows[r]);
  }

  firebasePut_(FIREBASE_PATH_DATA, out);
  Logger.log(`Bootstrap: pushed ${Object.keys(out).length} member rows to Firebase (skipped ${skipped} row(s) with blank/invalid Membership No).`);
}

/**
 * Installable onEdit trigger handler: upserts just the row(s) that were
 * actually touched. Covers typing into a new row and editing any field of
 * an existing row.
 */
function handleEditTrigger(e) {
  if (!e || !e.range) return; // defensive
  const sheet = e.range.getSheet();
  if (sheet.getName() !== TAB_NAME) return; // ignore edits on other tabs

  const startRow = e.range.getRow();
  const numRows = e.range.getNumRows();
  const lastCol = sheet.getLastColumn();

  for (let r = startRow; r < startRow + numRows; r++) {
    if (r === 1) continue; // header row

    const rowValues = sheet.getRange(r, 1, 1, lastCol).getValues()[0];
    const id = rowValues[0];

    if (isBlankOrInvalidId_(id)) {
      // Caution: if Membership No (col A) is blank while other columns in
      // this row are being edited, skip writing it rather than pushing a
      // record under a missing/invalid key. It'll sync normally once col A
      // is filled in (that edit will fire this same handler again).
      Logger.log(`Row ${r}: skipped sync (blank/invalid Membership No).`);
      continue;
    }

    firebasePut_(`${FIREBASE_PATH_DATA}/${firebaseKeyForId_(id)}`, rowToRecord_(rowValues));
  }
}

/**
 * Installable onChange trigger handler: catches structural changes (row
 * insert/delete/sort/paste-that-adds-rows) that onEdit can miss. Reconciles
 * by Membership No only - cheap (one column read + one shallow Firebase
 * read), and doesn't require guessing what changed.
 */
function handleChangeTrigger(e) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TAB_NAME);
  const rows = sheet.getDataRange().getValues();

  const sheetIds = new Set();
  const rowsById = {};
  for (let r = 1; r < rows.length; r++) {
    const id = rows[r][0];
    if (isBlankOrInvalidId_(id)) continue;
    const key = firebaseKeyForId_(id);
    sheetIds.add(key);
    rowsById[key] = rows[r];
  }

  const existing = firebaseShallowGet_(FIREBASE_PATH_DATA);
  const firebaseKeys = new Set(existing ? Object.keys(existing) : []);

  // In Firebase but no longer in the sheet -> remove. Your workflow says
  // rows shouldn't normally be deleted, so this should rarely fire - it's
  // a safety net for the rare/accidental case.
  let removed = 0;
  firebaseKeys.forEach(key => {
    if (!sheetIds.has(key)) { firebaseDelete_(`${FIREBASE_PATH_DATA}/${key}`); removed++; }
  });

  // In the sheet but missing from Firebase -> push. Catches inserted rows
  // that onEdit might not have individually caught (e.g. a paste that
  // extends the sheet).
  let added = 0;
  sheetIds.forEach(key => {
    if (!firebaseKeys.has(key)) { firebasePut_(`${FIREBASE_PATH_DATA}/${key}`, rowToRecord_(rowsById[key])); added++; }
  });

  Logger.log(`Reconcile: ${sheetIds.size} rows in sheet, ${firebaseKeys.size} were in Firebase, removed ${removed}, added ${added}.`);
}

function installIbraSyncTriggers() {
  // Remove any triggers this setup previously created, so re-running it
  // doesn't stack up duplicates.
  const handlerNames = ['handleEditTrigger', 'handleChangeTrigger'];
  ScriptApp.getProjectTriggers().forEach(t => {
    if (handlerNames.indexOf(t.getHandlerFunction()) !== -1) ScriptApp.deleteTrigger(t);
  });

  const ss = SpreadsheetApp.openById(SHEET_ID);
  ScriptApp.newTrigger('handleEditTrigger').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('handleChangeTrigger').forSpreadsheet(ss).onChange().create();

  Logger.log('Installed onEdit (handleEditTrigger) + onChange (handleChangeTrigger) triggers.');
}

// ---- Firebase auth + REST helpers ----
// (Uses the OAuth2 for Apps Script library - see setup step 1.)

function getFirebaseOAuthService_() {
  const props = PropertiesService.getScriptProperties();
  const clientEmail = props.getProperty('FIREBASE_CLIENT_EMAIL');
  const privateKey = props.getProperty('FIREBASE_PRIVATE_KEY');
  if (!clientEmail || !privateKey) {
    throw new Error('Missing FIREBASE_CLIENT_EMAIL / FIREBASE_PRIVATE_KEY script properties - see setup steps in the comment at the top of this file.');
  }

  return OAuth2.createService('FirebaseServiceAccount')
    .setTokenUrl('https://oauth2.googleapis.com/token')
    .setPrivateKey(privateKey.replace(/\\n/g, '\n')) // Script Properties flattens real newlines into \n text
    .setIssuer(clientEmail)
    .setPropertyStore(PropertiesService.getScriptProperties())
    .setScope('https://www.googleapis.com/auth/firebase.database https://www.googleapis.com/auth/userinfo.email');
}

function getFirebaseAccessToken_() {
  const service = getFirebaseOAuthService_();
  if (!service.hasAccess()) {
    throw new Error('Firebase service-account auth failed: ' + service.getLastError());
  }
  return service.getAccessToken();
}

function firebaseDbUrl_() {
  const dbUrl = PropertiesService.getScriptProperties().getProperty('FIREBASE_DB_URL');
  if (!dbUrl) throw new Error('Missing FIREBASE_DB_URL script property - see setup steps at the top of this file.');
  return dbUrl.replace(/\/$/, '');
}

/** Writes `value` to the given database path, replacing whatever was there. */
function firebasePut_(path, value) {
  const token = getFirebaseAccessToken_();
  const res = UrlFetchApp.fetch(`${firebaseDbUrl_()}/${path}.json`, {
    method: 'put',
    contentType: 'application/json',
    payload: JSON.stringify(value),
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() >= 300) {
    throw new Error(`Firebase write to "${path}" failed (HTTP ${res.getResponseCode()}): ${res.getContentText()}`);
  }
}

/** Deletes whatever is at the given database path. */
function firebaseDelete_(path) {
  const token = getFirebaseAccessToken_();
  const res = UrlFetchApp.fetch(`${firebaseDbUrl_()}/${path}.json`, {
    method: 'delete',
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() >= 300) {
    throw new Error(`Firebase delete of "${path}" failed (HTTP ${res.getResponseCode()}): ${res.getContentText()}`);
  }
}

/** Reads just the immediate child key NAMES at a path (not their contents) - cheap, used for id-set reconciliation. Returns null if the path has no children. */
function firebaseShallowGet_(path) {
  const token = getFirebaseAccessToken_();
  const res = UrlFetchApp.fetch(`${firebaseDbUrl_()}/${path}.json?shallow=true`, {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() >= 300) {
    throw new Error(`Firebase shallow read of "${path}" failed (HTTP ${res.getResponseCode()}): ${res.getContentText()}`);
  }
  const text = res.getContentText();
  return text === 'null' ? null : JSON.parse(text);
}
