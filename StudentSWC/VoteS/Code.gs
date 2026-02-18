// ==========================================
// ระบบนับคะแนนการเลือกตั้งสภานักเรียน
// Google Apps Script Backend — Refactored
// ==========================================

// ===== CONFIG =====
const SPREADSHEET_ID = "1ULX4LVskkS9HkQAvOvfNLeHgbv9siE1LYAYWNawwVYA";

// Sheet Names
const SHEET_NAME = "VoteData";
const RAW_VOTES_SHEET_NAME = "RawVotes";
const LOG_SHEET_NAME = "AccessLogs";
const PARTY_CONFIG_SHEET_NAME = "PartyConfig";
const ELECTION_CONFIG_SHEET_NAME = "ElectionConfig";
const ADMIN_USERS_SHEET_NAME = "AdminUsers";
const VOTED_HASHES_SHEET_NAME = "VotedHashes";
const VOTE_TOKENS_SHEET_NAME = "VoteTokens";

// ===== WEB APP ENTRY =====

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || '';

  // === JSON API MODE ===
  if (action) {
    return handleApiGet(action, e.parameter);
  }

  // === LEGACY HTML PAGE SERVING ===
  let page = 'index';
  if (e && e.parameter && e.parameter.page) {
    page = e.parameter.page;
  }
  const validPages = ['index', 'ballot', 'counting', 'admin', 'ranking'];
  if (!validPages.includes(page)) {
    page = 'index';
  }

  try {
    const template = HtmlService.createTemplateFromFile(page);
    return template.evaluate()
      .setTitle('ระบบนับคะแนนการเลือกตั้งสภานักเรียน')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  } catch (err) {
    Logger.log('Template evaluate error for page ' + page + ': ' + err.message);
    return HtmlService.createHtmlOutputFromFile(page)
      .setTitle('ระบบนับคะแนนการเลือกตั้งสภานักเรียน')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  }
}

// === JSON API: GET (read-only operations) ===
function handleApiGet(action, params) {
  var result;
  try {
    switch (action) {
      case 'getVoteData':       result = getVoteData(); break;
      case 'getElectionConfig': result = getElectionConfig(); break;
      case 'getPartyConfig':    result = getPartyConfig(); break;
      case 'getUsers':          result = getUsers(); break;
      case 'getAccessLogs':     result = getAccessLogs(parseInt(params.limit) || 100); break;
      default: result = { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { error: err.message };
  }
  return jsonResponse(result);
}

// === JSON API: POST (write/mutate operations) ===
function doPost(e) {
  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse({ error: 'Invalid JSON body' });
  }
  var action = body.action || '';
  return handleApiPost(action, body);
}

function handleApiPost(action, body) {
  var result;
  try {
    switch (action) {
      case 'submitVote':           result = submitVote(body.party, body.token); break;
      case 'updateVote':           result = updateVote(body.partyNum, body.delta); break;
      case 'updateSpecial':        result = updateSpecial(body.type, body.delta); break;
      case 'editStat':             result = editStat(body.statId, body.value); break;
      case 'verifyPassword':       result = verifyAdminPassword(body.password, body.name); break;
      case 'generateToken':        result = { success: true, token: generateVoteToken() }; break;
      case 'addParty':             result = addParty(body.name, body.logoId, body.color, body.leader); break;
      case 'updateParty':          result = updateParty(body.id, body.updates); break;
      case 'deleteParty':          result = deleteParty(body.id); break;
      case 'addUser':              result = addUser(body.username, body.password, body.role, body.displayName); break;
      case 'updateUser':           result = updateUser(body.username, body.updates); break;
      case 'deleteUser':           result = deleteUser(body.username); break;
      case 'lockResults':          result = lockResults(body.teacherName, body.confirmCode); break;
      case 'unlockResults':        result = unlockResults(body.confirmCode); break;
      case 'resetData':            result = resetData(); break;
      case 'updateElectionConfig': result = updateElectionConfig(body.updates); break;
      default: result = { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { error: err.message };
  }
  return jsonResponse(result);
}

// === JSON response helper ===
function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

// ===== SPREADSHEET HELPER (cached) =====

let _ssCache = null;
function getSS() {
  if (!_ssCache) {
    _ssCache = SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  return _ssCache;
}

// ===== SHEET INITIALIZERS =====

function getElectionConfigSheet() {
  const ss = getSS();
  let sheet = ss.getSheetByName(ELECTION_CONFIG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(ELECTION_CONFIG_SHEET_NAME);
    const data = [
      ['key', 'value'],
      ['electionName', 'การเลือกตั้งสภานักเรียน'],
      ['academicYear', '2569'],
      ['schoolName', 'โรงเรียนศรีวิชัยวิทยา'],
      ['schoolLogoId', '1l-r_ZGA-N4rKipi5zz4DKOoMJL-W82zW'],
      ['councilLogoId', '1MTw6a6RfcCU-QTIQQ_EyhHLB_8CE9vHy'],
      ['votingEnabled', 'true'],
      ['hashSalt', Utilities.getUuid()],
      ['developerName', 'นายธีรเดช เดชสูงเนิน']
    ];
    sheet.getRange(1, 1, data.length, 2).setValues(data);
    sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#f3f4f6');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getAdminUsersSheet() {
  const ss = getSS();
  let sheet = ss.getSheetByName(ADMIN_USERS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(ADMIN_USERS_SHEET_NAME);
    const headers = ['username', 'passwordHash', 'role', 'displayName', 'active', 'createdAt'];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f3f4f6');
    sheet.setFrozenRows(1);

    // Default users
    const now = new Date().toLocaleString('th-TH');
    const salt = _getHashSalt();
    sheet.appendRow(['advisor', _hashPassword('admin2569', salt), 'advisor', 'ครูที่ปรึกษา', 'true', now]);
    sheet.appendRow(['committee', _hashPassword('vote2569', salt), 'committee', 'กรรมการสภา', 'true', now]);
  }
  return sheet;
}

function getPartyConfigSheet() {
  const ss = getSS();
  let sheet = ss.getSheetByName(PARTY_CONFIG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(PARTY_CONFIG_SHEET_NAME);
    const configData = [
      ['id', 'name', 'logoId', 'color', 'leader'],
      ['1', 'พรรค WANT TO TRY', '1EiYNuV1-q0tHLyaW7v_tksP1AXge-yAE', '#E63946', ''],
      ['2', 'พรรค BETTER FUTURE', '1ONxc38OT-JR224YMv74d-xg3YbNmXg00', '#457B9D', ''],
      ['3', 'พรรคคำว่าแร่ แพ้คำว่ารัก', '1UHLwn30CbToqTl58BSgz4J-kp_P33fPy', '#6A4C93', '']
    ];
    sheet.getRange(1, 1, configData.length, 5).setValues(configData);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#f3f4f6');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getSheet() {
  const ss = getSS();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const initialData = [
      ['key', 'value'],
      ['totalEligible', '0'],
      ['isLocked', 'false'],
      ['lockedBy', ''],
      ['lockedAt', '']
    ];
    sheet.getRange(1, 1, initialData.length, 2).setValues(initialData);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getRawVotesSheet() {
  const ss = getSS();
  let sheet = ss.getSheetByName(RAW_VOTES_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(RAW_VOTES_SHEET_NAME);
    sheet.appendRow(['Timestamp', 'Party', 'Value', 'Type', 'RefID']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getLogSheet() {
  const ss = getSS();
  let sheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET_NAME);
    sheet.appendRow(['Timestamp', 'Name', 'Role', 'Status', 'Message']);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#f3f4f6');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Legacy — still used by resetData() to clear hashes
function _getVotedHashesSheet() {
  const ss = getSS();
  let sheet = ss.getSheetByName(VOTED_HASHES_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(VOTED_HASHES_SHEET_NAME);
    sheet.appendRow(['Hash', 'Timestamp']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function _getVoteTokensSheet() {
  const ss = getSS();
  let sheet = ss.getSheetByName(VOTE_TOKENS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(VOTE_TOKENS_SHEET_NAME);
    sheet.appendRow(['Token', 'CreatedAt', 'Used', 'UsedAt']);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#f3f4f6');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ===== ELECTION CONFIG =====

function getElectionConfig() {
  const sheet = getElectionConfigSheet();
  const data = sheet.getDataRange().getValues();
  const config = {};
  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][0]).trim();
    let value = data[i][1];
    if (value === 'true') value = true;
    else if (value === 'false') value = false;
    if (key) config[key] = value;
  }
  return config;
}

function updateElectionConfig(updates) {
  if (!updates || typeof updates !== 'object') {
    return { success: false, message: 'ข้อมูลไม่ถูกต้อง' };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getElectionConfigSheet();
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const key = String(data[i][0]).trim();
      if (key in updates) {
        sheet.getRange(i + 1, 2).setValue(String(updates[key]));
      }
    }

    // Add new keys that don't exist yet
    const existingKeys = data.map(r => String(r[0]).trim());
    for (const key in updates) {
      if (!existingKeys.includes(key)) {
        sheet.appendRow([key, String(updates[key])]);
      }
    }

    return { success: true, message: 'อัปเดตการตั้งค่าเรียบร้อย' };
  } catch (e) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

// ===== PARTY CONFIG CRUD =====

function getPartyConfig() {
  const sheet = getPartyConfigSheet();
  const data = sheet.getDataRange().getValues();
  const config = {};

  const defaultColors = ['#E63946', '#457B9D', '#6A4C93', '#2A9D8F', '#E76F51', '#F4A261', '#264653', '#D62828', '#023E8A', '#38B000'];

  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][0]).trim();
    if (!id) continue;
    config[id] = {
      name: data[i][1] || 'ไม่ระบุชื่อพรรค',
      logoId: data[i][2] || '',
      color: data[i][3] || defaultColors[(parseInt(id) - 1) % defaultColors.length],
      leader: data[i][4] || ''
    };
  }

  return config;
}

function addParty(name, logoId, color, leader) {
  if (!name || !String(name).trim()) {
    return { success: false, message: 'กรุณากรอกชื่อพรรค' };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getPartyConfigSheet();
    const data = sheet.getDataRange().getValues();

    // Find max ID
    let maxId = 0;
    for (let i = 1; i < data.length; i++) {
      const id = parseInt(data[i][0]) || 0;
      if (id > maxId) maxId = id;
    }

    const newId = maxId + 1;
    const defaultColors = ['#E63946', '#457B9D', '#6A4C93', '#2A9D8F', '#E76F51', '#F4A261'];
    const assignedColor = color || defaultColors[(newId - 1) % defaultColors.length];

    sheet.appendRow([String(newId), name, logoId || '', assignedColor, leader || '']);

    return { success: true, message: 'เพิ่มพรรคเรียบร้อย', id: newId };
  } catch (e) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

function updateParty(id, updates) {
  if (!id) {
    return { success: false, message: 'ไม่ระบุ ID พรรค' };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getPartyConfigSheet();
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(id)) {
        if (updates.name !== undefined) sheet.getRange(i + 1, 2).setValue(updates.name);
        if (updates.logoId !== undefined) sheet.getRange(i + 1, 3).setValue(updates.logoId);
        if (updates.color !== undefined) sheet.getRange(i + 1, 4).setValue(updates.color);
        if (updates.leader !== undefined) sheet.getRange(i + 1, 5).setValue(updates.leader);
        return { success: true, message: 'อัปเดตพรรคเรียบร้อย' };
      }
    }
    return { success: false, message: 'ไม่พบพรรคนี้' };
  } catch (e) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteParty(id) {
  if (!id) {
    return { success: false, message: 'ไม่ระบุ ID พรรค' };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    // Check if votes exist for this party
    const votes = aggregateVotes();
    if (votes['party' + id] && votes['party' + id] > 0) {
      return { success: false, message: 'ไม่สามารถลบพรรคที่มีคะแนนอยู่แล้ว' };
    }

    const sheet = getPartyConfigSheet();
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(id)) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'ลบพรรคเรียบร้อย' };
      }
    }
    return { success: false, message: 'ไม่พบพรรคนี้' };
  } catch (e) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

// ===== ADMIN USERS CRUD =====

function _hashPassword(password, salt) {
  const raw = password + ':' + salt;
  const hash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
  return hash.map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
}

function _getHashSalt() {
  const config = getElectionConfig();
  return config.hashSalt || 'default-salt-2569';
}

function getUsers() {
  const sheet = getAdminUsersSheet();
  const data = sheet.getDataRange().getValues();
  const users = [];
  for (let i = 1; i < data.length; i++) {
    users.push({
      username: data[i][0],
      role: data[i][2],
      displayName: data[i][3],
      active: String(data[i][4]) === 'true',
      createdAt: data[i][5]
    });
  }
  return users;
}

function addUser(username, password, role, displayName) {
  // Validate input before acquiring lock
  if (!username || !password || !role) {
    return { success: false, message: 'กรุณากรอกข้อมูลให้ครบ' };
  }
  const validRoles = ['advisor', 'committee'];
  if (!validRoles.includes(role)) {
    return { success: false, message: 'บทบาทไม่ถูกต้อง' };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getAdminUsersSheet();
    const data = sheet.getDataRange().getValues();

    // Check duplicate
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === username.trim().toLowerCase()) {
        return { success: false, message: 'ชื่อผู้ใช้นี้มีอยู่แล้ว' };
      }
    }

    const salt = _getHashSalt();
    const hash = _hashPassword(password, salt);
    const now = new Date().toLocaleString('th-TH');
    sheet.appendRow([username.trim(), hash, role, displayName || username, 'true', now]);

    return { success: true, message: 'เพิ่มผู้ใช้เรียบร้อย' };
  } catch (e) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

function updateUser(username, updates) {
  if (!username) {
    return { success: false, message: 'ไม่ระบุชื่อผู้ใช้' };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getAdminUsersSheet();
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === username.trim()) {
        if (updates.displayName !== undefined) sheet.getRange(i + 1, 4).setValue(updates.displayName);
        if (updates.role !== undefined) sheet.getRange(i + 1, 3).setValue(updates.role);
        if (updates.active !== undefined) sheet.getRange(i + 1, 5).setValue(String(updates.active));
        if (updates.password) {
          const salt = _getHashSalt();
          sheet.getRange(i + 1, 2).setValue(_hashPassword(updates.password, salt));
        }
        return { success: true, message: 'อัปเดตผู้ใช้เรียบร้อย' };
      }
    }
    return { success: false, message: 'ไม่พบผู้ใช้นี้' };
  } catch (e) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteUser(username) {
  if (!username) {
    return { success: false, message: 'ไม่ระบุชื่อผู้ใช้' };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getAdminUsersSheet();
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === username.trim()) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'ลบผู้ใช้เรียบร้อย' };
      }
    }
    return { success: false, message: 'ไม่พบผู้ใช้นี้' };
  } catch (e) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

function verifyAdminPassword(password, name) {
  const userName = name || "ไม่ระบุชื่อ";
  const sheet = getAdminUsersSheet();
  const data = sheet.getDataRange().getValues();
  const salt = _getHashSalt();
  const inputHash = _hashPassword(password, salt);

  for (let i = 1; i < data.length; i++) {
    const storedHash = String(data[i][1]).trim();
    const role = String(data[i][2]).trim();
    const displayName = String(data[i][3]).trim();
    const active = String(data[i][4]).trim();

    if (storedHash === inputHash && active === 'true') {
      logAccess(userName || displayName, role, 'Success', 'เข้าระบบสำเร็จ');
      return { success: true, role: role, displayName: displayName };
    }
  }

  logAccess(userName, 'unknown', 'Failed', 'รหัส PIN ผิดพลาด');
  return { success: false, message: "รหัส PIN ไม่ถูกต้อง" };
}

// ===== VOTE DATA (read / write / aggregate) =====

function readSheetData() {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const result = {};
  for (let i = 1; i < data.length; i++) {
    const key = data[i][0];
    let value = data[i][1];
    if (value === 'true') value = true;
    else if (value === 'false') value = false;
    else if (!isNaN(value) && value !== '') value = Number(value);
    result[key] = value;
  }
  return result;
}

function writeSheetData(data) {
  const sheet = getSheet();
  const existingData = sheet.getDataRange().getValues();
  for (let i = 1; i < existingData.length; i++) {
    const key = existingData[i][0];
    if (key in data) {
      sheet.getRange(i + 1, 2).setValue(String(data[key]));
    }
  }
}

function aggregateVotes() {
  const sheet = getRawVotesSheet();
  const lastRow = sheet.getLastRow();
  const partyConfig = getPartyConfig();
  const partyIds = Object.keys(partyConfig);

  // Initialize totals
  const totals = { noVote: 0, invalidBallot: 0 };
  partyIds.forEach(id => { totals['party' + id] = 0; });

  if (lastRow <= 1) return totals;

  const values = sheet.getRange(2, 2, lastRow - 1, 2).getValues(); // Party, Value

  for (let i = 0; i < values.length; i++) {
    const party = String(values[i][0]);
    const val = Number(values[i][1]) || 0;

    if (party === 'noVote') totals.noVote += val;
    else if (party === 'invalidBallot') totals.invalidBallot += val;
    else if (partyIds.includes(party)) totals['party' + party] += val;
  }

  return totals;
}

// ===== VOTE TOKEN SYSTEM =====

function generateVoteToken() {
  const token = Utilities.getUuid();
  const sheet = _getVoteTokensSheet();
  sheet.appendRow([token, new Date().toLocaleString('th-TH'), false, '']);
  return token;
}

function _validateAndUseToken(token) {
  if (!token) return { valid: false, message: 'ไม่มี token' };

  const sheet = _getVoteTokensSheet();
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(token)) {
      // Found token — check if already used
      if (data[i][2] === true || String(data[i][2]).toLowerCase() === 'true') {
        return { valid: false, message: 'บัตรลงคะแนนนี้ถูกใช้งานแล้ว กรุณารอบัตรใหม่' };
      }
      // Mark as used
      sheet.getRange(i + 1, 3).setValue(true);
      sheet.getRange(i + 1, 4).setValue(new Date().toLocaleString('th-TH'));
      return { valid: true };
    }
  }

  return { valid: false, message: 'ไม่พบบัตรลงคะแนนนี้ในระบบ' };
}

// ===== VOTING =====

function submitVote(party, token) {
  // Read-only checks outside lock to reduce lock hold time
  const electionConfig = getElectionConfig();
  if (electionConfig.votingEnabled === false) {
    return { success: false, message: 'ระบบโหวตปิดอยู่ กรุณาติดต่อผู้ดูแลระบบ' };
  }

  const partyConfig = getPartyConfig();
  if (party !== 'noVote' && !partyConfig[party]) {
    return { success: false, message: 'พรรคที่เลือกไม่ถูกต้อง' };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    // Check system lock status (must be inside lock for consistency)
    const statusData = readSheetData();
    if (statusData.isLocked) {
      return { success: false, message: 'ผลคะแนนถูกล็อคแล้ว ไม่สามารถลงคะแนนเพิ่มได้' };
    }

    // Validate token (one-time use)
    const tokenResult = _validateAndUseToken(token);
    if (!tokenResult.valid) {
      return { success: false, message: tokenResult.message };
    }

    // Append vote
    const rawSheet = getRawVotesSheet();
    rawSheet.appendRow([new Date(), party, 1, 'online', token.substring(0, 8)]);

    return { success: true, message: 'บันทึกคะแนนเรียบร้อย' };
  } catch (e) {
    return { success: false, message: 'ระบบไม่ว่าง กรุณาลองใหม่อีกครั้ง: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

// ===== COUNTING (manual vote updates) =====

function updateVote(partyNum, delta) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const statusData = readSheetData();
    if (statusData.isLocked) {
      return { success: false, message: 'ผลคะแนนถูกล็อคแล้ว ไม่สามารถแก้ไขได้' };
    }

    const rawSheet = getRawVotesSheet();
    rawSheet.appendRow([new Date(), String(partyNum), delta, 'manual', 'admin']);

    return { success: true, data: getVoteData() };
  } catch (e) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

function updateSpecial(type, delta) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const statusData = readSheetData();
    if (statusData.isLocked) {
      return { success: false, message: 'ผลคะแนนถูกล็อคแล้ว ไม่สามารถแก้ไขได้' };
    }

    const rawSheet = getRawVotesSheet();
    rawSheet.appendRow([new Date(), type, delta, 'manual', 'admin']);

    return { success: true, data: getVoteData() };
  } catch (e) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

function editStat(statId, value) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const data = readSheetData();

    // Check lock status (single read, no redundant getVoteData)
    if (data.isLocked) {
      return { success: false, message: 'ผลคะแนนถูกล็อคแล้ว ไม่สามารถแก้ไขได้' };
    }

    if (statId === 'totalEligible') {
      data[statId] = Math.max(0, parseInt(value) || 0);
      writeSheetData(data);
      return { success: true, data: getVoteData() };
    }

    // For party/special stats — compute delta and append to RawVotes
    const currentVotes = aggregateVotes();

    let currentVal = 0;
    let type = statId;

    // Dynamic party handling
    if (statId.startsWith('party')) {
      const partyId = statId.replace('party', '');
      currentVal = currentVotes['party' + partyId] || 0;
      type = partyId;
    } else {
      currentVal = currentVotes[statId] || 0;
    }

    const newVal = Math.max(0, parseInt(value) || 0);
    const delta = newVal - currentVal;

    if (delta !== 0) {
      const rawSheet = getRawVotesSheet();
      rawSheet.appendRow([new Date(), type, delta, 'manualset', 'admin']);
    }

    return { success: true, data: getVoteData() };
  } catch (e) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

// ===== PUBLIC API =====

function getVoteData() {
  const meta = readSheetData();
  const votes = aggregateVotes();
  const data = { ...meta, ...votes };
  data.config = getPartyConfig();
  data.electionConfig = getElectionConfig();
  return data;
}

// ===== LOCK / UNLOCK / RESET =====

function lockResults(teacherName, confirmCode) {
  // Verify confirm code is a valid advisor password
  const verifyResult = verifyAdminPassword(confirmCode, teacherName);
  if (!verifyResult.success || verifyResult.role !== 'advisor') {
    return { success: false, message: 'รหัสยืนยันไม่ถูกต้อง หรือไม่ใช่สิทธิ์ครูที่ปรึกษา' };
  }

  if (!teacherName || teacherName.trim() === '') {
    return { success: false, message: 'กรุณาใส่ชื่อครูที่ปรึกษา' };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const data = readSheetData();
    data.isLocked = true;
    data.lockedBy = teacherName.trim();
    data.lockedAt = new Date().toLocaleString('th-TH');
    writeSheetData(data);
    logAccess(teacherName, 'advisor', 'Lock', 'ล็อคผลคะแนน');
    return { success: true, message: 'ผลคะแนนถูกล็อคเรียบร้อยแล้ว' };
  } catch (e) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

function unlockResults(confirmCode) {
  const verifyResult = verifyAdminPassword(confirmCode, 'ผู้ดูแลระบบ');
  if (!verifyResult.success || verifyResult.role !== 'advisor') {
    return { success: false, message: 'รหัสยืนยันไม่ถูกต้อง' };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const data = readSheetData();
    data.isLocked = false;
    data.lockedBy = '';
    data.lockedAt = '';
    writeSheetData(data);
    logAccess('ผู้ดูแลระบบ', 'advisor', 'Unlock', 'ปลดล็อคผลคะแนน');
    return { success: true, message: 'ปลดล็อคผลคะแนนเรียบร้อย' };
  } catch (e) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

function resetData() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const data = readSheetData();
    if (data.isLocked) {
      return { success: false, message: 'ไม่สามารถรีเซ็ตได้ เพราะผลคะแนนถูกล็อคแล้ว' };
    }

    // Reset VoteData
    const resetValues = {
      totalEligible: 0,
      isLocked: false,
      lockedBy: '',
      lockedAt: ''
    };
    writeSheetData(resetValues);

    // Clear RawVotes
    const rawSheet = getRawVotesSheet();
    const lastRowRaw = rawSheet.getLastRow();
    if (lastRowRaw > 1) {
      rawSheet.deleteRows(2, lastRowRaw - 1);
    }

    // Clear VotedHashes (legacy)
    const hashSheet = _getVotedHashesSheet();
    const lastRowHash = hashSheet.getLastRow();
    if (lastRowHash > 1) {
      hashSheet.deleteRows(2, lastRowHash - 1);
    }

    // Clear VoteTokens
    const tokenSheet = _getVoteTokensSheet();
    const lastRowToken = tokenSheet.getLastRow();
    if (lastRowToken > 1) {
      tokenSheet.deleteRows(2, lastRowToken - 1);
    }

    logAccess('ผู้ดูแลระบบ', 'advisor', 'Reset', 'รีเซ็ตข้อมูลทั้งหมด');
    return { success: true, message: 'รีเซ็ตข้อมูลเรียบร้อยแล้ว' };
  } catch (e) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

// ===== ACCESS LOGS =====

function logAccess(name, role, status, message) {
  try {
    const sheet = getLogSheet();
    sheet.appendRow([new Date(), name, role, status, message || '']);
  } catch (e) {
    console.error("Logging failed: " + e.toString());
  }
}

function getAccessLogs(limit) {
  const sheet = getLogSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const maxRows = Math.min(limit || 50, lastRow - 1);
  const startRow = Math.max(2, lastRow - maxRows + 1);
  const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, 5).getValues();

  const logs = [];
  for (let i = data.length - 1; i >= 0; i--) {
    logs.push({
      timestamp: data[i][0] ? new Date(data[i][0]).toLocaleString('th-TH') : '',
      name: data[i][1],
      role: data[i][2],
      status: data[i][3],
      message: data[i][4]
    });
  }
  return logs;
}
