/**
 * ===================================================
 * WEB APP DASHBOARD - LOGIN + ROLE BASED ACCESS
 * ===================================================
 * Sheet "Users": username | password | role | nama
 * Sheet "Menus": menu_id | label | icon | url | role_access
 * ===================================================
 */

const SHEET_ID = '11kal0trMILl7jmmQwuboFU9adr634ljIGvdAvDCx4a8';
const SESSION_DURATION_MS = 8 * 60 * 60 * 1000; // 8 jam

// ===== LOGO =====
// Tempel link gambar LANGSUNG di sini (mis. https://situs.com/logo.png).
// Gunakan URL gambar publik (BUKAN link Google Drive). Kosongkan ('') untuk menyembunyikan.
const LOGO_URL = 'https://els.id/wp-content/uploads/2023/08/ELS-ID-oren-1024x316.png';        // logo Dashboard
const LOGIN_LOGO_URL = 'https://i.ibb.co.com/bjtDdtMQ/Logo.png'; // logo Login

/* ============ ROUTING ============ */

function doGet(e) {
  const token = (e && e.parameter) ? e.parameter.token : null;

  if (token) {
    const session = validateSession(token);
    if (session) {
      try {
        return serveDashboard(session, token);
      } catch (err) {
        return errorPage('serveDashboard', err);
      }
    }
    // token ada tapi tidak valid / expired -> balik ke login
    return serveLogin('Session berakhir, silakan login kembali.');
  }

  try {
    return serveLogin();
  } catch (err) {
    return errorPage('serveLogin', err);
  }
}

function serveLogin(notice) {
  const template = HtmlService.createTemplateFromFile('Login');
  template.notice = notice || '';
  template.logoUrl = LOGIN_LOGO_URL;
  return template.evaluate()
    .setTitle('Login - Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function serveDashboard(session, token) {
  const template = HtmlService.createTemplateFromFile('Dashboard');
  template.userData = session;
  template.menus = getMenusForRole(session.role);
  template.token = token;
  template.logoUrl = LOGO_URL;

  return template.evaluate()
    .setTitle('Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function errorPage(where, err) {
  return HtmlService.createHtmlOutput(
    '<pre style="padding:20px;font-family:monospace;white-space:pre-wrap;">' +
    'ERROR di ' + where + ':\n' + err.message + '\n\n' + err.stack +
    '</pre>'
  );
}

/** URL web app yang benar (.../exec). Dipakai untuk redirect setelah login. */
function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

/* ============ AUTH ============ */

function login(username, password) {
  if (!username || !password) {
    return { success: false, message: 'Username dan password wajib diisi' };
  }

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Users');
  if (!sheet) return { success: false, message: 'Sheet "Users" tidak ditemukan' };

  const data = sheet.getDataRange().getValues();
  const inputUser = String(username).trim();
  const inputPass = String(password).trim();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const sheetUser = String(row[0]).trim();
    const sheetPass = String(row[1]).trim();
    const role = String(row[2]).trim();
    const nama = String(row[3]).trim();

    if (sheetUser === inputUser && sheetPass === inputPass) {
      cleanupExpiredSessions(); // jaga agar properties tidak menumpuk
      const token = generateToken();
      saveSession(token, {
        username: sheetUser,
        role: role,
        nama: nama,
        loginTime: Date.now()
      });
      return { success: true, token: token, role: role, nama: nama };
    }
  }

  return { success: false, message: 'Username atau password salah' };
}

function logout(token) {
  if (token) {
    PropertiesService.getScriptProperties().deleteProperty('session_' + token);
  }
  return { success: true };
}

/** Ganti password user yang sedang login. */
function changePassword(token, oldPassword, newPassword) {
  const session = validateSession(token);
  if (!session) return { success: false, message: 'Session tidak valid, silakan login ulang' };
  if (!oldPassword || !newPassword) return { success: false, message: 'Password lama & baru wajib diisi' };
  if (String(newPassword).trim().length < 4) return { success: false, message: 'Password baru minimal 4 karakter' };

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Users');
  if (!sheet) return { success: false, message: 'Sheet "Users" tidak ditemukan' };

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[0]).trim() === session.username) {
      if (String(row[1]).trim() !== String(oldPassword).trim()) {
        return { success: false, message: 'Password lama salah' };
      }
      sheet.getRange(i + 1, 2).setValue(String(newPassword).trim()); // kolom B = password
      return { success: true };
    }
  }
  return { success: false, message: 'User tidak ditemukan' };
}

function generateToken() {
  return Utilities.getUuid();
}

/* ============ SESSION (PERSISTEN via PropertiesService) ============ */

function saveSession(token, sessionData) {
  PropertiesService.getScriptProperties()
    .setProperty('session_' + token, JSON.stringify(sessionData));
}

function validateSession(token) {
  if (!token) return null;
  const props = PropertiesService.getScriptProperties();
  const data = props.getProperty('session_' + token);
  if (!data) return null;

  let session;
  try {
    session = JSON.parse(data);
  } catch (e) {
    props.deleteProperty('session_' + token);
    return null;
  }

  if (Date.now() - session.loginTime > SESSION_DURATION_MS) {
    props.deleteProperty('session_' + token);
    return null;
  }
  return session;
}

/** Hapus semua session yang sudah kadaluarsa. */
function cleanupExpiredSessions() {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  const now = Date.now();
  Object.keys(all).forEach(function (key) {
    if (key.indexOf('session_') !== 0) return;
    try {
      const s = JSON.parse(all[key]);
      if (now - s.loginTime > SESSION_DURATION_MS) {
        props.deleteProperty(key);
      }
    } catch (e) {
      props.deleteProperty(key);
    }
  });
}

/* ============ MENUS ============ */

function getMenusForRole(role) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Menus');
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const menus = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[1]) continue; // skip baris kosong (tidak ada label)

    const roleAccess = String(row[4] || '').split(',').map(function (r) { return r.trim(); });

    if (roleAccess.indexOf(role) !== -1 || roleAccess.indexOf('All') !== -1) {
      menus.push({
        id: row[0],
        label: String(row[1]),
        icon: String(row[2] || ''),
        url: String(row[3] || '#')
      });
    }
  }
  return menus;
}

/* ============ ADMIN ============ */

function requireAdmin(token) {
  const session = validateSession(token);
  if (!session || session.role !== 'Admin') return null;
  return session;
}

function getAllMenus(token) {
  if (!requireAdmin(token)) return { success: false, message: 'Akses ditolak' };

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Menus');
  const data = sheet.getDataRange().getValues();
  const menus = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[1]) continue;
    menus.push({
      rowIndex: i + 1,
      id: row[0],
      label: String(row[1]),
      icon: String(row[2] || ''),
      url: String(row[3] || ''),
      roleAccess: String(row[4] || '')
    });
  }
  return { success: true, menus: menus };
}

function updateMenuAccess(token, rowIndex, roleAccess) {
  if (!requireAdmin(token)) return { success: false, message: 'Akses ditolak' };
  const ss = SpreadsheetApp.openById(SHEET_ID);
  ss.getSheetByName('Menus').getRange(rowIndex, 5).setValue(roleAccess);
  return { success: true };
}

function addMenu(token, label, icon, url, roleAccess) {
  if (!requireAdmin(token)) return { success: false, message: 'Akses ditolak' };
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Menus');
  const lastRow = sheet.getLastRow();
  sheet.appendRow([lastRow, label, icon, url, roleAccess]);
  return { success: true };
}

function deleteMenu(token, rowIndex) {
  if (!requireAdmin(token)) return { success: false, message: 'Akses ditolak' };
  const ss = SpreadsheetApp.openById(SHEET_ID);
  ss.getSheetByName('Menus').deleteRow(rowIndex);
  return { success: true };
}
