import './styles.css';
import * as XLSX from 'xlsx';
import { initializeApp } from 'firebase/app';
import {
  get,
  getDatabase,
  push,
  ref,
  remove,
  set,
  update,
} from 'firebase/database';
import {
  ArrowLeft,
  Banknote,
  Bell,
  Box,
  Building2,
  Camera,
  ChevronRight,
  createIcons,
  Download,
  FileUp,
  KeyRound,
  ListRestart,
  LogOut,
  MapPin,
  Package,
  Phone,
  Plus,
  Save,
  Search,
  Settings,
  ShieldCheck,
  Trash2,
  Upload,
  User,
  Users,
  Weight,
} from 'lucide';
import branches from '../domex_branches.json';
import logoUrl from '../domex logo.png';

const STORAGE_KEY = 'domex-pickups-v1';
const SETTINGS_KEY = 'domex-settings-v1';
const FIREBASE_CONFIG_KEY = 'domex-firebase-config-v1';
const SESSION_KEY = 'domex-session-v1';
const ADMIN_USERNAME = 'madu';
const ADMIN_PASSWORD = '2006';
const DEFAULT_FIREBASE_CONFIG = {
  apiKey: 'AIzaSyCzwo_V8pJHpsPdOHlspQal5T41z7ao-eQ',
  authDomain: 'dome-d-entry.firebaseapp.com',
  databaseURL: 'https://dome-d-entry-default-rtdb.firebaseio.com',
  projectId: 'dome-d-entry',
  storageBucket: 'dome-d-entry.firebasestorage.app',
  messagingSenderId: '419925262525',
  appId: '1:419925262525:web:128cf84c68d52aaa6fb7d4',
  measurementId: 'G-9KW13SD9QW',
};
const BASE_EXCEL_ONLY_DEFAULTS = {
  Reference: 0,
  Remark: 0,
  Exchange: 0,
};
const HEADERS = [
  'TrackingNumber',
  'Reference',
  'PackageDescription',
  'ReceiverName',
  'ReceiverAddress',
  'ReceiverCity',
  'ReceiverContactNo',
  'NoOfPcs',
  'Kilo',
  'Gram',
  'Amount',
  'Exchange',
  'Remark',
];

const FIELD_META = {
  TrackingNumber: { label: 'Tracking No', type: 'text', span: 'small', icon: 'package' },
  ReceiverName: { label: 'Receiver Name', type: 'text', span: 'medium', icon: 'user' },
  ReceiverAddress: { label: 'Receiver Address', type: 'textarea', span: 'wide', icon: 'map-pin' },
  ReceiverCity: { label: 'Receiver City', type: 'text', span: 'medium', icon: 'search' },
  ReceiverContactNo: { label: 'Contact No', type: 'tel', span: 'small', icon: 'phone' },
  NoOfPcs: { label: 'Pieces', type: 'number', span: 'tiny', min: 0, icon: 'box' },
  Kilo: { label: 'Kilo', type: 'number', span: 'tiny', min: 0, icon: 'weight' },
  Gram: { label: 'Gram', type: 'number', span: 'tiny', min: 0, icon: 'weight' },
  Amount: { label: 'Amount', type: 'number', span: 'small', min: 0, icon: 'banknote' },
};

const DETAIL_FIELDS = Object.keys(FIELD_META).filter((header) => header !== 'TrackingNumber');
const numericFields = new Set(['NoOfPcs', 'Kilo', 'Gram', 'Amount']);
const cityToOutstation = branches.city_to_outstation || {};
const cityNames = Object.keys(cityToOutstation).sort((a, b) => a.localeCompare(b));
const normalizedCityMap = new Map(cityNames.map((city) => [normalizeSearch(city), city]));

let settings = loadSettings();
let rows = loadRows();
let firebaseConfig = loadFirebaseConfig();
let firebaseDb = null;
let currentUser = loadSession();
let editingId = null;
let query = '';
let currentView = 'entry';
let formStep = 'tracking';
let draftRecord = emptyRecord();
let scannerStream = null;
let scannerFrame = null;
let adminUsers = [];
let adminSubmissions = [];
let userSubmissions = [];
let appMessage = '';
let uploadPopup = null;

initFirebase();

if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('/sw.js').catch(() => {});
  });
}

function emptyRecord() {
  return Object.fromEntries(
    HEADERS.map((header) => [header, excelOnlyDefaults()[header] ?? ''])
  );
}

function excelOnlyDefaults() {
  return {
    ...BASE_EXCEL_ONLY_DEFAULTS,
    PackageDescription: settings.defaultPackageDescription || 0,
  };
}

function loadSettings() {
  try {
    const parsed = JSON.parse(localStorage.getItem(SETTINGS_KEY) || '{}');
    return {
      defaultPackageDescription: parsed.defaultPackageDescription || '0',
    };
  } catch {
    return { defaultPackageDescription: '0' };
  }
}

function saveSettings() {
  localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings));
}

function loadFirebaseConfig() {
  try {
    const parsed = JSON.parse(localStorage.getItem(FIREBASE_CONFIG_KEY) || 'null');
    return parsed && parsed.apiKey ? parsed : DEFAULT_FIREBASE_CONFIG;
  } catch {
    return DEFAULT_FIREBASE_CONFIG;
  }
}

function saveFirebaseConfig(config) {
  firebaseConfig = config;
  localStorage.setItem(FIREBASE_CONFIG_KEY, JSON.stringify(config));
  initFirebase();
}

function initFirebase() {
  if (!firebaseConfig?.apiKey) return;
  try {
    const app = initializeApp(firebaseConfig);
    firebaseDb = getDatabase(app);
  } catch {
    firebaseDb = null;
  }
}

function loadSession() {
  try {
    return JSON.parse(localStorage.getItem(SESSION_KEY) || 'null');
  } catch {
    return null;
  }
}

function saveSession(user) {
  currentUser = user;
  localStorage.setItem(SESSION_KEY, JSON.stringify(user));
}

function clearSession() {
  currentUser = null;
  localStorage.removeItem(SESSION_KEY);
}

function loadRows() {
  try {
    const parsed = JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]');
    return Array.isArray(parsed) ? parsed.map(normalizeRow) : [];
  } catch {
    return [];
  }
}

function normalizeRow(row) {
  const normalized = emptyRecord();
  for (const header of HEADERS) normalized[header] = row?.[header] ?? normalized[header];
  for (const [header, value] of Object.entries(excelOnlyDefaults())) normalized[header] = value;
  normalized.id = row?.id || crypto.randomUUID();
  return normalized;
}

function saveRows() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(rows));
  if (currentUser?.role === 'user' && firebaseDb) {
    set(ref(firebaseDb, `drafts/${accountKey(currentUser.username)}`), {
      rows: rows.map((row) => normalizeRow(row)),
      updatedAt: Date.now(),
    }).catch(() => {});
  }
}

async function loadUserDraftRows(username) {
  try {
    const snap = await get(ref(firebaseDb, `drafts/${accountKey(username)}`));
    if (!snap.exists()) return loadRows();
    const draft = snap.val();
    return Array.isArray(draft.rows) ? draft.rows.map(normalizeRow) : [];
  } catch {
    return loadRows();
  }
}

async function clearUserDraftRows() {
  rows = [];
  localStorage.setItem(STORAGE_KEY, JSON.stringify(rows));
  if (currentUser?.role === 'user' && firebaseDb) {
    await remove(ref(firebaseDb, `drafts/${accountKey(currentUser.username)}`)).catch(() => {});
  }
}

function filteredRows() {
  const needle = query.trim().toLowerCase();
  if (!needle) return rows;
  return rows.filter((row) =>
    HEADERS.some((header) => String(row[header] || '').toLowerCase().includes(needle))
  );
}

async function render() {
  stopScanner();
  if (!firebaseDb) {
    renderFirebaseSetup();
    return;
  }
  if (!currentUser) {
    renderLogin();
    return;
  }
  if (currentUser.role === 'admin' && currentView === 'admin') {
    renderAdmin();
    return;
  }
  if (currentView === 'settings') {
    renderSettings();
    return;
  }

  if (currentUser?.role === 'user') {
    await loadUserSubmissions();
  }

  const current = draftRecord || emptyRecord();
  const records = filteredRows();
  const pendingUploads = userSubmissions.filter((item) => item.status !== 'exported');
  const exportedUploads = userSubmissions.filter((item) => item.status === 'exported');
  const totals = rows.reduce(
    (acc, row) => {
      acc.pieces += Number(row.NoOfPcs) || 0;
      acc.amount += Number(row.Amount) || 0;
      return acc;
    },
    { pieces: 0, amount: 0 }
  );

  document.querySelector('#app').innerHTML = `
    <main class="shell">
      <section class="brand-hero" aria-label="Application header">
        <div class="brand-mark">
          <img src="${logoUrl}" alt="Domex" />
        </div>
        <div class="hero-title">
          <p class="eyebrow">Fast Entry</p>
          <h1>${editingId ? 'Edit Pickup' : 'New Pickup'}</h1>
        </div>
        <button id="settings-btn" class="hero-scan" type="button" title="Settings">
          <i data-lucide="${currentUser.role === 'admin' ? 'bell' : 'upload'}"></i>
        </button>
      </section>

      <section class="app-card">
        ${messageTemplate()}
        ${uploadPopupTemplate()}
        <div class="top-actions">
          <button id="user-settings-btn" class="secondary compact-action" type="button" title="Default description settings">
            <i data-lucide="settings"></i>
            <span>Settings</span>
          </button>
          <label class="file-button" title="Import an existing Excel file">
            <i data-lucide="file-up"></i>
            <span>Import</span>
            <input id="import-file" type="file" accept=".xlsx,.xls" />
          </label>
          <button id="upload-btn" class="primary compact-action" type="button" title="Upload rows to admin">
            <i data-lucide="upload"></i>
            <span>Upload</span>
          </button>
          ${
            currentUser.role === 'admin'
              ? `<button id="admin-btn" class="secondary compact-action" type="button" title="Admin panel">
                  <i data-lucide="bell"></i>
                  <span>Admin</span>
                </button>`
              : ''
          }
          <button id="logout-btn" class="secondary compact-action" type="button" title="Logout">
            <i data-lucide="log-out"></i>
            <span>Logout</span>
          </button>
        </div>

        <section class="stats" aria-label="Entry summary">
          <div><span>Total Entries</span><strong>${rows.length}</strong></div>
          <div><span>Total Pieces</span><strong>${totals.pieces}</strong></div>
          <div><span>Total Amount</span><strong>${formatNumber(totals.amount)}</strong></div>
          <div><span>Pending</span><strong>${pendingUploads.length}</strong></div>
          <div><span>Exported</span><strong>${exportedUploads.length}</strong></div>
        </section>

        <div class="entry-layout">
          <form id="entry-form" class="entry-form" autocomplete="off">
            <div class="form-title">
              <div>
                <p class="form-kicker">${editingId ? 'Update row' : 'Fast entry'}</p>
                <h2>${formStep === 'tracking' ? 'Scan Tracking' : 'New Pickup'}</h2>
              </div>
              <button id="reset-form" class="ghost icon-only" type="button" title="Clear form">
                <i data-lucide="list-restart"></i>
              </button>
            </div>
            ${formStep === 'tracking' ? trackingStepTemplate(current) : detailsStepTemplate(current)}
          </form>

          <section class="table-panel" aria-label="Saved pickup entries">
            <div class="table-toolbar">
              <label class="search">
                <i data-lucide="search"></i>
                <input id="search-input" type="search" value="${escapeHtml(query)}" placeholder="Search entries" />
              </label>
            </div>
            <div class="table-wrap">
              <table>
                <thead>
                  <tr>
                    <th>Tracking No</th>
                    <th>Receiver</th>
                    <th>City</th>
                    <th>Contact</th>
                    <th>Pieces</th>
                    <th>Amount</th>
                    <th aria-label="Actions"></th>
                  </tr>
                </thead>
                <tbody>
                  ${
                    records.length
                      ? records.map(rowTemplate).join('')
                      : `<tr><td class="empty" colspan="7">No pickup entries yet.</td></tr>`
                  }
                </tbody>
              </table>
            </div>
          </section>
        </div>
        <section class="settings-panel user-history-panel">
          <div class="panel-heading">
            <div><p class="form-kicker">Your uploads</p><h2>Pending and Exported</h2></div>
          </div>
          <div class="history-columns">
            <div>
              <h3>Pending</h3>
              <div class="submission-list">
                ${pendingUploads.map(userSubmissionTemplate).join('') || '<p class="empty">No pending uploads.</p>'}
              </div>
            </div>
            <div>
              <h3>Exported</h3>
              <div class="submission-list">
                ${exportedUploads.map(userSubmissionTemplate).join('') || '<p class="empty">No exported uploads yet.</p>'}
              </div>
            </div>
          </div>
        </section>
      </section>
    </main>
  `;

  createIcons({
    icons: {
      ArrowLeft,
      Banknote,
      Bell,
      Box,
      Building2,
      Camera,
      ChevronRight,
      Download,
      FileUp,
      KeyRound,
      ListRestart,
      LogOut,
      MapPin,
      Package,
      Phone,
      Plus,
      Save,
      Search,
      Settings,
      ShieldCheck,
      Trash2,
      Upload,
      User,
      Users,
      Weight,
    },
  });
  bindEvents();
}

function uploadPopupTemplate() {
  if (!uploadPopup) return '';
  return `
    <div class="modal-backdrop" role="dialog" aria-modal="true" aria-labelledby="upload-success-title">
      <div class="success-modal">
        <div class="success-icon">
          <i data-lucide="shield-check"></i>
        </div>
        <div>
          <p class="form-kicker">Upload Complete</p>
          <h2 id="upload-success-title">Successfully Uploaded</h2>
          <p>${escapeHtml(uploadPopup.count)} rows uploaded to the admin notification panel.</p>
        </div>
        <div class="modal-actions">
          <button id="keep-entry-btn" class="secondary" type="button">Keep Entries</button>
          <button id="clear-entry-btn" class="primary" type="button">
            <i data-lucide="trash-2"></i>
            <span>Clear Entry</span>
          </button>
        </div>
      </div>
    </div>
  `;
}

function renderFirebaseSetup() {
  document.querySelector('#app').innerHTML = `
    <main class="shell settings-shell">
      <section class="brand-hero setup-hero">
        <div class="brand-mark"><img src="${logoUrl}" alt="Domex" /></div>
        <div class="hero-title">
          <p class="eyebrow">Firebase Setup</p>
          <h1>Connect Database</h1>
        </div>
      </section>
      <form id="firebase-form" class="settings-panel auth-panel">
        <div>
          <p class="form-kicker">Required once</p>
          <h2>Paste Firebase web config JSON</h2>
        </div>
        ${messageTemplate()}
        <label class="field wide">
          <span>Firebase config</span>
          <div class="input-shell textarea-shell">
            <i data-lucide="settings"></i>
            <textarea id="firebaseConfig" rows="9" required placeholder='{"apiKey":"...","authDomain":"...","projectId":"..."}'></textarea>
          </div>
        </label>
        <p class="settings-note">Create a Firebase web app, copy its config object, and paste it here. This app uses Firestore collections named users and submissions.</p>
        <div class="settings-actions">
          <button class="primary" type="submit"><i data-lucide="save"></i><span>Save Firebase</span></button>
        </div>
      </form>
    </main>
  `;
  createIcons({ icons: { Save, Settings } });
  document.querySelector('#firebase-form').addEventListener('submit', (event) => {
    event.preventDefault();
    try {
      const config = JSON.parse(document.querySelector('#firebaseConfig').value);
      saveFirebaseConfig(config);
      appMessage = 'Firebase connected. Login as admin to add users.';
      render();
    } catch {
      appMessage = 'Invalid Firebase config JSON.';
      renderFirebaseSetup();
    }
  });
}

function renderLogin() {
  document.querySelector('#app').innerHTML = `
    <main class="shell settings-shell">
      <section class="brand-hero setup-hero">
        <div class="brand-mark"><img src="${logoUrl}" alt="Domex" /></div>
        <div class="hero-title">
          <p class="eyebrow">Secure Entry</p>
          <h1>Login</h1>
        </div>
      </section>
      <section class="app-card">
        ${messageTemplate()}
        <form id="login-form" class="settings-panel auth-panel">
          <div>
            <p class="form-kicker">Username and password</p>
            <h2>Account Login</h2>
          </div>
          <label class="field wide">
            <span>Username</span>
            <div class="input-shell"><i data-lucide="user"></i><input id="loginUsername" required autocomplete="username" /></div>
          </label>
          <label class="field wide">
            <span>Password</span>
            <div class="input-shell"><i data-lucide="key-round"></i><input id="loginPassword" type="password" required autocomplete="current-password" /></div>
          </label>
          <div class="settings-actions">
            <button class="primary" type="submit"><i data-lucide="chevron-right"></i><span>Login</span></button>
          </div>
        </form>
      </section>
    </main>
  `;
  createIcons({ icons: { ChevronRight, KeyRound, User } });
  document.querySelector('#login-form').addEventListener('submit', handleLogin);
}

async function renderAdmin() {
  await loadAdminData();
  const pendingCount = adminSubmissions.filter((item) => item.status !== 'exported').length;
  document.querySelector('#app').innerHTML = `
    <main class="shell">
      <section class="brand-hero">
        <div class="brand-mark"><img src="${logoUrl}" alt="Domex" /></div>
        <div class="hero-title">
          <p class="eyebrow">Admin Panel</p>
          <h1>Notifications</h1>
        </div>
        <button id="back-entry-btn" class="hero-scan" type="button" title="Entry">
          <i data-lucide="arrow-left"></i>
        </button>
      </section>
      <section class="app-card admin-grid">
        ${messageTemplate()}
        <section class="stats">
          <div><span>Users</span><strong>${adminUsers.length}</strong></div>
          <div><span>Uploads</span><strong>${adminSubmissions.length}</strong></div>
          <div><span>Pending</span><strong>${pendingCount}</strong></div>
        </section>
        <section class="settings-panel">
          <div class="panel-heading">
            <div><p class="form-kicker">Accounts</p><h2>Add User</h2></div>
          </div>
          <form id="user-form" class="user-form-grid">
            <label class="field"><span>Username</span><div class="input-shell"><i data-lucide="user"></i><input id="newUsername" required /></div></label>
            <label class="field"><span>Password</span><div class="input-shell"><i data-lucide="key-round"></i><input id="newPassword" required /></div></label>
            <button class="primary" type="submit"><i data-lucide="plus"></i><span>Add User</span></button>
          </form>
          <div class="mini-list">
            ${adminUsers.map(userTemplate).join('') || '<p class="empty">No users yet.</p>'}
          </div>
        </section>
        <section class="settings-panel">
          <div class="panel-heading">
            <div><p class="form-kicker">Uploads from users</p><h2>Notification Panel</h2></div>
            <div class="submission-actions">
              <button id="admin-settings-btn" class="secondary" type="button"><i data-lucide="settings"></i><span>Settings</span></button>
              <button id="logout-btn" class="secondary" type="button"><i data-lucide="log-out"></i><span>Logout</span></button>
            </div>
          </div>
          <div class="submission-list">
            ${adminSubmissions.map(submissionTemplate).join('') || '<p class="empty">No uploaded data yet.</p>'}
          </div>
        </section>
      </section>
    </main>
  `;
  createIcons({
    icons: { ArrowLeft, Bell, Download, KeyRound, LogOut, Plus, Settings, ShieldCheck, Trash2, User, Users },
  });
  bindAdminEvents();
}

function renderSettings() {
  document.querySelector('#app').innerHTML = `
    <main class="shell settings-shell">
      <section class="topbar" aria-label="Settings header">
        <div>
          <p class="eyebrow">Domex</p>
          <h1>Settings</h1>
        </div>
        <button id="back-btn" class="secondary" type="button" title="Back to data entry">
          <i data-lucide="arrow-left"></i>
          <span>Back</span>
        </button>
      </section>

      <form id="settings-form" class="settings-panel" autocomplete="off">
        <div>
          <p class="form-kicker">Excel defaults</p>
          <h2>Default Package Description</h2>
        </div>
        <label class="field wide">
          <span>PackageDescription value</span>
          <textarea id="defaultPackageDescription" name="defaultPackageDescription" rows="4">${escapeHtml(
            settings.defaultPackageDescription
          )}</textarea>
        </label>
        <p class="settings-note">
          This value is hidden from the entry form and written to the Excel PackageDescription column for every exported row.
        </p>
        <div class="settings-actions">
          <button class="primary" type="submit">
            <i data-lucide="save"></i>
            <span>Save Settings</span>
          </button>
        </div>
      </form>
    </main>
  `;

  createIcons({
    icons: { ArrowLeft, Save },
  });
  bindSettingsEvents();
}

function fieldTemplate(header, value) {
  if (header === 'ReceiverCity') return cityFieldTemplate(value);

  const meta = FIELD_META[header];
  const common = `
    id="${header}"
    name="${header}"
    ${meta.min !== undefined ? `min="${meta.min}"` : ''}
    required
    value="${escapeHtml(value)}"
  `;

  return `
    <label class="field ${meta.span}">
      <span>${meta.label}</span>
      ${
        meta.type === 'textarea'
          ? `<div class="input-shell textarea-shell">
              <i data-lucide="${meta.icon}"></i>
              <textarea id="${header}" name="${header}" rows="3" required>${escapeHtml(value)}</textarea>
            </div>`
          : `<div class="input-shell">
              <i data-lucide="${meta.icon}"></i>
              <input ${common} type="${meta.type}" />
            </div>`
      }
    </label>
  `;
}

function cityFieldTemplate(value) {
  const selectedCity = normalizedCityMap.get(normalizeSearch(value)) || '';
  return `
    <label class="field city-field medium">
      <span>Receiver City</span>
      <div class="input-shell">
        <i data-lucide="search"></i>
        <input
          id="ReceiverCitySearch"
          type="search"
          autocomplete="off"
          value="${escapeHtml(selectedCity)}"
          placeholder="Type to search city"
          aria-describedby="city-error"
        />
      </div>
      <input id="ReceiverCity" name="ReceiverCity" type="hidden" required value="${escapeHtml(selectedCity)}" />
      <div id="city-options-panel" class="city-options" role="listbox" hidden>
        ${cityOptionsTemplate(selectedCity)}
      </div>
      <small id="city-error" class="field-error">Select a city from the list.</small>
    </label>
  `;
}

function cityOptionsTemplate(searchTerm) {
  const options = cityMatches(searchTerm).slice(0, 24);
  if (!options.length) return '<button type="button" disabled>No matching city</button>';
  return options
    .map(
      (city) => `
        <button type="button" data-city="${escapeHtml(city)}" role="option">
          <i data-lucide="building-2"></i>
          ${escapeHtml(city)}
        </button>
      `
    )
    .join('');
}

function cityMatches(searchTerm) {
  const needle = normalizeSearch(searchTerm);
  if (!needle) return cityNames.slice(0, 24);
  return cityNames.filter((city) => normalizeSearch(city).includes(needle));
}

function trackingStepTemplate(current) {
  return `
    <div class="step-indicator" aria-label="Entry step">
      <span class="active">1 Tracking</span>
      <span>2 Details</span>
    </div>
    <div class="scanner-panel">
      <div class="scanner-view">
        <video id="scanner-video" playsinline muted></video>
        <div class="scan-frame" aria-hidden="true"></div>
        <p id="scanner-status">Tap Scan Barcode, or type the tracking number manually.</p>
      </div>
      <button id="scan-btn" class="secondary" type="button">
        <i data-lucide="camera"></i>
        <span>Scan Barcode</span>
      </button>
    </div>
    <label class="field wide tracking-input">
      <span>Tracking No</span>
      <input id="TrackingNumber" name="TrackingNumber" type="text" inputmode="text" value="${escapeHtml(
        current.TrackingNumber
      )}" placeholder="Scan or type tracking number" required />
    </label>
    <div class="form-actions wizard-actions">
      <button id="next-step" class="primary" type="button">
        <i data-lucide="chevron-right"></i>
        <span>Next</span>
      </button>
    </div>
  `;
}

function detailsStepTemplate(current) {
  return `
    <div class="step-indicator" aria-label="Entry step">
      <span>1 Tracking</span>
      <span class="active">2 Details</span>
    </div>
    <div class="tracking-summary">
      <div class="tracking-badge">
        <i data-lucide="package"></i>
      </div>
      <div>
        <span>Tracking No</span>
        <strong>${escapeHtml(current.TrackingNumber || '-')}</strong>
      </div>
      <i class="tracking-ok" data-lucide="shield-check"></i>
    </div>
    <div class="form-grid">
      ${DETAIL_FIELDS.map((header) => fieldTemplate(header, current[header])).join('')}
    </div>
    <div class="form-actions wizard-actions split-actions">
      <button id="back-step" class="secondary" type="button">
        <i data-lucide="arrow-left"></i>
        <span>Back</span>
      </button>
      <button class="primary" type="submit">
        <i data-lucide="${editingId ? 'save' : 'plus'}"></i>
        <span>${editingId ? 'Save Row' : 'Add Row'}</span>
      </button>
    </div>
  `;
}

function rowTemplate(row) {
  return `
    <tr>
      <td>${escapeHtml(row.TrackingNumber)}</td>
      <td>${escapeHtml(row.ReceiverName)}</td>
      <td>${escapeHtml(row.ReceiverCity)}</td>
      <td>${escapeHtml(row.ReceiverContactNo)}</td>
      <td>${escapeHtml(row.NoOfPcs)}</td>
      <td>${escapeHtml(row.Amount)}</td>
      <td class="row-actions">
        <button data-edit="${row.id}" type="button">Edit</button>
        <button data-delete="${row.id}" class="danger icon-only" type="button" title="Delete row">
          <i data-lucide="trash-2"></i>
        </button>
      </td>
    </tr>
  `;
}

function messageTemplate() {
  if (!appMessage) return '';
  const message = appMessage;
  appMessage = '';
  return `<div class="app-message">${escapeHtml(message)}</div>`;
}

function userTemplate(user) {
  return `
    <div class="mini-row">
      <div>
        <strong>${escapeHtml(user.username)}</strong>
        <span>${user.active === false ? 'Disabled' : 'Active'}</span>
      </div>
      <button class="danger icon-only" type="button" data-delete-user="${escapeHtml(user.username)}" title="Delete user">
        <i data-lucide="trash-2"></i>
      </button>
    </div>
  `;
}

function submissionTemplate(item) {
  const date = item.createdAt ? new Date(item.createdAt).toLocaleString() : 'Uploaded';
  const canDelete = item.status === 'exported';
  return `
    <div class="submission-card">
      <div>
        <p class="form-kicker">${escapeHtml(item.status || 'pending')}</p>
        <h3>${escapeHtml(item.user || 'user')} uploaded ${item.count || item.rows?.length || 0} rows</h3>
        <span>${escapeHtml(date)}</span>
      </div>
      <div class="submission-actions">
        <button class="primary" type="button" data-export-submission="${escapeHtml(item.id)}">
          <i data-lucide="download"></i>
          <span>Export Excel</span>
        </button>
        <button class="secondary" type="button" data-mark-exported="${escapeHtml(item.id)}">
          <i data-lucide="shield-check"></i>
          <span>Mark Exported</span>
        </button>
        ${
          canDelete
            ? `<button class="danger" type="button" data-delete-submission="${escapeHtml(item.id)}">
                <i data-lucide="trash-2"></i>
                <span>Delete</span>
              </button>`
            : ''
        }
      </div>
    </div>
  `;
}

function userSubmissionTemplate(item) {
  const date = item.createdAt ? new Date(item.createdAt).toLocaleString() : 'Uploaded';
  return `
    <div class="submission-card compact-submission">
      <div>
        <p class="form-kicker">${escapeHtml(item.status || 'pending')}</p>
        <h3>${item.count || item.rows?.length || 0} rows</h3>
        <span>${escapeHtml(date)}</span>
      </div>
    </div>
  `;
}

function bindEvents() {
  const form = document.querySelector('#entry-form');

  form.addEventListener('submit', (event) => {
    event.preventDefault();
    if (formStep !== 'details') return;
    const data = Object.fromEntries(new FormData(form).entries());
    const invalidField = firstInvalidDetailField(data);
    if (invalidField) {
      showValidationMessage(invalidField);
      return;
    }
    for (const field of numericFields) data[field] = data[field] === '' ? '' : Number(data[field]);
    Object.assign(data, excelOnlyDefaults(), { TrackingNumber: draftRecord.TrackingNumber });

    if (editingId) {
      rows = rows.map((row) => (row.id === editingId ? normalizeRow({ ...row, ...data }) : row));
    } else {
      rows = [normalizeRow({ ...data, id: crypto.randomUUID() }), ...rows];
    }
    editingId = null;
    formStep = 'tracking';
    draftRecord = emptyRecord();
    saveRows();
    render();
  });

  document.querySelector('#next-step')?.addEventListener('click', () => {
    const trackingNumber = document.querySelector('#TrackingNumber').value.trim();
    if (!trackingNumber) {
      document.querySelector('#TrackingNumber').focus();
      setScannerStatus('Tracking number is required.');
      return;
    }
    draftRecord = normalizeRow({ ...draftRecord, TrackingNumber: trackingNumber });
    formStep = 'details';
    render();
    document.querySelector('#ReceiverName')?.focus();
  });

  document.querySelector('#back-step')?.addEventListener('click', () => {
    draftRecord = normalizeRow({
      ...draftRecord,
      ...Object.fromEntries(new FormData(form).entries()),
    });
    formStep = 'tracking';
    render();
    document.querySelector('#TrackingNumber')?.focus();
  });

  document.querySelector('#scan-btn')?.addEventListener('click', startScanner);
  bindCityPicker();

  document.querySelector('#settings-btn').addEventListener('click', () => {
    stopScanner();
    if (currentUser.role === 'admin') {
      currentView = 'admin';
      render();
    } else {
      uploadRows();
    }
  });

  document.querySelector('#admin-btn')?.addEventListener('click', () => {
    currentView = 'admin';
    render();
  });

  document.querySelector('#logout-btn')?.addEventListener('click', () => {
    clearSession();
    currentView = 'entry';
    render();
  });

  document.querySelector('#admin-settings-btn')?.addEventListener('click', () => {
    currentView = 'settings';
    render();
  });

  document.querySelector('#user-settings-btn')?.addEventListener('click', () => {
    currentView = 'settings';
    render();
  });

  document.querySelector('#reset-form').addEventListener('click', () => {
    stopScanner();
    editingId = null;
    formStep = 'tracking';
    draftRecord = emptyRecord();
    render();
  });

  document.querySelector('#search-input').addEventListener('input', (event) => {
    query = event.target.value;
    render();
  });

  document.querySelector('#upload-btn').addEventListener('click', uploadRows);
  document.querySelector('#import-file').addEventListener('change', importWorkbook);

  document.querySelector('#keep-entry-btn')?.addEventListener('click', () => {
    uploadPopup = null;
    render();
  });

  document.querySelector('#clear-entry-btn')?.addEventListener('click', async () => {
    await clearUserDraftRows();
    editingId = null;
    formStep = 'tracking';
    draftRecord = emptyRecord();
    uploadPopup = null;
    appMessage = 'Entries cleared.';
    render();
  });

  document.querySelectorAll('[data-edit]').forEach((button) => {
    button.addEventListener('click', () => {
      editingId = button.dataset.edit;
      draftRecord = normalizeRow(rows.find((row) => row.id === editingId) || emptyRecord());
      formStep = 'tracking';
      render();
      document.querySelector('#TrackingNumber').focus();
    });
  });

  document.querySelectorAll('[data-delete]').forEach((button) => {
    button.addEventListener('click', () => {
      rows = rows.filter((row) => row.id !== button.dataset.delete);
      if (editingId === button.dataset.delete) editingId = null;
      saveRows();
      render();
    });
  });
}

function bindCityPicker() {
  const searchInput = document.querySelector('#ReceiverCitySearch');
  const hiddenInput = document.querySelector('#ReceiverCity');
  const optionsPanel = document.querySelector('#city-options-panel');
  if (!searchInput || !hiddenInput || !optionsPanel) return;

  const refreshOptions = () => {
    optionsPanel.innerHTML = cityOptionsTemplate(searchInput.value);
    createIcons({ icons: { Building2 } });
    bindCityOptionButtons();
  };

  searchInput.addEventListener('input', () => {
    const exactCity = normalizedCityMap.get(normalizeSearch(searchInput.value));
    hiddenInput.value = exactCity && exactCity === hiddenInput.value ? exactCity : '';
    searchInput.classList.toggle('invalid', !hiddenInput.value);
    refreshOptions();
    optionsPanel.hidden = !searchInput.value.trim();
  });

  searchInput.addEventListener('focus', () => {
    if (searchInput.value.trim() && !hiddenInput.value) {
      refreshOptions();
      optionsPanel.hidden = false;
    }
  });
  bindCityOptionButtons();
}

function bindCityOptionButtons() {
  const searchInput = document.querySelector('#ReceiverCitySearch');
  const hiddenInput = document.querySelector('#ReceiverCity');
  if (!searchInput || !hiddenInput) return;
  document.querySelectorAll('[data-city]').forEach((button) => {
    button.addEventListener('click', () => {
      const city = button.dataset.city;
      searchInput.value = city;
      hiddenInput.value = city;
      searchInput.classList.remove('invalid');
      document.querySelector('#city-options-panel').hidden = true;
      document.querySelector('#ReceiverContactNo')?.focus();
    });
  });
}

function firstInvalidDetailField(data) {
  for (const header of DETAIL_FIELDS) {
    const value = String(data[header] ?? '').trim();
    if (!value) return header;
  }
  if (!normalizedCityMap.has(normalizeSearch(data.ReceiverCity))) return 'ReceiverCity';
  return '';
}

function showValidationMessage(header) {
  if (header === 'ReceiverCity') {
    const searchInput = document.querySelector('#ReceiverCitySearch');
    searchInput?.classList.add('invalid');
    searchInput?.focus();
    return;
  }
  const field = document.querySelector(`#${header}`);
  field?.focus();
  field?.reportValidity?.();
}

async function startScanner() {
  stopScanner();
  const video = document.querySelector('#scanner-video');
  const input = document.querySelector('#TrackingNumber');
  if (!video || !input) return;

  if (!navigator.mediaDevices?.getUserMedia) {
    setScannerStatus('Camera is not available. Type the tracking number manually.');
    input.focus();
    return;
  }

  if (!('BarcodeDetector' in window)) {
    setScannerStatus('Barcode scan is not supported in this browser. Type the tracking number manually.');
    input.focus();
    return;
  }

  try {
    setScannerStatus('Opening camera...');
    scannerStream = await navigator.mediaDevices.getUserMedia({
      video: { facingMode: { ideal: 'environment' } },
      audio: false,
    });
    video.srcObject = scannerStream;
    await video.play();
    setScannerStatus('Point the camera at the barcode.');
    const detector = new BarcodeDetector();

    const scan = async () => {
      if (!scannerStream || formStep !== 'tracking') return;
      try {
        const codes = await detector.detect(video);
        const value = codes?.[0]?.rawValue;
        if (value) {
          input.value = value.trim();
          draftRecord = normalizeRow({ ...draftRecord, TrackingNumber: input.value });
          setScannerStatus('Barcode scanned. Tap Next.');
          stopScanner(false);
          return;
        }
      } catch {
        setScannerStatus('Could not read barcode yet. Keep it inside the box.');
      }
      scannerFrame = requestAnimationFrame(scan);
    };

    scannerFrame = requestAnimationFrame(scan);
  } catch {
    setScannerStatus('Camera permission denied. Type the tracking number manually.');
    input.focus();
  }
}

async function handleLogin(event) {
  event.preventDefault();
  const username = document.querySelector('#loginUsername').value.trim();
  const password = document.querySelector('#loginPassword').value;

  if (username === ADMIN_USERNAME && password === ADMIN_PASSWORD) {
    saveSession({ username, role: 'admin' });
    currentView = 'admin';
    appMessage = 'Admin login successful.';
    render();
    return;
  }

  try {
    const snap = await get(ref(firebaseDb, `users/${accountKey(username)}`));
    if (!snap.exists()) throw new Error('missing');
    const user = snap.val();
    if (user.active === false || user.password !== password) throw new Error('invalid');
    saveSession({ username: user.username, role: 'user' });
    rows = await loadUserDraftRows(user.username);
    localStorage.setItem(STORAGE_KEY, JSON.stringify(rows));
    draftRecord = emptyRecord();
    formStep = 'tracking';
    currentView = 'entry';
    appMessage = 'Login successful.';
    render();
  } catch {
    appMessage = 'Invalid username or password.';
    renderLogin();
  }
}

async function uploadRows() {
  if (!rows.length) {
    appMessage = 'No rows to upload.';
    render();
    return;
  }
  try {
    const uploadRows = rows.map((row) => {
      const normalized = normalizeRow(row);
      return Object.fromEntries(HEADERS.map((header) => [header, normalized[header] ?? '']));
    });
    await push(ref(firebaseDb, 'submissions'), {
      user: currentUser.username,
      rows: uploadRows,
      count: uploadRows.length,
      status: 'pending',
      createdAt: Date.now(),
    });
    uploadPopup = { count: uploadRows.length };
    render();
  } catch {
    appMessage = 'Upload failed. Check Firebase connection/rules.';
    render();
  }
}

async function loadAdminData() {
  try {
    const usersSnap = await get(ref(firebaseDb, 'users'));
    adminUsers = usersSnap.exists()
      ? Object.entries(usersSnap.val())
          .map(([id, value]) => ({ id, ...value }))
          .sort((a, b) => String(a.username).localeCompare(String(b.username)))
      : [];
  } catch {
    adminUsers = [];
  }

  try {
    const submissionsSnap = await get(ref(firebaseDb, 'submissions'));
    adminSubmissions = submissionsSnap.exists()
      ? Object.entries(submissionsSnap.val())
          .map(([id, value]) => ({ id, ...value }))
          .sort((a, b) => Number(b.createdAt || 0) - Number(a.createdAt || 0))
      : [];
  } catch {
    adminSubmissions = [];
  }
}

async function loadUserSubmissions() {
  if (!currentUser?.username || !firebaseDb) {
    userSubmissions = [];
    return;
  }
  try {
    const submissionsSnap = await get(ref(firebaseDb, 'submissions'));
    userSubmissions = submissionsSnap.exists()
      ? Object.entries(submissionsSnap.val())
          .map(([id, value]) => ({ id, ...value }))
          .filter((item) => item.user === currentUser.username)
          .sort((a, b) => Number(b.createdAt || 0) - Number(a.createdAt || 0))
      : [];
  } catch {
    userSubmissions = [];
  }
}

async function addUserAccount(event) {
  event.preventDefault();
  const username = document.querySelector('#newUsername').value.trim();
  const password = document.querySelector('#newPassword').value.trim();
  if (!username || !password) return;
  try {
    await set(ref(firebaseDb, `users/${accountKey(username)}`), {
      username,
      password,
      active: true,
      createdAt: Date.now(),
    });
    appMessage = `User ${username} added.`;
    renderAdmin();
  } catch {
    appMessage = 'Could not add user. Check Firebase rules.';
    renderAdmin();
  }
}

async function deleteUserAccount(username) {
  try {
    await remove(ref(firebaseDb, `users/${accountKey(username)}`));
    appMessage = `User ${username} deleted.`;
    renderAdmin();
  } catch {
    appMessage = 'Could not delete user.';
    renderAdmin();
  }
}

async function markSubmissionExported(id) {
  try {
    await update(ref(firebaseDb, `submissions/${id}`), {
      status: 'exported',
      exportedAt: Date.now(),
    });
    appMessage = 'Submission marked as exported.';
    renderAdmin();
  } catch {
    appMessage = 'Could not update submission.';
    renderAdmin();
  }
}

async function deleteExportedSubmission(id) {
  const submission = adminSubmissions.find((item) => item.id === id);
  if (!submission || submission.status !== 'exported') {
    appMessage = 'Only exported uploads can be deleted.';
    renderAdmin();
    return;
  }
  try {
    await remove(ref(firebaseDb, `submissions/${id}`));
    appMessage = 'Exported upload deleted.';
    renderAdmin();
  } catch {
    appMessage = 'Could not delete upload.';
    renderAdmin();
  }
}

function exportSubmission(id) {
  const submission = adminSubmissions.find((item) => item.id === id);
  if (!submission) return;
  exportRowsAsWorkbook(
    submission.rows || [],
    `Domex_${submission.user || 'User'}_${new Date().toISOString().slice(0, 10)}.xlsx`
  );
}

function stopScanner(clearStatus = true) {
  if (scannerFrame) cancelAnimationFrame(scannerFrame);
  scannerFrame = null;
  if (scannerStream) {
    scannerStream.getTracks().forEach((track) => track.stop());
  }
  scannerStream = null;
  const video = document.querySelector('#scanner-video');
  if (video) video.srcObject = null;
  if (clearStatus) setScannerStatus('');
}

function setScannerStatus(message) {
  const status = document.querySelector('#scanner-status');
  if (status && message) status.textContent = message;
}

function bindSettingsEvents() {
  document.querySelector('#back-btn').addEventListener('click', () => {
    currentView = currentUser?.role === 'admin' ? 'admin' : 'entry';
    render();
  });

  document.querySelector('#settings-form').addEventListener('submit', (event) => {
    event.preventDefault();
    const data = Object.fromEntries(new FormData(event.currentTarget).entries());
    settings.defaultPackageDescription = String(data.defaultPackageDescription || '').trim() || '0';
    rows = rows.map(normalizeRow);
    saveSettings();
    saveRows();
    currentView = currentUser?.role === 'admin' ? 'admin' : 'entry';
    render();
  });
}

function bindAdminEvents() {
  document.querySelector('#back-entry-btn')?.addEventListener('click', () => {
    currentView = 'entry';
    render();
  });

  document.querySelector('#logout-btn')?.addEventListener('click', () => {
    clearSession();
    currentView = 'entry';
    render();
  });

  document.querySelector('#user-form')?.addEventListener('submit', addUserAccount);

  document.querySelectorAll('[data-delete-user]').forEach((button) => {
    button.addEventListener('click', () => deleteUserAccount(button.dataset.deleteUser));
  });

  document.querySelectorAll('[data-export-submission]').forEach((button) => {
    button.addEventListener('click', () => exportSubmission(button.dataset.exportSubmission));
  });

  document.querySelectorAll('[data-mark-exported]').forEach((button) => {
    button.addEventListener('click', () => markSubmissionExported(button.dataset.markExported));
  });

  document.querySelectorAll('[data-delete-submission]').forEach((button) => {
    button.addEventListener('click', () => deleteExportedSubmission(button.dataset.deleteSubmission));
  });
}

function exportWorkbook() {
  const exportRows = rows
    .slice()
    .reverse()
    .map((row) => {
      const normalized = normalizeRow(row);
      return Object.fromEntries(HEADERS.map((header) => [header, normalized[header] ?? '']));
    });
  exportRowsAsWorkbook(exportRows, `Domex_Pickups_${new Date().toISOString().slice(0, 10)}.xlsx`);
}

function exportRowsAsWorkbook(exportRows, fileName) {
  const worksheet = XLSX.utils.json_to_sheet(exportRows, { header: HEADERS, skipHeader: false });
  worksheet['!cols'] = HEADERS.map((header) => ({ wch: Math.max(header.length + 2, 14) }));
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Pickups');
  XLSX.writeFile(workbook, fileName);
}

async function importWorkbook(event) {
  const [file] = event.target.files;
  if (!file) return;
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const imported = XLSX.utils.sheet_to_json(sheet, { defval: '' });
  rows = imported.map((row) => {
    const normalized = normalizeRow(row);
    return normalized;
  });
  editingId = null;
  saveRows();
  render();
}

function formatNumber(value) {
  return new Intl.NumberFormat('en-LK', { maximumFractionDigits: 2 }).format(value || 0);
}

function normalizeSearch(value) {
  return String(value || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function accountKey(value) {
  return normalizeSearch(value).replace(/[.#$/[\]]/g, '_');
}

function escapeHtml(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}

render();
