import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js";
import { getAuth, onAuthStateChanged, signOut, createUserWithEmailAndPassword, setPersistence, browserSessionPersistence, updatePassword, reauthenticateWithCredential, EmailAuthProvider } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js";
import { getFirestore, collection, onSnapshot, addDoc, deleteDoc, doc, updateDoc, setDoc, getDoc, getDocs, writeBatch, query, where } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js";

const firebaseConfig = {
    apiKey: "AIzaSyCAKw7quIPloQtMMdt4pp3aeezAKzx39hA",
    authDomain: "radio-system-6aaad.firebaseapp.com",
    projectId: "radio-system-6aaad",
    storageBucket: "radio-system-6aaad.firebasestorage.app",
    messagingSenderId: "528567280152",
    appId: "1:528567280152:web:724638a00da58fc2c2196d"
};

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);
const secondaryApp = initializeApp(firebaseConfig, "Secondary");
const secondaryAuth = getAuth(secondaryApp);

setPersistence(auth, browserSessionPersistence);

let currentUser = null;
let isAdmin = false;
let isReadOnly = false;
let devicesData = [];
let storeData = [];
let purchasesData = [];
let archiveData = [];
let usersData = [];
let currentSearchResults = [];
let currentSearchType = '';

window.cleanText = function(text) {
    if (!text) return '';
    return String(text).trim().replace(/\s+/g, ' ');
};

function cleanArabic(text) {
    if (!text) return '';
    return cleanText(text)
        .replace(/[يى]/g, 'ي')
        .replace(/[ةه]/g, 'ة')
        .replace(/[أإآ]/g, 'ا')
        .replace(/ؤ/g, 'و')
        .replace(/ئ/g, 'ي');
}
window.cleanArabic = cleanArabic;

function normalizeHeader(text) {
    if (!text) return '';
    return cleanArabic(String(text).trim())
        .replace(/[''`]/g, '')
        .replace(/\s+/g, ' ')
        .trim();
}

function mapExcelColumns(json) {
    if (!json.length) return { data: json, mapping: {} };
    var headers = Object.keys(json[0]);
    var normalizedHeaders = headers.map(function(h) { return normalizeHeader(h); });

    var fieldAliases = {
        serial:       ['رقم التسلسل', 'تسلسل', 'serial', 'رقم'],
        deviceName:   ['اسم الجهاز', 'اسم جهاز', 'جهاز', 'device name', 'device'],
        deviceType:   ['نوع الجهاز', 'نوع جهاز', 'type', 'device type', 'النوع'],
        id:           ['id', 'آيدي', 'ايدي', 'رقم تعريفي'],
        location:     ['مكان العمل', 'المكان', 'مكان', 'القسم', 'location', 'workplace'],
        date:         ['تاريخ التسليم', 'التاريخ', 'تاريخ', 'date', 'تاريخ الاستلام'],
        receiverName: ['المستلم', 'اسم المستلم', 'المستلم اسم', 'receiver', 'اسم الشخص المستلم'],
        empId:        ['رقم التوظيف', 'رقم الموظف', 'الرقم الوظيفي', 'emp id', 'employee id', 'رقم'],
        jobTitle:     ['الصفة', 'الصفة الوظيفية', 'المسمى الوظيفي', 'الوظيفة', 'job title', 'الصفه', 'الصفه الوظيفيه'],
        department:   ['الإدارة', 'الادارة', 'القسم', 'department'],
        phone:        ['الهاتف', 'رقم الهاتف', 'جوال', 'الجوال', 'phone', 'mobile', 'رقم الجوال'],
        plateNum:     ['نوع ورقم السيارة', 'رقم اللوحة', 'اللوحة', 'plate', 'plate number', 'رقم السيارة', 'السيارة'],
        notes:        ['الملاحظات', 'ملاحظات', 'notes', 'ملاحظة']
    };

    var mapping = {};
    var usedIndices = {};

    Object.keys(fieldAliases).forEach(function(field) {
        var aliases = fieldAliases[field].map(function(a) { return normalizeHeader(a); });
        for (var ai = 0; ai < aliases.length; ai++) {
            for (var hi = 0; hi < normalizedHeaders.length; hi++) {
                if (usedIndices[hi]) continue;
                if (normalizedHeaders[hi] === aliases[ai]) {
                    mapping[field] = headers[hi];
                    usedIndices[hi] = true;
                    return;
                }
            }
        }
    });

    Object.keys(fieldAliases).forEach(function(field) {
        if (mapping[field]) return;
        var aliases = fieldAliases[field].map(function(a) { return normalizeHeader(a); });
        for (var hi = 0; hi < normalizedHeaders.length; hi++) {
            if (usedIndices[hi]) continue;
            for (var ai = 0; ai < aliases.length; ai++) {
                if (normalizedHeaders[hi].includes(aliases[ai]) || aliases[ai].includes(normalizedHeaders[hi])) {
                    if (aliases[ai].length >= 2 && normalizedHeaders[hi].length >= 2) {
                        mapping[field] = headers[hi];
                        usedIndices[hi] = true;
                        return;
                    }
                }
            }
        }
    });

    return { data: json, mapping: mapping };
}

function getCellValue(row, mapping, field, fallbacks) {
    if (mapping[field]) {
        var val = row[mapping[field]];
        if (val !== undefined && val !== null && val !== '') return val;
    }
    if (fallbacks) {
        for (var fi = 0; fi < fallbacks.length; fi++) {
            var fb = fallbacks[fi];
            if (mapping[fb]) {
                var val2 = row[mapping[fb]];
                if (val2 !== undefined && val2 !== null && val2 !== '') return val2;
            }
        }
    }
    if (fallbacks) {
        for (var fi2 = 0; fi2 < fallbacks.length; fi2++) {
            var headers = Object.keys(row);
            var normalizedRowHeaders = headers.map(function(h) { return normalizeHeader(h); });
            var normalizedFallback = normalizeHeader(fallbacks[fi2]);
            for (var hi = 0; hi < normalizedRowHeaders.length; hi++) {
                if (normalizedRowHeaders[hi] === normalizedFallback) {
                    var val3 = row[headers[hi]];
                    if (val3 !== undefined && val3 !== null && val3 !== '') return val3;
                }
            }
        }
    }
    return '';
}

function formatExcelDate(value) {
    if (!value && value !== 0) return '';
    if (value instanceof Date && !isNaN(value)) {
        const y = value.getFullYear();
        const m = String(value.getMonth() + 1).padStart(2, '0');
        const d = String(value.getDate()).padStart(2, '0');
        return y + '-' + m + '-' + d;
    }
    if (typeof value === 'number' && value > 40000 && value < 60000) {
        const d = new Date((value - 25569) * 86400 * 1000);
        return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
    }
    let str = String(value).trim();
    if (str.includes('/')) {
        const parts = str.split('/');
        if (parts[0].length === 4) return str;
        if (parts.length === 3) return parts[2] + '-' + parts[1].padStart(2, '0') + '-' + parts[0].padStart(2, '0');
    }
    return str;
}
window.formatExcelDate = formatExcelDate;

function showLoading() { var el = document.getElementById('loading-overlay'); if (el) el.style.display = 'flex'; }
function hideLoading() { var el = document.getElementById('loading-overlay'); if (el) el.style.display = 'none'; }

window.showToast = function(message, type) {
    type = type || 'success';
    var container = document.getElementById('toast-container');
    if (!container) return;
    var toast = document.createElement('div');
    var icons = { success: '\u2705', error: '\u274C', info: '\u2139\uFE0F' };
    toast.className = 'toast toast-' + type;
    toast.innerHTML = (icons[type] || '') + ' ' + message;
    container.appendChild(toast);
    setTimeout(function() { toast.remove(); }, 4000);
};

onAuthStateChanged(auth, async function(user) {
    if (user) {
        const userDoc = await getDoc(doc(db, "users", user.uid));
        if (userDoc.exists()) {
            currentUser = { uid: user.uid, ...userDoc.data() };
            setupUI();
            setupRealtimeListeners();
        } else {
            signOut(auth);
        }
    } else {
        window.location.href = "index.html";
    }
});

function setupUI() {
    if (document.getElementById('user-display-name')) document.getElementById('user-display-name').textContent = currentUser.name;
    if (document.getElementById('user-display-role')) document.getElementById('user-display-role').textContent = currentUser.role;
    if (document.getElementById('user-display-empid')) document.getElementById('user-display-empid').textContent = currentUser.empId;

    isAdmin = currentUser.role === '\u0645\u0633\u0626\u0648\u0644 \u0627\u0644\u0645\u0646\u0638\u0648\u0645\u0629' || currentUser.empId === '34285';
    isReadOnly = currentUser.permission === 'readonly';

    if (isAdmin) {
        if (document.getElementById('admin-add-user-card')) document.getElementById('admin-add-user-card').style.display = 'block';
        if (document.getElementById('admin-users-list-card')) document.getElementById('admin-users-list-card').style.display = 'block';
        if (document.getElementById('tab-archive')) document.getElementById('tab-archive').style.display = 'block';
        if (document.getElementById('admin-backup-card')) document.getElementById('admin-backup-card').style.display = 'block';
    }
    if (isReadOnly) {
        var hideBtns = ['add-device-btn', 'return-device-btn', 'upload-excel-btn', 'clear-data-btn', 'add-store-btn', 'add-purchase-btn', 'clear-archive-btn'];
        hideBtns.forEach(function(id) { if (document.getElementById(id)) document.getElementById(id).style.display = 'none'; });
    }
}

var idleTimeout = 2 * 60 * 1000;
var idleTimer;

function resetIdleTimer() {
    clearTimeout(idleTimer);
    idleTimer = setTimeout(function() {
        showToast('\u062a\u0645 \u062a\u0633\u062c\u064a\u0644 \u0627\u0644\u062e\u0631\u0648\u062c \u062a\u0644\u0642\u0627\u0626\u064a\u0627\u064b \u0644\u0639\u062f\u0645 \u0627\u0644\u0646\u0634\u0627\u0637', 'info');
        setTimeout(function() { signOut(auth); }, 800);
    }, idleTimeout);
}
['mousemove', 'keydown', 'mousedown', 'touchstart', 'scroll', 'click', 'touchmove'].forEach(function(evt) {
    window.addEventListener(evt, resetIdleTimer, { passive: true });
});
resetIdleTimer();

function updateStats() {
    var devEl = document.getElementById('stat-devices');
    var storeEl = document.getElementById('stat-store');
    var purEl = document.getElementById('stat-purchases');
    var usrEl = document.getElementById('stat-users');
    if (devEl) devEl.textContent = devicesData.length;
    if (storeEl) storeEl.textContent = storeData.length;
    if (purEl) purEl.textContent = purchasesData.length;
    if (usrEl) usrEl.textContent = usersData.length;
}

function setupRealtimeListeners() {
    onSnapshot(collection(db, "devices"), function(snapshot) {
        devicesData = snapshot.docs.map(function(d) { return { firebaseId: d.id, ...d.data() }; }).sort(function(a, b) { return (a.createdAt || 0) - (b.createdAt || 0); });
        loadDevicesTable();
        updateStats();
    });
    onSnapshot(collection(db, "store"), function(snapshot) {
        storeData = snapshot.docs.map(function(d) { return { firebaseId: d.id, ...d.data() }; }).sort(function(a, b) { return (b.createdAt || 0) - (a.createdAt || 0); });
        loadStoreTable();
        updateStats();
    });
    onSnapshot(collection(db, "purchases"), function(snapshot) {
        purchasesData = snapshot.docs.map(function(d) { return { firebaseId: d.id, ...d.data() }; }).sort(function(a, b) { return (b.createdAt || 0) - (a.createdAt || 0); });
        loadPurchasesTable();
        updateStats();
    });
    onSnapshot(collection(db, "archive"), function(snapshot) {
        archiveData = snapshot.docs.map(function(d) { return { firebaseId: d.id, ...d.data() }; }).sort(function(a, b) { return (b.createdAt || 0) - (a.createdAt || 0); });
        loadArchiveTable();
    });
    onSnapshot(collection(db, "users"), function(snapshot) {
        usersData = snapshot.docs.map(function(d) { return { firebaseId: d.id, ...d.data() }; });
        renderUsersTable();
        updateStats();
    });
}

function getAuditNote(item) {
    if (item.lastModifiedBy && item.lastModifiedAt) {
        return '<br><span style="font-size:10px; color:#e74c3c; display:block; margin-top:4px;">\u062a\u0639\u062f\u064a\u0644: ' + item.lastModifiedBy + '<br>' + item.lastModifiedAt + '</span>';
    }
    return '';
}

function safeVal(v) { return (!v || v === 'undefined' || v === 'null') ? '' : v; }

window.loadDevicesTable = function() {
    var tableBody = document.getElementById('table-body');
    if (!tableBody) return;
    tableBody.innerHTML = "";
    devicesData.forEach(function(item, index) {
        var auditStr = '<span style="color: #7f8c8d; font-size: 12px; font-weight:bold;">\u0625\u0636\u0627\u0641\u0629: ' + (item.addedBy || '') + '</span>' + getAuditNote(item);
        var actions = isReadOnly ? '<span style="color:#7f8c8d;font-size:12px;">\u0642\u0631\u0627\u0621\u0629 \u0641\u0642\u0637</span>' : '<button class="action-icon-btn" onclick="openEditDevice(' + index + ')">\u270F\uFE0F</button><button class="action-icon-btn" onclick="deleteDevice(' + index + ')">\uD83D\uDDD1\uFE0F</button>';
        tableBody.innerHTML += '<tr><td>' + (index + 1) + '</td><td class="right-click-cell" oncontextmenu="returnRightClick(event, ' + index + ')">' + safeVal(item.serial) + '</td><td class="right-click-cell" oncontextmenu="returnRightClick(event, ' + index + ')">' + safeVal(item.deviceName) + '</td><td>' + safeVal(item.deviceType) + '</td><td>' + safeVal(item.id) + '</td><td>' + safeVal(item.location) + '</td><td>' + safeVal(item.date) + '</td><td class="right-click-cell" oncontextmenu="returnRightClick(event, ' + index + ')">' + safeVal(item.receiverName) + '</td><td>' + safeVal(item.empId) + '</td><td>' + safeVal(item.jobTitle) + '</td><td>' + safeVal(item.department) + '</td><td>' + safeVal(item.phone) + '</td><td>' + safeVal(item.plateNum) + '</td><td>' + safeVal(item.notes) + '</td><td>' + auditStr + '</td><td>' + actions + '</td></tr>';
    });
};

window.loadStoreTable = function() {
    var storeTableBody = document.getElementById('store-table-body');
    if (!storeTableBody) return;
    storeTableBody.innerHTML = "";
    storeData.forEach(function(item, index) {
        var auditStr = '<span style="color: #7f8c8d; font-size: 12px; font-weight:bold;">\u0625\u0636\u0627\u0641\u0629: ' + (item.addedBy || '') + '</span>' + getAuditNote(item);
        var actions = isReadOnly ? '<span style="color:#7f8c8d;font-size:12px;">\u0642\u0631\u0627\u0621\u0629 \u0641\u0642\u0637</span>' : '<button class="action-icon-btn" onclick="openEditStore(' + index + ')">\u270F\uFE0F</button><button class="action-icon-btn" onclick="deleteStore(' + index + ')">\uD83D\uDDD1\uFE0F</button>';
        storeTableBody.innerHTML += '<tr><td>' + (index + 1) + '</td><td class="right-click-cell" oncontextmenu="handoverRightClick(event, ' + index + ')">' + (item.serial || '') + '</td><td>' + (item.deviceName || '') + '</td><td class="right-click-cell" oncontextmenu="handoverRightClick(event, ' + index + ')">' + (item.deviceType || '') + '</td><td><span class="status-badge">' + (item.status || '') + '</span></td><td>' + (item.notes || '') + '</td><td>' + auditStr + '</td><td>' + actions + '</td></tr>';
    });
};

window.loadPurchasesTable = function() {
    var tableBody = document.getElementById('purchases-table-body');
    if (!tableBody) return;
    tableBody.innerHTML = "";
    purchasesData.forEach(function(item, index) {
        var auditStr = '<span style="color: #7f8c8d; font-size: 12px; font-weight:bold;">\u0625\u0636\u0627\u0641\u0629: ' + (item.addedBy || '') + '</span>' + getAuditNote(item);
        var actions = isReadOnly ? '<span style="color:#7f8c8d;font-size:12px;">\u0642\u0631\u0627\u0621\u0629 \u0641\u0642\u0637</span>' : '<button class="action-icon-btn" onclick="openEditPurchase(' + index + ')">\u270F\uFE0F</button><button class="action-icon-btn" onclick="deletePurchase(' + index + ')">\uD83D\uDDD1\uFE0F</button>';
        tableBody.innerHTML += '<tr><td>' + (index + 1) + '</td><td>' + (item.orderNum || '') + '</td><td>' + (item.method || '') + '</td><td>' + (item.deviceType || '') + '</td><td>' + (item.col1 || '') + '</td><td>' + (item.col2 || '') + '</td><td>' + (item.col3 || '') + '</td><td>' + auditStr + '</td><td>' + actions + '</td></tr>';
    });
};

window.loadArchiveTable = function() {
    var tableBody = document.getElementById('archive-table-body');
    if (!tableBody) return;
    tableBody.innerHTML = "";
    archiveData.forEach(function(item, index) {
        var details = item.source === "\u0627\u0644\u0645\u0634\u062a\u0631\u064a\u0627\u062a" ? '\u0627\u0644\u0637\u0644\u0628\u064a\u0629: ' + (item.data ? item.data.orderNum : '') : '\u0627\u0644\u062a\u0633\u0644\u0633\u0644: ' + (item.data ? item.data.serial : '') + ' | \u0627\u0644\u062c\u0647\u0627\u0632: ' + (item.data ? item.data.deviceName : '');
        var deleteBtn = isReadOnly ? '' : '<button class="action-icon-btn" onclick="deleteArchiveItem(' + index + ')">\uD83D\uDDD1\uFE0F</button>';
        tableBody.innerHTML += '<tr><td>' + (index + 1) + '</td><td><span style="background:#e74c3c;color:white;padding:3px 8px;border-radius:3px;font-size:12px;">' + (item.source || '') + '</span></td><td>' + details + '</td><td style="font-family:monospace;">' + (item.deletedAt || '') + '</td><td style="font-weight:bold; color:#2c3e50;">' + (item.deletedBy || '') + '</td><td>' + deleteBtn + '</td></tr>';
    });
};

window.renderUsersTable = function() {
    if (!isAdmin) return;
    var tbody = document.getElementById('users-table-body');
    if (!tbody) return;
    tbody.innerHTML = '';
    usersData.forEach(function(u) {
        var deleteBtn = (u.empId === currentUser.empId) ? '<span style="color:#7f8c8d;font-size:14px;">\u0623\u0646\u062a</span>' : '<button class="action-icon-btn" onclick="deleteUser(\'' + u.firebaseId + '\')">\uD83D\uDDD1\uFE0F</button>';
        tbody.innerHTML += '<tr><td>' + (u.name || '') + '</td><td>' + (u.empId || '') + '</td><td>' + (u.role || '') + '</td><td>' + deleteBtn + '</td></tr>';
    });
};

window.executeSearch = function(section) {
    var inputId = section === 'devices' ? 'devices-search-input' : section === 'store' ? 'store-search-input' : 'purchases-search-input';
    var queryStr = document.getElementById(inputId).value;
    if (!queryStr.trim()) { alert('\u0627\u0644\u0631\u062c\u0627\u0621 \u0643\u062a\u0627\u0628\u0629 \u0643\u0644\u0645\u0629 \u0644\u0644\u0628\u062d\u062b.'); return; }
    var cleanQuery = cleanArabic(queryStr);
    var terms = cleanQuery.split(/\s+\u0648\s+|\s+/).filter(function(t) { return t.trim() !== ''; });
    currentSearchType = section;
    currentSearchResults = [];
    var dataSource = section === 'devices' ? devicesData : section === 'store' ? storeData : purchasesData;
    dataSource.forEach(function(item) {
        var isMatch = terms.some(function(term) {
            var q = term.toLowerCase();
            if (section === 'devices') return (cleanArabic(item.serial || "").toLowerCase().includes(q) || cleanArabic(item.deviceName || "").toLowerCase().includes(q) || cleanArabic(item.deviceType || "").toLowerCase().includes(q) || cleanArabic(item.location || "").toLowerCase().includes(q) || cleanArabic(item.receiverName || "").toLowerCase() === q || cleanArabic(item.receiverName || "").toLowerCase().startsWith(q) || String(item.empId || "").toLowerCase() === q || String(item.empId || "").toLowerCase().startsWith(q) || cleanArabic(item.notes || "").toLowerCase().includes(q));
            else if (section === 'store') return (cleanArabic(item.serial || "").toLowerCase().includes(q) || cleanArabic(item.deviceName || "").toLowerCase().includes(q) || cleanArabic(item.deviceType || "").toLowerCase().includes(q) || cleanArabic(item.notes || "").toLowerCase().includes(q));
            else return ((item.orderNum || "").toLowerCase().includes(q) || cleanArabic(item.method || "").toLowerCase().includes(q) || cleanArabic(item.deviceType || "").toLowerCase().includes(q) || cleanArabic(item.notes || "").toLowerCase().includes(q));
        });
        if (isMatch) currentSearchResults.push(item);
    });
    renderSearchResults();
};

function renderSearchResults() {
    var container = document.getElementById('search-results-container');
    var modal = document.getElementById('search-results-modal');
    if (!container || !modal) return;
    if (currentSearchResults.length === 0) {
        container.innerHTML = '<p style="text-align:center;padding:30px;color:#7f8c8d;font-size:16px;">\u0644\u0645 \u064a\u062a\u0645 \u0627\u0644\u0639\u062b\u0648\u0631 \u0639\u0644\u0649 \u0646\u062a\u0627\u0626\u062c.</p>';
        modal.style.display = 'flex';
        return;
    }
    var html = '';
    if (currentSearchType === 'devices') {
        var deviceHeaders = [
            { key: 'serial', label: 'رقم التسلسل' },
            { key: 'deviceName', label: 'اسم الجهاز' },
            { key: 'deviceType', label: 'النوع' },
            { key: 'id', label: 'ID' },
            { key: 'location', label: 'مكان العمل' },
            { key: 'date', label: 'تاريخ التسليم' },
            { key: 'receiverName', label: 'المستلم' },
            { key: 'empId', label: 'رقم التوظيف' },
            { key: 'jobTitle', label: 'الصفة' },
            { key: 'department', label: 'الإدارة' },
            { key: 'phone', label: 'الهاتف' },
            { key: 'plateNum', label: 'اللوحة' },
            { key: 'notes', label: 'الملاحظات' }
        ];
        html = '<table class="search-devices-table" style="width:100%;text-align:right;"><thead><tr>';
        deviceHeaders.forEach(function(h) { html += '<th>' + h.label + '</th>'; });
        if (!isReadOnly) html += '<th>إجراء</th>';
        html += '</tr></thead><tbody>';
        currentSearchResults.forEach(function(item) {
            var origIdx = -1;
            for (var oi = 0; oi < devicesData.length; oi++) {
                if (devicesData[oi].firebaseId === item.firebaseId) { origIdx = oi; break; }
            }
            html += '<tr' + (origIdx !== -1 ? ' data-orig-idx="' + origIdx + '"' : '') + '>';
            deviceHeaders.forEach(function(h) { html += '<td>' + (item[h.key] || '') + '</td>'; });
            if (!isReadOnly) {
                var actionBtns = '';
                if (origIdx !== -1) {
                    actionBtns = '<button class="action-icon-btn" onclick="searchReturnDevice(' + origIdx + ')" title="استرجاع">🔄</button> <button class="action-icon-btn" onclick="openEditDevice(' + origIdx + ')" title="تعديل">✏️</button> <button class="action-icon-btn" onclick="searchDeleteDevice(' + origIdx + ')" title="حذف">🗑️</button>';
                }
                html += '<td style="white-space:nowrap;">' + actionBtns + '</td>';
            }
            html += '</tr>';
        });
    } else if (currentSearchType === 'store') {
        var storeHeaders = [
            { key: 'serial', label: 'رقم التسلسل' },
            { key: 'deviceName', label: 'اسم الجهاز' },
            { key: 'deviceType', label: 'النوع' },
            { key: 'status', label: 'الحالة' },
            { key: 'notes', label: 'الملاحظات' }
        ];
        html = '<table class="search-store-table" style="width:100%;text-align:right;"><thead><tr>';
        storeHeaders.forEach(function(h) { html += '<th>' + h.label + '</th>'; });
        html += '</tr></thead><tbody>';
        currentSearchResults.forEach(function(item) {
            html += '<tr>';
            storeHeaders.forEach(function(h) { html += '<td>' + (item[h.key] || '') + '</td>'; });
            html += '</tr>';
        });
    } else {
        var purchaseHeaders = [
            { key: 'orderNum', label: 'رقم الطلبية' },
            { key: 'method', label: 'طريقة الشراء' },
            { key: 'deviceType', label: 'نوع الجهاز' },
            { key: 'notes', label: 'الملاحظات' }
        ];
        html = '<table style="width:100%;text-align:right;"><thead><tr>';
        purchaseHeaders.forEach(function(h) { html += '<th>' + h.label + '</th>'; });
        html += '</tr></thead><tbody>';
        currentSearchResults.forEach(function(item) {
            html += '<tr>';
            purchaseHeaders.forEach(function(h) { html += '<td>' + (item[h.key] || '') + '</td>'; });
            html += '</tr>';
        });
    }
    html += '</tbody></table>';
    container.innerHTML = html;
    modal.style.display = 'flex';
}
window.renderSearchResults = renderSearchResults;

window.searchReturnDevice = function(index) {
    var searchModal = document.getElementById('search-results-modal');
    if (searchModal) searchModal.style.display = 'none';
    if (typeof executeReturnModal === 'function') {
        executeReturnModal(index);
    } else {
        var d = devicesData[index];
        if (!d) return;
        if (confirm('تأكيد استرجاع (' + (d.serial || d.deviceName) + ') للمخزن؟')) {
            showLoading();
            addDoc(collection(db, "store"), {
                serial: d.serial,
                deviceName: d.deviceName,
                deviceType: d.deviceType,
                status: "مسترجع للمخزن",
                notes: "مسترجع من: " + d.receiverName,
                addedBy: currentUser.name,
                createdAt: Date.now()
            }).then(function() {
                return deleteDoc(doc(db, "devices", d.firebaseId));
            }).then(function() {
                hideLoading();
                showToast('تم استرجاع الجهاز للمخزن');
            }).catch(function(err) {
                hideLoading();
                showToast('خطأ في الاسترجاع', 'error');
            });
        }
    }
};

window.searchDeleteDevice = function(index) {
    var searchModal = document.getElementById('search-results-modal');
    if (searchModal) searchModal.style.display = 'none';
    deleteDevice(index);
};

window.openSearchPrintSettings = function() {
    var container = document.getElementById('print-columns-container');
    if (!container) { return; }
    var allColumns = {
        devices: [
            { key: 'serial', label: 'رقم التسلسل' },
            { key: 'deviceName', label: 'اسم الجهاز' },
            { key: 'deviceType', label: 'النوع' },
            { key: 'id', label: 'ID' },
            { key: 'location', label: 'مكان العمل' },
            { key: 'date', label: 'تاريخ التسليم' },
            { key: 'receiverName', label: 'المستلم' },
            { key: 'empId', label: 'رقم التوظيف' },
            { key: 'jobTitle', label: 'الصفة' },
            { key: 'department', label: 'الإدارة' },
            { key: 'phone', label: 'الهاتف' },
            { key: 'plateNum', label: 'اللوحة' },
            { key: 'notes', label: 'الملاحظات' }
        ],
        store: [
            { key: 'serial', label: 'رقم التسلسل' },
            { key: 'deviceName', label: 'اسم الجهاز' },
            { key: 'deviceType', label: 'النوع' },
            { key: 'status', label: 'الحالة' },
            { key: 'notes', label: 'الملاحظات' }
        ],
        purchases: [
            { key: 'orderNum', label: 'رقم الطلبية' },
            { key: 'method', label: 'طريقة الشراء' },
            { key: 'deviceType', label: 'نوع الجهاز' },
            { key: 'notes', label: 'الملاحظات' }
        ]
    };
    var cols = allColumns[currentSearchType] || allColumns.devices;
    var html = '';
    cols.forEach(function(col) {
        html += '<label style="display:flex;align-items:center;gap:8px;cursor:pointer;padding:6px 10px;border:1px solid #e2e8f0;border-radius:6px;font-size:14px;"><input type="checkbox" data-key="' + col.key + '" data-header="' + col.label + '" checked> ' + col.label + '</label>';
    });
    container.innerHTML = html;
    var modal = document.getElementById('print-settings-modal');
    if (modal) modal.style.display = 'flex';
};

window.deleteDevice = async function(index) {
    if (isReadOnly) return;
    if (!confirm('\u0647\u0644 \u0623\u0646\u062a \u0645\u062a\u0623\u0643\u062f \u0645\u0646 \u0645\u0633\u062d \u0627\u0644\u062c\u0647\u0627\u0632 \u0644\u0644\u0623\u0631\u0634\u064a\u0641\u061f')) return;
    showLoading();
    var item = devicesData[index];
    var id = item.firebaseId;
    var itemCopy = Object.assign({}, item);
    delete itemCopy.firebaseId;
    await addDoc(collection(db, "archive"), { source: '\u0627\u0644\u0623\u062c\u0647\u0632\u0629 \u0627\u0644\u0645\u0633\u0644\u0645\u0629', data: itemCopy, deletedAt: new Date().toLocaleString('ar-EG'), deletedBy: currentUser.name, createdAt: Date.now() });
    await deleteDoc(doc(db, "devices", id));
    hideLoading();
    showToast('\u062a\u0645 \u0646\u0642\u0644 \u0627\u0644\u062c\u0647\u0627\u0632 \u0625\u0644\u0649 \u0627\u0644\u0623\u0631\u0634\u064a\u0641', 'info');
};

window.deleteStore = async function(index) {
    if (isReadOnly) return;
    if (!confirm('\u0645\u0633\u062d \u0644\u0644\u0645\u062e\u0632\u0646\u061f')) return;
    showLoading();
    var item = storeData[index];
    var id = item.firebaseId;
    var itemCopy = Object.assign({}, item);
    delete itemCopy.firebaseId;
    await addDoc(collection(db, "archive"), { source: '\u0627\u0644\u0645\u062e\u0632\u0646 \u0627\u0644\u0639\u0627\u0645', data: itemCopy, deletedAt: new Date().toLocaleString('ar-EG'), deletedBy: currentUser.name, createdAt: Date.now() });
    await deleteDoc(doc(db, "store", id));
    hideLoading();
    showToast('\u062a\u0645 \u0627\u0644\u062d\u0630\u0641 \u0645\u0646 \u0627\u0644\u0645\u062e\u0632\u0646', 'info');
};

window.deletePurchase = async function(index) {
    if (isReadOnly) return;
    if (!confirm('\u0645\u0633\u062d \u0637\u0644\u0628 \u0627\u0644\u0634\u0631\u0627\u0621\u061f')) return;
    showLoading();
    var item = purchasesData[index];
    var id = item.firebaseId;
    var itemCopy = Object.assign({}, item);
    delete itemCopy.firebaseId;
    await addDoc(collection(db, "archive"), { source: '\u0627\u0644\u0645\u0634\u062a\u0631\u064a\u0627\u062a', data: itemCopy, deletedAt: new Date().toLocaleString('ar-EG'), deletedBy: currentUser.name, createdAt: Date.now() });
    await deleteDoc(doc(db, "purchases", id));
    hideLoading();
    showToast('\u062a\u0645 \u062d\u0630\u0641 \u0637\u0644\u0628 \u0627\u0644\u0634\u0631\u0627\u0621', 'info');
};

window.deleteUser = async function(firebaseId) {
    if (confirm('\u0645\u062a\u0623\u0643\u062f \u0645\u0646 \u0645\u0633\u062d \u0627\u0644\u0645\u0633\u062a\u062e\u062f\u0645\u061f')) {
        showLoading();
        await deleteDoc(doc(db, "users", firebaseId));
        hideLoading();
    }
};

window.deleteArchiveItem = async function(index) {
    if (isReadOnly) return;
    if (!confirm('\u0647\u0644 \u0623\u0646\u062a \u0645\u062a\u0623\u0643\u062f \u0645\u0646 \u062d\u0630\u0641 \u0647\u0630\u0627 \u0627\u0644\u0633\u062c\u0644 \u0645\u0646 \u0627\u0644\u0623\u0631\u0634\u064a\u0641\u061f')) return;
    showLoading();
    try {
        await deleteDoc(doc(db, "archive", archiveData[index].firebaseId));
        showToast('\u062a\u0645 \u062d\u0630\u0641 \u0627\u0644\u0633\u062c\u0644 \u0645\u0646 \u0627\u0644\u0623\u0631\u0634\u064a\u0641', 'info');
    } catch (err) {
        showToast('\u062d\u062f\u062b \u062e\u0637\u0623 \u0623\u062b\u0646\u0627\u0621 \u0627\u0644\u062d\u0630\u0641', 'error');
    }
    hideLoading();
};

window.openEditDevice = function(index) {
    var d = devicesData[index];
    document.getElementById('edit-device-index').value = index;
    document.getElementById('edit-modal-serial').value = d.serial || '';
    document.getElementById('edit-modal-deviceName').value = d.deviceName || '';
    document.getElementById('edit-modal-deviceType').value = d.deviceType || '';
    document.getElementById('edit-modal-id').value = d.id || '';
    document.getElementById('edit-modal-location').value = d.location || '';
    document.getElementById('edit-modal-date').value = d.date || '';
    document.getElementById('edit-modal-receiverName').value = d.receiverName || '';
    document.getElementById('edit-modal-empId').value = d.empId || '';
    document.getElementById('edit-modal-jobTitle').value = d.jobTitle || '';
    document.getElementById('edit-modal-department').value = d.department || '';
    document.getElementById('edit-modal-phone').value = d.phone || '';
    document.getElementById('edit-modal-plateNum').value = d.plateNum || '';
    document.getElementById('edit-modal-notes').value = d.notes || '';
    document.getElementById('edit-device-modal').style.display = 'flex';
};

window.openEditStore = function(index) {
    var d = storeData[index];
    document.getElementById('edit-store-index').value = index;
    document.getElementById('edit-store-modal-serial').value = d.serial || '';
    document.getElementById('edit-store-modal-deviceName').value = d.deviceName || '';
    document.getElementById('edit-store-modal-deviceType').value = d.deviceType || '';
    document.getElementById('edit-store-modal-status').value = d.status || '';
    document.getElementById('edit-store-modal-notes').value = d.notes || "";
    document.getElementById('edit-store-modal').style.display = 'flex';
};

window.openEditPurchase = function(index) {
    var d = purchasesData[index];
    document.getElementById('edit-purchase-index').value = index;
    document.getElementById('edit-purchase-modal-orderNum').value = d.orderNum || '';
    document.getElementById('edit-purchase-modal-method').value = d.method || '';
    document.getElementById('edit-purchase-modal-deviceType').value = d.deviceType || '';
    document.getElementById('edit-purchase-modal-col1').value = d.col1 || "";
    document.getElementById('edit-purchase-modal-col2').value = d.col2 || "";
    document.getElementById('edit-purchase-modal-col3').value = d.col3 || "";
    document.getElementById('edit-purchase-modal').style.display = 'flex';
};

if (!isReadOnly) {
    window.returnRightClick = async function(e, index) {
        e.preventDefault();
        var d = devicesData[index];
        if (confirm('\u0627\u0633\u062a\u0631\u062c\u0627\u0639 (' + d.serial + ') \u0644\u0644\u0645\u062e\u0632\u0646\u061f')) {
            showLoading();
            await addDoc(collection(db, "store"), { serial: d.serial, deviceName: d.deviceName, deviceType: d.deviceType, status: "\u0645\u0633\u062a\u0631\u062c\u0639 \u0644\u0644\u0645\u062e\u0632\u0646", notes: "\u0645\u0633\u062a\u0631\u062c\u0639 \u0645\u0646: " + d.receiverName, addedBy: currentUser.name, createdAt: Date.now() });
            await deleteDoc(doc(db, "devices", d.firebaseId));
            hideLoading();
            showToast('\u062a\u0645 \u0627\u0633\u062a\u0631\u062c\u0627\u0639 \u0627\u0644\u062c\u0647\u0627\u0632 \u0644\u0644\u0645\u062e\u0632\u0646');
        }
    };
    window.handoverRightClick = function(e, index) {
        e.preventDefault();
        document.getElementById('handover-store-index').value = index;
        document.getElementById('handover-form').reset();
        document.getElementById('handover-modal').style.display = 'flex';
    };
    window.executeReturnModal = async function(index) {
        var d = devicesData[index];
        if (confirm('\u062a\u0623\u0643\u064a\u062f \u0627\u0644\u0627\u0633\u062a\u0631\u062c\u0627\u0639\u061f')) {
            showLoading();
            await addDoc(collection(db, "store"), { serial: d.serial, deviceName: d.deviceName, deviceType: d.deviceType, status: "\u0645\u0633\u062a\u0631\u062c\u0639 \u0644\u0644\u0645\u062e\u0632\u0646", notes: "\u0645\u0633\u062a\u0631\u062c\u0639 \u0645\u0646: " + d.receiverName, addedBy: currentUser.name, createdAt: Date.now() });
            await deleteDoc(doc(db, "devices", d.firebaseId));
            document.getElementById('return-modal').style.display = 'none';
            hideLoading();
            showToast('\u062a\u0645 \u0627\u0633\u062a\u0631\u062c\u0627\u0639 \u0627\u0644\u062c\u0647\u0627\u0632 \u0644\u0644\u0645\u062e\u0632\u0646');
        }
    };
}

if (document.getElementById('add-device-form')) {
    document.getElementById('add-device-form').addEventListener('submit', async function(e) {
        e.preventDefault();
        showLoading();
        await addDoc(collection(db, "devices"), { serial: cleanText(document.getElementById('modal-serial').value), deviceName: cleanText(document.getElementById('modal-deviceName').value), deviceType: cleanText(document.getElementById('modal-deviceType').value), id: cleanText(document.getElementById('modal-id').value), location: cleanText(document.getElementById('modal-location').value), date: document.getElementById('modal-date').value, receiverName: cleanText(document.getElementById('modal-receiverName').value), empId: cleanText(document.getElementById('modal-empId').value), jobTitle: cleanText(document.getElementById('modal-jobTitle').value), department: cleanText(document.getElementById('modal-department').value), phone: cleanText(document.getElementById('modal-phone').value), plateNum: cleanText(document.getElementById('modal-plateNum').value), notes: cleanText(document.getElementById('modal-notes').value), addedBy: currentUser.name, createdAt: Date.now() });
        this.reset();
        document.getElementById('device-modal').style.display = 'none';
        hideLoading();
        showToast('\u062a\u0645 \u0625\u0636\u0627\u0641\u0629 \u0627\u0644\u062c\u0647\u0627\u0632 \u0628\u0646\u062c\u0627\u062d');
    });
}
if (document.getElementById('add-store-form')) {
    document.getElementById('add-store-form').addEventListener('submit', async function(e) {
        e.preventDefault();
        showLoading();
        var finalStatus = document.getElementById('store-modal-status').value === '\u0623\u062e\u0631\u0649' ? cleanText(document.getElementById('store-modal-custom-status').value) : document.getElementById('store-modal-status').value;
        await addDoc(collection(db, "store"), { serial: cleanText(document.getElementById('store-modal-serial').value), deviceName: cleanText(document.getElementById('store-modal-deviceName').value), deviceType: cleanText(document.getElementById('store-modal-deviceType').value), status: finalStatus, notes: cleanText(document.getElementById('store-modal-notes').value), addedBy: currentUser.name, createdAt: Date.now() });
        this.reset();
        document.getElementById('store-modal').style.display = 'none';
        hideLoading();
        showToast('\u062a\u0645\u062a \u0627\u0644\u0625\u0636\u0627\u0641\u0629 \u0644\u0644\u0645\u062e\u0632\u0646');
    });
}
if (document.getElementById('add-purchase-form')) {
    document.getElementById('add-purchase-form').addEventListener('submit', async function(e) {
        e.preventDefault();
        showLoading();
        await addDoc(collection(db, "purchases"), { orderNum: cleanText(document.getElementById('purchase-modal-orderNum').value), method: cleanText(document.getElementById('purchase-modal-method').value), deviceType: cleanText(document.getElementById('purchase-modal-deviceType').value), col1: cleanText(document.getElementById('purchase-modal-col1').value), col2: cleanText(document.getElementById('purchase-modal-col2').value), col3: cleanText(document.getElementById('purchase-modal-col3').value), addedBy: currentUser.name, createdAt: Date.now() });
        this.reset();
        document.getElementById('purchase-modal').style.display = 'none';
        hideLoading();
        showToast('\u062a\u0645 \u0625\u0636\u0627\u0641\u0629 \u0637\u0644\u0628 \u0627\u0644\u0634\u0631\u0627\u0621');
    });
}
if (document.getElementById('edit-device-form')) {
    document.getElementById('edit-device-form').addEventListener('submit', async function(e) {
        e.preventDefault();
        showLoading();
        var item = devicesData[document.getElementById('edit-device-index').value];
        await updateDoc(doc(db, "devices", item.firebaseId), { serial: cleanText(document.getElementById('edit-modal-serial').value), deviceName: cleanText(document.getElementById('edit-modal-deviceName').value), deviceType: cleanText(document.getElementById('edit-modal-deviceType').value), id: cleanText(document.getElementById('edit-modal-id').value), location: cleanText(document.getElementById('edit-modal-location').value), date: document.getElementById('edit-modal-date').value, receiverName: cleanText(document.getElementById('edit-modal-receiverName').value), empId: cleanText(document.getElementById('edit-modal-empId').value), jobTitle: cleanText(document.getElementById('edit-modal-jobTitle').value), department: cleanText(document.getElementById('edit-modal-department').value), phone: cleanText(document.getElementById('edit-modal-phone').value), plateNum: cleanText(document.getElementById('edit-modal-plateNum').value), notes: cleanText(document.getElementById('edit-modal-notes').value), lastModifiedBy: currentUser.name, lastModifiedAt: new Date().toLocaleString('ar-EG') });
        document.getElementById('edit-device-modal').style.display = 'none';
        hideLoading();
        showToast('\u062a\u0645 \u062a\u0639\u062f\u064a\u0644 \u0627\u0644\u062c\u0647\u0627\u0632 \u0628\u0646\u062c\u0627\u062d');
    });
}
if (document.getElementById('edit-store-form')) {
    document.getElementById('edit-store-form').addEventListener('submit', async function(e) {
        e.preventDefault();
        showLoading();
        var item = storeData[document.getElementById('edit-store-index').value];
        await updateDoc(doc(db, "store", item.firebaseId), { serial: cleanText(document.getElementById('edit-store-modal-serial').value), deviceName: cleanText(document.getElementById('edit-store-modal-deviceName').value), deviceType: cleanText(document.getElementById('edit-store-modal-deviceType').value), status: cleanText(document.getElementById('edit-store-modal-status').value), notes: cleanText(document.getElementById('edit-store-modal-notes').value), lastModifiedBy: currentUser.name, lastModifiedAt: new Date().toLocaleString('ar-EG') });
        document.getElementById('edit-store-modal').style.display = 'none';
        hideLoading();
        showToast('\u062a\u0645 \u062a\u0639\u062f\u064a\u0644 \u0627\u0644\u0645\u062e\u0632\u0646');
    });
}
if (document.getElementById('edit-purchase-form')) {
    document.getElementById('edit-purchase-form').addEventListener('submit', async function(e) {
        e.preventDefault();
        showLoading();
        var item = purchasesData[document.getElementById('edit-purchase-index').value];
        await updateDoc(doc(db, "purchases", item.firebaseId), { orderNum: cleanText(document.getElementById('edit-purchase-modal-orderNum').value), method: cleanText(document.getElementById('edit-purchase-modal-method').value), deviceType: cleanText(document.getElementById('edit-purchase-modal-deviceType').value), col1: cleanText(document.getElementById('edit-purchase-modal-col1').value), col2: cleanText(document.getElementById('edit-purchase-modal-col2').value), col3: cleanText(document.getElementById('edit-purchase-modal-col3').value), lastModifiedBy: currentUser.name, lastModifiedAt: new Date().toLocaleString('ar-EG') });
        document.getElementById('edit-purchase-modal').style.display = 'none';
        hideLoading();
        showToast('\u062a\u0645 \u062a\u0639\u062f\u064a\u0644 \u0637\u0644\u0628 \u0627\u0644\u0634\u0631\u0627\u0621');
    });
}

if (document.getElementById('handover-form')) {
    document.getElementById('handover-form').addEventListener('submit', async function(e) {
        e.preventDefault();
        showLoading();
        var s = storeData[document.getElementById('handover-store-index').value];
        await addDoc(collection(db, "devices"), { serial: s.serial, deviceName: s.deviceName, deviceType: s.deviceType, id: cleanText(document.getElementById('handover-id').value), location: cleanText(document.getElementById('handover-location').value), date: document.getElementById('handover-date').value, receiverName: cleanText(document.getElementById('handover-receiverName').value), empId: cleanText(document.getElementById('handover-empId').value), jobTitle: cleanText(document.getElementById('handover-jobTitle').value), department: cleanText(document.getElementById('handover-department').value), phone: cleanText(document.getElementById('handover-phone').value), plateNum: cleanText(document.getElementById('handover-plateNum').value), notes: cleanText(document.getElementById('handover-notes').value) || "\u0645\u0646 \u0627\u0644\u0645\u062e\u0632\u0646", addedBy: currentUser.name, createdAt: Date.now() });
        await deleteDoc(doc(db, "store", s.firebaseId));
        document.getElementById('handover-modal').style.display = 'none';
        hideLoading();
        showToast('\u062a\u0645 \u062a\u0633\u0644\u064a\u0645 \u0627\u0644\u062c\u0647\u0627\u0632 \u0628\u0646\u062c\u0627\u062d');
    });
}

var tabDevices = document.getElementById('tab-devices');
var tabStore = document.getElementById('tab-store');
var tabPurchases = document.getElementById('tab-purchases');
var tabArchive = document.getElementById('tab-archive');
var tabProfile = document.getElementById('tab-profile');
var sections = { devices: document.getElementById('devices-section'), store: document.getElementById('store-section'), purchases: document.getElementById('purchases-section'), archive: document.getElementById('archive-section'), profile: document.getElementById('profile-section') };

function switchTab(activeTab, activeSection) {
    [tabDevices, tabStore, tabPurchases, tabArchive, tabProfile].forEach(function(t) { if (t) t.classList.remove('active'); });
    Object.values(sections).forEach(function(s) { if (s) s.classList.remove('active'); });
    activeTab.classList.add('active');
    activeSection.classList.add('active');
}
if (tabDevices) tabDevices.addEventListener('click', function() { switchTab(tabDevices, sections.devices); });
if (tabStore) tabStore.addEventListener('click', function() { switchTab(tabStore, sections.store); });
if (tabPurchases) tabPurchases.addEventListener('click', function() { switchTab(tabPurchases, sections.purchases); });
if (tabArchive) tabArchive.addEventListener('click', function() { switchTab(tabArchive, sections.archive); });
if (tabProfile) tabProfile.addEventListener('click', function() { switchTab(tabProfile, sections.profile); });

var deviceModal = document.getElementById('device-modal');
var storeModal = document.getElementById('store-modal');
var purchaseModal = document.getElementById('purchase-modal');

if (!isReadOnly) {
    if (document.getElementById('add-device-btn')) document.getElementById('add-device-btn').addEventListener('click', function() { deviceModal.style.display = 'flex'; });
    if (document.getElementById('add-store-btn')) document.getElementById('add-store-btn').addEventListener('click', function() { storeModal.style.display = 'flex'; });
    if (document.getElementById('add-purchase-btn')) document.getElementById('add-purchase-btn').addEventListener('click', function() { purchaseModal.style.display = 'flex'; });
}

var closeBtns = ['close-device-modal', 'close-store-modal', 'close-purchase-modal', 'close-handover-modal', 'close-edit-device-modal', 'close-edit-store-modal', 'close-edit-purchase-modal', 'close-search-results', 'close-print-settings', 'close-return-modal'];
closeBtns.forEach(function(id) {
    if (document.getElementById(id)) document.getElementById(id).addEventListener('click', function() { this.closest('.modal').style.display = 'none'; });
});

var returnModal = document.getElementById('return-modal');
var returnSearchInput = document.getElementById('return-search-input');
var returnDeviceResults = document.getElementById('return-device-results');

if (!isReadOnly && document.getElementById('return-device-btn')) {
    document.getElementById('return-device-btn').addEventListener('click', function() {
        if (devicesData.length === 0) { alert("\u0644\u0627 \u062a\u0648\u062c\u062f \u0623\u062c\u0647\u0632\u0629 \u0645\u0633\u0644\u0651\u0645\u0629 \u062d\u0627\u0644\u064a\u0627\u064b!"); return; }
        if (returnSearchInput) returnSearchInput.value = '';
        populateReturnResults('');
        if (returnModal) returnModal.style.display = 'flex';
    });
}

function populateReturnResults(filterText) {
    var q = cleanArabic(filterText).toLowerCase().trim();
    if (!q) {
        if (returnDeviceResults) returnDeviceResults.innerHTML = '<p style="text-align:center; padding:20px; color:#7f8c8d;">\u0627\u0643\u062a\u0628 \u0644\u0644\u0628\u062d\u062b...</p>';
        return;
    }
    var html = '<table class="search-devices-table" style="width:100%; text-align:right;"><thead><tr><th>\u0631\u0642\u0645 \u0627\u0644\u062a\u0633\u0644\u0633\u0644</th><th>\u0627\u0633\u0645 \u0627\u0644\u062c\u0647\u0627\u0632</th><th>\u0627\u0633\u0645 \u0627\u0644\u0645\u0633\u062a\u0644\u0645</th><th>\u0631\u0642\u0645 \u0627\u0644\u062a\u0648\u0638\u064a\u0641</th><th>\u0625\u062c\u0631\u0627\u0621</th></tr></thead><tbody>';
    var matchCount = 0;
    devicesData.forEach(function(item, index) {
        if ((item.serial && cleanArabic(item.serial).toLowerCase().includes(q)) || (item.deviceName && cleanArabic(item.deviceName).toLowerCase().includes(q)) || (item.receiverName && cleanArabic(item.receiverName).toLowerCase().includes(q)) || (item.empId && item.empId.toLowerCase().includes(q))) {
            html += '<tr><td>' + (item.serial || '') + '</td><td>' + (item.deviceName || '') + '</td><td style="font-weight:bold; color:#2980b9;">' + (item.receiverName || '') + '</td><td>' + (item.empId || '') + '</td><td><button class="return-btn" onclick="executeReturnModal(' + index + ')" style="padding: 6px 12px; font-size:12px;">\uD83D\uDD04 \u0627\u0633\u062a\u0631\u062c\u0627\u0639</button></td></tr>';
            matchCount++;
        }
    });
    html += '</tbody></table>';
    if (returnDeviceResults) returnDeviceResults.innerHTML = matchCount === 0 ? '<p style="text-align:center; color:#e74c3c;">\u0644\u0627 \u062a\u0648\u062c\u062f \u0646\u062a\u0627\u0626\u062c.</p>' : html;
}

if (returnSearchInput) returnSearchInput.addEventListener('input', function(e) { populateReturnResults(e.target.value); });

if (document.getElementById('excel-upload')) {
    document.getElementById('excel-upload').addEventListener('change', function(e) {
        var file = e.target.files[0];
        if (!file) return;
        showLoading();
        var reader = new FileReader();
        reader.onload = function(event) {
            setTimeout(async function() {
                try {
                    var data = new Uint8Array(event.target.result);
                    var workbook = XLSX.read(data, { type: 'array', cellDates: true });
                    var importedCount = 0;
                    var skippedCount = 0;
                    var errorCount = 0;

                    for (var si = 0; si < workbook.SheetNames.length; si++) {
                        var sheetName = workbook.SheetNames[si];
                        var json = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" });
                        if (!json.length) continue;
                        var mapped = mapExcelColumns(json);
                        var colMap = mapped.mapping;
                        var batch = writeBatch(db);
                        var opCount = 0;
                        var batchPromises = [];
                        var timestamp = Date.now();

                        for (var i = 0; i < json.length; i++) {
                            var row = json[i];
                            var serial = cleanText(getCellValue(row, colMap, 'serial'));
                            var deviceName = cleanText(getCellValue(row, colMap, 'deviceName'));
                            if (!serial && !deviceName) { skippedCount++; continue; }

                            try {
                                var newDocRef = doc(collection(db, "devices"));
                                batch.set(newDocRef, {
                                    serial: serial,
                                    deviceName: deviceName,
                                    deviceType: cleanText(getCellValue(row, colMap, 'deviceType')),
                                    id: cleanText(getCellValue(row, colMap, 'id')),
                                    location: cleanText(getCellValue(row, colMap, 'location')),
                                    date: formatExcelDate(getCellValue(row, colMap, 'date')),
                                    receiverName: cleanText(getCellValue(row, colMap, 'receiverName')),
                                    empId: cleanText(getCellValue(row, colMap, 'empId') != null ? String(getCellValue(row, colMap, 'empId')) : ''),
                                    jobTitle: cleanText(getCellValue(row, colMap, 'jobTitle')),
                                    department: cleanText(getCellValue(row, colMap, 'department')),
                                    phone: cleanText(getCellValue(row, colMap, 'phone') != null ? String(getCellValue(row, colMap, 'phone')) : ''),
                                    plateNum: cleanText(getCellValue(row, colMap, 'plateNum')),
                                    notes: cleanText(getCellValue(row, colMap, 'notes')) || "\u0644\u0627 \u062a\u0648\u062c\u062f \u0645\u0644\u0627\u062d\u0638\u0627\u062a",
                                    addedBy: currentUser.name,
                                    createdAt: timestamp
                                });
                                importedCount++;
                                opCount++;
                                timestamp++;
                            } catch (rowErr) {
                                errorCount++;
                            }

                            if (opCount >= 490) {
                                batchPromises.push(batch.commit());
                                batch = writeBatch(db);
                                opCount = 0;
                            }
                        }
                        if (opCount > 0) batchPromises.push(batch.commit());
                        await Promise.all(batchPromises);
                    }

                    var msg = '\u062a\u0645 \u0625\u0636\u0627\u0641\u0629 ' + importedCount + ' \u062c\u0647\u0627\u0632 \u0628\u0646\u062c\u0627\u062d';
                    if (skippedCount > 0) msg += ' | \u062a\u0645 \u062a\u062e\u0637\u064a ' + skippedCount + ' \u0635\u0641 \u0641\u0627\u0631\u063a';
                    if (errorCount > 0) msg += ' | \u0641\u0634\u0644 \u0631\u0641\u0639 ' + errorCount + ' \u062c\u0647\u0627\u0632';
                    hideLoading();
                    showToast(msg);
                    document.getElementById('excel-upload').value = '';
                } catch (error) {
                    hideLoading();
                    showToast('\u062e\u0637\u0623 \u0641\u064a \u0642\u0631\u0627\u0621\u0629 \u0627\u0644\u0645\u0644\u0641: ' + error.message, 'error');
                }
            }, 500);
        };
        reader.readAsArrayBuffer(file);
    });
}

if (document.getElementById('add-user-form')) {
    document.getElementById('add-user-form').addEventListener('submit', async function(e) {
        e.preventDefault();
        showLoading();
        var empId = cleanText(document.getElementById('new-user-empid').value);
        var pass = document.getElementById('new-user-password').value;
        try {
            var userCred = await createUserWithEmailAndPassword(secondaryAuth, empId + '@radio.local', pass);
            await setDoc(doc(db, "users", userCred.user.uid), { name: cleanText(document.getElementById('new-user-name').value), empId: empId, role: cleanText(document.getElementById('new-user-role').value), permission: document.getElementById('new-user-permission').value });
            secondaryAuth.signOut();
            showToast('\u062a\u0645 \u0625\u0636\u0627\u0641\u0629 \u0627\u0644\u0645\u0633\u062a\u062e\u062f\u0645 \u0628\u0646\u062c\u0627\u062d');
            this.reset();
        } catch (err) {
            showToast('\u062d\u062f\u062b \u062e\u0637\u0623: \u0642\u062f \u064a\u0643\u0648\u0646 \u0631\u0642\u0645 \u0627\u0644\u062a\u0648\u0638\u064a\u0641 \u0645\u0633\u062a\u062e\u062f\u0645 \u0628\u0627\u0644\u0641\u0639\u0644', 'error');
        }
        hideLoading();
    });
}

if (document.getElementById('change-password-form')) {
    document.getElementById('change-password-form').addEventListener('submit', async function(e) {
        e.preventDefault();
        var currentPass = document.getElementById('current-password').value;
        var newPass = document.getElementById('new-password').value;
        var confirmPass = document.getElementById('confirm-password').value;
        if (newPass !== confirmPass) { showToast('\u0643\u0644\u0645\u0629 \u0627\u0644\u0633\u0631 \u0627\u0644\u062c\u062f\u064a\u062f\u0629 \u0644\u0627 \u062a\u062a\u0637\u0627\u0628\u0642 \u0645\u0639 \u0627\u0644\u062a\u0623\u0643\u064a\u062f', 'error'); return; }
        if (newPass.length < 6) { showToast('\u0643\u0644\u0645\u0629 \u0627\u0644\u0633\u0631 \u062c\u062f\u064a\u062f\u0629 6 \u0623\u062d\u0631\u0641 \u0639\u0644\u0649 \u0627\u0644\u0623\u0642\u0644', 'error'); return; }
        if (!currentPass) { showToast('\u0623\u062f\u062e\u0644 \u0643\u0644\u0645\u0629 \u0627\u0644\u0633\u0631 \u0627\u0644\u062d\u0627\u0644\u064a\u0629', 'error'); return; }
        showLoading();
        try {
            var email = currentUser.empId + '@radio.local';
            var credential = EmailAuthProvider.credential(email, currentPass);
            await reauthenticateWithCredential(auth.currentUser, credential);
            await updatePassword(auth.currentUser, newPass);
            showToast('\u062a\u0645 \u062a\u063a\u064a\u064a\u0631 \u0643\u0644\u0645\u0629 \u0627\u0644\u0633\u0631 \u0628\u0646\u062c\u0627\u062d');
            this.reset();
        } catch (err) {
            var msg = '\u062e\u0637\u0623 \u0641\u064a \u062a\u063a\u064a\u064a\u0631 \u0643\u0644\u0645\u0629 \u0627\u0644\u0633\u0631';
            if (err.code === 'auth/wrong-password' || err.code === 'auth/invalid-credential') msg = '\u0643\u0644\u0645\u0629 \u0627\u0644\u0633\u0631 \u0627\u0644\u062d\u0627\u0644\u064a\u0629 \u063a\u064a\u0631 \u0635\u062d\u064a\u062d\u0629';
            showToast(msg, 'error');
        }
        hideLoading();
    });
}

if (document.getElementById('logout-btn')) {
    document.getElementById('logout-btn').addEventListener('click', function() { signOut(auth); });
}

if (document.getElementById('clear-data-btn')) {
    document.getElementById('clear-data-btn').addEventListener('click', async function() {
        if (isReadOnly) return showToast('\u0639\u0630\u0631\u0627\u064b\u00a0\u0644\u0627 \u062a\u0645\u0644\u0643 \u0635\u0644\u0627\u062d\u064a\u0629 \u062d\u0630\u0641 \u0627\u0644\u0628\u064a\u0627\u0646\u0627\u062a', 'error');
        if (!devicesData || devicesData.length === 0) return alert('\u0644\u0627 \u062a\u0648\u062c\u062f \u0628\u064a\u0627\u0646\u0627\u062a \u0623\u062c\u0647\u0632\u0629 \u0644\u0645\u0633\u062d\u0647\u0627 \u062d\u0627\u0644\u064a\u0627\u064b.');
        if (!confirm('\u26a0\ufe0f \u062a\u0646\u0628\u064a\u0647: \u0633\u064a\u062a\u0645 \u0646\u0642\u0644 \u062c\u0645\u064a\u0639 \u0627\u0644\u0623\u062c\u0647\u0632\u0629 \u0625\u0644\u0649 \u0627\u0644\u0623\u0631\u0634\u064a\u0641. \u0647\u0644 \u0623\u0646\u062a \u0645\u062a\u0623\u0643\u062f\u061f')) return;

        showLoading();
        try {
            var batch = writeBatch(db);
            var operationCount = 0;
            var batchesPromises = [];
            var timestamp = Date.now();
            var deletedAtStr = new Date().toLocaleString('ar-EG');

            for (var idx = 0; idx < devicesData.length; idx++) {
                var item = devicesData[idx];
                var id = item.firebaseId;
                var itemCopy = Object.assign({}, item);
                delete itemCopy.firebaseId;

                var newArchiveRef = doc(collection(db, "archive"));
                batch.set(newArchiveRef, { source: '\u0627\u0644\u0623\u062c\u0647\u0632\u0629 \u0627\u0644\u0645\u0633\u0644\u0645\u0629', data: itemCopy, deletedAt: deletedAtStr, deletedBy: currentUser.name, createdAt: timestamp });
                operationCount++;

                if (id) {
                    batch.delete(doc(db, "devices", id));
                    operationCount++;
                }

                if (operationCount >= 490) {
                    batchesPromises.push(batch.commit());
                    batch = writeBatch(db);
                    operationCount = 0;
                }
            }
            if (operationCount > 0) batchesPromises.push(batch.commit());

            await Promise.all(batchesPromises);
            showToast('\u062a\u0645 \u0646\u0642\u0644 \u062c\u0645\u064a\u0639 \u0627\u0644\u0623\u062c\u0647\u0632\u0629 \u0625\u0644\u0649 \u0627\u0644\u0623\u0631\u0634\u064a\u0641 \u0648\u0645\u0633\u062d \u0627\u0644\u0642\u0627\u0626\u0645\u0629 \u0628\u0646\u062c\u0627\u062d');
        } catch (error) {
            console.error("\u062e\u0637\u0623 \u0623\u062b\u0646\u0627\u0621 \u0645\u0633\u062d \u0627\u0644\u0628\u064a\u0627\u0646\u0627\u062a: ", error);
            showToast('\u062d\u062f\u062b \u062e\u0637\u0623 \u0623\u062b\u0646\u0627\u0621 \u0645\u062d\u0627\u0648\u0644\u0629 \u0645\u0633\u062d \u0627\u0644\u0628\u064a\u0627\u0646\u0627\u062a', 'error');
        } finally {
            hideLoading();
        }
    });
}

if (document.getElementById('clear-archive-btn')) {
    document.getElementById('clear-archive-btn').addEventListener('click', async function() {
        if (isReadOnly) return showToast('\u0639\u0630\u0631\u0627\u064b\u00a0\u0644\u0627 \u062a\u0645\u0644\u0643 \u0635\u0644\u0627\u062d\u064a\u0629 \u062d\u0630\u0641', 'error');
        if (!archiveData || archiveData.length === 0) return alert('\u0644\u0627 \u062a\u0648\u062c\u062f \u0633\u062c\u0644\u0627\u062a \u0641\u064a \u0627\u0644\u0623\u0631\u0634\u064a\u0641 \u0644\u0645\u0633\u062d\u0647\u0627.');
        if (!confirm('\u26a0\ufe0f \u062a\u0646\u0628\u064a\u0647: \u0633\u064a\u062a\u0645 \u062d\u0630\u0641 \u062c\u0645\u064a\u0639 \u0633\u062c\u0644\u0627\u062a \u0627\u0644\u0623\u0631\u0634\u064a\u0641 (' + archiveData.length + ' \u0633\u062c\u0644). \u0644\u0627 \u064a\u0645\u0646 \u0627\u0644\u062a\u0631\u0627\u062c\u0639. \u0647\u0644 \u0623\u0646\u062a \u0645\u062a\u0623\u0643\u062f\u061f')) return;

        showLoading();
        try {
            var batch = writeBatch(db);
            var opCount = 0;
            var batches = [];

            for (var i = 0; i < archiveData.length; i++) {
                batch.delete(doc(db, "archive", archiveData[i].firebaseId));
                opCount++;
                if (opCount >= 490) {
                    batches.push(batch.commit());
                    batch = writeBatch(db);
                    opCount = 0;
                }
            }
            if (opCount > 0) batches.push(batch.commit());

            await Promise.all(batches);
            showToast('\u062a\u0645 \u0645\u0633\u062d \u0627\u0644\u0623\u0631\u0634\u064a\u0641 \u0628\u0627\u0644\u0643\u0627\u0645\u0644');
        } catch (err) {
            showToast('\u062d\u062f\u062b \u062e\u0637\u0623 \u0623\u062b\u0646\u0627\u0621 \u0645\u0633\u062d \u0627\u0644\u0623\u0631\u0634\u064a\u0641', 'error');
        }
        hideLoading();
    });
}

var adminBackupBtn = document.getElementById('admin-system-backup-btn');
if (adminBackupBtn) {
    adminBackupBtn.addEventListener('click', function() {
        showLoading();
        setTimeout(function() {
            var workbook = XLSX.utils.book_new();
            var devicesRows = devicesData.map(function(d) { return { '\u0631\u0642\u0645 \u0627\u0644\u062a\u0633\u0644\u0633\u0644': d.serial, '\u0627\u0633\u0645 \u0627\u0644\u062c\u0647\u0627\u0632': d.deviceName, '\u0646\u0648\u0639 \u0627\u0644\u062c\u0647\u0627\u0632': d.deviceType, ID: d.id, '\u0645\u0643\u0627\u0646 \u0639\u0645\u0644 \u0627\u0644\u062c\u0647\u0627\u0632': d.location, '\u062a\u0627\u0631\u064a\u062e \u0627\u0644\u062a\u0633\u0644\u064a\u0645': d.date, '\u0627\u0644\u0645\u0633\u062a\u0644\u0645': d.receiverName, '\u0631\u0642\u0645 \u0627\u0644\u062a\u0648\u0638\u064a\u0641': d.empId, '\u0627\u0644\u0635\u0641\u0629': d.jobTitle, '\u0627\u0644\u0625\u062f\u0627\u0631\u0629': d.department, '\u0627\u0644\u0647\u0627\u062a\u0641': d.phone, '\u0627\u0644\u0644\u0648\u062d\u0629': d.plateNum, '\u0627\u0644\u0645\u0644\u0627\u062d\u0638\u0627\u062a': d.notes }; });
            if (devicesRows.length) XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(devicesRows), "\u0627\u0644\u0623\u062c\u0647\u0632\u0629 \u0627\u0644\u0645\u0633\u0644\u0645\u0629");
            var storeRows = storeData.map(function(s) { return { '\u0631\u0642\u0645 \u0627\u0644\u062a\u0633\u0644\u0633\u0644': s.serial, '\u0627\u0633\u0645 \u0627\u0644\u062c\u0647\u0627\u0632': s.deviceName, '\u0646\u0648\u0639 \u0627\u0644\u062c\u0647\u0627\u0632': s.deviceType, '\u0627\u0644\u062d\u0627\u0644\u0629': s.status, '\u0627\u0644\u0645\u0644\u0627\u062d\u0638\u0627\u062a': s.notes }; });
            if (storeRows.length) XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(storeRows), "\u0627\u0644\u0645\u062e\u0632\u0646 \u0627\u0644\u0639\u0627\u0645");
            var purchaseRows = purchasesData.map(function(p) { return { '\u0631\u0642\u0645 \u0627\u0644\u0637\u0644\u0628\u064a\u0629': p.orderNum, '\u0637\u0631\u064a\u0642\u0629 \u0627\u0644\u0634\u0631\u0627\u0621': p.method, '\u0646\u0648\u0639 \u0627\u0644\u062c\u0647\u0627\u0632': p.deviceType, '\u0625\u0636\u0627\u0641\u064a 1': p.col1, '\u0625\u0636\u0627\u0641\u064a 2': p.col2, '\u0625\u0636\u0627\u0641\u064a 3': p.col3 }; });
            if (purchaseRows.length) XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(purchaseRows), "\u0627\u0644\u0645\u0634\u062a\u0631\u064a\u0627\u062a");
            var archiveRows = archiveData.map(function(a) { return { '\u0627\u0644\u0645\u0635\u062f\u0631': a.source, '\u0627\u0644\u062a\u0641\u0635\u064a\u0644': a.data ? JSON.stringify(a.data) : '', '\u062a\u0627\u0631\u064a\u062e \u0627\u0644\u062d\u0630\u0641': a.deletedAt, '\u0627\u0644\u0645\u0633\u0626\u0648\u0644': a.deletedBy }; });
            if (archiveRows.length) XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(archiveRows), "\u0633\u062c\u0644 \u0627\u0644\u0623\u0631\u0634\u064a\u0641 \u0648\u0627\u0644\u0645\u062d\u0630\u0648\u0641\u0627\u062a");
            XLSX.writeFile(workbook, '\u0627\u0644\u0646\u0633\u062e\u0629_\u0627\u0644\u0627\u062d\u062a\u064a\u0627\u0637\u064a\u0629_' + new Date().toISOString().split('T')[0] + '.xlsx');
            hideLoading();
            showToast('\u062a\u0645 \u062a\u062d\u0645\u064a\u0644 \u0627\u0644\u0646\u0633\u062e\u0629 \u0627\u0644\u0627\u062d\u062a\u064a\u0627\u0637\u064a\u0629 \u0628\u0646\u062c\u0627\u062d');
        }, 1000);
    });
}

var printModal = document.getElementById('print-settings-modal');
var confirmPrintBtn = document.getElementById('confirm-print-btn');
var exportExcelBtn = document.getElementById('export-excel-btn');
var printColsContainer = document.getElementById('print-columns-container');
var currentPrintColumns = [];

function openMainReport(section) {
    currentSearchResults = section === 'devices' ? devicesData : section === 'store' ? storeData : purchasesData;
    currentSearchType = section;
    if (currentSearchResults.length === 0) { alert('\u0644\u0627 \u062a\u0648\u062c\u062f \u0628\u064a\u0627\u0646\u0627\u062a \u0644\u0637\u0628\u0627\u0639\u062a\u0647\u0627!'); return; }
    openSearchPrintSettings();
}
window.openMainReport = openMainReport;

if (confirmPrintBtn) {
    confirmPrintBtn.addEventListener('click', function() {
        var checked = document.querySelectorAll('#print-columns-container input[type=checkbox]:checked');
        if (checked.length === 0) { alert('\u0627\u062e\u062a\u0631 \u0639\u0645\u0648\u062f\u0627\u064b \u0648\u0627\u062d\u062f\u0627\u064b \u0639\u0644\u0649 \u0627\u0644\u0623\u0642\u0644!'); return; }
        var headers = [];
        var keys = [];
        checked.forEach(function(cb) { headers.push(cb.dataset.header); keys.push(cb.dataset.key); });
        var html = '<table style="width:100%;border-collapse:collapse;text-align:right;"><thead><tr>';
        headers.forEach(function(h) { html += '<th style="border:1px solid #333;padding:8px;background:#f0f0f0;">' + h + '</th>'; });
        html += '</tr></thead><tbody>';
        currentSearchResults.forEach(function(item) {
            html += '<tr>';
            keys.forEach(function(k) { html += '<td style="border:1px solid #333;padding:8px;">' + (item[k] || '') + '</td>'; });
            html += '</tr>';
        });
        html += '</tbody></table>';
        var win = window.open('', '_blank');
        win.document.write('<html><head><title>\u0637\u0628\u0627\u0639\u0629</title></head><body dir="rtl">' + html + '</body></html>');
        win.document.close();
        win.print();
        printModal.style.display = 'none';
    });
}

if (exportExcelBtn) {
    exportExcelBtn.addEventListener('click', function() {
        var checked = document.querySelectorAll('#print-columns-container input[type=checkbox]:checked');
        if (checked.length === 0) { alert('\u064a\u062c\u0628 \u0627\u062e\u062a\u064a\u0627\u0631 \u0639\u0645\u0648\u062f \u0648\u0627\u062d\u062f \u0639\u0644\u0649 \u0627\u0644\u0623\u0642\u0644!'); return; }
        var headers = [];
        var keys = [];
        checked.forEach(function(cb) { headers.push(cb.dataset.header); keys.push(cb.dataset.key); });
        var rows = currentSearchResults.map(function(item) {
            var row = {};
            keys.forEach(function(k, i) { row[headers[i]] = item[k] || ''; });
            return row;
        });
        var ws = XLSX.utils.json_to_sheet(rows);
        var wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "\u0627\u0644\u0628\u062d\u062b");
        XLSX.writeFile(wb, "\u0646\u062a\u0627\u0626\u062c_\u0627\u0644\u0628\u062d\u062b.xlsx");
        printModal.style.display = 'none';
    });
}

/* ===== Context Menu for Search Results ===== */
(function() {
    var ctxMenu = document.getElementById('context-menu');
    var ctxRowIndex = -1;

    document.addEventListener('click', function() { ctxMenu.style.display = 'none'; });

    document.getElementById('search-results-container').addEventListener('contextmenu', function(e) {
        var row = e.target.closest('tr');
        if (!row || !row.dataset.origIdx) return;
        e.preventDefault();
        ctxRowIndex = parseInt(row.dataset.origIdx, 10);
        ctxMenu.style.left = e.clientX + 'px';
        ctxMenu.style.top = e.clientY + 'px';
        ctxMenu.style.display = 'block';
    });

    document.getElementById('ctx-return').addEventListener('click', function() {
        if (ctxRowIndex === -1) return;
        ctxMenu.style.display = 'none';
        if (typeof searchReturnDevice === 'function') searchReturnDevice(ctxRowIndex);
    });

    document.getElementById('ctx-edit').addEventListener('click', function() {
        if (ctxRowIndex === -1) return;
        ctxMenu.style.display = 'none';
        if (typeof openEditDevice === 'function') openEditDevice(ctxRowIndex);
    });

    document.getElementById('ctx-delete').addEventListener('click', function() {
        if (ctxRowIndex === -1) return;
        ctxMenu.style.display = 'none';
        if (typeof searchDeleteDevice === 'function') searchDeleteDevice(ctxRowIndex);
    });
})();
