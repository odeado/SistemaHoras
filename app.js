// Local Hours-Extras Application Logic
// Handles calculations, localStorage, UI updates, and ExcelJS export.

document.addEventListener('DOMContentLoaded', () => {
  // Purge legacy mock data from client local storage if present
  try {
    for (let i = 0; i < localStorage.length; i++) {
      const k = localStorage.key(i);
      if (k && k.startsWith("horas_extras_data_")) {
        const stored = localStorage.getItem(k);
        if (stored && stored.includes("macrena saavedra")) {
          localStorage.clear();
          location.reload();
          return;
        }
      }
    }
  } catch (e) {
    console.error("Storage purge error:", e);
  }

  // --- Firebase Configuration ---
  const firebaseConfig = {
    apiKey: "AIzaSyCKO3c4WitANxhIC5Wc0IjXF8X8mi68nF8",
    authDomain: "gestion-usuarios-equipos.firebaseapp.com",
    databaseURL: "https://gestion-usuarios-equipos-default-rtdb.firebaseio.com",
    projectId: "gestion-usuarios-equipos",
    storageBucket: "gestion-usuarios-equipos.firebasestorage.app",
    messagingSenderId: "124302724881",
    appId: "1:124302724881:web:4a68cac2ae22d7b21c06b0",
    measurementId: "G-3Q5DRB8DL5"
  };

  let db = null;
  let isFirebaseInitialized = false;

  try {
    if (typeof firebase !== 'undefined') {
      firebase.initializeApp(firebaseConfig);
      db = firebase.database();
      isFirebaseInitialized = true;
    }
  } catch (e) {
    console.error("Firebase initialization failed", e);
  }

  // --- Constants & Config ---
  const DAY_NAMES = ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"];
  const MONTH_NAMES = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", 
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
  ];
  
  // Standard daily normal hours (42-hour contract)
  const STANDARD_HOURS = {
    1: 8.5, // lunes
    2: 7.5, // martes
    3: 7.5, // miércoles
    4: 8.0, // jueves
    5: 8.0, // viernes
    6: 2.5, // sábado
    0: 0.0  // domingo
  };

  // --- State Variables ---
  let currentYear = new Date().getFullYear();
  let currentMonth = new Date().getMonth(); // Dynamic current month (0-indexed)
  let completeWeek = true;
  let daysData = []; // Array of day objects
  let registros = []; // Array of all historical flat records
  let employeeInfo = {
    name: "MARCO GARRIDO",
    role: "",
    area: "Sistemas",
    card: "",
    manager: "Carlos Rodriguez"
  };

  // --- UI Elements ---
  const employeeNameInput = document.getElementById('employee-name');
  const employeeRoleInput = document.getElementById('employee-role');
  const employeeAreaInput = document.getElementById('employee-area');
  const employeeCardInput = document.getElementById('employee-card');
  const employeeManagerInput = document.getElementById('employee-manager');
  const periodYearSelect = document.getElementById('period-year');
  const periodMonthSelect = document.getElementById('period-month');
  const completeWeekCheckbox = document.getElementById('complete-week');
  const dateRangeDisplay = document.getElementById('date-range-display');
  const payrollTbody = document.getElementById('payroll-tbody');
  
  // Metric displays
  const metricTotalWorked = document.getElementById('metric-total-worked');
  const metricNormal = document.getElementById('metric-normal');
  const metricExtras = document.getElementById('metric-extras');
  const metricExtras50 = document.getElementById('metric-extras-50');
  const metricExtras100 = document.getElementById('metric-extras-100');
  
  // Table Footer elements
  const footTpoPerm = document.getElementById('foot-tpo-perm');
  const footHorasH = document.getElementById('foot-horas-h');
  const footHorasMin = document.getElementById('foot-horas-min');
  const footHorasTotalDec = document.getElementById('foot-horas-total-dec');
  const footSummaryTotal = document.getElementById('foot-summary-total');
  const footSummaryNorm = document.getElementById('foot-summary-norm');
  const footSummaryExtras = document.getElementById('foot-summary-extras');
  
  // Buttons
  const btnExportExcel = document.getElementById('btn-export-excel');
  const btnExportBackup = document.getElementById('btn-export-backup');
  const btnImportBackup = document.getElementById('btn-import-backup');
  const fileImportInput = document.getElementById('file-import');
  const btnClearData = document.getElementById('btn-clear-data');
  const themeToggleBtn = document.getElementById('theme-toggle');
  const statusMessage = document.getElementById('status-message');

  // Form View Selectors
  const tabFormBtn = document.getElementById('tab-form');
  const tabTableBtn = document.getElementById('tab-table');
  const panelFormContent = document.getElementById('panel-form-content');
  const panelTableContent = document.getElementById('panel-table-content');
  const formDaySelect = document.getElementById('form-day-select');
  const formIsFeriado = document.getElementById('form-is-feriado');
  const formDeCorrido = document.getElementById('form-de-corrido');
  const formCustomHours = document.getElementById('form-custom-hours');
  const formHoursContainer = document.getElementById('form-hours-container');
  const formComment = document.getElementById('form-comment');
  const formEntTime = document.getElementById('form-ent-time');
  const formSalTime = document.getElementById('form-sal-time');
  const formBtnSaveNext = document.getElementById('form-btn-save-next');
  const formBtnClear = document.getElementById('form-btn-clear');

  // Sidebar Mobile Selectors
  const sidebar = document.querySelector('.sidebar');
  const sidebarToggle = document.getElementById('sidebar-toggle');
  const sidebarBackdrop = document.getElementById('sidebar-backdrop');

  // --- Initialization ---

  function initApp() {
    // 1. Load theme preference
    const savedTheme = localStorage.getItem('payroll_theme') || 'light';
    document.documentElement.setAttribute('data-theme', savedTheme);
    themeToggleBtn.textContent = savedTheme === 'dark' ? '☀️' : '🌙';

    // Set up Firebase connection status indicator
    if (isFirebaseInitialized) {
      const syncStatusEl = document.getElementById('sync-status');
      const syncLabel = syncStatusEl.querySelector('.label');
      
      db.ref(".info/connected").on("value", (snap) => {
        if (snap.val() === true) {
          syncStatusEl.className = 'sync-indicator';
          syncLabel.textContent = 'Sincronizado';
        } else {
          syncStatusEl.className = 'sync-indicator offline';
          syncLabel.textContent = 'Sin conexión';
        }
      });
    }

    // 2. Setup Year and Month select options dynamically around current time
    const currentLocalYear = new Date().getFullYear();
    periodYearSelect.innerHTML = '';
    for (let y = currentLocalYear - 2; y <= currentLocalYear + 2; y++) {
      const opt = document.createElement('option');
      opt.value = y;
      opt.textContent = y;
      if (y === currentYear) opt.selected = true;
      periodYearSelect.appendChild(opt);
    }

    // 3. Load initial values from controls
    periodMonthSelect.value = currentMonth;
    currentYear = parseInt(periodYearSelect.value);
    currentMonth = parseInt(periodMonthSelect.value);
    completeWeek = completeWeekCheckbox.checked;

    // 4. Load from LocalStorage/Firebase
    loadData();

    // 5. Setup Event Listeners
    setupEventListeners();
  }

  // --- Event Listeners ---
  function setupEventListeners() {
    // Theme toggle
    themeToggleBtn.addEventListener('click', () => {
      const currentTheme = document.documentElement.getAttribute('data-theme');
      const newTheme = currentTheme === 'dark' ? 'light' : 'dark';
      document.documentElement.setAttribute('data-theme', newTheme);
      themeToggleBtn.textContent = newTheme === 'dark' ? '☀️' : '🌙';
      localStorage.setItem('payroll_theme', newTheme);
    });

    // Employee Info Change
    [employeeNameInput, employeeRoleInput, employeeAreaInput, employeeCardInput, employeeManagerInput].forEach(input => {
      input.addEventListener('input', () => {
        employeeInfo.name = employeeNameInput.value;
        employeeInfo.role = employeeRoleInput.value;
        employeeInfo.area = employeeAreaInput.value;
        employeeInfo.card = employeeCardInput.value;
        employeeInfo.manager = employeeManagerInput.value;
        saveLocalOnly();
      });
      
      input.addEventListener('change', () => {
        saveToFirebaseOnly();
        if (input === employeeNameInput && isFirebaseInitialized) {
          setupFirebaseSync();
        }
      });
    });

    // Period Config Change
    periodYearSelect.addEventListener('change', handlePeriodChange);
    periodMonthSelect.addEventListener('change', handlePeriodChange);
    completeWeekCheckbox.addEventListener('change', handlePeriodChange);

    // Actions
    btnExportExcel.addEventListener('click', exportToExcel);
    btnExportBackup.addEventListener('click', exportBackup);
    btnImportBackup.addEventListener('click', () => fileImportInput.click());
    fileImportInput.addEventListener('change', importBackup);
    btnClearData.addEventListener('click', clearCurrentPeriodData);

    // Setup Form View Event Listeners
    setupFormEventListeners();
  }

  function handlePeriodChange() {
    currentYear = parseInt(periodYearSelect.value);
    currentMonth = parseInt(periodMonthSelect.value);
    completeWeek = completeWeekCheckbox.checked;
    
    // Reset day selector to the first day of the new period
    formDaySelect.value = "0";
    
    loadData();
  }

  // --- Data Management (Local Storage) ---

  const FERIADOS = {
    '2026-01-01': 'Año Nuevo',
    '2026-04-03': 'Viernes Santo',
    '2026-04-04': 'Sábado Santo',
    '2026-05-01': 'Día del Trabajo',
    '2026-05-21': 'Día de las Glorias Navales',
    '2026-06-29': 'San Pedro y San Pablo',
    '2026-07-16': 'Día de la Virgen del Carmen',
    '2026-08-15': 'Asunción de la Virgen',
    '2026-09-18': 'Independencia Nacional',
    '2026-09-19': 'Día de las Glorias del Ejército',
    '2026-10-12': 'Día del Encuentro de Dos Mundos',
    '2026-10-31': 'Día de las Iglesias Evangélicas',
    '2026-11-01': 'Día de Todos los Santos',
    '2026-12-08': 'Inmaculada Concepción',
    '2026-12-25': 'Navidad'
  };

  const BASE_HORARIOS_SPANISH = {
    'Lunes': { e1: '09:00', s1: '13:30', e2: '15:00', s2: '19:00', base: 8.5 },
    'Martes': { e1: '09:30', s1: '13:30', e2: '15:00', s2: '18:30', base: 7.5 },
    'Miércoles': { e1: '09:30', s1: '13:30', e2: '15:00', s2: '18:30', base: 7.5 },
    'Jueves': { e1: '09:30', s1: '13:30', e2: '15:00', s2: '19:00', base: 8.0 },
    'Viernes': { e1: '09:30', s1: '13:30', e2: '15:00', s2: '19:00', base: 8.0 },
    'Sábado': { e1: '10:00', s1: '12:30', e2: '12:30', s2: '12:30', base: 2.5 },
    'Domingo': { e1: '', s1: '', e2: '', s2: '', base: 0.0 }
  };

  function timeToMin(t) {
    if (!t) return 0;
    const p = t.split(':');
    return parseInt(p[0], 10) * 60 + parseInt(p[1], 10);
  }

  function minToHrs(m) { return m / 60; }

  function getDiaSemanaRealLocal(fechaStr) {
    if (!fechaStr) return '—';
    const d = new Date(fechaStr + 'T00:00:00');
    const dias = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'];
    return dias[d.getDay()];
  }

  function esFeriadoReal(fechaStr) {
    if (!fechaStr) return false;
    return FERIADOS[fechaStr] !== undefined;
  }

  function getSemana(diaNum, mes, año) {
    const fecha = new Date(año, mes, diaNum);
    const diaDelMes = fecha.getDate();
    if (diaDelMes >= 16 && diaDelMes <= 22) return 1;
    if (diaDelMes >= 23 && diaDelMes <= 29) return 2;
    if (diaDelMes >= 30 && diaDelMes <= 31) return 3;
    if (diaDelMes >= 1 && diaDelMes <= 5) return 3;
    if (diaDelMes >= 6 && diaDelMes <= 12) return 4;
    if (diaDelMes >= 13 && diaDelMes <= 16) return 5;
    return 1;
  }

  function calcularHorasDia(fechaStr, esCorrido, horaFinalStr, esFeriadoFlag, entradaManual, salidaManual, salidaSabado) {
    const diaNombre = getDiaSemanaRealLocal(fechaStr);
    const esDom = diaNombre === 'Domingo';
    const esSab = diaNombre === 'Sábado';
    const esFer = esFeriadoFlag || esFeriadoReal(fechaStr);

    if (esDom || esFer) {
      const entrada = entradaManual || '';
      const salida = salidaManual || '';
      if (!entrada || !salida) {
        return { total: 0, normales: 0, extras: 0 };
      }
      const totalMin = timeToMin(salida) - timeToMin(entrada);
      const total = totalMin > 0 ? minToHrs(totalMin) : 0;
      return { total: total, normales: 0, extras: total };
    }

    if (esSab) {
      const entrada = '10:00';
      const salida = salidaSabado || '12:30';
      const totalMin = timeToMin(salida) - timeToMin(entrada);
      const total = totalMin > 0 ? minToHrs(totalMin) : 0;
      const base = 2.5;
      const extras = total > base ? total - base : 0;
      const normales = total - extras;
      return { total: total, normales: normales, extras: extras };
    }

    const base = BASE_HORARIOS_SPANISH[diaNombre] || BASE_HORARIOS_SPANISH['Lunes'];
    const final = horaFinalStr || '19:00';

    let e1, s1, e2, s2;
    if (esCorrido) {
      e1 = base.e1;
      s1 = base.s1;
      e2 = base.s1;
      s2 = final;
    } else {
      e1 = base.e1;
      s1 = base.s1;
      e2 = base.e2;
      s2 = final;
    }

    const b1 = timeToMin(s1) - timeToMin(e1);
    const b2 = timeToMin(s2) - timeToMin(e2);
    let totalMin = b1 + b2;
    if (totalMin < 0) totalMin = 0;
    const total = minToHrs(totalMin);

    const baseMin = base.base * 60;
    const extraMin = totalMin > baseMin ? totalMin - baseMin : 0;
    const extras = minToHrs(extraMin);
    const normales = total - extras;

    return { total, normales, extras };
  }

  function migrateLegacyRecords() {
    let modified = false;
    if (!Array.isArray(registros)) return;
    const lengthBefore = registros.length;

    registros = registros.map(r => {
      const dayName = getDiaSemanaRealLocal(r.fecha);
      if (dayName === 'Lunes') {
        const hasComment = r.motivo && r.motivo.trim() !== '';
        if (!hasComment && (r.s2 === '19:30' || r.e1 === '09:30')) {
          r.e1 = '09:00';
          r.s1 = '13:30';
          r.e2 = '15:00';
          r.s2 = '19:00';
          r.total = 8.5;
          r.normales = 8.5;
          r.extras = 0;
          r.esCorrido = false;
          modified = true;
        }
      }
      return r;
    });

    registros = registros.filter(r => {
      const dayName = getDiaSemanaRealLocal(r.fecha);
      const isDom = r.esDomingo || dayName === 'Domingo';
      const isSab = r.esSabado || dayName === 'Sábado';
      const isFer = !!r.esFeriado;
      
      const hasComment = r.motivo && r.motivo.trim() !== '';
      if (hasComment) return true;

      if (isDom || isFer) {
        return !!(r.e1 && r.s1 && r.e1 !== '' && r.s1 !== '');
      }

      const base = BASE_HORARIOS_SPANISH[dayName] || BASE_HORARIOS_SPANISH['Lunes'];
      const matchesStandard = (r.e1 === base.e1 && r.s2 === base.s2 && !r.esCorrido);
      return !matchesStandard;
    });

    if (modified || registros.length !== lengthBefore) {
      try {
        localStorage.setItem('horasExtras_registros', JSON.stringify(registros));
      } catch(e) {}
      saveToFirebaseOnly();
    }
  }

  function getStorageKey() {
    return 'horasExtras_registros';
  }

  function loadData() {
    // 1. Load employee global details
    const savedEmployee = localStorage.getItem('employee_global_info');
    if (savedEmployee) {
      employeeInfo = JSON.parse(savedEmployee);
      employeeNameInput.value = employeeInfo.name || '';
      employeeRoleInput.value = employeeInfo.role || '';
      employeeAreaInput.value = employeeInfo.area || '';
      employeeCardInput.value = employeeInfo.card || '';
      employeeManagerInput.value = employeeInfo.manager || '';
    }

    // 2. Generate date sequence for the selected period
    const dates = generatePeriodDates(currentYear, currentMonth, completeWeek);
    
    // Show period range text
    if (dates.length > 0) {
      const firstDate = dates[0].dateObj;
      const lastDate = dates[dates.length - 1].dateObj;
      dateRangeDisplay.textContent = formatDateStr(firstDate) + ' al ' + formatDateStr(lastDate);
    }

    // 3. Load local storage 'horasExtras_registros'
    try {
      const localData = localStorage.getItem('horasExtras_registros');
      if (localData) {
        registros = JSON.parse(localData);
        migrateLegacyRecords();
        registros.sort((a, b) => {
          if (a.año !== b.año) return a.año - b.año;
          if (a.mes !== b.mes) return a.mes - b.mes;
          return a.dia - b.dia;
        });
      } else {
        registros = [];
      }
    } catch(e) {
      registros = [];
    }

    // 4. Populate daysData from registros
    syncDaysDataFromRegistros(dates);

    // 5. Render local data immediately for instant visibility
    renderTable();
    populateFormDaySelect();
    loadFormDayData();

    // 6. If Firebase is initialized, start cloud sync in background
    if (isFirebaseInitialized) {
      setupFirebaseSync();
    }
  }

  function syncDaysDataFromRegistros(dates) {
    daysData = dates.map(d => {
      const dayOfWeek = d.dayOfWeek;
      const dateKey = d.dateKey; // YYYY-MM-DD
      const parts = dateKey.split('-');
      const year = parseInt(parts[0], 10);
      const month = parseInt(parts[1], 10) - 1; // 0-indexed month
      const dayNum = parseInt(parts[2], 10);

      // Search in registros
      const matched = registros.find(r => r.dia === dayNum && r.mes === month && r.año === year);

      let shifts = ['', '', '', '', '', '', '', ''];
      let comment = '';
      let isFeriado = false;
      let isCorrido = false;
      let isRegistrado = false;

      if (matched) {
        isRegistrado = true;
        isFeriado = !!matched.esFeriado;
        isCorrido = !!matched.esCorrido;
        comment = matched.motivo || '';
        shifts = [
          matched.e1 || '', matched.s1 || '',
          matched.e2 || '', matched.s2 || '',
          '', '', '', ''
        ];
      } else {
        // Defaults
        isFeriado = esFeriadoReal(dateKey);
        if (!isFeriado && dayOfWeek !== 0) {
          const start = getStandardStartTime(dayOfWeek);
          const end = getStandardEndTime(dayOfWeek);
          if (dayOfWeek === 6) {
            shifts = [start, end, '', '', '', '', '', ''];
          } else {
            shifts = [start, '13:30', '15:00', end, '', '', '', ''];
          }
        }
      }

      return {
        dateKey,
        dayNum,
        dayName: d.dayName,
        dayOfWeek,
        dateObj: d.dateObj,
        isFeriado,
        deCorrido: isCorrido,
        comment,
        shifts,
        isRegistrado
      };
    });
  }

  function generateEmptyDays(dates) {
    syncDaysDataFromRegistros(dates);
  }

  function saveData() {
    saveLocalOnly();
    saveToFirebaseOnly();
  }

  function saveLocalOnly() {
    // 1. Save employee global details
    localStorage.setItem('employee_global_info', JSON.stringify(employeeInfo));

    // 2. Translate daysData of current period into registros flat list
    daysData.forEach(day => {
      const parts = day.dateKey.split('-');
      const year = parseInt(parts[0], 10);
      const month = parseInt(parts[1], 10) - 1;
      const dayNum = parseInt(parts[2], 10);

      // Check if the day is custom/modified:
      const isDom = day.dayOfWeek === 0;
      const isFer = day.isFeriado;
      const standardFeriado = esFeriadoReal(day.dateKey);
      const stdStart = getStandardStartTime(day.dayOfWeek);
      const stdEnd = getStandardEndTime(day.dayOfWeek);
      
      let isModified = false;
      if (day.comment && day.comment.trim() !== '') {
        isModified = true;
      }
      if (isFer !== standardFeriado) {
        isModified = true;
      }

      // Check if shifts are custom:
      if (isDom || isFer) {
        if (day.shifts[0] || day.shifts[1]) {
          isModified = true;
        }
      } else if (day.dayOfWeek === 6) {
        if ((day.shifts[0] && day.shifts[0] !== stdStart) || (day.shifts[1] && day.shifts[1] !== stdEnd)) {
          isModified = true;
        }
      } else {
        if ((day.shifts[0] && day.shifts[0] !== stdStart) || (day.shifts[3] && day.shifts[3] !== stdEnd)) {
          isModified = true;
        }
        if (day.shifts[2] === '13:30') {
          // deCorrido is true
          isModified = true;
        }
      }

      if (isModified) {
        let e1 = day.shifts[0] || stdStart;
        let s1 = '';
        let e2 = '';
        let s2 = '';
        const deCorrido = (day.shifts[2] === '13:30');

        if (isDom || isFer) {
          e1 = day.shifts[0] || '09:30';
          s1 = day.shifts[1] || '15:30';
          e2 = s1;
          s2 = s1;
        } else if (day.dayOfWeek === 6) {
          s1 = day.shifts[1] || '12:30';
          e2 = s1;
          s2 = s1;
        } else {
          s1 = '13:30';
          e2 = deCorrido ? '13:30' : '15:00';
          s2 = day.shifts[3] || stdEnd;
        }

        const h = calcularHorasDia(day.dateKey, deCorrido, s2 || stdEnd, isFer, e1 || stdStart, s1 || stdEnd, s1 || stdEnd);

        const registro = {
          dia: dayNum,
          fecha: day.dateKey,
          mes: month,
          año: year,
          semana: getSemana(dayNum, month, year),
          e1: e1,
          s1: s1,
          e2: e2,
          s2: s2,
          total: h.total,
          normales: h.normales,
          extras: h.extras,
          esCorrido: deCorrido,
          esManual: isDom || isFer,
          esDomingo: isDom,
          esSabado: day.dayOfWeek === 6,
          esFeriado: isFer,
          motivo: day.comment || '',
          horaFinal: s2 || stdEnd,
          entradaManual: e1 || stdStart,
          salidaManual: s1 || stdEnd,
          sabadoSalida: s1 || stdEnd
        };

        const idx = registros.findIndex(r => r.dia === dayNum && r.mes === month && r.año === year);
        if (idx !== -1) {
          registros[idx] = registro;
        } else {
          registros.push(registro);
        }
      } else {
        // Remove from registros if it was there
        registros = registros.filter(r => !(r.dia === dayNum && r.mes === month && r.año === year));
      }
    });

    registros.sort((a, b) => {
      if (a.año !== b.año) return a.año - b.año;
      if (a.mes !== b.mes) return a.mes - b.mes;
      return a.dia - b.dia;
    });

    try {
      localStorage.setItem('horasExtras_registros', JSON.stringify(registros));
    } catch(e) {}
  }

  function sanitizeFirebaseKey(key) {
    if (!key) return "ANONIMO";
    return key.replace(/[\.\$\#\[\]\/]/g, '_').trim().replace(/\s+/g, '_').toUpperCase();
  }

  function saveToFirebaseOnly() {
    if (!isFirebaseInitialized) return;
    const sanitizedName = sanitizeFirebaseKey(employeeInfo.name || "MARCO_GARRIDO");
    const dbPath = 'horas_extras/' + sanitizedName + '/registros';
    
    setSyncState('syncing');
    
    db.ref(dbPath).set(registros)
    .then(() => {
      setSyncState('online');
    })
    .catch((err) => {
      console.error("Firebase write failed", err);
      setSyncState('offline');
    });
  }

  let firebaseListener = null;

  function setupFirebaseSync() {
    if (!isFirebaseInitialized) return;

    if (firebaseListener) {
      firebaseListener.off();
    }

    const sanitizedName = sanitizeFirebaseKey(employeeInfo.name || "MARCO_GARRIDO");
    const dbPath = 'horas_extras/' + sanitizedName + '/registros';
    firebaseListener = db.ref(dbPath);

    setSyncState('syncing');

    firebaseListener.on('value', (snapshot) => {
      const val = snapshot.val();
      if (val) {
        if (JSON.stringify(registros) !== JSON.stringify(val)) {
          registros = val;
          migrateLegacyRecords();
          registros.sort((a, b) => {
            if (a.año !== b.año) return a.año - b.año;
            if (a.mes !== b.mes) return a.mes - b.mes;
            return a.dia - b.dia;
          });
          try {
            localStorage.setItem('horasExtras_registros', JSON.stringify(registros));
          } catch (e) {}
          
          const dates = generatePeriodDates(currentYear, currentMonth, completeWeek);
          syncDaysDataFromRegistros(dates);
          renderTable();
          populateFormDaySelect();
          loadFormDayData();
        }
        setSyncState('online');
      } else {
        if (registros.length > 0) {
          saveToFirebaseOnly();
        } else {
          setSyncState('online');
        }
      }
    }, (error) => {
      console.error("Firebase sync error:", error);
      setSyncState('offline');
    });
  }

  function loadLocalStorageBackup(stored, dates) {
    const datesSeq = generatePeriodDates(currentYear, currentMonth, completeWeek);
    syncDaysDataFromRegistros(datesSeq);
    renderTable();
    populateFormDaySelect();
    loadFormDayData();
  }

  function loadMockOrEmptyLocal(dates) {
    syncDaysDataFromRegistros(dates);
  }
  // --- Form View Handlers ---
  function setupFormEventListeners() {
    // Dropdown change
    formDaySelect.addEventListener('change', loadFormDayData);
 
    // Sidebar Mobile Toggle
    if (sidebarToggle && sidebar && sidebarBackdrop) {
      sidebarToggle.addEventListener('click', () => {
        sidebar.classList.toggle('show');
        sidebarBackdrop.classList.toggle('show');
      });
 
      sidebarBackdrop.addEventListener('click', () => {
        sidebar.classList.remove('show');
        sidebarBackdrop.classList.remove('show');
      });
    }
 
    // Toggle custom hours visibility
    formCustomHours.addEventListener('change', () => {
      formHoursContainer.style.display = formCustomHours.checked ? 'grid' : 'none';
      saveFormDayData();
    });
 
    // Time input auto-formatting while typing and local real-time updating
    [formEntTime, formSalTime].forEach(input => {
      input.addEventListener('input', (e) => {
        formatTimeInput(e);
        const val = input.value.trim();
        // If a valid time is fully typed (length 5), update locally in real-time
        if (val.length === 5) {
          const regex = /^([0-9]|0[0-9]|1[0-9]|2[0-3]):[0-5][0-9]$/;
          if (regex.test(val)) {
            saveFormDayDataLocal();
          }
        }
      });
    });
 
    // Save form data when fields change (blur/exit)
    [formEntTime, formSalTime, formComment, formIsFeriado, formDeCorrido].forEach(input => {
      input.addEventListener('change', () => {
        // Validate time format if it's a time input
        if (input === formEntTime || input === formSalTime) {
          const val = input.value.trim();
          if (val !== "") {
            let formattedVal = val;
            if (val.length === 4 && val.indexOf(':') === 1) {
              formattedVal = '0' + val;
            }
            const regex = /^([0-9]|0[0-9]|1[0-9]|2[0-3]):[0-5][0-9]$/;
            if (regex.test(formattedVal)) {
              input.value = formattedVal;
            } else {
              showStatus("Formato de hora inválido (HH:MM)", "danger");
              input.value = "";
            }
          }
        }
        saveFormDayData(); // Writes to Firebase
        if (input === formIsFeriado) {
          loadFormDayData();
        }
      });
    });
 
    // Auto-parse times from motive/comment in real time, updating locally
    formComment.addEventListener('input', () => {
      parseTimesFromComment();
      saveFormDayDataLocal(); // Local real-time update
    });
 
    // Clear day in form
    formBtnClear.addEventListener('click', () => {
      const dayIdx = parseInt(formDaySelect.value);
      if (isNaN(dayIdx)) return;
      if (confirm("¿Limpiar turnos y comentario de este día?")) {
        daysData[dayIdx].shifts = ['', '', '', '', '', '', '', ''];
        daysData[dayIdx].comment = '';
        daysData[dayIdx].isFeriado = false;
        saveData();
        loadFormDayData();
        updateTotalsAndLabels();
        showStatus("Día limpiado", "success");
      }
    });
 
    // Save and Next Day
    formBtnSaveNext.addEventListener('click', () => {
      const dayIdx = parseInt(formDaySelect.value);
      if (isNaN(dayIdx)) return;
      
      saveFormDayData(); // Ensure saved to Firebase
      
      // Advance dropdown
      if (dayIdx < daysData.length - 1) {
        formDaySelect.value = String(dayIdx + 1);
        loadFormDayData();
        showStatus("Día guardado, avanzando...", "success");
      } else {
        showStatus("Último día del período guardado.", "success");
      }
    });
  }

  function populateFormDaySelect() {
    const selectedVal = formDaySelect.value;
    
    // If the dropdown already has the correct number of options, just update the text labels to avoid rebuilding DOM on sync
    if (formDaySelect.options.length === daysData.length) {
      daysData.forEach((day, idx) => {
        const opt = formDaySelect.options[idx];
        const calc = calculateDayHours(day);
        let workedStr = calc.hasHours ? ` (${calc.timeStr} hrs)` : '';
        if (day.isFeriado) workedStr += ' [Feriado]';
        opt.textContent = `${day.dayNum} - ${day.dayName}${workedStr}`;
      });
      return;
    }
    
    formDaySelect.innerHTML = '';
    
    daysData.forEach((day, idx) => {
      const opt = document.createElement('option');
      opt.value = idx;
      
      const calc = calculateDayHours(day);
      let workedStr = calc.hasHours ? ` (${calc.timeStr} hrs)` : '';
      if (day.isFeriado) workedStr += ' [Feriado]';
      
      opt.textContent = `${day.dayNum} - ${day.dayName}${workedStr}`;
      
      if (selectedVal !== "" && parseInt(selectedVal) === idx) {
        opt.selected = true;
      }
      formDaySelect.appendChild(opt);
    });
    
    if (formDaySelect.value === "") {
      formDaySelect.value = "0";
    }
  }

  function getStandardStartTime(dayOfWeek) {
    switch (dayOfWeek) {
      case 1: return "09:00"; // lunes
      case 6: return "10:00"; // sábado
      default: return "09:30"; // martes a viernes, domingo
    }
  }

  function getStandardEndTime(dayOfWeek) {
    switch (dayOfWeek) {
      case 1: return "19:00"; // lunes
      case 2:
      case 3: return "18:30"; // martes, miércoles
      case 4:
      case 5: return "19:00"; // jueves, viernes
      case 6: return "12:30"; // sábado
      default: return "15:30"; // domingo
    }
  }

  function resetDayToDefault(day) {
    day.comment = '';
    day.isFeriado = false;
    let shifts = ['', '', '', '', '', '', '', ''];
    if (day.dayOfWeek !== 0) {
      const start = getStandardStartTime(day.dayOfWeek);
      const end = getStandardEndTime(day.dayOfWeek);
      if (day.dayOfWeek === 6) {
        shifts = [start, end, '', '', '', '', '', ''];
      } else {
        shifts = [start, '13:30', '15:00', end, '', '', '', ''];
      }
    }
    day.shifts = shifts;
  }

  function parseTimesFromComment() {
    const commentVal = formComment.value.trim();
    if (!commentVal) return;

    // 0. Detect continuous work (de corrido)
    if (/de corrido|sin colaci[oó]n|trabajo continuo/i.test(commentVal)) {
      formDeCorrido.checked = true;
    }

    // 1. Check for "de HH:MM a HH:MM"
    const deAPattern = /(?:de|desde)\s+(?:las\s+)?(\d{1,2})[:.](\d{2})\s+(?:a|hasta)\s+(?:las\s+)?(\d{1,2})[:.](\d{2})/i;
    const deAMatch = commentVal.match(deAPattern);
    if (deAMatch) {
      const entH = deAMatch[1].padStart(2, '0');
      const entM = deAMatch[2];
      const salH = deAMatch[3].padStart(2, '0');
      const salM = deAMatch[4];
      formEntTime.value = `${entH}:${entM}`;
      formSalTime.value = `${salH}:${salM}`;
      formCustomHours.checked = true;
      formHoursContainer.style.display = 'grid';
      return;
    }

    // 2. Check for "hasta las HH:MM"
    const hastaPattern = /(?:hasta las|hasta)\s+(\d{1,2})[:.](\d{2})/i;
    const hastaMatch = commentVal.match(hastaPattern);
    if (hastaMatch) {
      const salH = hastaMatch[1].padStart(2, '0');
      const salM = hastaMatch[2];
      formSalTime.value = `${salH}:${salM}`;
      return;
    }

    // 3. Check for "de HH a HH"
    const deAHourPattern = /(?:de|desde)\s+(?:las\s+)?(\d{1,2})\s+(?:a|hasta)\s+(?:las\s+)?(\d{1,2})/i;
    const deAHourMatch = commentVal.match(deAHourPattern);
    if (deAHourMatch) {
      const entH = deAHourMatch[1].padStart(2, '0');
      const salH = deAHourMatch[2].padStart(2, '0');
      formEntTime.value = `${entH}:00`;
      formSalTime.value = `${salH}:00`;
      formCustomHours.checked = true;
      formHoursContainer.style.display = 'grid';
      return;
    }

    // 4. Check for "hasta las HH"
    const hastaHourPattern = /(?:hasta las|hasta)\s+(\d{1,2})(?!\d)/i;
    const hastaHourMatch = commentVal.match(hastaHourPattern);
    if (hastaHourMatch) {
      const salH = hastaHourMatch[1].padStart(2, '0');
      formSalTime.value = `${salH}:00`;
      return;
    }
  }

  function loadFormDayData() {
    const idx = parseInt(formDaySelect.value);
    if (isNaN(idx) || !daysData[idx]) return;
    
    const day = daysData[idx];
    
    // Display full Spanish date under day dropdown
    const daysOfWeekSpanish = ["Domingo", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"];
    const monthsSpanish = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
    
    const dObj = day.dateObj;
    const dayNameStr = daysOfWeekSpanish[dObj.getDay()];
    const monthNameStr = monthsSpanish[dObj.getMonth()];
    const fullDateText = `${dayNameStr}, ${dObj.getDate()} de ${monthNameStr} de ${dObj.getFullYear()}`;
    const fullDateDisplayEl = document.getElementById('form-full-date-display');
    if (fullDateDisplayEl) {
      fullDateDisplayEl.textContent = fullDateText;
    }
    
    // 1. Set comment/motive directly (ticket field is removed)
    formComment.value = day.comment || "";
    
    // 2. Set Feriado checkbox
    formIsFeriado.checked = !!day.isFeriado;
    
    // 3. Set De Corrido checkbox
    let deCorrido = false;
    let entVal = "";
    let salVal = "";
    let hasCustomHours = false;
    
    const stdStart = getStandardStartTime(day.dayOfWeek);
    const stdEnd = getStandardEndTime(day.dayOfWeek);

    const isRestDay = (day.dayOfWeek === 0 || day.isFeriado);

    if (isRestDay) {
      // Sunday/Holiday: 1 shift (index 0, 1)
      entVal = day.shifts[0] || "09:30";
      salVal = day.shifts[1] || "15:30";
      deCorrido = true; // Always continuous on rest days
      hasCustomHours = true; // Always show entrance for Sunday/Holiday
      
      const deCorridoWrapper = document.getElementById('form-de-corrido-wrapper');
      if (deCorridoWrapper) deCorridoWrapper.style.display = 'none';
    } else if (day.dayOfWeek === 6) {
      // Saturday: 1 shift (index 0, 1)
      entVal = day.shifts[0] || "";
      salVal = day.shifts[1] || "";
      deCorrido = true;
      hasCustomHours = (entVal !== "" && entVal !== stdStart);
      
      const deCorridoWrapper = document.getElementById('form-de-corrido-wrapper');
      if (deCorridoWrapper) deCorridoWrapper.style.display = 'none';
    } else {
      // Weekdays: 2 shifts
      entVal = day.shifts[0] || "";
      if (day.shifts[2] === "13:30") {
        deCorrido = true;
      }
      salVal = day.shifts[3] || day.shifts[1] || "";
      hasCustomHours = (entVal !== "" && entVal !== stdStart);
      
      const deCorridoWrapper = document.getElementById('form-de-corrido-wrapper');
      if (deCorridoWrapper) deCorridoWrapper.style.display = 'flex';
    }
    
    formDeCorrido.checked = deCorrido;
    formCustomHours.checked = hasCustomHours;
    formHoursContainer.style.display = hasCustomHours ? 'grid' : 'none';
    
    // Set text input values for hours
    formEntTime.value = entVal || stdStart;
    formSalTime.value = salVal || stdEnd;
  }

  function saveFormDayDataLocal() {
    const dayIdx = parseInt(formDaySelect.value);
    if (isNaN(dayIdx) || !daysData[dayIdx]) return;
    
    const day = daysData[dayIdx];
    const isFeriado = formIsFeriado.checked;
    const deCorrido = formDeCorrido.checked;
    const motivo = formComment.value.trim();
    const hasCustomHours = formCustomHours.checked;
    
    let entVal = formEntTime.value.trim();
    let salVal = formSalTime.value.trim();
    
    day.isFeriado = isFeriado;
    day.comment = motivo;
    
    const stdStart = getStandardStartTime(day.dayOfWeek);
    const stdEnd = getStandardEndTime(day.dayOfWeek);
    
    const isRestDay = (isFeriado || day.dayOfWeek === 0);
    const finalEnt = entVal || stdStart;
    const finalSal = salVal || stdEnd;
    
    if (isRestDay) {
      // Only save shift on Sunday/Holiday if there is actual worked exit time or motive
      if (salVal || motivo) {
        day.shifts = [finalEnt, finalSal, "", "", "", "", "", ""];
      } else {
        day.shifts = ["", "", "", "", "", "", "", ""];
      }
    } else if (day.dayOfWeek === 6) {
      if (hasCustomHours || (salVal && salVal !== stdEnd) || motivo) {
        day.shifts = [finalEnt, finalSal, "", "", "", "", "", ""];
      } else {
        day.shifts = [stdStart, stdEnd, "", "", "", "", "", ""];
      }
    } else {
      if (deCorrido) {
        day.shifts = [finalEnt, "13:30", "13:30", finalSal, "", "", "", ""];
      } else {
        day.shifts = [finalEnt, "13:30", "15:00", finalSal, "", "", "", ""];
      }
    }
    
    saveLocalOnly();
    renderTable();
    
    // Update current option label in select dropdown
    const opt = formDaySelect.options[dayIdx];
    if (opt) {
      const calc = calculateDayHours(day);
      let workedStr = calc.hasHours ? ` (${calc.timeStr} hrs)` : '';
      if (day.isFeriado) workedStr += ' [Feriado]';
      opt.textContent = `${day.dayNum} - ${day.dayName}${workedStr}`;
    }
  }

  function saveFormDayData() {
    saveFormDayDataLocal();
    saveToFirebaseOnly();
  }

  function updateTotalsAndLabels() {
    renderTable(); 
    populateFormDaySelect();
  }

  // --- Date Math Helpers ---
  function generatePeriodDates(year, month, extendToSunday) {
    // Period starts on the 16th of the PREVIOUS month
    // If currentMonth is May (4), previous month is April (3)
    let startYear = year;
    let startMonth = month - 1;
    if (startMonth < 0) {
      startMonth = 11;
      startYear--;
    }

    const startDate = new Date(startYear, startMonth, 16);
    const endDate = new Date(year, month, 16);
    
    // Generate dates sequence
    const dates = [];
    let current = new Date(startDate);
    
    while (current <= endDate) {
      dates.push(createDateItem(new Date(current)));
      current.setDate(current.getDate() + 1);
    }

    // Extend to Sunday if requested
    if (extendToSunday) {
      let last = dates[dates.length - 1].dateObj;
      let dayOfWeek = last.getDay(); // 0 is Sunday, 6 is Saturday
      
      while (dayOfWeek !== 0) { // Keep adding days until Sunday is reached
        let nextDay = new Date(last);
        nextDay.setDate(nextDay.getDate() + 1);
        dates.push(createDateItem(nextDay));
        last = nextDay;
        dayOfWeek = last.getDay();
      }
    }

    return dates;
  }

  function createDateItem(d) {
    const dayOfWeek = d.getDay();
    const dayNum = d.getDate();
    const dayName = DAY_NAMES[dayOfWeek];
    // Key format: YYYY-MM-DD
    const dateKey = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(dayNum).padStart(2, '0')}`;
    
    return {
      dateKey,
      dayNum,
      dayName,
      dayOfWeek,
      dateObj: d
    };
  }

  function formatDateStr(d) {
    return `${d.getDate()} de ${MONTH_NAMES[d.getMonth()]} de ${d.getFullYear()}`;
  }

  // --- Time Calculation Engines ---
  // Parse time "HH:MM" -> minutes
  function timeToMin(timeStr) {
    if (!timeStr || !timeStr.trim()) return null;
    const parts = timeStr.trim().split(':');
    if (parts.length < 2) return null;
    const hrs = parseInt(parts[0], 10);
    const mins = parseInt(parts[1], 10);
    if (isNaN(hrs) || isNaN(mins)) return null;
    return hrs * 60 + mins;
  }

  // Format minutes -> time string "HH:MM"
  function minToTimeStr(totalMins) {
    if (totalMins === null || totalMins < 0) return "";
    const hrs = Math.floor(totalMins / 60);
    const mins = Math.round(totalMins % 60);
    return `${String(hrs).padStart(2, '0')}:${String(mins).padStart(2, '0')}`;
  }

  // Calculate day worked minutes and decimal hours
  function calculateDayHours(day) {
    let totalMins = 0;
    
    // Loop through 4 shifts (entries at indices 0, 2, 4, 6 and exits at 1, 3, 5, 7)
    for (let i = 0; i < 8; i += 2) {
      const entStr = day.shifts[i];
      const salStr = day.shifts[i + 1];
      
      const ent = timeToMin(entStr);
      const sal = timeToMin(salStr);
      
      if (ent !== null && sal !== null) {
        if (sal >= ent) {
          totalMins += (sal - ent);
        } else {
          // Wrap around midnight (unlikely but supported)
          totalMins += ((1440 - ent) + sal);
        }
      }
    }

    if (totalMins === 0) {
      return {
        hasHours: false,
        timeStr: "",
        h: 0,
        minDec: 0,
        totalDec: 0
      };
    }

    const h = Math.floor(totalMins / 60);
    const mins = totalMins % 60;
    const minDec = parseFloat((mins / 60).toFixed(4));
    // Match excel: HOUR(K) + MINUTE(K)/60
    const totalDec = parseFloat((h + minDec).toFixed(4));

    return {
      hasHours: true,
      timeStr: minToTimeStr(totalMins),
      h,
      minDec,
      totalDec
    };
  }

  // Group days into weeks, returning list of weeks with indices
  // Weeks run Monday to Sunday
  function groupWeeks(days) {
    const weeks = [];
    let currentWeek = [];
    
    days.forEach((day, index) => {
      currentWeek.push({ day, index });
      
      // If it's Sunday (0) or the last day in the array, close the week
      if (day.dayOfWeek === 0 || index === days.length - 1) {
        weeks.push(currentWeek);
        currentWeek = [];
      }
    });

    return weeks;
  }

  // --- Render Table & UI Metrics ---
  function renderTable() {
    payrollTbody.innerHTML = '';

    // Calculate daily durations
    const dayCalcs = daysData.map(d => calculateDayHours(d));
    
    // Group into weeks
    const weeks = groupWeeks(daysData);

    // Track weekly summaries
    const weeklySummaries = {};

    weeks.forEach((week, weekIdx) => {
      // Find Sunday row in this week (if any) or default to last day of week
      const sundayItem = week.find(w => w.day.dayOfWeek === 0) || week[week.length - 1];
      const sundayIdx = sundayItem.index;

      // 1. Calculate weekly total worked decimal hours (Sum of N)
      let weeklyTotalDec = 0;
      week.forEach(w => {
        weeklyTotalDec += dayCalcs[w.index].totalDec;
      });
      weeklyTotalDec = parseFloat(weeklyTotalDec.toFixed(4));

      // 2. Calculate weekly normal hours (Sum of standard hours for active days)
      let weeklyNorm = 0;
      week.forEach(w => {
        const d = w.day;
        if (d.dayOfWeek === 0 || d.isFeriado) {
          // Sunday and Holidays have 0 normal hours
          weeklyNorm += 0;
        } else {
          weeklyNorm += STANDARD_HOURS[d.dayOfWeek];
        }
      });
      weeklyNorm = parseFloat(weeklyNorm.toFixed(4));

      // 3. Calculate weekly extras (Total - Norm)
      let weeklyExtras = Math.max(0, weeklyTotalDec - weeklyNorm);
      weeklyExtras = parseFloat(weeklyExtras.toFixed(4));

      // Save summary for Sunday's index
      weeklySummaries[sundayIdx] = {
        total: weeklyTotalDec,
        norm: weeklyNorm,
        extras: weeklyExtras
      };
    });

    // Populate rows
    daysData.forEach((day, rIdx) => {
      const calc = dayCalcs[rIdx];
      const isSunday = day.dayOfWeek === 0;
      const isSaturday = day.dayOfWeek === 6;
      const isFeriado = day.isFeriado;
      
      // Weekly summary for this row
      const summary = weeklySummaries[rIdx];

      // Format shifts for mobile view
      let horarioText = "";
      const activeShifts = [];
      for (let s = 0; s < 8; s += 2) {
        const ent = day.shifts[s];
        const sal = day.shifts[s + 1];
        if (ent && sal) {
          activeShifts.push(`${ent}→${sal}`);
        }
      }
      horarioText = activeShifts.length > 0 ? activeShifts.join(" ") : "-";

      // Determine Row CSS class
      let rowClass = 'day-row';
      if (isSunday || isFeriado) {
        if (calc.hasHours) {
          rowClass += ' rest-day-worked'; // Red background
        } else {
          rowClass += ' rest-day-free'; // Yellow background
        }
      } else if (isSaturday) {
        rowClass += ' saturday-row';
      }

      const tr = document.createElement('tr');
      tr.className = rowClass;
      tr.dataset.index = rIdx;

      // Columns HTML structure
      let html = `
        <td class="center font-bold" style="font-size: 0.85rem;">${day.dayNum}</td>
        <td class="font-medium" style="text-transform: capitalize;">${day.dayName}</td>
        <td class="center col-horario-mobile" style="display: none; font-size: 0.75rem; white-space: nowrap;">${horarioText}</td>
      `;

      // Shifts Inputs (1ª to 4ª Ent/Sal)
      for (let s = 0; s < 8; s++) {
        html += `
          <td class="col-shift">
            <input type="text" 
              class="time-input" 
              placeholder="--:--" 
              value="${day.shifts[s]}" 
              data-day-idx="${rIdx}" 
              data-shift-idx="${s}"
              maxlength="5"
            >
          </td>
        `;
      }

      // Calculations Columns (K, L, M, N)
      html += `
        <td class="center font-semibold col-calc-details" style="background-color: rgba(0,0,0,0.01);">${calc.timeStr}</td>
        <td class="center col-calc-details" style="background-color: rgba(0,0,0,0.01);">${calc.h || ""}</td>
        <td class="center col-calc-details" style="background-color: rgba(0,0,0,0.01);">${calc.hasHours ? calc.minDec.toFixed(2) : ""}</td>
        <td class="center font-semibold" style="background-color: rgba(0,0,0,0.02);">${calc.hasHours ? calc.totalDec.toFixed(2) : "0.00"}</td>
      `;

      // Weekly summaries (O, P, Q) shown on Sundays (or end of week)
      if (summary) {
        html += `
          <td class="col-summary" style="background-color: #f1f5f9; text-align: center; font-weight: 700; padding: 4px;" title="Total Acumulado de la Semana">
            <span style="font-size: 0.6rem; display: block; color: #64748b; font-weight: 800; margin-bottom: 2px;">TOTAL SEM.</span>
            ${summary.total.toFixed(2)}
          </td>
          <td class="col-summary" style="background-color: #f1f5f9; text-align: center; padding: 4px;" title="Horas Base de la Semana">
            <span style="font-size: 0.6rem; display: block; color: #64748b; font-weight: 800; margin-bottom: 2px;">BASE SEM.</span>
            ${summary.norm.toFixed(2)}
          </td>
          <td class="col-summary" style="background-color: ${summary.extras > 0 ? '#fee2e2' : '#f1f5f9'}; color: ${summary.extras > 0 ? '#ef4444' : 'inherit'}; font-weight: 800; text-align: center; padding: 4px;" title="Horas Extras de la Semana (Total - Base)">
            <span style="font-size: 0.6rem; display: block; color: ${summary.extras > 0 ? '#b91c1c' : '#64748b'}; font-weight: 800; margin-bottom: 2px;">EXTRA SEM.</span>
            ${summary.extras.toFixed(2)}
          </td>
        `;
      } else {
        html += `
          <td class="col-summary" style="background-color: rgba(0,0,0,0.01);"></td>
          <td class="col-summary" style="background-color: rgba(0,0,0,0.01);"></td>
          <td class="col-summary" style="background-color: rgba(0,0,0,0.01);"></td>
        `;
      }

      // Comments & Feriado toggle button
      html += `
        <td class="col-comment">
          <input type="text" 
            class="comment-input" 
            placeholder="Motivo de horas extraordinarias..." 
            value="${day.comment || (isFeriado ? (FERIADOS[day.dateKey] || 'Feriado') : '')}"
            data-day-idx="${rIdx}"
          >
        </td>
        <td class="center" style="padding: 4px !important;">
          <div style="display: flex; gap: 4px; justify-content: center; align-items: center;">
            <button class="btn-toggle-feriado" data-day-idx="${rIdx}" title="${isFeriado ? 'Marcar como Laborable' : 'Marcar como Feriado'}" style="padding: 4px 6px; font-size: 0.75rem;">
              ${isFeriado ? '☀️ Lab' : '🎈 Fer'}
            </button>
            <button class="btn-delete-day" data-day-idx="${rIdx}" title="Limpiar Día" style="background-color: #fee2e2; color: #b91c1c; border: 1px solid #fecaca; padding: 4px 6px; border-radius: 4px; cursor: pointer; font-size: 0.75rem; line-height: 1;">
              🗑️
            </button>
          </div>
        </td>
      `;

      tr.innerHTML = html;
      payrollTbody.appendChild(tr);
    });

    // --- Totals Row and Cards Calculation ---
    let totalDecimalHours = 0; // Sum of N
    let totalNormalHours = 0;  // Sum of weekly P
    let totalExtraHours = 0;   // Sum of weekly Q
    let overtime100Hours = 0;  // Sunday and Holiday worked hours

    // Sum individual daily worked hours for decimal sum
    dayCalcs.forEach((c, idx) => {
      totalDecimalHours += c.totalDec;
      
      const day = daysData[idx];
      if (day.dayOfWeek === 0) {
        overtime100Hours += c.totalDec;
      } else if (day.isFeriado) {
        let holidayOvertimeMins = 0;
        for (let i = 4; i < 8; i += 2) {
          const ent = timeToMin(day.shifts[i]);
          const sal = timeToMin(day.shifts[i + 1]);
          if (ent && sal && sal >= ent) {
            holidayOvertimeMins += (sal - ent);
          }
        }
        overtime100Hours += minToHrs(holidayOvertimeMins);
      }
    });

    // Sum weekly columns from summaries
    Object.values(weeklySummaries).forEach(s => {
      totalNormalHours += s.norm;
      totalExtraHours += s.extras;
    });

    // 50% overtime is: total extra hours - 100% overtime hours (min 0)
    let overtime50Hours = Math.max(0, totalExtraHours - overtime100Hours);

    // Update Totals Cards
    metricTotalWorked.textContent = totalDecimalHours.toFixed(2);
    metricNormal.textContent = totalNormalHours.toFixed(2);
    metricExtras.textContent = totalExtraHours.toFixed(2);
    metricExtras50.textContent = overtime50Hours.toFixed(2);
    metricExtras100.textContent = overtime100Hours.toFixed(2);

    // Update Table Footer
    // Let's format the aggregate total worked minutes as HH:MM
    const totalMinutes = Math.round(totalDecimalHours * 60);
    footTpoPerm.textContent = minToTimeStr(totalMinutes);
    footHorasH.textContent = Math.floor(totalMinutes / 60);
    footHorasMin.textContent = ((totalMinutes % 60) / 60).toFixed(2);
    
    footHorasTotalDec.textContent = totalDecimalHours.toFixed(2);
    footSummaryTotal.textContent = totalDecimalHours.toFixed(2); // Since sum of N matches sum of O
    footSummaryNorm.textContent = totalNormalHours.toFixed(2);
    footSummaryExtras.textContent = totalExtraHours.toFixed(2);

    // Adjust footer colspans dynamically for mobile view
    const isMobile = window.innerWidth <= 768;
    const footLabel = document.querySelector('tfoot tr td:first-child');
    if (footLabel) {
      footLabel.colSpan = isMobile ? 3 : 10;
    }

    // --- Inputs Event Attachment ---
    // Handle time-input keyup, focusout
    const inputs = payrollTbody.querySelectorAll('.time-input');
    inputs.forEach(inp => {
      // Validate time input (add colon automatically, limit characters)
      inp.addEventListener('input', formatTimeInput);
      inp.addEventListener('change', handleTimeChange);
    });

    // Handle comment inputs change
    const commentInputs = payrollTbody.querySelectorAll('.comment-input');
    commentInputs.forEach(inp => {
      inp.addEventListener('change', handleCommentChange);
    });

    // Handle Feriado toggle button click
    const toggleFeriadoBtns = payrollTbody.querySelectorAll('.btn-toggle-feriado');
    toggleFeriadoBtns.forEach(btn => {
      btn.addEventListener('click', handleFeriadoToggle);
    });

    // Handle Delete Day button click
    const deleteDayBtns = payrollTbody.querySelectorAll('.btn-delete-day');
    deleteDayBtns.forEach(btn => {
      btn.addEventListener('click', handleDeleteDayClick);
    });
  }

  // Auto formats HH:MM while typing
  function formatTimeInput(e) {
    let val = e.target.value.replace(/[^0-9]/g, '');
    if (val.length > 4) val = val.substring(0, 4);
    
    if (val.length > 2) {
      e.target.value = val.substring(0, 2) + ':' + val.substring(2);
    } else {
      e.target.value = val;
    }
  }

  function handleTimeChange(e) {
    const dayIdx = parseInt(e.target.dataset.dayIdx);
    const shiftIdx = parseInt(e.target.dataset.shiftIdx);
    const val = e.target.value.trim();

    // Basic format validation
    if (val !== "") {
      const regex = /^([0-9]|0[0-9]|1[0-9]|2[0-3]):[0-5][0-9]$/;
      // If single digit hour was typed, format it properly, e.g. 9:30 -> 09:30
      let formattedVal = val;
      if (val.length === 4 && val.indexOf(':') === 1) {
        formattedVal = '0' + val;
      }

      if (regex.test(formattedVal)) {
        daysData[dayIdx].shifts[shiftIdx] = formattedVal;
        e.target.value = formattedVal;
      } else {
        // Invalid, clear input and alert
        showStatus("Formato de hora inválido. Use HH:MM (ej. 09:30 o 18:00)", "danger");
        daysData[dayIdx].shifts[shiftIdx] = "";
        e.target.value = "";
      }
    } else {
      daysData[dayIdx].shifts[shiftIdx] = "";
    }

    saveData();
    renderTable();
  }

  function handleCommentChange(e) {
    const dayIdx = parseInt(e.target.dataset.dayIdx);
    daysData[dayIdx].comment = e.target.value;
    saveData();
  }

  function handleFeriadoToggle(e) {
    const dayIdx = parseInt(e.target.dataset.dayIdx);
    daysData[dayIdx].isFeriado = !daysData[dayIdx].isFeriado;
    saveData();
    renderTable();
  }

  function handleDeleteDayClick(e) {
    const btn = e.target.closest('.btn-delete-day');
    if (!btn) return;
    const dayIdx = parseInt(btn.dataset.dayIdx);
    const day = daysData[dayIdx];
    if (confirm(`¿Está seguro de que desea limpiar todos los turnos y comentarios registrados para el día ${day.dayNum}?`)) {
      resetDayToDefault(day);
      saveData();
      renderTable();
      populateFormDaySelect();
      loadFormDayData();
      showStatus(`Día ${day.dayNum} restablecido a su horario base`, "success");
    }
  }

  // --- Show Status Messages ---
  function showStatus(text, type = "success") {
    statusMessage.textContent = text;
    statusMessage.style.display = 'block';
    statusMessage.className = `status-msg ${type}`;
    
    setTimeout(() => {
      statusMessage.style.display = 'none';
    }, 4000);
  }

  // --- Clear Period Data ---
  function clearCurrentPeriodData() {
    if (confirm("¿Está seguro de que desea limpiar todos los turnos y comentarios de este período?")) {
      daysData.forEach(d => {
        d.shifts = ['', '', '', '', '', '', '', ''];
        d.comment = '';
        d.isFeriado = false;
      });
      saveData();
      renderTable();
      showStatus("Datos del período limpiados con éxito.", "success");
    }
  }

  // --- Backup (Export/Import JSON) ---
  function exportBackup() {
    const backupData = {
      version: "1.0",
      year: currentYear,
      month: currentMonth,
      employeeInfo: employeeInfo,
      days: daysData.map(d => ({
        dateKey: d.dateKey,
        isFeriado: d.isFeriado,
        comment: d.comment,
        shifts: d.shifts
      }))
    };

    const blob = new Blob([JSON.stringify(backupData, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `Respaldo_Horas_Extras_${employeeInfo.name.replace(/\s+/g, '_')}_${currentYear}_${currentMonth + 1}.json`;
    a.click();
    URL.revokeObjectURL(url);
    showStatus("Copia de respaldo exportada.", "success");
  }

  function importBackup(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(evt) {
      try {
        const parsed = JSON.parse(evt.target.result);
        
        if (parsed.employeeInfo) {
          employeeInfo = parsed.employeeInfo;
          // Save global employee info
          localStorage.setItem('employee_global_info', JSON.stringify(employeeInfo));
          
          employeeNameInput.value = employeeInfo.name || '';
          employeeRoleInput.value = employeeInfo.role || '';
          employeeAreaInput.value = employeeInfo.area || '';
          employeeCardInput.value = employeeInfo.card || '';
          employeeManagerInput.value = employeeInfo.manager || '';
        }

        if (parsed.days) {
          // Merge in imported days
          daysData.forEach(d => {
            const matched = parsed.days.find(p => p.dateKey === d.dateKey);
            if (matched) {
              d.isFeriado = matched.isFeriado || false;
              d.comment = matched.comment || '';
              d.shifts = matched.shifts || ['', '', '', '', '', '', '', ''];
            }
          });
        }
        
        saveData();
        renderTable();
        showStatus("Copia de respaldo importada con éxito.", "success");
      } catch (err) {
        console.error(err);
        showStatus("Error al leer el archivo de respaldo.", "danger");
      }
    };
    reader.readAsText(file);
    e.target.value = ""; // Reset file input
  }

  // --- High-Fidelity ExcelJS Export ---
  async function exportToExcel() {
    try {
      showStatus('Generando planilla Excel...', 'success');
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Hoja1', {
        views: [{ showGridLines: true }]
      });

      // Column widths
      const columns = [
        { key: 'A', width: 13 },
        { key: 'B', width: 13 },
        { key: 'C', width: 13 },
        { key: 'D', width: 13 },
        { key: 'E', width: 13 },
        { key: 'F', width: 13 },
        { key: 'G', width: 13 },
        { key: 'H', width: 13 },
        { key: 'I', width: 13 },
        { key: 'J', width: 13 },
        { key: 'K', width: 13 },
        { key: 'L', width: 13 },
        { key: 'M', width: 13 },
        { key: 'N', width: 13 },
        { key: 'O', width: 13 },
        { key: 'P', width: 13 },
        { key: 'Q', width: 6.57 },
        { key: 'R', width: 118.43 }
      ];
      worksheet.columns = columns.map(c => ({ header: '', key: c.key, width: c.width }));

      // Helper to parse HH:MM to Excel serial time
      const parseTimeToExcel = (timeStr) => {
        if (!timeStr) return null;
        const parts = timeStr.split(':');
        const hours = parseInt(parts[0], 10);
        const minutes = parseInt(parts[1], 10) || 0;
        return (hours * 60 + minutes) / 1440;
      };

      const borderThinBox = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };

      const borderThinOpenRight = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' }
      };

      const borderMedium = {
        top: { style: 'medium' },
        left: { style: 'medium' },
        bottom: { style: 'medium' },
        right: { style: 'medium' }
      };

      const fillWhite = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF' } };
      const fillGreen = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'CCFFCC' } };
      const fillCyan = { type: 'pattern', pattern: 'solid', fgColor: { argb: '00FFFF' } };
      const fillIvory = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFCC' } };
      const fillRed = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0000' } };
      const fillYellow = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF00' } };
      const fillLightBlue = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'CCFFFF' } };
      const fillTan = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCC99' } };
      const fillPink = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF99CC' } };
      const fillPaleBlue = { type: 'pattern', pattern: 'solid', fgColor: { argb: '99CCFF' } };

      const fillRange = (rangeStartCol, rangeStartRow, rangeEndCol, rangeEndRow, fillObj) => {
        for (let row = rangeStartRow; row <= rangeEndRow; row++) {
          for (let col = rangeStartCol; col <= rangeEndCol; col++) {
            worksheet.getRow(row).getCell(col).fill = fillObj;
          }
        }
      };

      const setCellBorder = (cell, top, left, bottom, right) => {
        cell.border = {
          top: top ? { style: top } : undefined,
          left: left ? { style: left } : undefined,
          bottom: bottom ? { style: bottom } : undefined,
          right: right ? { style: right } : undefined
        };
      };
      const setBorder = setCellBorder;

      // Row 1
      worksheet.getCell('A1').value = 'EMELNOR S.A.';
      worksheet.getCell('A1').font = { name: 'Arial', size: 10, bold: true };
      worksheet.getCell('A1').fill = fillWhite;

      // Row 2
      worksheet.mergeCells('A2:R2');
      const cellA2 = worksheet.getCell('A2');
      cellA2.value = 'CONVENIO Y CONTROL DE HORAS EXTRAORDINARIAS';
      cellA2.font = { name: 'Arial', size: 14, bold: true };
      cellA2.alignment = { horizontal: 'center' };
      worksheet.getRow(2).height = 18;

      // Row 3
      for (let col = 1; col <= 18; col++) {
        const cell = worksheet.getRow(3).getCell(col);
        cell.font = { name: 'Arial', size: 14, bold: true };
        cell.alignment = { horizontal: 'center' };
      }
      worksheet.getRow(3).height = 18;

      // Row 4
      worksheet.mergeCells('A4:C4');
      worksheet.getCell('A4').value = 'En  Antofagasta a        ';
      worksheet.getCell('A4').font = { name: 'Arial', size: 12 };

      worksheet.mergeCells('D4:H4');
      const cellD4 = worksheet.getCell('D4');
      cellD4.value = daysData[0] ? daysData[0].dateObj : new Date();
      cellD4.font = { name: 'Arial', size: 12, bold: true };
      cellD4.numFmt = 'dd" de "mmmm" del "yyyy';
      cellD4.alignment = { horizontal: 'center' };
      fillRange(4, 4, 8, 4, fillGreen);

      worksheet.getCell('I4').value = 'se suscribe  el  presente convenio , mediante el cual el trabajador acuerda laborar horas   extraordinarias, de acuerdo a';
      worksheet.getCell('I4').font = { name: 'Arial', size: 12 };
      worksheet.getRow(4).height = 18;

      // Row 5
      worksheet.mergeCells('A5:R5');
      worksheet.getRow(5).height = 15.75;

      // Row 6
      const padZero = (n) => String(n).padStart(2, '0');
      const startD = daysData[0] ? daysData[0].dateObj : new Date();
      const endD = daysData[daysData.length - 1] ? daysData[daysData.length - 1].dateObj : new Date();
      const startStr = padZero(startD.getDate()) + '/' + padZero(startD.getMonth() + 1) + '/' + startD.getFullYear();
      const endStr = padZero(endD.getDate()) + '/' + padZero(endD.getMonth() + 1) + '/' + endD.getFullYear();

      worksheet.mergeCells('A6:R6');
      const cellA6 = worksheet.getCell('A6');
      cellA6.value = 'Periodo de  Vigencia del Convenio: ' + startStr + ' al ' + endStr;
      cellA6.font = { name: 'Arial', size: 12, bold: true };
      cellA6.alignment = { horizontal: 'center' };
      cellA6.numFmt = 'mmmm" del "yyyy';
      fillRange(1, 6, 18, 6, fillGreen);
      worksheet.getRow(6).height = 15.75;

      // Row 7
      worksheet.mergeCells('A7:C7');
      const cellA7 = worksheet.getCell('A7');
      cellA7.value = 'Funcionario(a):';
      cellA7.font = { name: 'Arial', size: 12, bold: true };
      cellA7.alignment = { horizontal: 'center' };
      cellA7.numFmt = 'mmm-yy';

      worksheet.mergeCells('D7:I7');
      const cellD7 = worksheet.getCell('D7');
      cellD7.value = employeeInfo.name;
      cellD7.font = { name: 'Arial', size: 12, bold: true };
      fillRange(4, 7, 9, 7, fillCyan);

      worksheet.mergeCells('P7:Q7');
      const cellP7 = worksheet.getCell('P7');
      cellP7.value = 'Cargo:';
      cellP7.font = { name: 'Arial', size: 12, bold: true };
      cellP7.alignment = { horizontal: 'right' };
      cellP7.numFmt = 'mmm-yy';
      
      worksheet.getCell('R7').value = employeeInfo.role;
      worksheet.getCell('R7').font = { name: 'Arial', size: 12 };
      worksheet.getCell('R7').numFmt = 'mmm-yy';
      worksheet.getRow(7).height = 15.75;

      // Row 8
      worksheet.getCell('A8').value = 'Nº Tarjeta:';
      worksheet.getCell('A8').font = { name: 'Arial', size: 12, bold: true };
      worksheet.getCell('A8').fill = fillWhite;
      worksheet.getCell('A8').numFmt = 'mmm-yy';
      worksheet.getCell('B8').numFmt = 'mmm-yy';
      worksheet.getCell('B8').font = { name: 'Arial', size: 12, bold: true };

      worksheet.getCell('C8').value = employeeInfo.card ? parseInt(employeeInfo.card, 10) : '';
      worksheet.getCell('C8').fill = fillGreen;
      worksheet.getCell('C8').numFmt = '#,##0';
      worksheet.getCell('C8').font = { name: 'Arial', size: 12 };
      worksheet.getCell('C8').alignment = { horizontal: 'center' };
      setCellBorder(worksheet.getCell('C8'), null, null, 'medium', null);

      worksheet.getCell('E8').value = 'Área:';
      worksheet.getCell('E8').font = { name: 'Arial', size: 10, bold: true };

      worksheet.mergeCells('F8:O8');
      const cellF8 = worksheet.getCell('F8');
      cellF8.value = employeeInfo.area;
      cellF8.font = { name: 'Arial', size: 12 };
      fillRange(6, 8, 15, 8, fillWhite);
      for (let col = 6; col <= 15; col++) {
        setCellBorder(worksheet.getRow(8).getCell(col), null, null, 'medium', null);
      }
      
      worksheet.mergeCells('P8:Q8');
      const cellP8 = worksheet.getCell('P8');
      cellP8.value = 'Jefe directo:';
      cellP8.font = { name: 'Arial', size: 12, bold: true };
      cellP8.alignment = { horizontal: 'right' };
      cellP8.numFmt = 'mmm-yy';
      setCellBorder(worksheet.getCell('P8'), null, null, 'medium', null);
      setCellBorder(worksheet.getCell('Q8'), null, null, 'medium', null);

      worksheet.getCell('R8').value = employeeInfo.manager;
      worksheet.getCell('R8').font = { name: 'Arial', size: 12 };
      worksheet.getCell('R8').numFmt = 'mmm-yy';
      worksheet.getRow(8).height = 16.5;

      // Row 9 Headers
      worksheet.getCell('C9').value = '1ª';
      worksheet.getCell('D9').value = '1ª';
      worksheet.getCell('E9').value = '2ª';
      worksheet.getCell('F9').value = '2ª';
      worksheet.getCell('G9').value = '3ª';
      worksheet.getCell('H9').value = '3ª';
      worksheet.getCell('I9').value = '4ª';
      worksheet.getCell('J9').value = '4ª';
      worksheet.getCell('K9').value = 'TPO.';
      
      worksheet.mergeCells('L9:N9');
      worksheet.getCell('L9').value = 'HORAS';

      worksheet.mergeCells('O9:Q9');
      worksheet.getCell('O9').value = 'Nº DE HORAS';

      // Row 10 Headers
      worksheet.mergeCells('A10:B10');
      worksheet.getCell('A10').value = 'DÍA';
      
      worksheet.getCell('C10').value = 'ENT.';
      worksheet.getCell('D10').value = 'SAL.';
      worksheet.getCell('E10').value = 'ENT.';
      worksheet.getCell('F10').value = 'SAL.';
      worksheet.getCell('G10').value = 'ENT.';
      worksheet.getCell('H10').value = 'SAL.';
      worksheet.getCell('I10').value = 'ENT.';
      worksheet.getCell('J10').value = 'SAL.';
      worksheet.getCell('K10').value = 'PERM.';

      worksheet.mergeCells('L10:N10');
      worksheet.getCell('L10').value = 'PERM.';

      worksheet.getCell('O10').value = 'Total';
      worksheet.getCell('P10').value = 'Norm.';
      worksheet.getCell('Q10').value = 'Extras';
      worksheet.getCell('R10').value = 'MOTIVO DE HORAS EXTRAORDINARIAS';

      // Apply colors to Headers row 9 & 10
      fillRange(3, 9, 4, 9, fillTan);
      fillRange(3, 10, 4, 10, fillTan);
      fillRange(5, 9, 6, 9, fillLightBlue);
      fillRange(5, 10, 6, 10, fillLightBlue);
      fillRange(7, 9, 8, 9, fillPink);
      fillRange(7, 10, 8, 10, fillPink);
      fillRange(9, 9, 10, 9, fillPaleBlue);
      fillRange(9, 10, 10, 10, fillPaleBlue);
      
      worksheet.getCell('Q10').fill = fillIvory;

      // Format Rows 9 and 10 fonts/alignments
      const colNames = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R'];
      colNames.forEach(col => {
        const c9 = worksheet.getCell(col + '9');
        const c10 = worksheet.getCell(col + '10');

        const isAB9 = col === 'A' || col === 'B';
        c9.font = { name: 'Arial', size: 10, bold: !isAB9 && col !== 'R' };
        c10.font = { name: 'Arial', size: 10, bold: true };

        if (!isAB9 && col !== 'R') {
          c9.alignment = { horizontal: 'center' };
        }
        c10.alignment = { horizontal: 'center' };
      });

      // Precise Row 9 borders
      setBorder(worksheet.getCell('A9'), 'medium', 'medium', null, null);
      setBorder(worksheet.getCell('B9'), 'medium', null, null, 'medium');
      for (let c = 3; c <= 11; c++) {
        setBorder(worksheet.getRow(9).getCell(c), 'medium', 'medium', null, null);
      }
      setBorder(worksheet.getCell('L9'), 'medium', 'medium', null, 'medium');
      setBorder(worksheet.getCell('O9'), 'medium', 'medium', 'medium', 'medium');
      setBorder(worksheet.getCell('R9'), 'medium', 'medium', null, 'medium');

      // Precise Row 10 borders
      setBorder(worksheet.getCell('A10'), null, 'medium', 'medium', 'medium');
      for (let c = 3; c <= 11; c++) {
        setBorder(worksheet.getRow(10).getCell(c), null, 'medium', 'medium', null);
      }
      setBorder(worksheet.getCell('L10'), null, 'medium', 'medium', 'medium');
      
      setBorder(worksheet.getCell('O10'), 'medium', 'medium', 'medium', 'medium');
      setBorder(worksheet.getCell('P10'), 'medium', 'medium', 'medium', 'medium');
      setBorder(worksheet.getCell('Q10'), 'medium', 'medium', 'medium', 'medium');
      setBorder(worksheet.getCell('R10'), null, 'medium', 'medium', 'medium');

      worksheet.getRow(9).height = 15.75;
      worksheet.getRow(10).height = 15.75;

      // Populate data rows 11 to 42
      let r = 11;
      let firstRowOfWeek = 11;
      let weekBaseHours = 0;

      const totalRowsToGenerate = 32;
      for (let index = 0; index < totalRowsToGenerate; index++) {
        const isPadding = (index >= daysData.length);
        let diaNum = '';
        let diaNombre = '';
        let esDom = false;
        let esSab = false;
        let esFer = false;
        let dayData = null;

        if (!isPadding) {
          dayData = daysData[index];
          diaNum = dayData.dayNum;
          diaNombre = dayData.dayName;
          esDom = dayData.dayOfWeek === 0;
          esSab = dayData.dayOfWeek === 6;
          esFer = dayData.isFeriado;
        }

        const cellA = worksheet.getCell('A' + r);
        const cellB = worksheet.getCell('B' + r);
        const cellC = worksheet.getCell('C' + r);
        const cellD = worksheet.getCell('D' + r);
        const cellE = worksheet.getCell('E' + r);
        const cellF = worksheet.getCell('F' + r);
        const cellG = worksheet.getCell('G' + r);
        const cellH = worksheet.getCell('H' + r);
        const cellI = worksheet.getCell('I' + r);
        const cellJ = worksheet.getCell('J' + r);
        const cellK = worksheet.getCell('K' + r);
        const cellL = worksheet.getCell('L' + r);
        const cellM = worksheet.getCell('M' + r);
        const cellN = worksheet.getCell('N' + r);
        const cellR = worksheet.getCell('R' + r);

        cellA.value = diaNum;
        cellB.value = diaNombre ? diaNombre.toLowerCase() : '';

        if (!isPadding && dayData) {
          cellC.value = dayData.shifts[0] ? parseTimeToExcel(dayData.shifts[0]) : '';
          cellD.value = dayData.shifts[1] ? parseTimeToExcel(dayData.shifts[1]) : '';
          cellE.value = dayData.shifts[2] ? parseTimeToExcel(dayData.shifts[2]) : '';
          cellF.value = dayData.shifts[3] ? parseTimeToExcel(dayData.shifts[3]) : '';
          cellG.value = dayData.shifts[4] ? parseTimeToExcel(dayData.shifts[4]) : '';
          cellH.value = dayData.shifts[5] ? parseTimeToExcel(dayData.shifts[5]) : '';
          cellI.value = dayData.shifts[6] ? parseTimeToExcel(dayData.shifts[6]) : '';
          cellJ.value = dayData.shifts[7] ? parseTimeToExcel(dayData.shifts[7]) : '';
          cellR.value = dayData.comment || (dayData.isFeriado ? (FERIADOS[dayData.dateKey] || 'Feriado') : '');
        } else {
          cellC.value = '';
          cellD.value = '';
          cellE.value = '';
          cellF.value = '';
          cellG.value = '';
          cellH.value = '';
          cellI.value = '';
          cellJ.value = '';
          cellR.value = '';
        }

        cellK.value = { formula: '(' + 'D' + r + '-' + 'C' + r + ')+(' + 'F' + r + '-' + 'E' + r + ')+(' + 'H' + r + '-' + 'G' + r + ')+(' + 'J' + r + '-' + 'I' + r + ')' };
        cellL.value = { formula: 'HOUR(' + 'K' + r + ')' };
        cellM.value = { formula: '(MINUTE(' + 'K' + r + '))/60' };
        cellN.value = { formula: '+L' + r + '+M' + r };

        // Apply numFmt
        cellC.numFmt = 'h:mm';
        cellD.numFmt = 'h:mm';
        cellE.numFmt = 'h:mm';
        cellF.numFmt = 'h:mm';
        cellG.numFmt = 'h:mm';
        cellH.numFmt = 'h:mm';
        cellI.numFmt = 'h:mm';
        cellJ.numFmt = 'h:mm';
        cellK.numFmt = 'h:mm';
        cellL.numFmt = '#,##0;[Red]\(#,##0\)';
        cellM.numFmt = '#,##0.00;[Red]\(#,##0.00\)';
        cellN.numFmt = '#,##0.00;[Red]\(#,##0.00\)';

        // Arial 10pt regular for all cell rows
        const rowCells = [cellA, cellB, cellC, cellD, cellE, cellF, cellG, cellH, cellI, cellJ, cellK, cellL, cellM, cellN, cellR];
        rowCells.forEach(cell => {
          cell.font = { name: 'Arial', size: 10 };
        });
        cellN.font = { name: 'Arial', size: 10, bold: true };

        // Precise borders for daily row
        const isSundayOrHoliday = esDom || esFer;
        if (isSundayOrHoliday) {
          const boxCols = [cellA, cellB, cellJ, cellK, cellR];
          boxCols.forEach(c => c.border = borderThinBox);
          const openRightCols = [cellC, cellD, cellE, cellF, cellG, cellH, cellI, cellL, cellM, cellN];
          openRightCols.forEach(c => c.border = borderThinOpenRight);
        } else {
          const boxCols = [cellA, cellB, cellG, cellH, cellI, cellJ, cellK, cellR];
          boxCols.forEach(c => c.border = borderThinBox);
          const openRightCols = [cellC, cellD, cellE, cellF, cellL, cellM, cellN];
          openRightCols.forEach(c => c.border = borderThinOpenRight);
        }

        if (r === 35 || r === 42) {
          cellR.border = {
            top: undefined,
            bottom: undefined,
            left: undefined,
            right: undefined
          };
        }

        cellB.alignment = { horizontal: 'left' };
        const centerCells = [cellC, cellD, cellE, cellF, cellG, cellH, cellI, cellJ, cellK];
        centerCells.forEach(c => c.alignment = { horizontal: 'center' });

        // Fills
        cellA.fill = fillWhite;

        let fill = null;
        const workedAny = !isPadding && dayData && dayData.shifts[0] && dayData.shifts[1];

        if (isSundayOrHoliday) {
          if (workedAny) {
            fill = fillRed;
          } else {
            fill = fillYellow;
          }
        } else {
          fill = fillLightBlue;
        }

        if (isSundayOrHoliday) {
          // Fill C to Q with red/yellow on Sundays/Holidays
          for (let col = 3; col <= 17; col++) {
            worksheet.getRow(r).getCell(col).fill = fill;
          }
        } else {
          // Fill C to J with light blue, K to N with white
          for (let col = 3; col <= 10; col++) {
            worksheet.getRow(r).getCell(col).fill = fill;
          }
          for (let col = 11; col <= 14; col++) {
            worksheet.getRow(r).getCell(col).fill = fillWhite;
          }
        }

        if (!isPadding && dayData) {
          if (isSundayOrHoliday) {
            weekBaseHours += 0;
          } else {
            weekBaseHours += STANDARD_HOURS[dayData.dayOfWeek];
          }
        } else {
          weekBaseHours += 0;
        }

        const isLastRow = (index === totalRowsToGenerate - 1);
        if (esDom || isLastRow) {
          const cellO = worksheet.getCell('O' + r);
          const cellP = worksheet.getCell('P' + r);
          const cellQ = worksheet.getCell('Q' + r);

          cellO.value = { formula: 'SUM(N' + firstRowOfWeek + ':N' + r + ')' };
          cellP.value = weekBaseHours;
          cellQ.value = { formula: 'SUM(O' + r + '-P' + r + ')' };

          cellO.numFmt = '0.00';
          cellP.numFmt = '0.00';
          cellQ.numFmt = '0.00';

          const summaryCells = [cellO, cellP, cellQ];
          summaryCells.forEach(c => {
            c.font = { name: 'Arial', size: 10, bold: true };
            if (c === cellQ) {
              c.alignment = undefined;
            } else {
              c.alignment = { horizontal: 'center' };
            }
            if (isSundayOrHoliday) {
              c.border = {
                left: { style: 'thin' },
                right: { style: 'thin' },
                bottom: { style: 'thin' }
              };
            } else {
              c.border = borderThinBox;
            }
          });

          firstRowOfWeek = r + 1;
          weekBaseHours = 0;
        } else {
          const cellO = worksheet.getCell('O' + r);
          const cellP = worksheet.getCell('P' + r);
          const cellQ = worksheet.getCell('Q' + r);
          const emptyCells = [cellO, cellP, cellQ];
          emptyCells.forEach(c => {
            c.value = '';
            c.font = { name: 'Arial', size: 10 };
            c.border = borderThinBox;
          });
        }

        r++;
      }

      // Totals at bottom
      // Row 43 (empty Sunday row in template)
      for (let col = 1; col <= 18; col++) {
        const colName = colNames[col - 1];
        const cell = worksheet.getCell(colName + '43');
        cell.font = { name: 'Arial', size: 10 };
        cell.fill = fillYellow;
      }
      worksheet.getCell('A43').border = borderThinBox;

      // Row 44 (empty spacer row)
      worksheet.getRow(44).height = 15.75;
      for (let col = 1; col <= 18; col++) {
        const colName = colNames[col - 1];
        const cell = worksheet.getCell(colName + '44');
        cell.font = { name: 'Arial', size: 10 };
      }

      // Row 45 (Totals)
      worksheet.getRow(45).height = 15.75;
      for (let col = 1; col <= 18; col++) {
        const colName = colNames[col - 1];
        const cell = worksheet.getCell(colName + '45');
        cell.font = { name: 'Arial', size: 10 };
      }
      
      const cellN45 = worksheet.getCell('N45');
      cellN45.value = { formula: 'SUM(N7:N42)' };
      cellN45.font = { name: 'Arial', size: 10, bold: true };
      cellN45.numFmt = '0.00';
      cellN45.border = borderMedium;

      const cellO45 = worksheet.getCell('O45');
      cellO45.value = { formula: 'SUM(O11:O42)' };
      cellO45.font = { name: 'Arial', size: 10, bold: true };
      cellO45.numFmt = '0.00';
      cellO45.alignment = { horizontal: 'center' };
      cellO45.border = borderMedium;

      const cellP45 = worksheet.getCell('P45');
      cellP45.value = { formula: 'SUM(P11:P42)' };
      cellP45.font = { name: 'Arial', size: 10, bold: true };
      cellP45.numFmt = '0.00';
      cellP45.alignment = { horizontal: 'center' };
      cellP45.border = borderMedium;

      const cellQ45 = worksheet.getCell('Q45');
      cellQ45.value = { formula: 'SUM(Q13:Q42)' };
      cellQ45.font = { name: 'Arial', size: 10, bold: true };
      cellQ45.numFmt = '0.00';
      cellQ45.alignment = { horizontal: 'center' };
      cellQ45.border = borderMedium;

      setCellBorder(worksheet.getCell('A45'), 'medium', null, null, null);
      for (let col = 3; col <= 13; col++) {
        setCellBorder(worksheet.getRow(45).getCell(col), 'medium', null, null, null);
      }
      worksheet.getCell('K45').alignment = { horizontal: 'center' };

      // Row 46 (Percentage & Extras)
      worksheet.getRow(46).height = 15.75;
      for (let col = 1; col <= 18; col++) {
        const colName = colNames[col - 1];
        const cell = worksheet.getCell(colName + '46');
        cell.font = { name: 'Arial', size: 10 };
      }
      
      worksheet.getCell('N46').font = { name: 'Arial', size: 10, bold: true };
      worksheet.getCell('O46').font = { name: 'Arial', size: 10, bold: true };
      worksheet.getCell('P46').font = { name: 'Arial', size: 10, bold: true };
      worksheet.getCell('Q46').font = { name: 'Arial', size: 10, bold: true };
      
      worksheet.getCell('N46').numFmt = '0.00';
      worksheet.getCell('O46').numFmt = '0.00';
      worksheet.getCell('P46').numFmt = '0.00';
      worksheet.getCell('Q46').numFmt = '0.00';

      worksheet.getCell('K46').alignment = { horizontal: 'center' };
      worksheet.getCell('O46').alignment = { horizontal: 'center' };
      worksheet.getCell('P46').alignment = { horizontal: 'center' };
      worksheet.getCell('Q46').alignment = { horizontal: 'center' };

      worksheet.getCell('Q46').fill = fillYellow;
      worksheet.getCell('Q46').border = borderMedium;

      const cellR46 = worksheet.getCell('R46');
      cellR46.value = 0.5;
      cellR46.font = { name: 'Arial', size: 10, bold: true };
      cellR46.fill = fillYellow;
      cellR46.alignment = { horizontal: 'center' };
      cellR46.numFmt = '0%';
      cellR46.border = borderMedium;

      // Row 47
      for (let col = 1; col <= 18; col++) {
        const colName = colNames[col - 1];
        const cell = worksheet.getCell(colName + '47');
        cell.font = { name: 'Arial', size: 10 };
      }
      worksheet.getCell('L47').alignment = { horizontal: 'center' };
      worksheet.getCell('Q47').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCC00' } };
      
      for (let col = 1; col <= 17; col++) {
        const cell = worksheet.getRow(47).getCell(col);
        cell.border = { top: { style: 'medium' } };
      }
      worksheet.getCell('A47').border = { top: { style: 'medium' }, left: { style: 'medium' } };

      // Row 48
      for (let col = 1; col <= 18; col++) {
        const colName = colNames[col - 1];
        const cell = worksheet.getCell(colName + '48');
        cell.font = { name: 'Arial', size: 10 };
      }
      worksheet.getCell('A48').value = ' Primero: A través del presente documento, el  trabajador  quien  suscribe  el  presente convenio y   autorización de horas extraordinarias,    acepta  y  reconoce  que las horas   ';
      worksheet.getCell('A48').alignment = { horizontal: 'left' };
      worksheet.getCell('B48').alignment = { horizontal: 'left' };
      worksheet.getCell('L48').alignment = { horizontal: 'center' };
      worksheet.getCell('R48').alignment = { horizontal: 'left' };
      worksheet.getCell('Q48').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCC00' } };

      // Row 49
      for (let col = 1; col <= 18; col++) {
        const colName = colNames[col - 1];
        const cell = worksheet.getCell(colName + '49');
        cell.font = { name: 'Arial', size: 10 };
      }
      worksheet.getCell('A49').value = 'extraordinarias aquí expuestas  son  las  únicas   horas trabajadas  con  autorización  del  empleador  y consecuentemente  reconoce que son las que deben  ser canceladas como tal.';
      worksheet.getCell('A49').alignment = { horizontal: 'left' };
      worksheet.getCell('L49').alignment = { horizontal: 'center' };
      worksheet.getCell('Q49').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCC00' } };

      // Row 50
      for (let col = 1; col <= 18; col++) {
        const colName = colNames[col - 1];
        const cell = worksheet.getCell(colName + '50');
        cell.font = { name: 'Arial', size: 10 };
      }
      worksheet.getCell('A50').value = 'Segundo: Los abajo firmantes declaran que las horas de entrada y salida señaladas en el presente documento, son fiel reflejo de lo registrado en las tarjetas de control de asistencia utilizadas formal y oficialmente por el Empleador,';
      worksheet.getCell('A50').font = { name: 'Arial', size: 10, bold: true };
      for (let col = 1; col <= 17; col++) {
        worksheet.getRow(50).getCell(col).alignment = { horizontal: 'left' };
      }
      worksheet.getCell('Q50').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCC00' } };
      worksheet.getCell('A50').border = { left: { style: 'medium' } };

      // Row 51
      for (let col = 1; col <= 18; col++) {
        const colName = colNames[col - 1];
        const cell = worksheet.getCell(colName + '51');
        cell.font = { name: 'Arial', size: 10 };
      }
      worksheet.getCell('A51').value = 'expresando su conformidad con los cálculos y demás datos indicados. ';
      worksheet.getCell('L51').alignment = { horizontal: 'center' };
      worksheet.getCell('Q51').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCC00' } };
      worksheet.getCell('A51').border = { left: { style: 'medium' } };

      // Row 52
      for (let col = 1; col <= 18; col++) {
        const colName = colNames[col - 1];
        const cell = worksheet.getCell(colName + '52');
        cell.font = { name: 'Arial', size: 10 };
      }
      worksheet.getCell('R52').value = '________________________________';
      worksheet.getCell('L52').alignment = { horizontal: 'center' };
      worksheet.getCell('Q52').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCC00' } };
      worksheet.getCell('A52').border = { left: { style: 'medium' } };

      // Row 53
      worksheet.getRow(53).height = 15.75;
      for (let col = 1; col <= 18; col++) {
        const colName = colNames[col - 1];
        const cell = worksheet.getCell(colName + '53');
        cell.font = { name: 'Arial', size: 10 };
      }
      worksheet.getCell('R53').value = 'VºBº Jefe Depto. Recurso Humano';
      worksheet.getCell('R53').font = { name: 'Arial', size: 10, bold: true };
      worksheet.getCell('L53').alignment = { horizontal: 'center' };
      worksheet.getCell('R53').alignment = { horizontal: 'left' };
      worksheet.getCell('Q53').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCC00' } };
      worksheet.getCell('A53').border = { left: { style: 'medium' } };

      // Row 54
      for (let col = 1; col <= 18; col++) {
        const colName = colNames[col - 1];
        const cell = worksheet.getCell(colName + '54');
        cell.font = { name: 'Arial', size: 10 };
      }
      worksheet.getCell('L54').alignment = { horizontal: 'center' };
      worksheet.getCell('Q54').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCC00' } };
      worksheet.getCell('A54').border = { left: { style: 'medium' } };
      worksheet.getCell('R54').border = { top: { style: 'medium' }, right: { style: 'medium' } };

      // Row 55
      for (let col = 1; col <= 18; col++) {
        const colName = colNames[col - 1];
        const cell = worksheet.getCell(colName + '55');
        cell.font = { name: 'Arial', size: 10 };
      }
      worksheet.getCell('B55').value = 'Firma Funcionario';
      worksheet.getCell('J55').value = 'Firma Jefe Directo';
      worksheet.getCell('R55').alignment = { horizontal: 'left' };
      worksheet.getCell('Q55').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCC00' } };
      
      worksheet.getCell('B55').font = { name: 'Arial', size: 10, bold: true };
      worksheet.getCell('C55').font = { name: 'Arial', size: 10, bold: true };
      worksheet.getCell('D55').font = { name: 'Arial', size: 10, bold: true };
      worksheet.getCell('J55').font = { name: 'Arial', size: 10, bold: true };
      worksheet.getCell('K55').font = { name: 'Arial', size: 10, bold: true };
      worksheet.getCell('L55').font = { name: 'Arial', size: 10, bold: true };
      worksheet.getCell('M55').font = { name: 'Arial', size: 10, bold: true };
      worksheet.getCell('N55').font = { name: 'Arial', size: 10, bold: true };
      worksheet.getCell('O55').font = { name: 'Arial', size: 10, bold: true };

      worksheet.getCell('B55').border = { top: { style: 'thin' } };
      worksheet.getCell('C55').border = { top: { style: 'thin' } };
      worksheet.getCell('D55').border = { top: { style: 'thin' } };
      worksheet.getCell('J55').border = { top: { style: 'thin' } };
      worksheet.getCell('K55').border = { top: { style: 'thin' } };
      worksheet.getCell('L55').border = { top: { style: 'thin' } };
      worksheet.getCell('M55').border = { top: { style: 'thin' } };
      worksheet.getCell('N55').border = { top: { style: 'thin' } };
      worksheet.getCell('O55').border = { top: { style: 'thin' } };

      worksheet.getCell('A55').border = { left: { style: 'medium' } };
      worksheet.getCell('R55').border = { right: { style: 'medium' } };

      // Row 56
      for (let col = 1; col <= 18; col++) {
        const colName = colNames[col - 1];
        const cell = worksheet.getCell(colName + '56');
        cell.font = { name: 'Arial', size: 10 };
      }
      worksheet.getCell('K56').alignment = { horizontal: 'center' };
      worksheet.getCell('R56').font = { name: 'Arial', size: 10, bold: true };
      worksheet.getCell('R56').alignment = { horizontal: 'left' };
      worksheet.getCell('A56').border = { left: { style: 'medium' } };
      worksheet.getCell('R56').border = { right: { style: 'medium' } };

      // Row heights for data rows (rows 11 to 42)
      for (let rowIdx = 11; rowIdx <= 42; rowIdx++) {
        worksheet.getRow(rowIdx).height = 15.75;
      }
      // Medium bottom divider border on row 56
      for (let col = 1; col <= 18; col++) {
        const cell = worksheet.getRow(56).getCell(col);
        cell.border = {
          top: cell.border ? cell.border.top : undefined,
          left: cell.border ? cell.border.left : undefined,
          right: cell.border ? cell.border.right : undefined,
          bottom: { style: 'medium' }
        };
      }

      // Write Excel File Buffer & Download
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = window.URL.createObjectURL(blob);
      
      const anchor = document.createElement('a');
      anchor.href = url;
      anchor.download = 'Horas_Extras_' + employeeInfo.name.replace(/\s+/g, '_') + '_' + currentYear + '_' + (currentMonth + 1) + '.xlsx';
      anchor.click();
      
      window.URL.revokeObjectURL(url);
      showStatus('Planilla Excel exportada con éxito.', 'success');
    } catch (error) {
      console.error(error);
      showStatus('Error al exportar a Excel. Intente nuevamente.', 'danger');
    }
  }

  function getMockDataForMay2026() {
    return {
      days: [
        { dateKey: "2026-04-16", shifts: ['09:30', '13:30', '15:00', '19:00', '', '', '', ''], comment: "" },
        { dateKey: "2026-04-17", shifts: ['09:30', '13:30', '15:00', '19:00', '', '', '', ''], comment: "" },
        { dateKey: "2026-04-18", shifts: ['10:00', '15:20', '', '', '', '', '', ''], comment: "hasta las 15:20 asistiendo macrena saavedra problema conexión casa" },
        { dateKey: "2026-04-19", shifts: ['09:30', '15:30', '', '', '', '', '', ''], comment: "asistiendo de las 9:30 a los usuarios de rendic problema de conexión, hasta las 15:30" },
        { dateKey: "2026-04-20", shifts: ['09:30', '13:30', '15:00', '19:30', '', '', '', ''], comment: "" },
        { dateKey: "2026-04-21", shifts: ['09:30', '13:30', '15:00', '21:20', '', '', '', ''], comment: "asistiendo usuario denizard gallardo conexión casa, hasta las 21:20 luego marcelo cruz, cambio de ips" },
        { dateKey: "2026-04-22", shifts: ['09:30', '13:30', '15:00', '18:30', '', '', '', ''], comment: "" },
        { dateKey: "2026-04-23", shifts: ['09:30', '13:30', '13:30', '20:30', '', '', '', ''], comment: "de corrido asistiendo instalacion matta, conexión salon, luego hasta las 20:30 asistiendo alejandro ahumada proeblema coenxion milenium" },
        { dateKey: "2026-04-24", shifts: ['09:30', '13:30', '15:00', '19:00', '', '', '', ''], comment: "" },
        { dateKey: "2026-04-25", shifts: ['10:00', '12:30', '', '', '', '', '', ''], comment: "" },
        { dateKey: "2026-04-26", shifts: ['', '', '', '', '', '', '', ''], comment: "" },
        { dateKey: "2026-04-27", shifts: ['09:30', '13:30', '15:00', '19:30', '', '', '', ''], comment: "" },
        { dateKey: "2026-04-28", shifts: ['09:30', '13:30', '15:00', '18:30', '', '', '', ''], comment: "" },
        { dateKey: "2026-04-29", shifts: ['09:30', '13:30', '15:00', '18:30', '', '', '', ''], comment: "" },
        { dateKey: "2026-04-30", shifts: ['09:30', '13:30', '15:00', '19:00', '', '', '', ''], comment: "" },
        { dateKey: "2026-05-01", shifts: ['09:30', '13:30', '15:00', '19:00', '', '', '', ''], comment: "" },
        { dateKey: "2026-05-02", shifts: ['10:00', '12:30', '', '', '', '', '', ''], comment: "" },
        { dateKey: "2026-05-03", shifts: ['', '', '', '', '', '', '', ''], comment: "" },
        { dateKey: "2026-05-04", shifts: ['09:30', '13:30', '15:00', '19:30', '', '', '', ''], comment: "" },
        { dateKey: "2026-05-05", shifts: ['09:30', '13:30', '15:00', '21:10', '', '', '', ''], comment: "hasta las 21:10 asistiendo denizard gallardo probelma equipo arica, valpo" },
        { dateKey: "2026-05-06", shifts: ['09:30', '13:30', '15:00', '18:30', '', '', '', ''], comment: "" },
        { dateKey: "2026-05-07", shifts: ['09:30', '13:30', '15:00', '20:10', '', '', '', ''], comment: "hasta las 20:10 asistiendo eduardo campos, luego alejandro ahumada probelma conecion milenium rendic" },
        { dateKey: "2026-05-08", shifts: ['09:30', '13:30', '15:00', '19:00', '', '', '', ''], comment: "" },
        { dateKey: "2026-05-09", shifts: ['10:00', '12:30', '', '', '', '', '', ''], comment: "" },
        { dateKey: "2026-05-10", shifts: ['15:20', '17:20', '', '', '', '', '', ''], comment: "hasta las 17:20 asistiendo usuarios corte electrico rendic" },
        { dateKey: "2026-05-11", shifts: ['09:30', '13:30', '15:00', '19:30', '', '', '', ''], comment: "" },
        { dateKey: "2026-05-12", shifts: ['09:30', '13:30', '15:00', '18:30', '', '', '', ''], comment: "" },
        { dateKey: "2026-05-13", shifts: ['09:30', '13:30', '15:00', '18:30', '', '', '', ''], comment: "" },
        { dateKey: "2026-05-14", shifts: ['09:30', '13:30', '15:00', '19:00', '', '', '', ''], comment: "" },
        { dateKey: "2026-05-15", shifts: ['09:30', '13:30', '15:00', '19:00', '', '', '', ''], comment: "" },
        { dateKey: "2026-05-16", shifts: ['10:00', '12:30', '', '', '', '', '', ''], comment: "" },
        { dateKey: "2026-05-17", shifts: ['', '', '', '', '', '', '', ''], comment: "" }
      ]
    };
  }
  
  initApp();
});
