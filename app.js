// Local Hours-Extras Application Logic
// Handles calculations, localStorage, UI updates, and ExcelJS export.

document.addEventListener('DOMContentLoaded', () => {
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
  let currentYear = 2026;
  let currentMonth = 4; // Mayo (0-indexed)
  let completeWeek = true;
  let daysData = []; // Array of day objects
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
  const formTicket = document.getElementById('form-ticket');
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
  initApp();

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
    loadData();
  }

  // --- Data Management (Local Storage) ---
  function getStorageKey() {
    return `horas_extras_data_${currentYear}_${currentMonth}`;
  }

  function loadData() {
    const key = getStorageKey();
    const stored = localStorage.getItem(key);
    
    // Always load global employee details if available
    const savedEmployee = localStorage.getItem('employee_global_info');
    if (savedEmployee) {
      employeeInfo = JSON.parse(savedEmployee);
      employeeNameInput.value = employeeInfo.name || '';
      employeeRoleInput.value = employeeInfo.role || '';
      employeeAreaInput.value = employeeInfo.area || '';
      employeeCardInput.value = employeeInfo.card || '';
      employeeManagerInput.value = employeeInfo.manager || '';
    }

    // Generate date sequence
    const dates = generatePeriodDates(currentYear, currentMonth, completeWeek);
    
    // Show period range text
    if (dates.length > 0) {
      const firstDate = dates[0].dateObj;
      const lastDate = dates[dates.length - 1].dateObj;
      dateRangeDisplay.textContent = `${formatDateStr(firstDate)} al ${formatDateStr(lastDate)}`;
    }

    // Initialize daysData empty
    generateEmptyDays(dates);

    // If Firebase is initialized, start cloud sync. Otherwise, load local storage
    if (isFirebaseInitialized) {
      setupFirebaseSync();
    } else {
      loadLocalStorageBackup(stored, dates);
    }
  }

  function generateEmptyDays(dates) {
    daysData = dates.map(d => {
      const dayOfWeek = d.dayOfWeek;
      let shifts = ['', '', '', '', '', '', '', ''];
      
      if (dayOfWeek !== 0) { // Not Sunday
        const start = getStandardStartTime(dayOfWeek);
        const end = getStandardEndTime(dayOfWeek);
        if (dayOfWeek === 6) {
          shifts = [start, end, '', '', '', '', '', ''];
        } else {
          shifts = [start, '13:30', '15:00', end, '', '', '', ''];
        }
      }
      
      return {
        dateKey: d.dateKey,
        dayNum: d.dayNum,
        dayName: d.dayName,
        dayOfWeek: dayOfWeek,
        dateObj: d.dateObj,
        isFeriado: false,
        comment: '',
        shifts: shifts
      };
    });
  }

  function saveData() {
    saveLocalOnly();
    saveToFirebaseOnly();
  }

  function saveLocalOnly() {
    const key = getStorageKey();
    const dataToSave = {
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
    localStorage.setItem(key, JSON.stringify(dataToSave));
    localStorage.setItem('employee_global_info', JSON.stringify(employeeInfo));
  }

  function saveToFirebaseOnly() {
    if (!isFirebaseInitialized) return;
    const sanitizedName = sanitizeFirebaseKey(employeeInfo.name || "Anonimo");
    const dbPath = `horas_extras/${sanitizedName}/${currentYear}_${currentMonth}`;
    
    setSyncState('syncing');
    
    db.ref(dbPath).set({
      year: currentYear,
      month: currentMonth,
      employeeInfo: employeeInfo,
      days: daysData.map(d => ({
        dateKey: d.dateKey,
        isFeriado: d.isFeriado,
        comment: d.comment,
        shifts: d.shifts
      }))
    })
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

    const sanitizedName = sanitizeFirebaseKey(employeeInfo.name || "Anonimo");
    const dbPath = `horas_extras/${sanitizedName}/${currentYear}_${currentMonth}`;
    firebaseListener = db.ref(dbPath);

    setSyncState('syncing');

    firebaseListener.on('value', (snapshot) => {
      const val = snapshot.val();
      if (val) {
        // Sync from Firebase
        daysData = daysData.map(d => {
          const matched = val.days ? val.days.find(p => p.dateKey === d.dateKey) : null;
          return {
            ...d,
            isFeriado: matched ? matched.isFeriado : false,
            comment: matched ? matched.comment : '',
            shifts: matched ? matched.shifts : ['', '', '', '', '', '', '', '']
          };
        });
        const isFormFocused = document.activeElement && 
                              (document.activeElement.tagName === 'INPUT' || document.activeElement.tagName === 'SELECT') && 
                              document.activeElement.id.startsWith('form-');
        
        saveLocalOnly();
        renderTable();
        populateFormDaySelect();
        if (!isFormFocused) {
          loadFormDayData();
        }
        setSyncState('online');
      } else {
        // Cloud is empty for this period. Check if we have local storage data first
        const key = getStorageKey();
        const stored = localStorage.getItem(key);
        const dates = generatePeriodDates(currentYear, currentMonth, completeWeek);
        
        if (stored) {
          try {
            const parsed = JSON.parse(stored);
            daysData = dates.map(d => {
              const matched = parsed.days.find(p => p.dateKey === d.dateKey);
              return {
                ...d,
                isFeriado: matched ? matched.isFeriado : false,
                comment: matched ? matched.comment : '',
                shifts: matched ? matched.shifts : ['', '', '', '', '', '', '', '']
              };
            });
            saveToFirebaseOnly(); // Push local data to cloud
          } catch (e) {
            loadMockOrEmpty(dates);
          }
        } else {
          loadMockOrEmpty(dates);
        }
      }
    }, (err) => {
      console.error("Firebase read failed", err);
      setSyncState('offline');
    });
  }

  function loadMockOrEmpty(dates) {
    if (currentYear === 2026 && currentMonth === 4) {
      const mock = getMockDataForMay2026();
      daysData = dates.map(d => {
        const matched = mock.days.find(p => p.dateKey === d.dateKey);
        return {
          ...d,
          isFeriado: matched ? matched.isFeriado : false,
          comment: matched ? matched.comment : '',
          shifts: matched ? matched.shifts : ['', '', '', '', '', '', '', '']
        };
      });
      saveToFirebaseOnly();
    } else {
      generateEmptyDays(dates);
      renderTable();
      populateFormDaySelect();
      loadFormDayData();
    }
  }

  function loadLocalStorageBackup(stored, dates) {
    if (stored) {
      try {
        const parsed = JSON.parse(stored);
        daysData = dates.map(d => {
          const matched = parsed.days.find(p => p.dateKey === d.dateKey);
          return {
            ...d,
            isFeriado: matched ? matched.isFeriado : false,
            comment: matched ? matched.comment : '',
            shifts: matched ? matched.shifts : ['', '', '', '', '', '', '', '']
          };
        });
      } catch (e) {
        loadMockOrEmptyLocal(dates);
      }
    } else {
      loadMockOrEmptyLocal(dates);
    }
    renderTable();
    populateFormDaySelect();
    loadFormDayData();
  }

  function loadMockOrEmptyLocal(dates) {
    if (currentYear === 2026 && currentMonth === 4) {
      const mock = getMockDataForMay2026();
      daysData = dates.map(d => {
        const matched = mock.days.find(p => p.dateKey === d.dateKey);
        return {
          ...d,
          isFeriado: matched ? matched.isFeriado : false,
          comment: matched ? matched.comment : '',
          shifts: matched ? matched.shifts : ['', '', '', '', '', '', '', '']
        };
      });
      saveLocalOnly();
    } else {
      generateEmptyDays(dates);
    }
  }

  function setSyncState(state) {
    const syncStatusEl = document.getElementById('sync-status');
    if (!syncStatusEl) return;
    const syncLabel = syncStatusEl.querySelector('.label');
    syncStatusEl.className = `sync-indicator ${state === 'online' ? '' : state}`;
    if (state === 'online') {
      syncLabel.textContent = 'Sincronizado';
    } else if (state === 'offline') {
      syncLabel.textContent = 'Sin conexión';
    } else if (state === 'syncing') {
      syncLabel.textContent = 'Sincronizando...';
    }
  }

  function sanitizeFirebaseKey(key) {
    return key.replace(/[\.\$\#\[\]\/]/g, '_').trim().replace(/\s+/g, '_').toUpperCase();
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
    [formEntTime, formSalTime, formTicket, formComment, formIsFeriado, formDeCorrido].forEach(input => {
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
      });
    });
 
    // Auto-parse times from motive/comment in real time, updating locally
    formComment.addEventListener('input', () => {
      parseTimesFromComment();
      saveFormDayDataLocal(); // Local real-time update
    });
 
    // Also update locally when ticket changes in real-time
    formTicket.addEventListener('input', () => {
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
      case 6: return "10:00"; // sábado
      default: return "09:30"; // lunes a viernes, domingo
    }
  }

  function getStandardEndTime(dayOfWeek) {
    switch (dayOfWeek) {
      case 1: return "19:30"; // lunes
      case 2:
      case 3: return "18:30"; // martes, miércoles
      case 4:
      case 5: return "19:00"; // jueves, viernes
      case 6: return "12:30"; // sábado
      default: return ""; // domingo
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
    
    // 1. Separate ticket from comment/motive
    let ticketVal = "";
    let motiveVal = day.comment || "";
    
    // Match pattern: "Ticket 12345 - reason..." or "Tkt 12345: reason..."
    const ticketPattern = /^(?:Ticket|Tkt)\s*#?([a-zA-Z0-9_-]+)\s*[:-]\s*(.*)$/i;
    const ticketMatch = motiveVal.match(ticketPattern);
    if (ticketMatch) {
      ticketVal = ticketMatch[1];
      motiveVal = ticketMatch[2];
    }
    
    formTicket.value = ticketVal;
    formComment.value = motiveVal;
    
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
      entVal = day.shifts[0] || "";
      salVal = day.shifts[1] || "";
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
    const ticket = formTicket.value.trim();
    const motivo = formComment.value.trim();
    const hasCustomHours = formCustomHours.checked;
    
    let entVal = formEntTime.value.trim();
    let salVal = formSalTime.value.trim();
    
    day.isFeriado = isFeriado;
    
    let combinedComment = motivo;
    if (ticket) {
      combinedComment = `Ticket ${ticket} - ${motivo}`;
    }
    day.comment = combinedComment;
    
    const stdStart = getStandardStartTime(day.dayOfWeek);
    const stdEnd = getStandardEndTime(day.dayOfWeek);
    
    const isRestDay = (isFeriado || day.dayOfWeek === 0);
    const finalEnt = entVal || stdStart;
    const finalSal = salVal || stdEnd;
    
    if (isRestDay) {
      // Only save shift on Sunday/Holiday if there is actual worked exit time or motive/ticket
      if (salVal || motivo || ticket) {
        day.shifts = [finalEnt, finalSal, "", "", "", "", "", ""];
      } else {
        day.shifts = ["", "", "", "", "", "", "", ""];
      }
    } else if (day.dayOfWeek === 6) {
      if (hasCustomHours || (salVal && salVal !== stdEnd) || motivo || ticket) {
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
          <td class="col-summary">${summary.total.toFixed(2)}</td>
          <td class="col-summary">${summary.norm.toFixed(2)}</td>
          <td class="col-summary" style="color: ${summary.extras > 0 ? '#ef4444' : 'inherit'}; font-weight: 800;">
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
            value="${day.comment}"
            data-day-idx="${rIdx}"
          >
        </td>
        <td class="center" style="padding: 2px !important;">
          <button class="btn-toggle-feriado" data-day-idx="${rIdx}">
            ${isFeriado ? '☀️ Lab' : '🎈 Fer'}
          </button>
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
      // Sundays and holidays worked go to 100% overtime directly
      if (day.dayOfWeek === 0 || day.isFeriado) {
        overtime100Hours += c.totalDec;
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
      showStatus("Generando planilla Excel...", "success");
      const workbook = new ExcelJS.Workbook();
      const sheet = workbook.addWorksheet('Hoja1', {
        views: [{ showGridLines: true }]
      });

      // 1. Set column widths to match the template dimensions
      sheet.columns = [
        { width: 6 },   // A: DÍA (number)
        { width: 11 },  // B: DÍA (name)
        { width: 8 },   // C: 1° ENT.
        { width: 8 },   // D: 1° SAL.
        { width: 8 },   // E: 2° ENT.
        { width: 8 },   // F: 2° SAL.
        { width: 8 },   // G: 3° ENT.
        { width: 8 },   // H: 3° SAL.
        { width: 8 },   // I: 4° ENT.
        { width: 8 },   // J: 4° SAL.
        { width: 11 },  // K: TPO. PERM.
        { width: 6 },   // L: H
        { width: 9 },   // M: MIN. DEC.
        { width: 11 },  // N: TOTAL DEC.
        { width: 9 },   // O: Total
        { width: 9 },   // P: Norm.
        { width: 9 },   // Q: Extras
        { width: 45 }   // R: MOTIVO
      ];

      // Styling helpers
      const fontName = 'Arial';
      const borderThin = { style: 'thin', color: { rgb: 'CCCCCC' } };
      const borderMedium = { style: 'medium', color: { rgb: '000000' } };
      const doubleBottom = { style: 'double', color: { rgb: '000000' } };
      
      const borderAllThin = {
        left: borderThin, right: borderThin, top: borderThin, bottom: borderThin
      };

      const fillShift = {
        type: 'pattern', pattern: 'solid', fgColor: { rgb: 'E0F2F1' } // Ice blue
      };

      // --- Cell Writing & Formatting ---
      
      // Row 1: Company Logo Text
      sheet.getCell('A1').value = 'EMELNOR S.A.';
      sheet.getCell('A1').font = { name: fontName, size: 10, bold: true };

      // Row 2: Title (Merged A2:R2)
      sheet.mergeCells('A2:R2');
      const titleCell = sheet.getCell('A2');
      titleCell.value = 'CONVENIO Y CONTROL DE HORAS EXTRAORDINARIAS';
      titleCell.font = { name: fontName, size: 14, bold: true };
      titleCell.alignment = { horizontal: 'center', vertical: 'middle' };

      // Row 4: Agreement legal header text
      sheet.mergeCells('A4:C4');
      sheet.getCell('A4').value = 'En  Antofagasta a        ';
      sheet.getCell('A4').font = { name: fontName, size: 10 };
      sheet.getCell('A4').alignment = { horizontal: 'left' };

      // Write Date in D4 (Matches 16th of previous month)
      const prevDate = daysData[0].dateObj;
      sheet.getCell('D4').value = prevDate;
      sheet.getCell('D4').numFormat = 'yyyy-mm-dd';
      sheet.getCell('D4').font = { name: fontName, size: 10, bold: true };

      sheet.mergeCells('I4:R4');
      sheet.getCell('I4').value = 'se suscribe  el  presente convenio , mediante el cual el trabajador acuerda laborar horas   extraordinarias, de acuerdo a';
      sheet.getCell('I4').font = { name: fontName, size: 9 };

      // Row 6: Agreement Period info
      sheet.mergeCells('A6:R6');
      const startRangeStr = daysData[0].dateObj.toLocaleDateString('es-ES', { day: '2-digit', month: '2-digit', year: 'numeric' });
      const endRangeStr = daysData[daysData.length - 1].dateObj.toLocaleDateString('es-ES', { day: '2-digit', month: '2-digit', year: 'numeric' });
      sheet.getCell('A6').value = `Periodo de  Vigencia del Convenio: ${startRangeStr} al ${endRangeStr}`;
      sheet.getCell('A6').font = { name: fontName, size: 10, italic: true };

      // Row 7: Funcionario and Cargo
      sheet.getCell('A7').value = 'Funcionario(a):';
      sheet.getCell('A7').font = { name: fontName, size: 10, bold: true };
      
      // Cyan Highlight employee name in D7
      sheet.mergeCells('D7:H7');
      const empCell = sheet.getCell('D7');
      empCell.value = employeeInfo.name;
      empCell.font = { name: fontName, size: 10, bold: true };
      empCell.fill = {
        type: 'pattern', pattern: 'solid', fgColor: { rgb: '00FFFF' } // Solid Cyan
      };
      
      sheet.getCell('P7').value = 'Cargo:';
      sheet.getCell('P7').font = { name: fontName, size: 10, bold: true };
      sheet.getCell('Q7').value = employeeInfo.role;
      sheet.getCell('Q7').font = { name: fontName, size: 10 };

      // Row 8: Card, Area, Manager
      sheet.getCell('A8').value = 'Nº Tarjeta:';
      sheet.getCell('A8').font = { name: fontName, size: 10, bold: true };
      sheet.getCell('D8').value = employeeInfo.card;
      sheet.getCell('D8').font = { name: fontName, size: 10 };

      sheet.getCell('E8').value = 'Área:';
      sheet.getCell('E8').font = { name: fontName, size: 10, bold: true };
      sheet.getCell('F8').value = employeeInfo.area;
      sheet.getCell('F8').font = { name: fontName, size: 10 };

      sheet.getCell('P8').value = 'Jefe directo:';
      sheet.getCell('P8').font = { name: fontName, size: 10, bold: true };
      sheet.getCell('R8').value = employeeInfo.manager;
      sheet.getCell('R8').font = { name: fontName, size: 10 };

      // --- Table Headers (Rows 9 & 10) ---
      
      // Merge Column A&B headers
      sheet.mergeCells('A9:A10');
      sheet.getCell('A9').value = 'DÍA';
      sheet.getCell('A9').alignment = { horizontal: 'center', vertical: 'middle' };
      sheet.getCell('A9').font = { name: fontName, size: 9, bold: true };

      sheet.mergeCells('B9:B10'); // Just let it render Day Name header block
      
      // Shift Headers
      const shiftsHeaderCells = [
        ['C9', 'D9', '1º'],
        ['E9', 'F9', '2º'],
        ['G9', 'H9', '3º'],
        ['I9', 'J9', '4º']
      ];
      shiftsHeaderCells.forEach(([cell1, cell2, label]) => {
        sheet.mergeCells(`${cell1}:${cell2}`);
        const c = sheet.getCell(cell1);
        c.value = label;
        c.font = { name: fontName, size: 9, bold: true };
        c.alignment = { horizontal: 'center', vertical: 'middle' };
      });

      const shiftDetails = [
        ['C10', 'ENT.'], ['D10', 'SAL.'],
        ['E10', 'ENT.'], ['F10', 'SAL.'],
        ['G10', 'ENT.'], ['H10', 'SAL.'],
        ['I10', 'ENT.'], ['J10', 'SAL.']
      ];
      shiftDetails.forEach(([cell, label]) => {
        const c = sheet.getCell(cell);
        c.value = label;
        c.font = { name: fontName, size: 8, bold: true };
        c.alignment = { horizontal: 'center' };
      });

      // TPO PERM
      sheet.mergeCells('K9:K10');
      const tpoHeader = sheet.getCell('K9');
      tpoHeader.value = 'TPO. PERM.';
      tpoHeader.alignment = { horizontal: 'center', vertical: 'middle' };
      tpoHeader.font = { name: fontName, size: 8, bold: true };

      // HORAS PERM
      sheet.mergeCells('L9:N9');
      const hpHeader = sheet.getCell('L9');
      hpHeader.value = 'HORAS PERM.';
      hpHeader.alignment = { horizontal: 'center', vertical: 'middle' };
      hpHeader.font = { name: fontName, size: 9, bold: true };

      sheet.getCell('L10').value = 'H';
      sheet.getCell('M10').value = 'MIN. DEC.';
      sheet.getCell('N10').value = 'TOTAL DEC.';
      ['L10', 'M10', 'N10'].forEach(cell => {
        const c = sheet.getCell(cell);
        c.alignment = { horizontal: 'center' };
        c.font = { name: fontName, size: 8, bold: true };
      });

      // N° DE HORAS
      sheet.mergeCells('O9:Q9');
      const nhHeader = sheet.getCell('O9');
      nhHeader.value = 'Nº DE HORAS';
      nhHeader.alignment = { horizontal: 'center', vertical: 'middle' };
      nhHeader.font = { name: fontName, size: 9, bold: true };

      sheet.getCell('O10').value = 'Total';
      sheet.getCell('P10').value = 'Norm.';
      sheet.getCell('Q10').value = 'Extras';
      ['O10', 'P10', 'Q10'].forEach(cell => {
        const c = sheet.getCell(cell);
        c.alignment = { horizontal: 'center' };
        c.font = { name: fontName, size: 8, bold: true };
      });

      // MOTIVO
      sheet.mergeCells('R9:R10');
      const motHeader = sheet.getCell('R9');
      motHeader.value = 'MOTIVO DE HORAS EXTRAORDINARIAS';
      motHeader.alignment = { horizontal: 'center', vertical: 'middle' };
      motHeader.font = { name: fontName, size: 9, bold: true };

      // Apply borders and medium borders to table header rows
      for (let c = 1; c <= 18; c++) {
        sheet.getCell(9, c).border = { top: borderMedium, left: borderThin, right: borderThin };
        sheet.getCell(10, c).border = { bottom: borderMedium, left: borderThin, right: borderThin };
      }

      // --- Data Rows (Rows 11 to 42+) ---
      const startRowIdx = 11;
      const weeks = groupWeeks(daysData);

      daysData.forEach((day, idx) => {
        const r = startRowIdx + idx;
        
        // Col A: Day Num
        sheet.getCell(`A${r}`).value = day.dayNum;
        sheet.getCell(`A${r}`).font = { name: fontName, size: 9 };
        sheet.getCell(`A${r}`).alignment = { horizontal: 'center' };

        // Col B: Day Name
        sheet.getCell(`B${r}`).value = day.dayName;
        sheet.getCell(`B${r}`).font = { name: fontName, size: 9 };
        sheet.getCell(`B${r}`).alignment = { horizontal: 'left' };

        // Fill shifts
        for (let s = 0; s < 8; s++) {
          const colLet = String.fromCharCode(67 + s); // C to J
          const cell = sheet.getCell(`${colLet}${r}`);
          cell.value = day.shifts[s] || null;
          cell.alignment = { horizontal: 'center' };
          cell.font = { name: fontName, size: 9 };
          cell.fill = fillShift;
        }

        // Col K: TPO. PERM. Formula
        // =(D{row}-C{row})+(F{row}-E{row})+(H{row}-G{row})+(J{row}-I{row})
        sheet.getCell(`K${r}`).value = {
          formula: `=(D${r}-C${r})+(F${r}-E${r})+(H${r}-G${r})+(J${r}-I${r})`
        };
        sheet.getCell(`K${r}`).font = { name: fontName, size: 9 };
        sheet.getCell(`K${r}`).alignment = { horizontal: 'center' };
        sheet.getCell(`K${r}`).numFormat = '[h]:mm';

        // Col L: Hour Part Formula
        // =IF(K{row}="","",HOUR(K{row})) -> Excel doesn't like empty if not conditional, simply:
        sheet.getCell(`L${r}`).value = {
          formula: `=HOUR(K${r})`
        };
        sheet.getCell(`L${r}`).font = { name: fontName, size: 9 };
        sheet.getCell(`L${r}`).alignment = { horizontal: 'center' };

        // Col M: Minute Decimal Part Formula
        sheet.getCell(`M${r}`).value = {
          formula: `=(MINUTE(K${r}))/60`
        };
        sheet.getCell(`M${r}`).font = { name: fontName, size: 9 };
        sheet.getCell(`M${r}`).alignment = { horizontal: 'center' };
        sheet.getCell(`M${r}`).numFormat = '0.00';

        // Col N: Total Decimal
        sheet.getCell(`N${r}`).value = {
          formula: `=+L${r}+M${r}`
        };
        sheet.getCell(`N${r}`).font = { name: fontName, size: 9 };
        sheet.getCell(`N${r}`).alignment = { horizontal: 'center' };
        sheet.getCell(`N${r}`).numFormat = '0.00';

        // Col R: Comment (Motivo)
        sheet.getCell(`R${r}`).value = day.comment || null;
        sheet.getCell(`R${r}`).font = { name: fontName, size: 8 };

        // Apply base thin borders
        for (let c = 1; c <= 18; c++) {
          sheet.getCell(r, c).border = borderAllThin;
        }

        // Apply Saturday style
        if (day.dayOfWeek === 6) {
          const fillSat = { type: 'pattern', pattern: 'solid', fgColor: { rgb: 'F1F5F9' } };
          ['A', 'B', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R'].forEach(col => {
            sheet.getCell(`${col}${r}`).fill = fillSat;
          });
        }
      });

      // --- Weekly Summaries Rows (Sundays) ---
      weeks.forEach((week) => {
        // Find Sunday item in week
        const sundayItem = week.find(w => w.day.dayOfWeek === 0) || week[week.length - 1];
        const sunIdx = sundayItem.index;
        const sunRow = startRowIdx + sunIdx;
        const startWeekRow = startRowIdx + week[0].index;
        const endWeekRow = startRowIdx + week[week.length - 1].index;

        // Calculate normal hours standard for week
        let normHours = 0;
        week.forEach(w => {
          const d = w.day;
          if (d.dayOfWeek === 0 || d.isFeriado) {
            normHours += 0;
          } else {
            normHours += STANDARD_HOURS[d.dayOfWeek];
          }
        });

        // Write Sunday summaries
        // Col O: `=SUM(N{start}:N{end})`
        sheet.getCell(`O${sunRow}`).value = {
          formula: `=SUM(N${startWeekRow}:N${endWeekRow})`
        };
        sheet.getCell(`O${sunRow}`).font = { name: fontName, size: 9, bold: true };
        sheet.getCell(`O${sunRow}`).alignment = { horizontal: 'center' };
        sheet.getCell(`O${sunRow}`).numFormat = '0.00';

        // Col P: Normal hardcoded value
        sheet.getCell(`P${sunRow}`).value = normHours;
        sheet.getCell(`P${sunRow}`).font = { name: fontName, size: 9, bold: true };
        sheet.getCell(`P${sunRow}`).alignment = { horizontal: 'center' };
        sheet.getCell(`P${sunRow}`).numFormat = '0.00';

        // Col Q: `=SUM(O{sun}-P{sun})`
        sheet.getCell(`Q${sunRow}`).value = {
          formula: `=SUM(O${sunRow}-P${sunRow})`
        };
        sheet.getCell(`Q${sunRow}`).font = { name: fontName, size: 9, bold: true };
        sheet.getCell(`Q${sunRow}`).alignment = { horizontal: 'center' };
        sheet.getCell(`Q${sunRow}`).numFormat = '0.00';
      });

      // --- Apply Sunday and Holiday (Feriado) colors ---
      daysData.forEach((day, idx) => {
        const r = startRowIdx + idx;
        const isSunday = day.dayOfWeek === 0;
        const isFeriado = day.isFeriado;
        
        if (isSunday || isFeriado) {
          // Check if worked by analyzing shift inputs
          let hasHours = false;
          for (let s = 0; s < 8; s++) {
            if (day.shifts[s]) { hasHours = true; break; }
          }

          let colorFillHex = '';
          let colorFontHex = '';
          let isBold = false;

          if (hasHours) {
            colorFillHex = 'FF0000'; // RED
            colorFontHex = 'FFFFFF'; // WHITE
            isBold = true;
          } else {
            colorFillHex = 'FFFF00'; // YELLOW
            colorFontHex = '000000'; // BLACK
          }

          const fillObj = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { rgb: colorFillHex }
          };

          // Apply color to columns (Except shifts which stay ice blue/shaded but we can tint them too)
          // Actually, let's color ALL columns A to R to replicate the image where the entire row is colored!
          for (let c = 1; c <= 18; c++) {
            const cell = sheet.getCell(r, c);
            cell.fill = fillObj;
            cell.font = {
              name: fontName,
              size: cell.font && cell.font.size ? cell.font.size : 9,
              bold: isBold || (cell.font && cell.font.bold),
              color: { rgb: colorFontHex }
            };
          }
        }
      });

      // --- Totals Bottom Row (Row 45 in original, dynamic here) ---
      const totalsRow = startRowIdx + daysData.length + 2; // Leave 2 blank rows after the table
      
      // Label cell (Merged A to J)
      sheet.mergeCells(`A${totalsRow}:J${totalsRow}`);
      const labelCell = sheet.getCell(`A${totalsRow}`);
      labelCell.value = 'TOTALES PERÍODO';
      labelCell.font = { name: fontName, size: 10, bold: true };
      sheet.getCell(`B${totalsRow}`); // needed to touch merged cells sometimes
      
      // Formulas for sums
      const endDataRow = startRowIdx + daysData.length - 1;
      
      // Column K: Sum time values (not strictly necessary but matches footer)
      sheet.getCell(`K${totalsRow}`).value = {
        formula: `=SUM(K${startRowIdx}:K${endDataRow})`
      };
      sheet.getCell(`K${totalsRow}`).font = { name: fontName, size: 9, bold: true };
      sheet.getCell(`K${totalsRow}`).alignment = { horizontal: 'center' };
      sheet.getCell(`K${totalsRow}`).numFormat = '[h]:mm';

      // Column L: Sum hours part
      sheet.getCell(`L${totalsRow}`).value = {
        formula: `=SUM(L${startRowIdx}:L${endDataRow})`
      };
      sheet.getCell(`L${totalsRow}`).font = { name: fontName, size: 9, bold: true };
      sheet.getCell(`L${totalsRow}`).alignment = { horizontal: 'center' };

      // Column M: Sum min dec part
      sheet.getCell(`M${totalsRow}`).value = {
        formula: `=SUM(M${startRowIdx}:M${endDataRow})`
      };
      sheet.getCell(`M${totalsRow}`).font = { name: fontName, size: 9, bold: true };
      sheet.getCell(`M${totalsRow}`).alignment = { horizontal: 'center' };
      sheet.getCell(`M${totalsRow}`).numFormat = '0.00';

      // Column N: `=SUM(N7:N42)`
      sheet.getCell(`N${totalsRow}`).value = {
        formula: `=SUM(N${startRowIdx}:N${endDataRow})`
      };
      sheet.getCell(`N${totalsRow}`).font = { name: fontName, size: 10, bold: true };
      sheet.getCell(`N${totalsRow}`).alignment = { horizontal: 'center' };
      sheet.getCell(`N${totalsRow}`).numFormat = '0.00';

      // Column O: `=SUM(O11:O42)`
      sheet.getCell(`O${totalsRow}`).value = {
        formula: `=SUM(O${startRowIdx}:O${endDataRow})`
      };
      sheet.getCell(`O${totalsRow}`).font = { name: fontName, size: 10, bold: true };
      sheet.getCell(`O${totalsRow}`).alignment = { horizontal: 'center' };
      sheet.getCell(`O${totalsRow}`).numFormat = '0.00';

      // Column P: `=SUM(P11:P42)`
      sheet.getCell(`P${totalsRow}`).value = {
        formula: `=SUM(P${startRowIdx}:P${endDataRow})`
      };
      sheet.getCell(`P${totalsRow}`).font = { name: fontName, size: 10, bold: true };
      sheet.getCell(`P${totalsRow}`).alignment = { horizontal: 'center' };
      sheet.getCell(`P${totalsRow}`).numFormat = '0.00';

      // Column Q: `=SUM(Q11:Q42)` (Summing extras)
      sheet.getCell(`Q${totalsRow}`).value = {
        formula: `=SUM(Q${startRowIdx}:Q${endDataRow})`
      };
      sheet.getCell(`Q${totalsRow}`).font = { name: fontName, size: 10, bold: true };
      sheet.getCell(`Q${totalsRow}`).alignment = { horizontal: 'center' };
      sheet.getCell(`Q${totalsRow}`).numFormat = '0.00';

      // Formatting borders and yellow background for the totals row
      const fillTotals = {
        type: 'pattern', pattern: 'solid', fgColor: { rgb: 'FFFF00' } // Solid Yellow
      };
      
      for (let c = 1; c <= 17; c++) {
        const cell = sheet.getCell(totalsRow, c);
        cell.fill = fillTotals;
        cell.border = {
          top: borderMedium,
          bottom: doubleBottom,
          left: borderThin,
          right: borderThin
        };
      }

      // --- Row 46: 50% and 100% Labels (Row totalsRow + 1) ---
      const ratesRow = totalsRow + 1;
      
      // Column R: 0.5 (50% Surcharge box)
      const rCell = sheet.getCell(`R${ratesRow}`);
      rCell.value = 0.5;
      rCell.font = { name: fontName, size: 10, bold: true };
      rCell.numFormat = '0%';
      rCell.alignment = { horizontal: 'center' };
      rCell.border = borderAllThin;
      rCell.fill = {
        type: 'pattern', pattern: 'solid', fgColor: { rgb: 'FFFF00' } // Yellow
      };

      // Column S: 1.0 (100% Surcharge box)
      const sCell = sheet.getCell(`S${ratesRow}`);
      sCell.value = 1.0;
      sCell.font = { name: fontName, size: 10, bold: true };
      sCell.numFormat = '0%';
      sCell.alignment = { horizontal: 'center' };
      sCell.border = borderAllThin;
      sCell.fill = {
        type: 'pattern', pattern: 'solid', fgColor: { rgb: 'D1FAE5' } // Light Green
      };

      // --- Legal Text Paragraphs (Below the table) ---
      const legalStartRow = ratesRow + 2;
      
      const legalText1 = "Primero: A través del presente documento, el  trabajador  quien  suscribe  el  presente convenio y   autorización de horas extraordinarias,    acepta  y  reconoce  que las horas extraordinarias aquí expuestas  son  las  únicas   horas trabajadas  con  autorización  del  empleador  y consecuentemente  reconoce que son las que deben  ser canceladas como tal.";
      const legalText2 = "Segundo: Los abajo firmantes declaran que las horas de entrada y salida señaladas en el presente documento, son fiel reflejo de lo registrado en las tarjetas de control de asistencia utilizadas formal y oficialmente por el Empleador, expresando su conformidad con los cálculos y demás datos indicados.";

      sheet.mergeCells(`A${legalStartRow}:R${legalStartRow + 1}`);
      const lCell1 = sheet.getCell(`A${legalStartRow}`);
      lCell1.value = legalText1;
      lCell1.font = { name: fontName, size: 8, color: { rgb: '475569' } };
      lCell1.alignment = { wrapText: true, vertical: 'top' };

      sheet.mergeCells(`A${legalStartRow + 3}:R${legalStartRow + 4}`);
      const lCell2 = sheet.getCell(`A${legalStartRow + 3}`);
      lCell2.value = legalText2;
      lCell2.font = { name: fontName, size: 8, color: { rgb: '475569' } };
      lCell2.alignment = { wrapText: true, vertical: 'top' };

      // --- Signatures Section ---
      const sigRow = legalStartRow + 7;
      
      sheet.mergeCells(`B${sigRow}:D${sigRow}`);
      const fCell = sheet.getCell(`B${sigRow}`);
      fCell.value = 'Firma Funcionario';
      fCell.font = { name: fontName, size: 9, bold: true };
      fCell.alignment = { horizontal: 'center' };
      fCell.border = { top: { style: 'dashed', color: { rgb: '000000' } } };

      sheet.mergeCells(`J${sigRow}:M${sigRow}`);
      const mCell = sheet.getCell(`J${sigRow}`);
      mCell.value = 'Firma Jefe Directo';
      mCell.font = { name: fontName, size: 9, bold: true };
      mCell.alignment = { horizontal: 'center' };
      mCell.border = { top: { style: 'dashed', color: { rgb: '000000' } } };

      // --- Write Excel File Buffer & Download ---
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = window.URL.createObjectURL(blob);
      
      const anchor = document.createElement('a');
      anchor.href = url;
      anchor.download = `Horas_Extras_${employeeInfo.name.replace(/\s+/g, '_')}_${currentYear}_${currentMonth + 1}.xlsx`;
      anchor.click();
      
      window.URL.revokeObjectURL(url);
      showStatus("Planilla Excel exportada con éxito.", "success");
    } catch (error) {
      console.error(error);
      showStatus("Error al exportar a Excel. Intente nuevamente.", "danger");
    }
  }

  // --- Demonstration Mock Data ---
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
});
