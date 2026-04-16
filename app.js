/**
 * Terapia App - Firebase Firestore Edition
 * Sincronizzazione real-time su tutti i dispositivi.
 */

const APP_VERSION = 'v2026.04.17.001';

// ===== COSTANTI DI RIFERIMENTO =====
const GLUCOSE_MIN = 70;
const GLUCOSE_MAX = 180;
const GLUCOSE_AVG_LIMIT = 140;


// ===== FIREBASE =====
const firebaseConfig = {
  apiKey: "AIzaSyCLxYvfbNl6bMlot7xx1vOH-BjYGWgg3fM",
  authDomain: "terapia-dd6c3.firebaseapp.com",
  projectId: "terapia-dd6c3",
  storageBucket: "terapia-dd6c3.firebasestorage.app",
  messagingSenderId: "574810885658",
  appId: "1:574810885658:web:40c0e1139a23c702f1cba3"
};
firebase.initializeApp(firebaseConfig);
const auth = firebase.auth();
const db = firebase.firestore();

class TerapiaApp {
  constructor() {
    this.currentDate = new Date();
    this.currentFilter = 'all';
    this.entries = [];
    this.currentUserId = null;
    this.unsubscribeFirestore = null;
    this._migrationChecked = false;
    this.deferredInstallPrompt = null;
    this.isPwaInstalled = window.matchMedia('(display-mode: standalone)').matches || window.navigator.standalone === true;
    this.logLimit = 50; // Inizialmente mostra solo 50 elementi

    this.elements = {
      datePickerList: document.getElementById('datePickerList'),
      addEntryBtn: document.getElementById('addEntryBtn'),
      installAppBtn: document.getElementById('installAppBtn'),
      importBtn: document.getElementById('importBtn'),
      importInput: document.getElementById('importInput'),
      exportAllBtn: document.getElementById('exportAllBtn'),
      logTitle: document.getElementById('logTitle'),
      entryModal: document.getElementById('entryModal'),
      entryForm: document.getElementById('entryForm'),
      medFilterSelect: document.getElementById('medFilterSelect'),
      filterBtns: document.querySelectorAll('.filter-btns button'),
      logTbody: document.getElementById('log-tbody'),
      avgGlucose: document.getElementById('avgGlucose'),
      medsCount: document.getElementById('medsCount'),
      statusIndicator: document.getElementById('statusIndicator'),
      closeModalBtns: document.querySelectorAll('.close-modal, .cancel-btn'),
      entryTypeRadios: document.querySelectorAll('input[name="type"]'),
      entryValue: document.getElementById('entryValue'),
      unitInput: document.getElementById('unitInput'),
      unitsList: document.getElementById('unitsList'),
      medNameGroup: document.getElementById('med-name-group'),
      medName: document.getElementById('medName'),
      medsList: document.getElementById('medsList'),
      valueLabel: document.getElementById('value-label'),
      modalTitle: document.getElementById('modalTitle'),
      modalSub: document.getElementById('modalSub'),
      submitBtn: document.getElementById('submitBtn'),
      glucoseChartCanvas: document.getElementById('glucoseChart'),
      chartScopeBtns: document.querySelectorAll('#chartScopeBtns button'),
      chartSub: document.getElementById('chartSub'),
      syncRepoBtn: document.getElementById('syncRepoBtn'),
      deleteAllBtn: document.getElementById('deleteAllBtn'),
      appVersionMenu: document.getElementById('appVersionMenu'),
      appVersion: document.getElementById('appVersion'),
      appVersionToggle: document.getElementById('appVersionToggle'),
      glucoseChart: null,
      entryDay: document.getElementById('entryDay'),
      entryMonth: document.getElementById('entryMonth'),
      entryYear: document.getElementById('entryYear'),
      entryHH: document.getElementById('entryHH'),
      entryMM: document.getElementById('entryMM'),
      entryNote: document.getElementById('entryNote'),
      chartStartDate: document.getElementById('chartStartDate'),
      chartEndDate: document.getElementById('chartEndDate'),
      applyRangeBtn: document.getElementById('applyRangeBtn')
    };

    this.viewAll = true;
    this.chartScope = 'all';
    this.currentMedFilter = '';
    this.editingId = null;
    this.renderAppVersion();
    this.setupVersionToggle();
    this.setupPwaInstall();
    this.setupForceReset(); // New
    this.initFirebase();
    this.initDateRangeFields();
  }

  setupForceReset() {
    const btn = document.getElementById('forceResetApp');
    if (!btn) return;

    btn.onclick = async () => {
      if (confirm('Vuoi davvero resettare la cache? L\'app verrà ricaricata. I tuoi dati su Firebase NON verranno toccati.')) {
        // 1. Unregister Service Workers
        if ('serviceWorker' in navigator) {
          const regs = await navigator.serviceWorker.getRegistrations();
          for (let reg of regs) {
            await reg.unregister();
          }
        }
        
        // 2. Clear Caches
        if (window.caches) {
          const keys = await caches.keys();
          for (let key of keys) {
            await caches.delete(key);
          }
        }

        // 3. Clear LocalStorage related to app (optional)
        // localStorage.clear(); // Use with caution, maybe just clear specific ones

        alert('App resettata. L\'app verrà chiusa/ricaricata.');
        window.location.reload();
      }
    };
  }

  initDateRangeFields() {
    if (this.elements.chartStartDate && this.elements.chartEndDate) {
      const today = new Date();
      const firstOfMonth = new Date(today.getFullYear(), today.getMonth(), 1);
      
      this.elements.chartStartDate.value = firstOfMonth.toISOString().split('T')[0];
      this.elements.chartEndDate.value = today.toISOString().split('T')[0];
    }
  }

  renderAppVersion() {
    if (this.elements.appVersion) {
      this.elements.appVersion.textContent = APP_VERSION;
    }
  }

  setupVersionToggle() {
    const { appVersionMenu, appVersionToggle } = this.elements;
    if (!appVersionMenu || !appVersionToggle) return;

    const closeVersionPanel = () => {
      appVersionMenu.classList.remove('is-open');
      appVersionToggle.setAttribute('aria-expanded', 'false');
    };

    appVersionToggle.addEventListener('click', (event) => {
      event.stopPropagation();
      const isOpen = appVersionMenu.classList.toggle('is-open');
      appVersionToggle.setAttribute('aria-expanded', String(isOpen));
    });

    document.addEventListener('click', (event) => {
      if (!appVersionMenu.contains(event.target)) {
        closeVersionPanel();
      }
    });

    document.addEventListener('keydown', (event) => {
      if (event.key === 'Escape') {
        closeVersionPanel();
      }
    });
  }

  setupPwaInstall() {
    if (!this.elements.installAppBtn) return;

    window.addEventListener('beforeinstallprompt', (event) => {
      event.preventDefault();
      this.deferredInstallPrompt = event;
      this.updateInstallButton();
    });

    window.addEventListener('appinstalled', () => {
      this.isPwaInstalled = true;
      this.deferredInstallPrompt = null;
      this.updateInstallButton();
      this.showToast('App installata sul dispositivo');
    });

    this.updateInstallButton();
  }

  updateInstallButton() {
    if (!this.elements.installAppBtn) return;

    const shouldShow = Boolean(this.deferredInstallPrompt) && !this.isPwaInstalled;
    this.elements.installAppBtn.hidden = !shouldShow;
  }

  // ===== AUTH =====
  initFirebase() {
    // Abilita la persistenza offline con parametri che evitano blocchi tra schede
    db.enablePersistence({ synchronizeTabs: true }).catch(err => {
      if (err.code == 'failed-precondition') console.warn('Persistenza fallita: più schede aperte');
      else if (err.code == 'unimplemented') console.warn('Persistenza fallita: browser non supportato');
    });

    auth.onAuthStateChanged(user => {
      if (user) {
        this.currentUserId = user.uid;
        
        // Show loading screen while we wait for data sync
        const loadingOverlay = document.getElementById('loadingOverlay');
        if (loadingOverlay) {
          loadingOverlay.style.display = 'flex';
          loadingOverlay.style.opacity = '1';
        }

        this.showApp(user);
        this.setupEventListeners();
        this.updateMedsDataList();
        this.updateUnitsDataList('glucose');
        this.setupFirestoreListener();
      } else {
        this.showLoginScreen();
      }
    });
  }

  hideLoading() {
    const loadingOverlay = document.getElementById('loadingOverlay');
    if (!loadingOverlay || loadingOverlay.style.display === 'none') return;

    loadingOverlay.style.opacity = '0';
    setTimeout(() => {
        loadingOverlay.style.display = 'none';
        loadingOverlay.style.opacity = '1'; // Reset for next use
    }, 500);
  }

  showLoginScreen() {
    // If not logged in, we ensure the login overlay is visible and the loading is gone
    const loadingOverlay = document.getElementById('loadingOverlay');
    if (loadingOverlay) loadingOverlay.style.display = 'none';
    
    document.getElementById('loginOverlay').style.display = 'flex';
    this.updateSyncRepoButton();
    const googleBtn = document.getElementById('googleSignInBtn');
    if (googleBtn) googleBtn.onclick = () => this.signInWithGoogle();
  }

  showApp(user) {
    // Hide login and prepare app UI
    document.getElementById('loginOverlay').style.display = 'none';
    this.updateSyncRepoButton(user);
  }

  updateSyncRepoButton(user = auth.currentUser) {
    if (!this.elements.syncRepoBtn) return;

    if (user) {
      const name = user.displayName ? user.displayName.split(' ')[0] : user.email.split('@')[0];
      this.elements.syncRepoBtn.innerHTML = `<i data-lucide="cloud"></i>`;
      this.elements.syncRepoBtn.title = "";
    } else {
      this.elements.syncRepoBtn.innerHTML = `<i data-lucide="cloud"></i>`;
      this.elements.syncRepoBtn.title = "";
    }

    lucide.createIcons();
  }

  signInWithGoogle() {
    // Show loading while the popup is opening/processing
    const loadingOverlay = document.getElementById('loadingOverlay');
    if (loadingOverlay) {
      loadingOverlay.style.display = 'flex';
      loadingOverlay.style.opacity = '1';
    }

    const provider = new firebase.auth.GoogleAuthProvider();
    auth.signInWithPopup(provider).catch(err => {
      console.error(err);
      this.hideLoading();
      alert('Errore durante l\'accesso. Riprova.');
    });
  }

  // ===== FIRESTORE =====
  setupFirestoreListener() {
    if (this.unsubscribeFirestore) this.unsubscribeFirestore();
    this.unsubscribeFirestore = db
      .collection('users').doc(this.currentUserId).collection('entries')
      .onSnapshot(snapshot => {
        this.entries = snapshot.docs.map(doc => ({ ...doc.data(), id: doc.id }));
        this.updateUI();
        this.checkLocalStorageMigration();
        
        this.hideLoading();
      }, error => {
        console.error('Firestore error:', error);
        this.hideLoading();
        this.showToast('Errore di connessione al cloud', 'error');
      });
  }

  async checkLocalStorageMigration() {
    if (this._migrationChecked) return;
    this._migrationChecked = true;
    const localData = JSON.parse(localStorage.getItem('terapia_data')) || [];
    if (localData.length > 0 && this.entries.length === 0) {
      if (confirm(`Trovati ${localData.length} dati salvati localmente sul PC.\nVuoi importarli su Firebase per averli su tutti i dispositivi?`)) {
        try {
          const batch = db.batch();
          localData.forEach(entry => {
            const ref = db.collection('users').doc(this.currentUserId).collection('entries').doc(entry.id);
            batch.set(ref, entry);
          });
          await batch.commit();
          localStorage.removeItem('terapia_data');
          this.showToast(`${localData.length} dati migrati su Firebase! ☁️`);
        } catch (e) {
          console.error(e);
          this.showToast('Errore durante la migrazione', 'error');
        }
      }
    }
  }

  userCol() {
    return db.collection('users').doc(this.currentUserId).collection('entries');
  }

  // ===== EVENT LISTENERS =====
  setupEventListeners() {
    this.elements.addEntryBtn.onclick = () => {
      this.editingId = null;
      this.elements.modalTitle.innerText = "Nuovo Inserimento";
      this.elements.modalSub.innerText = "Aggiungi una lettura o un'assunzione";
      this.elements.submitBtn.innerText = "Salva Dato";
      this.elements.entryForm.reset();
      this.toggleEntryType('med');
      this.updateMedsDataList(); // Aggiorna elenco farmaci
      this.updateUnitsDataList('med'); // Aggiorna elenco unità
      this.elements.entryModal.style.display = 'flex';
      this.elements.entryDay.value = this.currentDate.getDate().toString().padStart(2, '0');
      this.elements.entryMonth.value = (this.currentDate.getMonth() + 1).toString().padStart(2, '0');
      this.elements.entryYear.value = this.currentDate.getFullYear();
      
      const now = new Date();
      this.elements.entryHH.value = now.getHours().toString().padStart(2, '0');
      this.elements.entryMM.value = now.getMinutes().toString().padStart(2, '0');
    };

    this.elements.exportAllBtn.onclick = () => this.downloadHistoryXLSX();
    if (this.elements.deleteAllBtn) this.elements.deleteAllBtn.onclick = () => this.deleteAllEntries();
    if (this.elements.installAppBtn) {
      this.elements.installAppBtn.onclick = () => this.installPwa();
    }
    this.elements.importBtn.onclick = () => this.elements.importInput.click();
    this.elements.importInput.onchange = (e) => this.handleImport(e);

    this.elements.filterBtns.forEach(btn => {
      if (btn.hasAttribute('data-scope')) return;
      btn.onclick = () => {
        const logBtns = Array.from(this.elements.filterBtns).filter(b => b.hasAttribute('data-filter'));
        logBtns.forEach(b => b.classList.remove('primary'));
        logBtns.forEach(b => b.classList.add('secondary'));
        btn.classList.remove('secondary');
        btn.classList.add('primary');
        this.currentFilter = btn.dataset.filter;
        this.logLimit = 50; // Reset limite quando si cambia filtro
        if (this.currentFilter === 'meds') {
          if (this.elements.medFilterSelect) {
            this.elements.medFilterSelect.style.display = 'block';
            this.populateMedFilter();
          }
        } else {
          if (this.elements.medFilterSelect) {
            this.elements.medFilterSelect.style.display = 'none';
            this.currentMedFilter = '';
            this.elements.medFilterSelect.value = '';
          }
        }
        this.renderLog();
      };
    });

    if (this.elements.medFilterSelect) {
      this.elements.medFilterSelect.onchange = (e) => {
        this.currentMedFilter = e.target.value;
        this.renderLog();
      };
    }

    this.elements.closeModalBtns.forEach(btn => {
      btn.onclick = () => this.elements.entryModal.style.display = 'none';
    });

    this.elements.entryTypeRadios.forEach(radio => {
      radio.onchange = (e) => this.toggleEntryType(e.target.value);
    });

    this.elements.chartScopeBtns.forEach(btn => {
      btn.onclick = () => {
        this.elements.chartScopeBtns.forEach(b => b.classList.remove('primary'));
        this.elements.chartScopeBtns.forEach(b => b.classList.add('secondary'));
        btn.classList.remove('secondary');
        btn.classList.add('primary');
        this.chartScope = btn.dataset.scope;
        this.viewAll = (this.chartScope === 'all');
        this.elements.chartSub.innerText = this.chartScope === 'day' ? "Valori rilevati oggi" : "Storico completo dei dati";
        this.updateDashboard();
        this.renderLog();
      };
    });

    this.elements.entryForm.onsubmit = (e) => {
      e.preventDefault();
      this.saveEntry();
    };

    window.onclick = (e) => {
      if (e.target === this.elements.entryModal) {
        this.elements.entryModal.style.display = 'none';
      }
    };

    // Gestione apertura automatica datalist (Suggerimenti) al primo click/focus
    [this.elements.unitInput, this.elements.medName].forEach(el => {
      if (el) {
        let lastAutoOpen = 0;
        const autoOpen = () => {
          const now = Date.now();
          // Evita di chiamare showPicker troppo frequentemente (es. click dopo focus)
          if (now - lastAutoOpen < 300) return;
          lastAutoOpen = now;
          if (typeof el.showPicker === 'function') {
            try { el.showPicker(); } catch(e) {}
          }
        };
        el.addEventListener('click', autoOpen);
        el.addEventListener('focus', autoOpen);
      }
    });

    // Sync button = cloud status / sign out
    this.elements.syncRepoBtn.onclick = () => {
      const user = auth.currentUser;
      if (user) {
        if (confirm('Vuoi davvero uscire e disconnettere il Cloud?')) {
          if (this.unsubscribeFirestore) this.unsubscribeFirestore();
          auth.signOut().then(() => {
            this.showToast('Cloud disconnesso', 'info');
          });
        }
      } else {
        // Se non è loggato, potremmo voler mostrare il login, 
        // ma seguiamo la richiesta dell'utente specifica sulla disconnessione.
        this.showToast('Cloud non attivo', 'info');
      }
    };
    this.elements.syncRepoBtn.oncontextmenu = (e) => {
      e.preventDefault();
      this.downloadHistoryXLSX();
    };

    if (this.elements.applyRangeBtn) {
      this.elements.applyRangeBtn.onclick = () => {
        const start = this.elements.chartStartDate.value;
        const end = this.elements.chartEndDate.value;
        if (!start || !end) {
          this.showToast('Seleziona entrambe le date', 'error');
          return;
        }
        this.chartScope = 'custom';
        this.viewAll = true; // Use global entries but filtered by date in the UI methods
        
        this.elements.chartScopeBtns.forEach(b => b.classList.remove('primary'));
        this.elements.chartScopeBtns.forEach(b => b.classList.add('secondary'));
        
        this.elements.chartSub.innerText = `Dati dal ${new Date(start).toLocaleDateString('it-IT')} al ${new Date(end).toLocaleDateString('it-IT')}`;
        this.updateDashboard();
        this.renderLog();
      };
    }

    // New resize listener for sticky headers
    window.addEventListener('resize', () => this.syncStickyHeaders());

    // Limita input numerici della data e ora
    [this.elements.entryDay, this.elements.entryMonth, this.elements.entryHH, this.elements.entryMM].forEach(el => {
      if (el) {
        el.addEventListener('input', (e) => {
          if (e.target.value.length > 2) {
            e.target.value = e.target.value.slice(0, 2);
          }
        });
      }
    });
    if (this.elements.entryYear) {
      this.elements.entryYear.addEventListener('input', (e) => {
        if (e.target.value.length > 4) {
          e.target.value = e.target.value.slice(0, 4);
        }
      });
    }
  }

  async installPwa() {
    if (this.isPwaInstalled) {
      this.showToast('App gia installata');
      return;
    }

    if (!this.deferredInstallPrompt) {
      this.showToast('Installazione non disponibile da questo browser', 'error');
      return;
    }

    this.deferredInstallPrompt.prompt();
    const { outcome } = await this.deferredInstallPrompt.userChoice;
    this.deferredInstallPrompt = null;
    this.updateInstallButton();

    if (outcome !== 'accepted') {
      this.showToast('Installazione annullata', 'error');
    }
  }

  // ===== UI METHODS =====
  toggleEntryType(type) {
    if (type === 'glucose') {
      this.elements.medNameGroup.style.display = 'none';
      this.elements.valueLabel.innerText = 'Valore Glicemia';
      this.elements.entryValue.placeholder = 'es. 110';
      this.updateUnitsDataList('glucose');
      this.elements.entryValue.required = true;
      this.elements.medName.required = false;
    } else {
      this.elements.medNameGroup.style.display = 'block';
      this.elements.valueLabel.innerText = 'Dose / Quantità Farmaco';
      this.elements.entryValue.placeholder = 'es. 1';
      this.updateUnitsDataList('med');
      this.elements.entryValue.required = false;
      this.elements.medName.required = true;
    }
  }

  changeDay(offset) {
    this.currentDate.setDate(this.currentDate.getDate() + offset);
    this.updateUI();
  }

  updateMedsDataList() {
    if (!this.elements.medsList) return;
    const names = [...new Set(this.entries.filter(e => e.type === 'med' && e.medName).map(e => e.medName))].sort();
    this.elements.medsList.innerHTML = names.map(name => `<option value="${name}">`).join('');
    this.populateMedFilter();
  }

  populateMedFilter() {
    if (!this.elements.medFilterSelect) return;
    const names = [...new Set(this.entries.filter(e => e.type === 'med' && e.medName).map(e => e.medName))].sort();
    const currentVal = this.currentMedFilter;
    this.elements.medFilterSelect.innerHTML = '<option value="">Tutti i farmaci</option>' +
      names.map(name => `<option value="${name}">${name}</option>`).join('');
    this.elements.medFilterSelect.value = currentVal;
  }

  updateUnitsDataList(type) {
    if (!this.elements.unitsList) return;
    const defaultUnits = type === 'glucose' ? ['mg/dL', 'mmol/L'] : ['UI', 'ml', 'compresse', 'gocce', 'mg', 'bustine'];
    const historyUnits = this.entries.filter(e => e.type === type && e.unit).map(e => e.unit);
    const allUnits = [...new Set([...defaultUnits, ...historyUnits])].sort();
    this.elements.unitsList.innerHTML = allUnits.map(u => `<option value="${u}">`).join('');
  }

  renderDateList() {
    if (!this.elements.datePickerList) return;
    const dateSet = new Set();
    const today = new Date();
    today.setHours(0,0,0,0);
    
    // Periodo di interesse: 15 giorni fa -> 7 giorni dopo
    const minDay = new Date(today);
    minDay.setDate(minDay.getDate() - 15);
    const maxDay = new Date(today);
    maxDay.setDate(maxDay.getDate() + 7);

    for(let i = -15; i <= 7; i++) {
      const d = new Date(today);
      d.setDate(d.getDate() + i);
      dateSet.add(d.toDateString());
    }
    
    // Aggiungi la data corrente se fuori range
    dateSet.add(this.currentDate.toDateString());
    
    // Aggiungi date rilevanti recenti (ultimi 30 giorni) o selezionate
    const thirtyDaysAgo = new Date(today);
    thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
    
    this.entries.forEach(e => { 
      const et = new Date(e.timestamp);
      if (et >= thirtyDaysAgo || et.toDateString() === this.currentDate.toDateString()) {
        dateSet.add(et.toDateString()); 
      }
    });

    const sortedDates = Array.from(dateSet).map(ds => new Date(ds)).sort((a,b) => a - b);
    this.elements.datePickerList.innerHTML = '';
    sortedDates.forEach(date => {
      const isToday = date.toDateString() === today.toDateString();
      const isActive = date.toDateString() === this.currentDate.toDateString();
      const item = document.createElement('div');
      item.className = `date-item ${isActive ? 'active' : ''} ${isToday ? 'today' : ''}`;
      const dayName = date.toLocaleDateString('it-IT', { weekday: 'short' });
      const monthName = date.toLocaleDateString('it-IT', { month: 'short' });
      const dayNum = date.getDate();
      item.innerHTML = `
        <span class="date-name">${dayName}</span>
        <span class="date-number">${dayNum}</span>
        <span class="date-month">${monthName}</span>
      `;
      item.onclick = () => {
        this.currentDate = date;
        this.viewAll = false;
        this.chartScope = 'day';
        this.updateUI();
      };
      this.elements.datePickerList.appendChild(item);
      if (isActive) {
        setTimeout(() => item.scrollIntoView({ behavior: 'smooth', inline: 'center', block: 'nearest' }), 100);
      }
    });
  }

  updateUI() {
    this.renderLog();
    this.renderDateList();
    this.updateDashboard();
  }

  getEntriesForDate(date) {
    const dateStr = date.toDateString();
    return this.entries.filter(e => new Date(e.timestamp).toDateString() === dateStr);
  }

  renderLog() {
    let filteredEntries;
    if (this.chartScope === 'custom' && this.elements.chartStartDate.value && this.elements.chartEndDate.value) {
      const startObj = new Date(this.elements.chartStartDate.value);
      startObj.setHours(0,0,0,0);
      const endObj = new Date(this.elements.chartEndDate.value);
      endObj.setHours(23,59,59,999);
      filteredEntries = this.entries.filter(e => e.timestamp >= startObj.getTime() && e.timestamp <= endObj.getTime());
    } else {
      filteredEntries = this.viewAll ? [...this.entries] : this.getEntriesForDate(this.currentDate);
    }
    
    let titleText = "Storico Completo";
    if (this.chartScope === 'custom') {
      const startStr = new Date(this.elements.chartStartDate.value).toLocaleDateString('it-IT', { day: '2-digit', month: '2-digit', year: 'numeric' });
      const endStr = new Date(this.elements.chartEndDate.value).toLocaleDateString('it-IT', { day: '2-digit', month: '2-digit', year: 'numeric' });
      titleText = `Dati dal ${startStr} al ${endStr}`;
    } else if (this.viewAll && this.entries.length > 0) {
      const minTimestamp = Math.min(...this.entries.map(e => e.timestamp));
      const startDate = new Date(minTimestamp).toLocaleDateString('it-IT', { day: '2-digit', month: '2-digit', year: 'numeric' });
      titleText = `Storico Completo dal ${startDate}`;
    }
    this.elements.logTitle.innerText = this.viewAll ? titleText : `Dati del ${this.currentDate.toLocaleDateString('it-IT', { day: '2-digit', month: '2-digit', year: 'numeric' })}`;

    if (this.currentFilter === 'glucose') {
      filteredEntries = filteredEntries.filter(e => e.type === 'glucose');
    } else if (this.currentFilter === 'meds') {
      filteredEntries = filteredEntries.filter(e => e.type === 'med');
      if (this.currentMedFilter) {
        filteredEntries = filteredEntries.filter(e => e.medName === this.currentMedFilter);
      }
    }

    filteredEntries.sort((a, b) => b.timestamp - a.timestamp);
    const totalFiltered = filteredEntries.length;
    const entriesToShow = filteredEntries.slice(0, this.logLimit);
    
    this.elements.logTbody.innerHTML = '';

    if (filteredEntries.length === 0) {
      this.elements.logTbody.innerHTML = `<tr><td colspan="5" style="text-align: center; color: var(--text-secondary); padding: 2rem;">Nessun dato trovato.</td></tr>`;
      return;
    }

    let lastDateTag = "";
    entriesToShow.forEach(entry => {
      const entryDate = new Date(entry.timestamp);
      const groupKey = `${entryDate.getFullYear()}-${String(entryDate.getMonth() + 1).padStart(2, '0')}-${String(entryDate.getDate()).padStart(2, '0')}`;
      const groupDayLabel = entryDate.toLocaleDateString('it-IT', { weekday: 'long' });
      const groupDateLabel = entryDate.toLocaleDateString('it-IT', { day: 'numeric', month: 'long', year: 'numeric' });
      if (this.viewAll && groupKey !== lastDateTag) {
        const headerTr = document.createElement('tr');
        headerTr.className = 'log-date-divider';
        headerTr.innerHTML = `
          <td colspan="6" class="log-date-divider-cell">
            <div class="log-date-divider-content">
              <span class="log-date-divider-day">${groupDayLabel}</span>
              <span class="log-date-divider-date">${groupDateLabel}</span>
            </div>
          </td>
        `;
        this.elements.logTbody.appendChild(headerTr);
        lastDateTag = groupKey;
      }
      const time = entryDate.toLocaleTimeString('it-IT', { hour: '2-digit', minute: '2-digit' });
      const tr = document.createElement('tr');
      tr.draggable = true;
      tr.classList.add('draggable-row');
      tr.dataset.id = entry.id;
      const isGlucose = entry.type === 'glucose';
      const medStyle = !isGlucose ? this.getMedBadgeStyle(entry.medName) : null;
      tr.innerHTML = `
        <td data-label="Data"><span class="date-inline-value">${entryDate.toLocaleDateString('it-IT', { day: '2-digit', month: '2-digit', year: 'numeric' })}</span></td>
        <td data-label="Ora"><span class="time-inline-value">${time}</span></td>
        <td data-label="Tipo"><span class="badge ${isGlucose ? 'badge-glucose' : 'badge-pill'}">${isGlucose ? 'Glicemia' : 'Farmaco'}</span></td>
        <td data-label="Dose">
          <div class="dose-inline" style="display: flex; align-items: baseline; gap: 6px;">
            <span class="dose-value" style="font-weight: 700; font-size: 1.1rem;">${entry.value}</span>
            <span class="dose-unit" style="color: var(--text-secondary); font-size: 0.825rem;">${entry.unit || (isGlucose ? 'mg/dL' : '')}</span>
            ${!isGlucose ? `<span class="badge" style="background: ${medStyle.bg}; color: ${medStyle.color}; border: 1px solid ${medStyle.border}; font-size: 0.725rem; padding: 0.15rem 0.6rem;">${entry.medName}</span>` : ''}
          </div>
        </td>
        <td data-label="Note" style="color: var(--text-secondary); font-size: 0.9rem;"><span class="note-inline">${entry.note || '-'}</span></td>
        <td data-label="Azioni" style="text-align: right;">
          <div class="row-actions">
            <i data-lucide="grip-vertical" style="width: 14px; height: 14px; color: var(--text-muted); cursor: grab;"></i>
            <button class="secondary icon-btn edit-btn" data-id="${entry.id}" style="padding: 0.25rem;" title="Modifica">
              <i data-lucide="edit-3" style="width: 16px; height: 16px;"></i>
            </button>
            <button class="secondary icon-btn delete-btn" data-id="${entry.id}" style="padding: 0.25rem;" title="Elimina">
              <i data-lucide="trash-2" style="width: 16px; height: 16px;"></i>
            </button>
          </div>
        </td>
      `;
      this.elements.logTbody.appendChild(tr);
    });

    // Bottone "Carica Altri" se ci sono più dati
    if (totalFiltered > this.logLimit) {
      const loadMoreTr = document.createElement('tr');
      loadMoreTr.className = 'load-more-row';
      loadMoreTr.innerHTML = `
        <td colspan="6" class="load-more-cell">
          <button id="loadMoreBtn" class="secondary">
            Carica altri (${totalFiltered - this.logLimit} rimanenti)
          </button>
        </td>
      `;
      this.elements.logTbody.appendChild(loadMoreTr);
      document.getElementById('loadMoreBtn').onclick = () => {
        this.logLimit += 100;
        this.renderLog();
      };
    }

    lucide.createIcons();
    this.setupDragAndDrop();

    document.querySelectorAll('.edit-btn').forEach(btn => {
      btn.onclick = () => this.editEntry(btn.dataset.id);
    });
    document.querySelectorAll('.delete-btn').forEach(btn => {
      btn.onclick = () => this.deleteEntry(btn.dataset.id);
    });
    
    // Recalibrate sticky headers after rendering
    this.syncStickyHeaders();
  }

  syncStickyHeaders() {
    const mainHeader = document.querySelector('header');
    const logHeader = document.querySelector('.log-header');
    const thead = document.querySelector('.log-table-container thead');
    
    if (!mainHeader || !logHeader || !thead) return;
    
    // Use requestAnimationFrame to ensure we measure after layout
    requestAnimationFrame(() => {
      const headerHeight = mainHeader.offsetHeight;
      logHeader.style.setProperty('top', `${headerHeight}px`, 'important');
      
      const logHeaderHeight = logHeader.offsetHeight;
      thead.style.setProperty('top', `${headerHeight + logHeaderHeight - 1}px`, 'important');
    });
  }

  setupDragAndDrop() {
    const rows = document.querySelectorAll('.draggable-row');
    let draggedId = null;
    rows.forEach(row => {
      row.addEventListener('dragstart', (e) => {
        draggedId = row.dataset.id;
        row.style.opacity = '0.4';
        e.dataTransfer.effectAllowed = 'move';
      });
      row.addEventListener('dragend', () => {
        row.style.opacity = '1';
        rows.forEach(r => r.classList.remove('drag-over'));
      });
      row.addEventListener('dragover', (e) => { e.preventDefault(); e.dataTransfer.dropEffect = 'move'; return false; });
      row.addEventListener('dragenter', () => { row.classList.add('drag-over'); });
      row.addEventListener('dragleave', () => { row.classList.remove('drag-over'); });
      row.addEventListener('drop', (e) => {
        e.stopPropagation();
        const targetId = row.dataset.id;
        if (draggedId !== targetId) this.reorderEntries(draggedId, targetId);
        return false;
      });
    });
  }

  async reorderEntries(sourceId, targetId) {
    let list = this.viewAll ? [...this.entries] : this.getEntriesForDate(this.currentDate);
    list.sort((a, b) => b.timestamp - a.timestamp);
    const sourceIdx = list.findIndex(e => e.id === sourceId);
    const targetIdx = list.findIndex(e => e.id === targetId);
    if (sourceIdx === -1 || targetIdx === -1) return;
    const [movedEntry] = list.splice(sourceIdx, 1);
    list.splice(targetIdx, 0, movedEntry);
    const prev = list[targetIdx - 1];
    const next = list[targetIdx + 1];
    let newTimestamp;
    if (prev && next) {
      newTimestamp = Math.floor((prev.timestamp + next.timestamp) / 2);
    } else if (prev) {
      newTimestamp = prev.timestamp - 1000;
    } else if (next) {
      newTimestamp = next.timestamp + 1000;
    } else {
      newTimestamp = movedEntry.timestamp;
    }
    try {
      await this.userCol().doc(sourceId).update({ timestamp: newTimestamp });
    } catch (e) {
      console.error('Reorder error:', e);
    }
    this.renderLog();
  }

  editEntry(id) {
    const entry = this.entries.find(e => e.id === id);
    if (!entry) return;
    this.editingId = id;
    this.elements.modalTitle.innerText = "Modifica Dato";
    this.elements.modalSub.innerText = "Aggiorna le informazioni dell'inserimento";
    this.elements.submitBtn.innerText = "Aggiorna Dato";
    const form = this.elements.entryForm;
    form.type.value = entry.type;
    this.toggleEntryType(entry.type);
    this.elements.entryValue.value = entry.value;
    this.elements.unitInput.value = entry.unit || (entry.type === 'glucose' ? 'mg/dL' : 'ml');
    this.elements.medName.value = entry.medName || '';
    this.elements.entryNote.value = entry.note || '';

    const d = new Date(entry.timestamp);
    this.elements.entryDay.value = d.getDate().toString().padStart(2, '0');
    this.elements.entryMonth.value = (d.getMonth() + 1).toString().padStart(2, '0');
    this.elements.entryYear.value = d.getFullYear();
    this.elements.entryHH.value = d.getHours().toString().padStart(2, '0');
    this.elements.entryMM.value = d.getMinutes().toString().padStart(2, '0');
    this.elements.entryModal.style.display = 'flex';
  }

  getFilteredEntries() {
    if (this.chartScope === 'all') {
      return this.entries;
    } else if (this.chartScope === 'day') {
      return this.getEntriesForDate(this.currentDate);
    } else if (this.chartScope === 'custom') {
      const startStr = this.elements.chartStartDate.value;
      const endStr = this.elements.chartEndDate.value;
      if (!startStr || !endStr) return this.entries;
      
      const start = new Date(startStr);
      start.setHours(0, 0, 0, 0);
      const end = new Date(endStr);
      end.setHours(23, 59, 59, 999);
      
      return this.entries.filter(e => {
        const d = new Date(e.timestamp);
        return d >= start && d <= end;
      });
    }
    return this.entries;
  }

  updateDashboard() {
    const dataSource = this.getFilteredEntries();
    const glucoseEntries = dataSource.filter(e => e.type === 'glucose');
    const medEntries = dataSource.filter(e => e.type === 'med');

    if (glucoseEntries.length > 0) {
      const avg = Math.round(glucoseEntries.reduce((acc, curr) => acc + parseFloat(curr.value || 0), 0) / glucoseEntries.length);
      this.elements.avgGlucose.innerText = avg;
      this.elements.avgGlucose.style.color = avg > GLUCOSE_AVG_LIMIT ? '#f43f5e' : (avg < GLUCOSE_MIN ? '#f59e0b' : '#10b981');
    } else {
      this.elements.avgGlucose.innerText = '--';
      this.elements.avgGlucose.style.color = 'var(--text-secondary)';
    }

    this.elements.medsCount.innerText = medEntries.length;

    // Stato Glicemico basato sul range selezionato
    if (glucoseEntries.length === 0) {
      this.elements.statusIndicator.innerText = 'Dati Assenti';
      this.elements.statusIndicator.style.color = 'var(--text-secondary)';
    } else {
      const filteredAvg = Math.round(glucoseEntries.reduce((acc, curr) => acc + parseFloat(curr.value || 0), 0) / glucoseEntries.length);
      if (filteredAvg > GLUCOSE_MAX || filteredAvg < GLUCOSE_MIN) {
        this.elements.statusIndicator.innerText = 'Attenzione';
        this.elements.statusIndicator.style.color = '#f43f5e';
      } else {
        this.elements.statusIndicator.innerText = 'Ottimo';
        this.elements.statusIndicator.style.color = '#10b981';
      }
    }

    this.elements.chartScopeBtns.forEach(btn => {
      const isActive = btn.dataset.scope === this.chartScope;
      btn.classList.toggle('primary', isActive);
      btn.classList.toggle('secondary', !isActive);
    });

    if (this.elements.applyRangeBtn) {
      const isCustom = this.chartScope === 'custom';
      this.elements.applyRangeBtn.classList.toggle('primary', isCustom);
      this.elements.applyRangeBtn.classList.toggle('secondary', !isCustom);
    }

    this.renderChart();
  }

  renderChart() {
    const canvas = this.elements.glucoseChartCanvas;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;

    if (this.elements.glucoseChart) {
      this.elements.glucoseChart.destroy();
    }

    const data = this.getFilteredEntries();
    const chartPoints = data
      .filter(e => e.type === 'glucose')
      .map(e => ({ x: e.timestamp, y: parseFloat(e.value) || 0 }))
      .sort((a, b) => a.x - b.x);

    let xMin, xMax;
    if (this.chartScope === 'day') {
      const d = new Date(this.currentDate);
      xMin = new Date(d.getFullYear(), d.getMonth(), d.getDate(), 0, 0, 0).getTime();
      xMax = new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1, 0, 0, 0).getTime();
    } else if (this.chartScope === 'custom') {
      xMin = new Date(this.elements.chartStartDate.value).setHours(0, 0, 0, 0);
      xMax = new Date(this.elements.chartEndDate.value).setHours(23, 59, 59, 999);
    } else if (chartPoints.length > 0) {
      xMin = Math.min(...chartPoints.map(p => p.x));
      xMax = Math.max(...chartPoints.map(p => p.x));
    } else {
      xMin = Date.now() - 3600000;
      xMax = Date.now() + 3600000;
    }
    const yMaxVal = Math.max(200, ...(chartPoints.length > 0 ? chartPoints.map(p => p.y) : [0])) + 20;

    const gradient = ctx.createLinearGradient(0, 0, 0, 300);
    gradient.addColorStop(0, 'rgba(16, 185, 129, 0.4)');
    gradient.addColorStop(1, 'rgba(16, 185, 129, 0)');

    this.elements.glucoseChart = new Chart(ctx, {
      type: 'line',
      data: {
        datasets: [
          {
            label: 'Base',
            data: [{ x: xMin, y: 0 }, { x: xMax, y: 0 }],
            borderColor: 'transparent', pointRadius: 0, fill: false, tension: 0, z: -2
          },
          {
            label: 'Fascia Bassa',
            data: [{ x: xMin, y: GLUCOSE_MIN }, { x: xMax, y: GLUCOSE_MIN }],
            borderColor: 'rgba(245, 158, 11, 0.2)',
            borderWidth: 1,
            pointRadius: 0,
            fill: 0, // Riempie verso la base (0)
            backgroundColor: 'rgba(244, 63, 94, 0.04)',
            tension: 0,
            z: -1
          },
          {
            label: 'Fascia Ideale',
            data: [{ x: xMin, y: GLUCOSE_MAX }, { x: xMax, y: GLUCOSE_MAX }],
            borderColor: 'rgba(16, 185, 129, 0.2)',
            borderWidth: 1,
            pointRadius: 0,
            fill: 1, // Riempie verso il 70
            backgroundColor: 'rgba(16, 185, 129, 0.08)',
            tension: 0,
            z: -1
          },
          {
            label: 'Fascia Alta',
            data: [{ x: xMin, y: yMaxVal }, { x: xMax, y: yMaxVal }],
            borderColor: 'transparent',
            borderWidth: 1,
            pointRadius: 0,
            fill: 2, // Riempie verso il 180
            backgroundColor: 'rgba(244, 63, 94, 0.04)',
            tension: 0,
            z: -1
          },
          {
            label: 'Glicemia (mg/dL)', data: chartPoints,
            borderColor: '#10b981', backgroundColor: 'transparent', borderWidth: 1.5, tension: 0, fill: false,
            pointBackgroundColor: '#fff', pointBorderColor: '#10b981', pointBorderWidth: 1,
            pointRadius: this.chartScope === 'all' ? 1.5 : 3, pointHoverRadius: this.chartScope === 'all' ? 3 : 5,
            pointHoverBackgroundColor: '#10b981', pointHoverBorderColor: '#fff', pointHoverBorderWidth: 2,
            z: 10
          }
        ]
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        animation: { duration: 1200, easing: 'easeInOutQuart' },
        interaction: { intersect: false, mode: 'index' },
        plugins: {
          legend: { display: false },
          tooltip: {
            backgroundColor: 'rgba(15, 23, 42, 0.9)', titleColor: '#fff', bodyColor: '#fff',
            borderColor: 'rgba(255, 255, 255, 0.1)', borderWidth: 1, padding: 12, boxPadding: 8, usePointStyle: true,
            callbacks: {
              label: (context) => `Glicemia: ${context.parsed.y} mg/dL`,
              title: (context) => {
                const d = new Date(context[0].parsed.x);
                return d.toLocaleDateString('it-IT', { day: '2-digit', month: '2-digit', year: 'numeric' }) + ' ' + d.toLocaleTimeString('it-IT', { hour: '2-digit', minute: '2-digit' });
              }
            }
          }
        },
        scales: {
          y: {
            min: 0,
            max: yMaxVal,
            position: 'left',
            grid: { color: 'rgba(255,255,255,0.05)', lineWidth: 0.3 },
            ticks: {
              stepSize: 10,
              autoSkip: false,
              color: '#94a3b8',
              font: { weight: '400', size: 11 },
              callback: (value) => {
                const v = Math.round(value);
                // Nasconde 70 e 180 da sinistra
                if (v === GLUCOSE_MIN || v === GLUCOSE_MAX) return null;
                if (v === 0 || v === 400 || v % 50 === 0) return v;
                return null;
              }
            }
          },
          y1: {
            min: 0,
            max: yMaxVal,
            position: 'right',
            grid: { display: false },
            ticks: {
              stepSize: 10,
              autoSkip: false,
              align: 'center',
              crossAlign: 'far',
              padding: 0,
              color: '#fff',
              font: { weight: '500', size: 12 },
              callback: (value) => {
                const v = Math.round(value);
                // Mostra SOLO 70 e 180 a destra
                if (v === GLUCOSE_MIN || v === GLUCOSE_MAX) return v;
                return null;
              }
            }
          },
          x: {
            type: 'linear', min: xMin, max: xMax,
            grid: { color: 'rgba(255,255,255,0.06)', lineWidth: 0.3 },
            ticks: {
              color: '#94a3b8',
              maxRotation: 90,
              minRotation: 90,
              autoSkip: true,
              maxTicksLimit: this.chartScope === 'day' ? 25 : 20,
              stepSize: this.chartScope === 'day' ? 3600000 : undefined,
              callback: (value) => {
                const d = new Date(value);
                if (this.chartScope === 'day') {
                  // Mostriamo solo le ore piene per pulizia ed evitare sovrapposizioni
                  if (d.getMinutes() === 0) {
                    return d.toLocaleTimeString('it-IT', { hour: '2-digit', minute: '2-digit' });
                  }
                  return null;
                }
                return d.toLocaleDateString('it-IT', { day: '2-digit', month: '2-digit', year: 'numeric' });
              }
            }
          }
        }
      }
    });
    const { entryDay, entryMonth, entryYear, entryHH, entryMM } = this.elements;
    if (entryDay && entryMonth && entryYear && entryHH && entryMM) {
      entryDay.oninput = (e) => { if (e.target.value.length >= 2) entryMonth.focus(); };
      entryMonth.oninput = (e) => { if (e.target.value.length >= 2) entryYear.focus(); };
      entryYear.oninput = (e) => { if (e.target.value.length >= 4) entryHH.focus(); };
      entryHH.oninput = (e) => { if (e.target.value.length >= 2) entryMM.focus(); };
    }
  }

  // ===== CRUD =====
  async saveEntry() {
    const form = this.elements.entryForm;
    const type = form.type.value;
    const day = parseInt(this.elements.entryDay.value) || 1;
    const month = parseInt(this.elements.entryMonth.value) || 1;
    const year = parseInt(this.elements.entryYear.value) || new Date().getFullYear();
    const hours = parseInt(this.elements.entryHH.value) || 0;
    const minutes = parseInt(this.elements.entryMM.value) || 0;
    const timestamp = new Date(year, month - 1, day, hours, minutes, 0, 0).getTime();

    const entryData = {
      type,
      value: this.elements.entryValue.value,
      unit: this.elements.unitInput.value,
      medName: type === 'med' ? this.elements.medName.value : '',
      note: this.elements.entryNote.value,
      timestamp
    };

    try {
      if (this.editingId) {
        await this.userCol().doc(this.editingId).update(entryData);
        this.editingId = null;
      } else {
        const id = Date.now().toString();
        await this.userCol().doc(id).set({ id, ...entryData });
      }
      form.reset();
      this.elements.entryModal.style.display = 'none';
      this.updateMedsDataList();
      this.updateUnitsDataList(type);
    } catch (e) {
      console.error(e);
      this.showToast('Errore durante il salvataggio', 'error');
    }
  }

  async deleteEntry(id) {
    if (confirm('Sei sicuro di voler eliminare questo dato?')) {
      try {
        await this.userCol().doc(id).delete();
      } catch (e) {
        console.error(e);
        this.showToast('Errore durante l\'eliminazione', 'error');
      }
    }
  }

  async deleteAllEntries() {
    if (this.entries.length === 0) {
      this.showToast('Non ci sono dati da eliminare.', 'info');
      return;
    }
    const count = this.entries.length;
    const confirmed = confirm(`SEI SICURO? Questa azione cancellerà DEFINITIVAMENTE tutti i ${count} dati salvati.\n\nNon sarà possibile annullare questa operazione.`);
    if (!confirmed) return;
    const secondConfirm = confirm('CONFERMA DEFINITIVA: Vuoi davvero svuotare tutto lo storico?');
    if (!secondConfirm) return;
    try {
      this.showToast('Eliminazione in corso...', 'info');
      const chunkSize = 400;
      const allEntries = [...this.entries];
      for (let i = 0; i < allEntries.length; i += chunkSize) {
        const chunk = allEntries.slice(i, i + chunkSize);
        const batch = db.batch();
        chunk.forEach(entry => {
          const ref = this.userCol().doc(entry.id);
          batch.delete(ref);
        });
        await batch.commit();
      }
      this.showToast('Tutti i dati sono stati eliminati.');
    } catch (e) {
      console.error(e);
      this.showToast('Errore durante l\'eliminazione', 'error');
    }
  }

  // ===== IMPORT / EXPORT =====
  exportToXLSXLegacy() {
    if (this.entries.length === 0) { alert('Nessun dato da esportare.'); return; }
    const rows = this.entries.map(e => {
      const d = new Date(e.timestamp);
      return {
        Data: d.toLocaleDateString('it-IT'),
        Orario: d.toLocaleTimeString('it-IT', { hour: '2-digit', minute: '2-digit' }),
        Tipo: e.type === 'glucose' ? 'Glicemia' : 'Farmaco',
        'Valore/Dose': e.value ?? '',
        'Unità': e.unit || (e.type === 'glucose' ? 'mg/dL' : ''),
        Farmaco: e.type === 'glucose' ? '' : (e.medName || ''),
        Note: e.note || ''
      };
    });
    const worksheet = XLSX.utils.json_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    worksheet['!cols'] = [
      { wch: 12 },
      { wch: 9 },
      { wch: 12 },
      { wch: 14 },
      { wch: 12 },
      { wch: 24 },
      { wch: 36 }
    ];
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Storico');
    XLSX.writeFile(workbook, `terapia_export_${new Date().toISOString().slice(0, 10)}.xlsx`);
    this.showToast('Esportazione XLSX completata');
  }

  downloadHistoryXLSX() {
    if (this.entries.length === 0) { alert('Nessun dato da esportare.'); return; }

    const rows = [...this.entries]
      .sort((a, b) => b.timestamp - a.timestamp)
      .map((entry) => {
        const date = new Date(entry.timestamp);
        return {
          Data: date.toLocaleDateString('it-IT'),
          Orario: date.toLocaleTimeString('it-IT', { hour: '2-digit', minute: '2-digit' }),
          Tipo: entry.type === 'glucose' ? 'Glicemia' : 'Farmaco',
          'Valore/Dose': entry.value ?? '',
          Unita: entry.unit || (entry.type === 'glucose' ? 'mg/dL' : ''),
          Farmaco: entry.type === 'glucose' ? '' : (entry.medName || ''),
          Note: entry.note || ''
        };
      });

    const worksheet = XLSX.utils.json_to_sheet(rows);
    worksheet['!cols'] = [
      { wch: 12 },
      { wch: 9 },
      { wch: 12 },
      { wch: 14 },
      { wch: 12 },
      { wch: 24 },
      { wch: 36 }
    ];

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Storico');
    XLSX.writeFile(workbook, `terapia_export_${new Date().toISOString().slice(0, 10)}.xlsx`);
    this.showToast('Esportazione XLSX completata');
  }

  async handleImport(event) {
    const file = event.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { 
          type: 'array', 
          cellDates: true,
          cellNF: false,
          cellText: false
        });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
        if (jsonData.length === 0) { this.showToast('Il file è vuoto', 'error'); return; }
        
        const importedEntries = this.mapJsonToEntries(jsonData);
        if (importedEntries.length > 0) {
          for (let i = 0; i < importedEntries.length; i += 400) {
            const chunk = importedEntries.slice(i, i + 400);
            const batch = db.batch();
            chunk.forEach(entry => {
              const ref = this.userCol().doc(entry.id);
              batch.set(ref, entry);
            });
            await batch.commit();
          }
          this.showToast(`${importedEntries.length} nuovi dati importati!`);
        } else {
          // Controlliamo se jsonData aveva righe ma importedEntries è 0
          const msg = jsonData.length > 0 ? 
            'Nessun nuovo dato compatibile trovato (forse duplicati o colonne errate).' : 
            'Nessun dato trovato nel file.';
          this.showToast(msg, 'info');
        }
      } catch (err) {
        console.error(err);
        this.showToast('L\'importazione è fallita.', 'error');
      }
      this.elements.importInput.value = '';
    };
    reader.readAsArrayBuffer(file);
  }

  mapJsonToEntries(data) {
    const results = [];
    const seen = new Set();
    const existing = new Set(this.entries.map(e => this.getFingerprint(e)));

    data.forEach((row, index) => {
      let glucose = null, medName = null, dVal = null, note = '', unit = '';
      let dateObj = new Date(), hasDate = false, hasTime = false, typeHint = null;

      // Piccolo offset per evitare fingerprint identici se manca l'orario preciso
      dateObj.setMilliseconds(index);

      Object.entries(row).forEach(([key, val]) => {
        if (val === "" || val === null || val === undefined) return;
        const k = key.toLowerCase().trim();
        const v = val.toString().trim();

        if (k.includes('tipo')) {
          if (v.toLowerCase().includes('glic')) typeHint = 'glucose';
          else if (v.toLowerCase().includes('farm')) typeHint = 'med';
        } else if (k.includes('data') || k.includes('giorno') || k.includes('date')) {
          const d = this.parseExcelDate(val);
          if (d) { dateObj.setFullYear(d.getFullYear(), d.getMonth(), d.getDate()); hasDate = true; }
        } else if (k.includes('ora') || k.includes('time') || k.includes('orario')) {
          const t = this.parseExcelTime(val);
          if (t) { dateObj.setHours(t.getHours(), t.getMinutes(), 0, 0); hasTime = true; }
        } else if (k.includes('farm') || k.includes('med') || (k.includes('nome') && !k.includes('utente') && v.length > 1)) {
          medName = v;
        } else if (k.includes('glic') || k.includes('glucos') || (k.includes('valore') && !k.includes('dose'))) {
          const n = parseFloat(v.replace(',', '.'));
          if (!isNaN(n)) glucose = n;
        } else if (k.includes('dose') || k.includes('quant')) {
          const n = parseFloat(v.replace(',', '.'));
          if (!isNaN(n)) dVal = n;
        } else if (k.includes('unit')) {
          unit = v;
        } else if (k.includes('note') || k.includes('comm') || k.includes('nota')) {
          note = v;
        }
      });

      if (!hasDate) {
        Object.values(row).forEach(v => {
          if (v instanceof Date && !hasDate) { dateObj.setFullYear(v.getFullYear(), v.getMonth(), v.getDate()); hasDate = true; }
        });
      }

      if (!typeHint) {
        if (medName) typeHint = 'med';
        else if (glucose !== null) typeHint = 'glucose';
      }

      const timestamp = dateObj.getTime();
      let e = null;

      if (typeHint === 'glucose' && (glucose !== null || dVal !== null)) {
        e = {
          id: 'imp_' + timestamp + '_' + Math.random().toString(36).substr(2, 5),
          type: 'glucose', value: (glucose || dVal).toString(), unit: unit || 'mg/dL', medName: '', note, timestamp
        };
      } else if (typeHint === 'med' && medName) {
        e = {
          id: 'imp_' + timestamp + '_' + Math.random().toString(36).substr(2, 5) + '_m',
          type: 'med', value: (dVal || glucose || 1).toString(), unit: unit || (medName.toLowerCase().includes('insulina') ? 'UI' : 'ml'), medName, note, timestamp
        };
      }

      if (e) {
        const fp = this.getFingerprint(e);
        if (!existing.has(fp) && !seen.has(fp)) {
          results.push(e);
          seen.add(fp);
        }
      }
    });
    return results;
  }

    getMedBadgeStyle(medName) {
    const name = (medName || '').toLowerCase();
    
    // Insuline - Rosa (Vibrante)
    if (name.includes('toujeo') || name.includes('insul') || name.includes('lantus') || name.includes('humalog')) {
      return { bg: 'rgba(236, 72, 153, 0.12)', color: '#f9a8d4', border: 'rgba(236, 72, 153, 0.2)' };
    }
    
    // Antidiabetici orali (SGLT2i / Altri) - Arancio/Ambra
    if (name.includes('invokana') || name.includes('jardiance') || name.includes('forxiga')) {
      return { bg: 'rgba(245, 158, 11, 0.12)', color: '#fcd34d', border: 'rgba(245, 158, 11, 0.2)' };
    }
    
    // Mounjaro / GLP-1 - Smeraldo (Salute metabolica)
    if (name.includes('mounjaro') || name.includes('ozempic') || name.includes('wegovy') || name.includes('trulicity')) {
      return { bg: 'rgba(16, 185, 129, 0.15)', color: '#6ee7b7', border: 'rgba(16, 185, 129, 0.25)' };
    }
    
    // Tovastibe / Statine / Colesterolo - Ciano/Azzurro
    if (name.includes('tovastibe') || name.includes('tavastibe') || name.includes('atorvast') || name.includes('ezetim')) {
       return { bg: 'rgba(6, 182, 212, 0.15)', color: '#67e8f9', border: 'rgba(6, 182, 212, 0.25)' };
    }

    // Metformina e simili - Viola
    if (name.includes('metformin') || name.includes('janumet')) {
      return { bg: 'rgba(139, 92, 246, 0.12)', color: '#c4b5fd', border: 'rgba(139, 92, 246, 0.2)' };
    }
    
    // Default - Indaco
    return { bg: 'rgba(99, 102, 241, 0.12)', color: '#a5b4fc', border: 'rgba(99, 102, 241, 0.2)' };
  }

  getFingerprint(e) {
    return `${e.timestamp}_${e.type}_${e.value}_${e.medName || ''}`;
  }

parseExcelDate(val) {
    if (val instanceof Date) return val;
    if (typeof val === 'number') return new Date(Math.round((val - 25569) * 86400 * 1000));
    if (typeof val === 'string') {
      const parts = val.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
      if (parts) {
        return new Date(parseInt(parts[3]), parseInt(parts[2]) - 1, parseInt(parts[1]));
      }
    }
    const d = new Date(val);
    return isNaN(d.getTime()) ? null : d;
  }

  parseExcelTime(val) {
    if (val instanceof Date) return val;
    if (typeof val === 'string') {
      const match = val.match(/(\d{1,2})[:.](\d{2})/);
      if (match) { const d = new Date(); d.setHours(parseInt(match[1]), parseInt(match[2]), 0, 0); return d; }
    }
    if (typeof val === 'number' && val < 1) {
      const totalSeconds = Math.round(val * 24 * 60 * 60);
      const d = new Date(); d.setHours(Math.floor(totalSeconds / 3600), Math.floor((totalSeconds % 3600) / 60), 0, 0);
      return d;
    }
    const d = new Date(val);
    return isNaN(d.getTime()) ? null : d;
  }

  showToast(message, type = 'success') {
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.innerHTML = `
      <i data-lucide="${type === 'success' ? 'check-circle' : 'alert-circle'}" style="width: 18px;"></i>
      <span>${message}</span>
    `;
    document.body.appendChild(toast);
    lucide.createIcons();
    setTimeout(() => toast.classList.add('show'), 100);
    setTimeout(() => { toast.classList.remove('show'); setTimeout(() => toast.remove(), 300); }, 3000);
  }
}

document.addEventListener('DOMContentLoaded', () => {
  window.app = new TerapiaApp();

  if ('serviceWorker' in navigator) {
    let refreshing = false;

    navigator.serviceWorker.addEventListener('controllerchange', () => {
      if (refreshing) return;
      refreshing = true;
      window.location.reload();
    });

    navigator.serviceWorker.register('./service-worker.js', { updateViaCache: 'none' })
      .then((registration) => registration.update())
      .catch((error) => {
        console.error('Service worker registration failed:', error);
      });
  }
});
