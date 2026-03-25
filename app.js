/**
 * Terapia App - Core Logic
 * Handles data management, date navigation and UI updates.
 */

class TerapiaApp {
  constructor() {
    this.currentDate = new Date();
    this.currentFilter = 'all';
    this.entries = JSON.parse(localStorage.getItem('terapia_data')) || [];
    
    // UI Elements
    this.elements = {
      datePickerList: document.getElementById('datePickerList'),
      addEntryBtn: document.getElementById('addEntryBtn'),
      importBtn: document.getElementById('importBtn'),
      importInput: document.getElementById('importInput'),
      viewAllBtn: document.getElementById('viewAllBtn'),
      exportAllBtn: document.getElementById('exportAllBtn'),
      logTitle: document.getElementById('logTitle'),
      entryModal: document.getElementById('entryModal'),
      entryForm: document.getElementById('entryForm'),
      filterBtns: document.querySelectorAll('.filter-btns button'),
      logTbody: document.getElementById('log-tbody'),
      avgGlucose: document.getElementById('avgGlucose'),
      medsCount: document.getElementById('medsCount'),
      statusIndicator: document.getElementById('statusIndicator'),
      closeModalBtns: document.querySelectorAll('.close-modal'),
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
      glucoseChart: null
    };

    this.viewAll = true; // Default to showing everything
    this.chartScope = 'all'; // Default to showing all history in chart
    this.editingId = null;
    this.init();
  }

  init() {
    this.setupEventListeners();
    this.updateMedsDataList();
    this.loadFromRepository().then(() => {
      this.updateUI();
    });
  }

  setupEventListeners() {
    this.elements.addEntryBtn.onclick = () => {
      this.editingId = null;
      this.elements.modalTitle.innerText = "Nuovo Inserimento";
      this.elements.modalSub.innerText = "Aggiungi una lettura o un'assunzione";
      this.elements.submitBtn.innerText = "Salva Dato";
      this.elements.entryForm.reset();
      this.elements.entryModal.style.display = 'flex';
      
      // Sync form with current dashboard date and time
      const dateStr = this.currentDate.toISOString().split('T')[0];
      document.getElementById('entryDate').value = dateStr;
      
      const now = new Date();
      this.elements.entryForm.entryTime.value = now.toLocaleTimeString('it-IT', { hour: '2-digit', minute: '2-digit', hour12: false });
    };

    this.elements.viewAllBtn.onclick = () => {
      this.viewAll = !this.viewAll;
      this.updateViewAllButton();
      this.renderLog();
    };

    this.updateViewAllButton();

    this.elements.exportAllBtn.onclick = () => this.exportToCSV();

    this.elements.importBtn.onclick = () => this.elements.importInput.click();
    this.elements.importInput.onchange = (e) => this.handleImport(e);

    this.elements.filterBtns.forEach(btn => {
      btn.onclick = () => {
        this.elements.filterBtns.forEach(b => b.classList.remove('primary'));
        this.elements.filterBtns.forEach(b => b.classList.add('secondary'));
        btn.classList.remove('secondary');
        btn.classList.add('primary');
        this.currentFilter = btn.dataset.filter;
        this.renderLog();
      };
    });

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
        this.elements.chartSub.innerText = this.chartScope === 'day' ? "Valori rilevati oggi" : "Storico completo dei dati";
        this.updateDashboard();
      };
    });

    this.elements.entryForm.onsubmit = (e) => {
      e.preventDefault();
      this.saveEntry();
    };

    // Close modal on click outside
    window.onclick = (e) => {
      if (e.target === this.elements.entryModal) {
        this.elements.entryModal.style.display = 'none';
      }
    };
    this.elements.syncRepoBtn.onclick = () => {
      if (confirm('Vuoi aggiornare i dati dal file data.json? Questo unirà i nuovi dati senza cancellare i tuoi.')) {
        this.loadFromRepository(true);
      }
    };

    // Right-click sync button to export
    this.elements.syncRepoBtn.oncontextmenu = (e) => {
      e.preventDefault();
      if (confirm('Vuoi generare una nuova versione di data.json da salvare nel repository?')) {
        this.exportToRepository();
      }
    };
  }

  toggleEntryType(type) {
    if (type === 'glucose') {
      this.elements.medNameGroup.style.display = 'none';
      this.elements.valueLabel.innerText = 'Valore Glicemia';
      this.elements.entryValue.placeholder = 'es. 110';
      this.elements.unitInput.value = 'mg/dL';
      this.elements.unitsList.innerHTML = '<option value="mg/dL"></option>';
      this.elements.entryValue.required = true;
      this.elements.medName.required = false;
    } else {
      this.elements.medNameGroup.style.display = 'block';
      this.elements.valueLabel.innerText = 'Dose / Quantità Farmaco';
      this.elements.entryValue.placeholder = 'es. 1';
      this.elements.unitInput.value = 'ml'; // Default
      this.elements.unitsList.innerHTML = `
        <option value="ml"></option>
        <option value="unità"></option>
        <option value="UI"></option>
      `;
      this.elements.entryValue.required = false;
      this.elements.medName.required = true;
    }
  }

  changeDay(offset) {
    this.currentDate.setDate(this.currentDate.getDate() + offset);
    this.updateUI();
  }

  updateViewAllButton() {
    this.elements.viewAllBtn.classList.toggle('primary', this.viewAll);
    this.elements.viewAllBtn.classList.toggle('secondary', !this.viewAll);
    this.elements.viewAllBtn.innerHTML = this.viewAll ? 
      `<i data-lucide="eye"></i> Solo Oggi` : 
      `<i data-lucide="calendar"></i> Tutti i Dati`;
    this.elements.logTitle.innerText = this.viewAll ? "Storico Completo" : "Cronologia Giornaliera";
    lucide.createIcons();
  }

  updateMedsDataList() {
    if (!this.elements.medsList) return;
    
    // Get unique med names from entries
    const names = [...new Set(this.entries
      .filter(e => e.type === 'med' && e.medName)
      .map(e => e.medName)
    )].sort();
    
    this.elements.medsList.innerHTML = names
      .map(name => `<option value="${name}">`)
      .join('');
  }

  renderDateList() {
    if (!this.elements.datePickerList) return;
    
    const dateSet = new Set();
    const today = new Date();
    today.setHours(0,0,0,0);

    for(let i = -7; i <= 7; i++) {
      const d = new Date(today);
      d.setDate(d.getDate() + i);
      dateSet.add(d.toDateString());
    }

    this.entries.forEach(e => {
      dateSet.add(new Date(e.timestamp).toDateString());
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
    let filteredEntries = this.viewAll ? [...this.entries] : this.getEntriesForDate(this.currentDate);
    
    // Apply type filter
    if (this.currentFilter === 'glucose') {
      filteredEntries = filteredEntries.filter(e => e.type === 'glucose');
    } else if (this.currentFilter === 'meds') {
      filteredEntries = filteredEntries.filter(e => e.type === 'med');
    }

    // Sort by time (reverse chronological)
    filteredEntries.sort((a, b) => b.timestamp - a.timestamp);

    this.elements.logTbody.innerHTML = '';
    
    if (filteredEntries.length === 0) {
      this.elements.logTbody.innerHTML = `<tr><td colspan="5" style="text-align: center; color: var(--text-secondary); padding: 2rem;">Nessun dato trovato.</td></tr>`;
      return;
    }

    let lastDateTag = "";
    filteredEntries.forEach(entry => {
      const entryDate = new Date(entry.timestamp);
      const dateTag = entryDate.toLocaleDateString('it-IT', { weekday: 'long', day: 'numeric', month: 'long' });
      
      // Add Date Header if it's the first entry for this date in the view
      if (this.viewAll && dateTag !== lastDateTag) {
        const headerTr = document.createElement('tr');
        headerTr.innerHTML = `
          <td colspan="6" style="background: rgba(255,255,255,0.02); padding: 0.6rem 1.5rem; text-transform: uppercase; font-size: 0.725rem; letter-spacing: 0.05em; color: var(--accent-primary); font-weight: 800;">
            ${dateTag}
          </td>
        `;
        this.elements.logTbody.appendChild(headerTr);
        lastDateTag = dateTag;
      }

      const time = entryDate.toLocaleTimeString('it-IT', { hour: '2-digit', minute: '2-digit' });
      const tr = document.createElement('tr');
      tr.draggable = true; 
      tr.classList.add('draggable-row');
      tr.dataset.id = entry.id;
      
      const isGlucose = entry.type === 'glucose';
      
      tr.innerHTML = `
        <td style="color: var(--text-secondary); font-size: 0.9rem; text-align: center;">${entryDate.toLocaleDateString('it-IT', { day: '2-digit', month: '2-digit' })}</td>
        <td style="font-weight: 600; text-align: center;">${time}</td>
        <td><span class="badge ${isGlucose ? 'badge-glucose' : 'badge-pill'}">${isGlucose ? 'Glicemia' : 'Farmaco'}</span></td>
        <td>
          <span style="font-weight: 700; font-size: 1.1rem;">${entry.value}</span> 
          <span style="color: var(--text-secondary); font-size: 0.825rem;">${entry.unit || (isGlucose ? 'mg/dL' : 'ml')}</span>
          ${!isGlucose ? `<span style="color: var(--text-muted); font-size: 0.8rem; margin-left: 0.5rem; background: rgba(255,255,255,0.03); padding: 0.2rem 0.5rem; border-radius: 6px; border: 1px solid var(--glass-border);">${entry.medName}</span>` : ''}
        </td>
        <td style="color: var(--text-secondary); font-size: 0.9rem;">${entry.note || '-'}</td>
        <td style="text-align: right;">
          <div style="display: flex; gap: 0.25rem; justify-content: flex-end; align-items: center;">
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

    // Re-initialize icons for new rows
    lucide.createIcons();

    // Setup drag and drop
    this.setupDragAndDrop();

    // Setup action buttons
    document.querySelectorAll('.edit-btn').forEach(btn => {
      btn.onclick = (e) => this.editEntry(btn.dataset.id);
    });
    
    document.querySelectorAll('.delete-btn').forEach(btn => {
      btn.onclick = (e) => this.deleteEntry(btn.dataset.id);
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

      row.addEventListener('dragend', (e) => {
        row.style.opacity = '1';
        rows.forEach(r => r.classList.remove('drag-over'));
      });

      row.addEventListener('dragover', (e) => {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';
        return false;
      });

      row.addEventListener('dragenter', (e) => {
        row.classList.add('drag-over');
      });

      row.addEventListener('dragleave', (e) => {
        row.classList.remove('drag-over');
      });

      row.addEventListener('drop', (e) => {
        e.stopPropagation();
        const targetId = row.dataset.id;
        if (draggedId !== targetId) {
          this.reorderEntries(draggedId, targetId);
        }
        return false;
      });
    });
  }

  reorderEntries(sourceId, targetId) {
    // Get the current list (preserving the sort order used for rendering)
    let list = this.viewAll ? [...this.entries] : this.getEntriesForDate(this.currentDate);
    list.sort((a, b) => b.timestamp - a.timestamp); // Most recent first
    
    const sourceIdx = list.findIndex(e => e.id === sourceId);
    const targetIdx = list.findIndex(e => e.id === targetId);
    
    if (sourceIdx === -1 || targetIdx === -1) return;

    // Move in local list
    const [movedEntry] = list.splice(sourceIdx, 1);
    list.splice(targetIdx, 0, movedEntry);
    
    // Calculate new timestamp to persist the order
    const prev = list[targetIdx - 1];
    const next = list[targetIdx + 1];
    
    let newTimestamp;
    if (prev && next) {
      // Put exactly in the middle
      newTimestamp = Math.floor((prev.timestamp + next.timestamp) / 2);
    } else if (prev) {
      // Put right below prev (slightly older)
      newTimestamp = prev.timestamp - 1000;
    } else if (next) {
      // Put right above next (slightly newer)
      newTimestamp = next.timestamp + 1000;
    } else {
      newTimestamp = movedEntry.timestamp;
    }

    // Find in main entries and update
    const mainIdx = this.entries.findIndex(e => e.id === sourceId);
    if (mainIdx !== -1) {
      this.entries[mainIdx].timestamp = newTimestamp;
    }
    
    localStorage.setItem('terapia_data', JSON.stringify(this.entries));
    this.renderLog();
  }

  editEntry(id) {
    const entry = this.entries.find(e => e.id === id);
    if (!entry) return;

    this.editingId = id;
    this.elements.modalTitle.innerText = "Modifica Dato";
    this.elements.modalSub.innerText = "Aggiorna le informazioni dell'inserimento";
    this.elements.submitBtn.innerText = "Aggiorna Dato";

    // Fill form
    const form = this.elements.entryForm;
    form.type.value = entry.type;
    this.toggleEntryType(entry.type);
    
    this.elements.entryValue.value = entry.value;
    this.elements.unitInput.value = entry.unit || (entry.type === 'glucose' ? 'mg/dL' : 'ml');
    this.elements.medName.value = entry.medName || '';
    document.getElementById('entryNote').value = entry.note || '';
    
    const d = new Date(entry.timestamp);
    document.getElementById('entryDate').value = d.toISOString().split('T')[0];
    form.entryTime.value = d.toLocaleTimeString('it-IT', { hour: '2-digit', minute: '2-digit', hour12: false });
    
    this.elements.entryModal.style.display = 'flex';
  }

  updateDashboard() {
    const dataSource = this.viewAll ? this.entries : this.getEntriesForDate(this.currentDate);
    const glucoseEntries = dataSource.filter(e => e.type === 'glucose');
    const medEntries = dataSource.filter(e => e.type === 'med');

    // Glucose Avg
    if (glucoseEntries.length > 0) {
      const avg = Math.round(glucoseEntries.reduce((acc, curr) => acc + parseFloat(curr.value || 0), 0) / glucoseEntries.length);
      this.elements.avgGlucose.innerText = avg;
      this.elements.avgGlucose.style.color = avg > 140 ? '#f43f5e' : (avg < 70 ? '#f59e0b' : '#10b981');
    } else {
      this.elements.avgGlucose.innerText = '--';
      this.elements.avgGlucose.style.color = 'var(--text-secondary)';
    }

    // Meds Count
    this.elements.medsCount.innerText = medEntries.length;

    // Status Indicator logic
    const dailyEntries = this.getEntriesForDate(this.currentDate);
    if (dailyEntries.length === 0) {
      this.elements.statusIndicator.innerText = 'Dati Assenti';
      this.elements.statusIndicator.style.color = 'var(--text-secondary)';
    } else {
      const dailyGlucose = dailyEntries.filter(e => e.type === 'glucose');
      const dailyAvg = dailyGlucose.length > 0 ? Math.round(dailyGlucose.reduce((acc, curr) => acc + parseFloat(curr.value || 0), 0) / dailyGlucose.length) : null;
      if (dailyAvg && (dailyAvg > 180 || dailyAvg < 70)) {
        this.elements.statusIndicator.innerText = 'Attenzione';
        this.elements.statusIndicator.style.color = '#f43f5e';
      } else {
        this.elements.statusIndicator.innerText = 'Ottimo';
        this.elements.statusIndicator.style.color = '#10b981';
      }
    }

    // Update scope buttons state
    this.elements.chartScopeBtns.forEach(btn => {
      const isActive = btn.dataset.scope === this.chartScope;
      btn.classList.toggle('primary', isActive);
      btn.classList.toggle('secondary', !isActive);
    });

    const chartData = this.chartScope === 'all' ? 
      this.entries.filter(e => e.type === 'glucose') : 
      dailyEntries.filter(e => e.type === 'glucose');

    this.renderChart(chartData);
  }

  renderChart(glucoseEntries) {
    if (!this.elements.glucoseChartCanvas) return;
    const ctx = this.elements.glucoseChartCanvas.getContext('2d');
    
    // Sort oldest first for chart
    const data = [...glucoseEntries].sort((a, b) => a.timestamp - b.timestamp);
    
    if (this.elements.glucoseChart) {
      this.elements.glucoseChart.destroy();
    }

    if (data.length === 0 && this.chartScope !== 'day') {
      return;
    }

    const chartPoints = data.map(e => ({
      x: e.timestamp,
      y: parseFloat(e.value)
    }));

    // Calculate bounds for X axis
    let xMin, xMax;
    if (this.chartScope === 'day') {
      const d = new Date(this.currentDate);
      xMin = new Date(d.getFullYear(), d.getMonth(), d.getDate(), 0, 0, 0).getTime();
      xMax = new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1, 0, 0, 0).getTime();
    } else if (data.length > 0) {
      const firstDate = new Date(Math.min(...data.map(e => e.timestamp)));
      const lastDate = new Date(Math.max(...data.map(e => e.timestamp)));
      xMin = new Date(firstDate.getFullYear(), firstDate.getMonth(), firstDate.getDate(), 0, 0, 0).getTime();
      xMax = new Date(lastDate.getFullYear(), lastDate.getMonth(), lastDate.getDate() + 1, 0, 0, 0).getTime();
    } else {
      xMin = Date.now() - 3600000;
      xMax = Date.now() + 3600000;
    }    // Create gradient
    const gradient = ctx.createLinearGradient(0, 0, 0, 300);
    gradient.addColorStop(0, 'rgba(16, 185, 129, 0.4)');
    gradient.addColorStop(1, 'rgba(16, 185, 129, 0)');

    this.elements.glucoseChart = new Chart(ctx, {
      type: 'line',
      data: {
        datasets: [{
          label: 'Glicemia (mg/dL)',
          data: chartPoints,
          borderColor: '#10b981',
          backgroundColor: gradient,
          borderWidth: 3,
          tension: 0,
          fill: true,
          pointBackgroundColor: '#fff',
          pointBorderColor: '#10b981',
          pointBorderWidth: 2,
          pointRadius: this.chartScope === 'all' ? 2 : 6,
          pointHoverRadius: this.chartScope === 'all' ? 4 : 8,
          pointHoverBackgroundColor: '#10b981',
          pointHoverBorderColor: '#fff',
          pointHoverBorderWidth: 2
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        animation: {
          duration: 1200,
          easing: 'easeInOutQuart'
        },
        interaction: {
          intersect: false,
          mode: 'index',
        },
        plugins: {
          legend: { display: false },
          tooltip: {
            backgroundColor: 'rgba(15, 23, 42, 0.9)',
            titleColor: '#fff',
            bodyColor: '#fff',
            borderColor: 'rgba(255, 255, 255, 0.1)',
            borderWidth: 1,
            padding: 12,
            boxPadding: 8,
            usePointStyle: true,
            callbacks: {
              label: (context) => `Glicemia: ${context.parsed.y} mg/dL`,
              title: (context) => {
                const d = new Date(context[0].parsed.x);
                return d.toLocaleDateString('it-IT') + ' ' + d.toLocaleTimeString('it-IT', { hour: '2-digit', minute: '2-digit' });
              }
            }
          }
        },
        scales: {
          y: {
            min: 60,
            grid: { color: 'rgba(255,255,255,0.05)' },
            ticks: { 
              color: '#94a3b8',
              font: { weight: '600' }
            }
          },
          x: {
            type: 'linear',
            min: xMin,
            max: xMax,
            bounds: 'ticks',
            grid: { color: 'rgba(255,255,255,0.06)' },
            ticks: { 
              color: '#94a3b8',
              maxRotation: 45,
              minRotation: 0,
              stepSize: this.chartScope === 'day' ? 3600000 * 2 : 3600000 * 24,
              callback: (value) => {
                const d = new Date(value);
                if (this.chartScope === 'day') {
                  return d.toLocaleTimeString('it-IT', { hour: '2-digit', minute: '2-digit' });
                } else {
                  return d.toLocaleDateString('it-IT', { day: '2-digit', month: '2-digit' });
                }
              }
            }
          }
        }
      }
    });
  }

  saveEntry() {
    const form = this.elements.entryForm;
    const type = form.type.value;
    const dateValue = document.getElementById('entryDate').value; // "YYYY-MM-DD"
    const timeValue = form.entryTime.value; // "HH:MM"
    
    // Create timestamp from selected date and time
    const [year, month, day] = dateValue.split('-').map(v => parseInt(v));
    const [hours, minutes] = timeValue.split(':').map(v => parseInt(v));
    
    const timestamp = new Date(year, month - 1, day, hours, minutes, 0, 0);

    const entryData = {
      type: type,
      value: this.elements.entryValue.value,
      unit: this.elements.unitInput.value,
      medName: type === 'med' ? this.elements.medName.value : '',
      note: document.getElementById('entryNote').value,
      timestamp: timestamp.getTime()
    };

    if (this.editingId) {
      const index = this.entries.findIndex(e => e.id === this.editingId);
      if (index !== -1) {
        this.entries[index] = { ...this.entries[index], ...entryData };
      }
      this.editingId = null;
    } else {
      this.entries.push({
        id: Date.now().toString(),
        ...entryData
      });
    }

    localStorage.setItem('terapia_data', JSON.stringify(this.entries));
    
    form.reset();
    this.elements.entryModal.style.display = 'none';
    this.updateMedsDataList();
    this.updateUI();
  }

  deleteEntry(id) {
    if (confirm('Sei sicuro di voler eliminare questo dato?')) {
      this.entries = this.entries.filter(e => e.id !== id);
      localStorage.setItem('terapia_data', JSON.stringify(this.entries));
      this.updateUI();
    }
  }

  exportToCSV() {
    if (this.entries.length === 0) {
      alert('Nessun dato da esportare.');
      return;
    }

    const headers = ['Data', 'Orario', 'Tipo', 'Valore/Dose', 'Note'];
    const rows = this.entries.map(e => {
      const d = new Date(e.timestamp);
      return [
        d.toLocaleDateString('it-IT'),
        d.toLocaleTimeString('it-IT', { hour: '2-digit', minute: '2-digit' }),
        e.type === 'glucose' ? 'Glicemia' : 'Farmaco',
        e.type === 'glucose' ? `${e.value} mg/dL` : `${e.value} (${e.medName})`,
        `"${e.note || ''}"`
      ];
    });

    const csvContent = [headers, ...rows].map(r => r.join(',')).join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    
    link.setAttribute('href', url);
    link.setAttribute('download', `terapia_export_${new Date().toISOString().slice(0, 10)}.csv`);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    this.showToast('Esportazione completata');
  }

  handleImport(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { 
          type: 'array',
          cellDates: true,
          cellNF: false,
          cellText: false
        });
        
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet);

        if (jsonData.length === 0) {
          this.showToast('Il file è vuoto', 'error');
          return;
        }

        const importedEntries = this.mapJsonToEntries(jsonData);
        if (importedEntries.length > 0) {
          this.entries = [...this.entries, ...importedEntries];
          localStorage.setItem('terapia_data', JSON.stringify(this.entries));
          this.updateUI();
          this.showToast(`${importedEntries.length} dati importati con successo!`);
        } else {
          this.showToast('Nessun dato compatibile trovato. Controlla le intestazioni delle colonne.', 'error');
        }
      } catch (err) {
        console.error(err);
        this.showToast('Errore durante l\'importazione', 'error');
      }
      this.elements.importInput.value = ''; // Reset input
    };
    reader.readAsArrayBuffer(file);
  }

  mapJsonToEntries(data) {
    const results = [];
    
    data.forEach(row => {
      let glucoseValue = null;
      let medName = null;
      let medDose = 1;
      let note = '';
      let dateObj = new Date();
      let hasDate = false;
      let hasTime = false;

      // Extract data from row keys
      Object.entries(row).forEach(([key, val]) => {
        if (!val && val !== 0) return;
        
        const k = key.toLowerCase().trim();
        
        // Value / Glucose
        if (k.includes('glic') || k.includes('glucos') || (k.includes('valore') && !k.includes('dose'))) {
          glucoseValue = parseFloat(val);
        } 
        // Medication
        else if (k.includes('farm') || k.includes('pastig') || k.includes('med') || k.includes('nome')) {
          medName = val.toString();
        }
        // Dose
        else if (k.includes('dose') || k.includes('quant')) {
          medDose = val;
        }
        // Notes
        else if (k.includes('note') || k.includes('comm')) {
          note = val.toString();
        }
        // Date handling
        else if (k.includes('data') || k.includes('giorno') || k.includes('date')) {
          const d = this.parseExcelDate(val);
          if (d) {
            dateObj.setFullYear(d.getFullYear(), d.getMonth(), d.getDate());
            hasDate = true;
          }
        }
        // Time handling
        else if (k.includes('ora') || k.includes('time')) {
          const t = this.parseExcelTime(val);
          if (t) {
            dateObj.setHours(t.getHours(), t.getMinutes(), 0, 0);
            hasTime = true;
          }
        }
      });

      // If we found a date/time in the same cell (common in Excel)
      if (!hasDate || !hasTime) {
        Object.values(row).forEach(val => {
          if (val instanceof Date) {
            if (!hasDate) {
              dateObj.setFullYear(val.getFullYear(), val.getMonth(), val.getDate());
              hasDate = true;
            }
            if (!hasTime && (val.getHours() !== 0 || val.getMinutes() !== 0)) {
              dateObj.setHours(val.getHours(), val.getMinutes(), 0, 0);
              hasTime = true;
            }
          }
        });
      }

      const timestamp = dateObj.getTime();

      // Create Glucose Entry
      if (glucoseValue !== null && !isNaN(glucoseValue)) {
        results.push({
          id: Date.now().toString() + Math.random().toString(36).substr(2, 9),
          type: 'glucose',
          value: glucoseValue,
          unit: 'mg/dL',
          medName: '',
          note: note,
          timestamp: timestamp
        });
      }

      // Create Medication Entry
      if (medName) {
        results.push({
          id: Date.now().toString() + Math.random().toString(36).substr(2, 9) + '_m',
          type: 'med',
          value: medDose,
          unit: 'ml', // Default for import
          medName: medName,
          note: glucoseValue !== null ? '' : note, // Only attach note to one if split
          timestamp: timestamp
        });
      }
    });

    return results;
  }

  parseExcelDate(val) {
    if (val instanceof Date) return val;
    if (typeof val === 'number') {
      // Excel serial date to JS Date
      return new Date(Math.round((val - 25569) * 86400 * 1000));
    }
    const d = new Date(val);
    return isNaN(d.getTime()) ? null : d;
  }

  parseExcelTime(val) {
    if (val instanceof Date) return val;
    if (typeof val === 'string') {
      const match = val.match(/(\d{1,2})[:.](\d{2})/);
      if (match) {
        const d = new Date();
        d.setHours(parseInt(match[1]), parseInt(match[2]), 0, 0);
        return d;
      }
    }
    if (typeof val === 'number' && val < 1) {
      // Excel serial time (fraction of day)
      const totalSeconds = Math.round(val * 24 * 60 * 60);
      const d = new Date();
      d.setHours(Math.floor(totalSeconds / 3600), Math.floor((totalSeconds % 3600) / 60), 0, 0);
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
    setTimeout(() => {
      toast.classList.remove('show');
      setTimeout(() => toast.remove(), 300);
    }, 3000);
  }

  loadDemoData() {
    if (this.entries.length > 0) return;

    const today = new Date();
    const demoEntries = [
      { id: '1', type: 'glucose', value: 105, unit: 'mg/dL', timestamp: new Date(today.setHours(8, 30)).getTime(), note: 'A digiuno' },
      { id: '2', type: 'med', value: 1, unit: 'unità', medName: 'Metformina 500mg', timestamp: new Date(today.setHours(8, 45)).getTime(), note: 'Dopo colazione' },
      { id: '3', type: 'glucose', value: 145, unit: 'mg/dL', timestamp: new Date(today.setHours(13, 15)).getTime(), note: 'Dopo pranzo' },
      { id: '4', type: 'med', value: 3, unit: 'ml', medName: 'Insulina', timestamp: new Date(today.setHours(13, 30)).getTime(), note: '' },
      { id: '5', type: 'glucose', value: 115, unit: 'mg/dL', timestamp: new Date(today.setHours(19, 45)).getTime(), note: 'Prima di cena' }
    ];

    this.entries = demoEntries;
    localStorage.setItem('terapia_data', JSON.stringify(this.entries));
    this.updateUI();
  }
  async loadFromRepository(showToast = false) {
    try {
      const response = await fetch('data.json?t=' + Date.now());
      if (response.ok) {
        const fileContent = await response.json();
        const initialCount = this.entries.length;
        
        // Merge entries, checking for duplicates by timestamp + value
        fileContent.forEach(fileEntry => {
          const exists = this.entries.some(e => 
            e.timestamp === fileEntry.timestamp && 
            e.type === fileEntry.type && 
            e.value == fileEntry.value
          );
          if (!exists) {
            this.entries.push(fileEntry);
          }
        });

        localStorage.setItem('terapia_data', JSON.stringify(this.entries));
        if (showToast) {
          this.showToast(`Sincronizzato: ${this.entries.length - initialCount} nuovi record.`);
          this.updateUI();
        }
      } else {
        if (showToast) this.showToast("Impossibile trovare data.json nel repository.", "error");
        if (this.entries.length === 0) this.loadDemoData();
      }
    } catch (e) {
      console.warn("SincronizzazioneRepository:", e);
      if (showToast) this.showToast("File data.json non accessibile.", "error");
      if (this.entries.length === 0) this.loadDemoData();
    }
  }

  exportToRepository() {
    try {
      const jsonContent = JSON.stringify(this.entries, null, 2);
      const blob = new Blob([jsonContent], { type: 'application/json' });
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      link.download = 'data.json';
      link.click();
      this.showToast("Database pronto. Salvalo nel repository sopra data.json");
    } catch (e) {
      this.showToast("Errore durante l'esportazione", "error");
    }
  }
}

// Start the app when the DOM is ready
document.addEventListener('DOMContentLoaded', () => {
  window.app = new TerapiaApp();
});
