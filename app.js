/**
 * Bulk Poster Generator Engine — Fixed Version
 * Canvas editor with Excel Merge Fields & ZIP batch export
 */

(function () {
  'use strict';

  // ──────────────────────────────────────────────────────────────────────────
  // Global Application State
  // ──────────────────────────────────────────────────────────────────────────
  const state = {
    canvas: null,
    ctx: null,
    bgImage: null,
    templateWidth: 1200,
    templateHeight: 630,
    zoomLevel: 1.0,

    mode: 'template',        // 'template' | 'preview'
    currentRecordIndex: 0,

    ticketFilter: 'all',     // 'all' | 'single' | 'multiple'
    modalTabFilter: 'all',   // 'all' | 'single' | 'multiple'
    modalSearchQuery: '',

    elements: [],            // All canvas elements
    selectedElementId: null,

    excelHeaders: [],
    excelRows: [],

    cancelExport: false,

    dragState: {
      isDragging: false,
      isResizing: false,
      startX: 0,
      startY: 0,
      elementStartX: 0,
      elementStartY: 0,
      elementStartW: 0,
      elementStartH: 0
    }
  };

  // ──────────────────────────────────────────────────────────────────────────
  // DOM References
  // ──────────────────────────────────────────────────────────────────────────
  const dom = {
    canvas:               document.getElementById('posterCanvas'),
    canvasWrapper:        document.getElementById('canvasWrapper'),
    canvasViewport:       document.getElementById('canvasViewport'),
    canvasOverlay:        document.getElementById('canvasOverlay'),

    templateUploadInput:  document.getElementById('templateUploadInput'),
    excelUploadInput:     document.getElementById('excelUploadInput'),
    btnLoadSample:        document.getElementById('btnLoadSample'),
    btnBatchExport:       document.getElementById('btnBatchExport'),

    btnAddText:           document.getElementById('btnAddText'),
    btnAddPhotoFrame:     document.getElementById('btnAddPhotoFrame'),
    mergeFieldsContainer: document.getElementById('mergeFieldsContainer'),
    excelStatusBadge:     document.getElementById('excelStatusBadge'),
    layersList:           document.getElementById('layersList'),

    btnPrevRecord:        document.getElementById('btnPrevRecord'),
    btnNextRecord:        document.getElementById('btnNextRecord'),
    btnGroupDuplicateNames: document.getElementById('btnGroupDuplicateNames'),
    filterTicketType:     document.getElementById('filterTicketType'),
    filteredCountBadge:   document.getElementById('filteredCountBadge'),
    recordDropdown:       document.getElementById('recordDropdown'),
    modeTemplateBtn:      document.getElementById('modeTemplateBtn'),
    modePreviewBtn:       document.getElementById('modePreviewBtn'),
    btnZoomIn:            document.getElementById('btnZoomIn'),
    btnZoomOut:           document.getElementById('btnZoomOut'),
    btnResetZoom:         document.getElementById('btnResetZoom'),
    zoomLevelDisplay:     document.getElementById('zoomLevel'),

    elementProperties:    document.getElementById('elementProperties'),
    noSelectionText:      document.getElementById('noSelectionText'),
    textGroup:            document.getElementById('textGroup'),
    propText:             document.getElementById('propText'),
    propFontFamily:       document.getElementById('propFontFamily'),
    propFontSize:         document.getElementById('propFontSize'),
    propAlign:            document.getElementById('propAlign'),
    propColor:            document.getElementById('propColor'),
    propColorText:        document.getElementById('propColorText'),
    propStrokeColor:      document.getElementById('propStrokeColor'),
    propStrokeColorText:  document.getElementById('propStrokeColorText'),
    propStrokeWidth:      document.getElementById('propStrokeWidth'),
    propShadowBlur:       document.getElementById('propShadowBlur'),
    propX:                document.getElementById('propX'),
    propY:                document.getElementById('propY'),
    btnBringForward:      document.getElementById('btnBringForward'),
    btnSendBackward:      document.getElementById('btnSendBackward'),
    btnDeleteElement:     document.getElementById('btnDeleteElement'),

    dataModal:            document.getElementById('dataModal'),
    btnViewDataModal:     document.getElementById('btnViewDataModal'),
    btnCloseDataModal:    document.getElementById('btnCloseDataModal'),
    btnCloseDataModal2:   document.getElementById('btnCloseDataModal2'),
    excelTableContainer:  document.getElementById('excelTableContainer'),
    modalCountAll:        document.getElementById('modalCountAll'),
    modalCountSingle:     document.getElementById('modalCountSingle'),
    modalCountMultiple:   document.getElementById('modalCountMultiple'),
    sideCountAll:         document.getElementById('sideCountAll'),
    sideCountSingle:      document.getElementById('sideCountSingle'),
    sideCountMultiple:    document.getElementById('sideCountMultiple'),
    modalSearchInput:     document.getElementById('modalSearchInput'),
    btnExportFilteredZip: document.getElementById('btnExportFilteredZip'),

    exportModal:          document.getElementById('exportModal'),
    exportStatusText:     document.getElementById('exportStatusText'),
    exportProgressBar:    document.getElementById('exportProgressBar'),
    btnCancelExport:      document.getElementById('btnCancelExport')
  };

  // ──────────────────────────────────────────────────────────────────────────
  // Initialisation
  // ──────────────────────────────────────────────────────────────────────────
  function init() {
    state.canvas = dom.canvas;
    state.ctx    = state.canvas.getContext('2d');

    setupEventListeners();
    loadSampleData();
  }

  // ──────────────────────────────────────────────────────────────────────────
  // Event Listeners
  // ──────────────────────────────────────────────────────────────────────────
  function setupEventListeners() {
    dom.templateUploadInput.addEventListener('change', handleTemplateUpload);
    dom.excelUploadInput.addEventListener('change', handleExcelUpload);
    dom.btnLoadSample.addEventListener('click', loadSampleData);
    dom.btnBatchExport.addEventListener('click', handleBatchExport);

    dom.btnAddText.addEventListener('click', () => addTextElement('អត្ថបទថ្មី'));
    dom.btnAddPhotoFrame.addEventListener('click', addPhotoFrameElement);

    dom.modeTemplateBtn.addEventListener('click', () => setMode('template'));
    dom.modePreviewBtn.addEventListener('click',  () => setMode('preview'));
    dom.btnPrevRecord.addEventListener('click',   () => navigateRecord(-1));
    dom.btnNextRecord.addEventListener('click',   () => navigateRecord(1));
    if (dom.btnGroupDuplicateNames) {
      dom.btnGroupDuplicateNames.addEventListener('click', () => {
        if (!state.excelRows.length) {
          alert('សូម Upload ឯកសារ Excel ជាមុនសិន!');
          return;
        }
        const countBefore = state.excelRows.length;
        state.excelRows = groupDuplicateRecords(state.excelRows, state.excelHeaders);
        const countAfter = state.excelRows.length;
        state.currentRecordIndex = 0;
        dom.excelStatusBadge.textContent = `${countAfter} Records (Grouped)`;
        updateRecordDropdown();
        renderExcelTableModal();
        renderCanvas();
        alert(`បង្រួមទិន្នន័យបានជោគជ័យ! (ពី ${countBefore} Records មកនៅ ${countAfter} Records)`);
      });
    }
    dom.recordDropdown.addEventListener('change', e => {
      if (e.target.value === 'template') {
        setMode('template');
      } else {
        state.currentRecordIndex = parseInt(e.target.value, 10);
        setMode('preview');
      }
    });

    dom.btnZoomIn.addEventListener('click',    () => changeZoom(0.1));
    dom.btnZoomOut.addEventListener('click',   () => changeZoom(-0.1));
    dom.btnResetZoom.addEventListener('click', fitZoomToScreen);

    // Inspector inputs — use 'input' + 'change' for cross-browser
    ['input', 'change'].forEach(evt => {
      dom.propText.addEventListener(evt,        onFormChange);
      dom.propFontFamily.addEventListener(evt,  onFormChange);
      dom.propFontSize.addEventListener(evt,    onFormChange);
      dom.propAlign.addEventListener(evt,       onFormChange);
      dom.propStrokeWidth.addEventListener(evt, onFormChange);
      dom.propShadowBlur.addEventListener(evt,  onFormChange);
      dom.propX.addEventListener(evt,           onFormChange);
      dom.propY.addEventListener(evt,           onFormChange);
    });

    dom.propColor.addEventListener('input', e => {
      dom.propColorText.value = e.target.value;
      onFormChange();
    });
    dom.propColorText.addEventListener('input', e => {
      if (/^#[0-9a-fA-F]{6}$/.test(e.target.value)) dom.propColor.value = e.target.value;
      onFormChange();
    });
    dom.propStrokeColor.addEventListener('input', e => {
      dom.propStrokeColorText.value = e.target.value;
      onFormChange();
    });
    dom.propStrokeColorText.addEventListener('input', e => {
      if (/^#[0-9a-fA-F]{6}$/.test(e.target.value)) dom.propStrokeColor.value = e.target.value;
      onFormChange();
    });

    dom.btnBringForward.addEventListener('click',  () => moveLayer(1));
    dom.btnSendBackward.addEventListener('click',  () => moveLayer(-1));
    dom.btnDeleteElement.addEventListener('click', deleteSelectedElement);

    dom.btnViewDataModal.addEventListener('click',    () => { renderExcelTableModal(); dom.dataModal.classList.remove('hidden'); });
    dom.btnCloseDataModal.addEventListener('click',   () => dom.dataModal.classList.add('hidden'));
    dom.btnCloseDataModal2.addEventListener('click',  () => dom.dataModal.classList.add('hidden'));
    dom.btnCancelExport.addEventListener('click',     () => { state.cancelExport = true; dom.exportModal.classList.add('hidden'); });

    // Sidebar Filter Buttons
    document.querySelectorAll('.btn-side-filter').forEach(btn => {
      btn.addEventListener('click', e => {
        const filterType = e.currentTarget.getAttribute('data-filter') || 'all';
        setTicketFilter(filterType);
      });
    });

    // Filter by Ticket Count Dropdown in Toolbar
    if (dom.filterTicketType) {
      dom.filterTicketType.addEventListener('change', e => {
        setTicketFilter(e.target.value);
      });
    }

    // Modal Real-time Search Input
    if (dom.modalSearchInput) {
      dom.modalSearchInput.addEventListener('input', e => {
        state.modalSearchQuery = e.target.value;
        renderExcelTableModal();
      });
    }

    // Modal Export Filtered ZIP
    if (dom.btnExportFilteredZip) {
      dom.btnExportFilteredZip.addEventListener('click', () => {
        dom.dataModal.classList.add('hidden');
        handleBatchExport(true);
      });
    }

    // Modal Filter Tab Buttons
    document.querySelectorAll('.filter-tab').forEach(tab => {
      tab.addEventListener('click', e => {
        const filterType = e.currentTarget.getAttribute('data-filter') || 'all';
        setTicketFilter(filterType);
      });
    });

    // Click on canvas viewport (deselect when clicking empty space)
    dom.canvasViewport.addEventListener('mousedown', e => {
      if (e.target === dom.canvasViewport || e.target === dom.canvasWrapper || e.target === dom.canvas) {
        state.selectedElementId = null;
        updateInspectorForm();
        renderOverlay();
      }
    });
  }

  // ──────────────────────────────────────────────────────────────────────────
  // Sample Template
  // ──────────────────────────────────────────────────────────────────────────
  function loadSampleData() {
    state.elements = [];
    state.selectedElementId = null;

    const sc = document.createElement('canvas');
    sc.width  = 1200;
    sc.height = 630;
    const s = sc.getContext('2d');

    // Parchment background
    const bg = s.createLinearGradient(0, 0, 1200, 630);
    bg.addColorStop(0,   '#fdfbf7');
    bg.addColorStop(0.5, '#f4ece1');
    bg.addColorStop(1,   '#e8d7c1');
    s.fillStyle = bg;
    s.fillRect(0, 0, 1200, 630);

    // Outer border
    s.strokeStyle = '#a17937'; s.lineWidth = 14;
    s.strokeRect(20, 20, 1160, 590);
    s.strokeStyle = '#d97706'; s.lineWidth = 4;
    s.strokeRect(32, 32, 1136, 566);

    // Title
    s.font      = "bold 38px 'Moul', serif";
    s.fillStyle = '#6b4712';
    s.textAlign = 'center';
    s.fillText('បុណ្យសឡាកភត្ត (ចាប់ឆ្នោត)', 600, 95);

    s.font      = "bold 32px 'Moul', serif";
    s.fillStyle = '#8d5b18';
    s.fillText('វត្តខេមររំសី (បឹងសាឡាង)', 600, 155);

    // White rounded name-field banner
    s.fillStyle   = 'rgba(255,255,255,0.75)';
    s.strokeStyle = '#a17937'; s.lineWidth = 3;
    roundRect(s, 450, 228, 650, 75, 30, true, true);

    // Labels
    s.font      = "bold 26px 'Moul', serif";
    s.fillStyle = '#5b360f';
    s.textAlign = 'right';
    s.fillText('នាមម្ចាស់ឆ្នោតលេខ', 1080, 215);

    s.textAlign = 'center';
    s.font      = "24px 'Moul', serif";
    s.fillText('ឆ្នោតលេខ', 250, 520);

    s.font      = "bold 22px 'Battambang', sans-serif";
    s.fillStyle = '#4a3111';
    s.fillText('កម្មវិធីប្រពៃណីធ្វើ', 800, 410);
    s.font      = "18px 'Battambang', sans-serif";
    s.fillText('ថ្ងៃសៅរ៍ ៨រោច ខែស្រាពណ៍ ឆ្នាំរោង អដ្ឋស័ក ព.ស. ២៥៧០', 800, 450);
    s.fillText('ត្រូវនឹងថ្ងៃទី០៥ ខែកញ្ញា ឆ្នាំ២០២៦', 800, 485);

    const img = new Image();
    img.onload = () => {
      state.bgImage        = img;
      state.templateWidth  = 1200;
      state.templateHeight = 630;
      resizeCanvas(1200, 630);

      const t = Date.now();
      state.elements = [
        {
          id: `el_${t}_1`, type: 'text',
          text: '{{ គោត្តនាម-នាម }}',
          x: 775, y: 268, width: 580, height: 60,
          fontFamily: "'Moul', Khmer, serif", fontSize: 30,
          color: '#5b360f', strokeColor: '#ffffff', strokeWidth: 0, shadowBlur: 0,
          align: 'center'
        },
        {
          id: `el_${t}_2`, type: 'text',
          text: '{{ ល.រ }}',
          x: 250, y: 568, width: 250, height: 50,
          fontFamily: "'Moul', Khmer, serif", fontSize: 36,
          color: '#d97706', strokeColor: '#ffffff', strokeWidth: 2, shadowBlur: 4,
          align: 'center'
        },
        {
          id: `el_${t}_3`, type: 'photoFrame',
          x: 250, y: 310, width: 180, height: 180, shape: 'circle'
        }
      ];

      state.excelHeaders = ['ល.រ', 'គោត្តនាម-នាម', 'លេខទូរស័ព្ទ', 'ព្រះសង្ឃ'];
      state.excelRows = [
        { 'ល.រ': '៦៩ - ៨៨',  'គោត្តនាម-នាម': 'ឯកឧត្តម ខែក សំអុន លោកជំទាវ ស៊ី វណ្ណថា', 'លេខទូរស័ព្ទ': '012832823', 'ព្រះសង្ឃ': '២០ អង្គ' },
        { 'ល.រ': '២៩ - ៣០',  'គោត្តនាម-នាម': 'ឧបាសិកា ចាន់ សុខេង',                   'លេខទូរស័ព្ទ': '089883685', 'ព្រះសង្ឃ': '២ អង្គ' },
        { 'ល.រ': '៥៥ - ៥៦',  'គោត្តនាម-នាម': 'លោកស្រី សុខ ស៊ីណាត',                   'លេខទូរស័ព្ទ': '011416069', 'ព្រះសង្ឃ': '២ អង្គ' },
        { 'ល.រ': '៦៣ - ៦៤',  'គោត្តនាម-នាម': 'ឧបាសក វ័រ គង់ ឧបាសិកា ឌី សុខ',          'លេខទូរស័ព្ទ': '070586505', 'ព្រះសង្ឃ': '២ អង្គ' },
        { 'ល.រ': '១',         'គោត្តនាម-នាម': 'ឧបាសក អ៊ុន សុផា',                       'លេខទូរស័ព្ទ': '0962216763', 'ព្រះសង្ឃ': '១ អង្គ' },
        { 'ល.រ': '២',         'គោត្តនាម-នាម': 'លោក បូរ និង ស៊ី រ៉ាត',                  'លេខទូរស័ព្ទ': '0964733670', 'ព្រះសង្ឃ': '១ អង្គ' },
        { 'ល.រ': '៣',         'គោត្តនាម-នាម': 'ឧបាសក សាន សន ឧបាសិកា ស៊ាន ក្រឹម',      'លេខទូរស័ព្ទ': '0964987762', 'ព្រះសង្ឃ': '១ អង្គ' },
        { 'ល.រ': '៤',         'គោត្តនាម-នាម': 'ឧបាសក ជា អន ឧបាសិកា ថេង កាក់',          'លេខទូរស័ព្ទ': '092767025', 'ព្រះសង្ឃ': '១ អង្គ' }
      ];

      dom.excelStatusBadge.textContent = `${state.excelRows.length} Sample Records`;
      dom.excelStatusBadge.classList.add('active');

      updateMergeFieldsList();
      updateRecordDropdown();
      renderExcelTableModal();
      renderLayersList();
      renderCanvas();
      fitZoomToScreen();
    };
    img.src = sc.toDataURL('image/png');
  }

  function roundRect(ctx, x, y, w, h, r, fill, stroke) {
    ctx.beginPath();
    ctx.moveTo(x + r, y);
    ctx.arcTo(x + w, y,     x + w, y + h, r);
    ctx.arcTo(x + w, y + h, x,     y + h, r);
    ctx.arcTo(x,     y + h, x,     y,     r);
    ctx.arcTo(x,     y,     x + w, y,     r);
    ctx.closePath();
    if (fill)   ctx.fill();
    if (stroke) ctx.stroke();
  }

  // ──────────────────────────────────────────────────────────────────────────
  // Upload Handlers
  // ──────────────────────────────────────────────────────────────────────────
  function handleTemplateUpload(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => {
      const img = new Image();
      img.onload = () => {
        state.bgImage        = img;
        state.templateWidth  = img.naturalWidth  || 1200;
        state.templateHeight = img.naturalHeight || 630;
        resizeCanvas(state.templateWidth, state.templateHeight);
        renderCanvas();
        fitZoomToScreen();
      };
      img.src = ev.target.result;
    };
    reader.readAsDataURL(file);
    e.target.value = '';  // allow re-upload of same file
  }

  // Raw Excel rows (array of arrays) stored for header-row picker
  let rawExcelRows = [];

  function handleExcelUpload(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const wb  = XLSX.read(new Uint8Array(ev.target.result), { type: 'array' });
        const ws  = wb.Sheets[wb.SheetNames[0]];

        // Read as array-of-arrays so we control the header row ourselves
        rawExcelRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

        if (!rawExcelRows.length) {
          alert('ឯកសារ Excel គ្មានទិន្នន័យ!');
          return;
        }

        // Auto-detect best header row: first row where ≥2 cells are non-empty strings
        const autoRow = detectHeaderRow(rawExcelRows);
        showHeaderPickerModal(rawExcelRows, autoRow);

      } catch (err) {
        alert('មានបញ្ហា: ' + err.message);
        console.error(err);
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  }

  /** Find the first row index where most cells look like text headers */
  function detectHeaderRow(rows) {
    let bestRow = 0;
    let bestScore = -1;
    for (let i = 0; i < Math.min(rows.length, 10); i++) {
      const row = rows[i];
      const score = row.filter(c => typeof c === 'string' && c.trim().length > 0).length;
      if (score > bestScore) { bestScore = score; bestRow = i; }
    }
    return bestRow;
  }

  /** Show a modal that previews the first N rows and lets user pick header row */
  function showHeaderPickerModal(rows, autoHeaderRow) {
    // Remove existing picker modal if any
    const existing = document.getElementById('headerPickerModal');
    if (existing) existing.remove();

    const preview = rows.slice(0, Math.min(rows.length, 8));
    const maxCols = Math.max(...preview.map(r => r.length));

    let tableHtml = `<table class="data-table" style="min-width:600px">`;
    tableHtml += `<thead><tr><th style="width:70px">Row</th><th colspan="${maxCols}">ខ្លឹមសារទិន្នន័យ (Preview)</th></tr></thead><tbody>`;

    preview.forEach((row, i) => {
      const isAuto = i === autoHeaderRow;
      tableHtml += `
        <tr class="${isAuto ? 'picker-auto-row' : ''}" id="picker-row-${i}" 
            style="cursor:pointer;${isAuto ? 'background:rgba(217,119,6,0.2);' : ''}">
          <td style="text-align:center;">
            <label style="cursor:pointer;display:flex;align-items:center;gap:6px;justify-content:center;">
              <input type="radio" name="headerRowPicker" value="${i}" ${isAuto ? 'checked' : ''}>
              <span style="font-size:0.75rem;color:#94a3b8">Row ${i + 1}</span>
            </label>
          </td>`;
      for (let c = 0; c < maxCols; c++) {
        const val = row[c] !== undefined ? String(row[c]) : '';
        tableHtml += `<td style="font-size:0.82rem;max-width:120px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${escapeHtml(val)}">${escapeHtml(val) || '<span style="color:#475569">—</span>'}</td>`;
      }
      tableHtml += '</tr>';
    });
    tableHtml += '</tbody></table>';

    const modal = document.createElement('div');
    modal.id = 'headerPickerModal';
    modal.className = 'modal-overlay';
    modal.innerHTML = `
      <div class="modal-card modal-large" style="max-height:90vh;">
        <div class="modal-header">
          <h2>📊 កំណត់ទម្រង់ ឯកសារ Excel (Header & Data)</h2>
          <button class="modal-close" id="hpClose">&times;</button>
        </div>
        <div class="modal-body">
          <div style="background:#0f172a;padding:14px;border-radius:10px;margin-bottom:14px;border:1px solid #334155;">
            <h4 style="color:#f59e0b;font-size:0.9rem;margin-bottom:8px;">សូមជ្រើសរើសប្រភេទ Header នៃឯកសារ Excel របស់អ្នក៖</h4>
            
            <label style="display:flex;align-items:center;gap:10px;margin-bottom:8px;cursor:pointer;font-size:0.88rem;">
              <input type="radio" name="headerModeOption" value="no_header">
              <span><strong>📌 Row 1 ជាទិន្នន័យស្រាប់ (គ្មាន Header Row)</strong> — បង្កើត Column 1, Column 2, Column 3... ស្វ័យប្រវត្តិ (រក្សាទុក Row 1 ជាទិន្នន័យ)</span>
            </label>
            
            <label style="display:flex;align-items:center;gap:10px;cursor:pointer;font-size:0.88rem;">
              <input type="radio" name="headerModeOption" value="has_header" checked>
              <span><strong>📋 ជ្រើសរើស Row ខាងក្រោមជា Header (ចំណងជើងជួរឈរ)</strong></span>
            </label>

            <label style="display:flex;align-items:center;gap:10px;margin-top:12px;cursor:pointer;font-size:0.88rem;color:#f59e0b;background:rgba(245,158,11,0.1);padding:8px 12px;border-radius:6px;border:1px dashed #f59e0b;">
              <input type="checkbox" id="hpAutoGroupOption" checked>
              <span><strong>🔀 បង្រួមឈ្មោះដូចគ្នាជា Poster តែមួយ</strong> (ឧទាហរណ៍៖ លេខ ៦៤ - ៨៨)</span>
            </label>
          </div>

          <p style="color:#94a3b8;font-size:0.82rem;margin-bottom:10px;">
            តារាងខាងក្រោមបង្ហាញទិន្នន័យ ៨ ជួរដំបូង៖
          </p>
          <div class="table-responsive">${tableHtml}</div>
        </div>
        <div class="modal-footer">
          <button class="btn btn-secondary" id="hpCancel">បោះបង់</button>
          <button class="btn btn-primary" id="hpApply">✔ បញ្ចូលទិន្នន័យ Excel (Apply)</button>
        </div>
      </div>`;

    document.body.appendChild(modal);
    lucide.createIcons();

    // Toggle row selection behavior based on headerModeOption
    const modeRadios = modal.querySelectorAll('input[name=headerModeOption]');
    modeRadios.forEach(r => {
      r.addEventListener('change', () => {
        const isHasHeader = modal.querySelector('input[name=headerModeOption]:checked').value === 'has_header';
        const rowsContainer = modal.querySelector('.table-responsive');
        if (isHasHeader) {
          rowsContainer.style.opacity = '1';
          rowsContainer.style.pointerEvents = 'auto';
        } else {
          rowsContainer.style.opacity = '0.4';
          rowsContainer.style.pointerEvents = 'none';
        }
      });
    });

    // Row click selects radio
    preview.forEach((_, i) => {
      const tr = document.getElementById(`picker-row-${i}`);
      if (tr) tr.addEventListener('click', () => {
        const hasHeaderRadio = modal.querySelector('input[name=headerModeOption][value=has_header]');
        hasHeaderRadio.checked = true;
        const radio = tr.querySelector('input[type=radio]');
        if (radio) radio.checked = true;
        document.querySelectorAll('#headerPickerModal tbody tr').forEach(r => r.style.background = '');
        tr.style.background = 'rgba(217,119,6,0.2)';
      });
    });

    document.getElementById('hpClose').addEventListener('click',  () => modal.remove());
    document.getElementById('hpCancel').addEventListener('click', () => modal.remove());
    document.getElementById('hpApply').addEventListener('click', () => {
      const headerMode = modal.querySelector('input[name=headerModeOption]:checked').value;
      const isAutoGroup = modal.querySelector('#hpAutoGroupOption') ? modal.querySelector('#hpAutoGroupOption').checked : true;
      if (headerMode === 'no_header') {
        applyExcelData(rawExcelRows, -1, isAutoGroup);
      } else {
        const selected = modal.querySelector('input[name=headerRowPicker]:checked');
        const hRow = selected ? parseInt(selected.value, 10) : autoHeaderRow;
        applyExcelData(rawExcelRows, hRow, isAutoGroup);
      }
      modal.remove();
    });
  }

  /** Convert Latin digits (0-9) to Khmer digits (០-៩) */
  function toKhmerDigits(val) {
    if (val === null || val === undefined) return '';
    const khmerDigits = ['០', '១', '២', '៣', '៤', '៥', '៦', '៧', '៨', '៩'];
    return String(val).replace(/[0-9]/g, ch => khmerDigits[parseInt(ch, 10)]);
  }

  /** Convert Khmer digits (០-៩) back to Latin digits (0-9) */
  function toLatinDigits(val) {
    if (val === null || val === undefined) return '';
    const map = { '០':'0', '១':'1', '២':'2', '៣':'3', '៤':'4', '៥':'5', '៦':'6', '៧':'7', '៨':'8', '៩':'9' };
    return String(val).replace(/[០-៩]/g, ch => map[ch] || ch);
  }

  /** Format ticket numbers e.g. ["1", "2", "3", "4"] -> "១-៤" or ["65", ..., "88"] -> "៦៥-៨៨" */
  function formatTicketNumbers(ticketArr) {
    if (!ticketArr || !ticketArr.length) return '';
    
    const cleanArr = [];
    ticketArr.forEach(tStr => {
      const latinStr = toLatinDigits(tStr);
      const parts = latinStr.split(/[,;/]+/).map(s => s.trim()).filter(Boolean);
      parts.forEach(p => {
        if (!cleanArr.includes(p)) cleanArr.push(p);
      });
    });

    if (!cleanArr.length) return '';
    if (cleanArr.length === 1) return toKhmerDigits(cleanArr[0]);

    const numObjects = cleanArr.map(t => {
      const match = t.match(/\d+/);
      return {
        original: t,
        num: match ? parseInt(match[0], 10) : null
      };
    });

    const allNumeric = numObjects.every(o => o.num !== null);

    if (allNumeric) {
      numObjects.sort((a, b) => a.num - b.num);

      const uniqueNums = [];
      numObjects.forEach(o => {
        if (!uniqueNums.includes(o.num)) uniqueNums.push(o.num);
      });

      const ranges = [];
      let start = uniqueNums[0];
      let prev = uniqueNums[0];

      for (let i = 1; i < uniqueNums.length; i++) {
        const curr = uniqueNums[i];
        if (curr === prev + 1) {
          prev = curr;
        } else {
          if (start === prev) {
            ranges.push(toKhmerDigits(start));
          } else {
            ranges.push(`${toKhmerDigits(start)}-${toKhmerDigits(prev)}`);
          }
          start = curr;
          prev = curr;
        }
      }

      if (start === prev) {
        ranges.push(toKhmerDigits(start));
      } else {
        ranges.push(`${toKhmerDigits(start)}-${toKhmerDigits(prev)}`);
      }

      if (ranges.length === 2 && uniqueNums.length === 2) {
        return ranges.join(' - ');
      }
      return ranges.join(', ');
    }

    return cleanArr.map(toKhmerDigits).join(' - ');
  }

  /** Normalize a name string for comparison: strip zero-width chars, collapse whitespace, NFC normalize */
  function normalizeName(name) {
    if (!name) return '';
    let s = String(name);
    // Unicode NFC normalization
    if (s.normalize) s = s.normalize('NFC');
    // Remove zero-width characters (ZWNJ, ZWJ, ZWSP, BOM, etc.)
    s = s.replace(/[\u200B\u200C\u200D\uFEFF\u00AD\u034F\u2028\u2029\u202A-\u202E\u2060-\u2064\u2066-\u206F]/g, '');
    // Collapse all whitespace (tabs, non-breaking spaces, etc.) into single space
    s = s.replace(/[\s\u00A0\u1680\u2000-\u200A\u3000]+/g, ' ');
    // Trim
    s = s.trim();
    return s;
  }

  /** Group duplicate names into single records (merging sequence numbers into "៦៥-៨៨") */
  function groupDuplicateRecords(rows, headers) {
    if (!rows || !rows.length) return rows;

    const nameHeader = headers.find(h => /ឈ្មោះ|name|owner|ម្ចាស់|គោត្តនាម|បេក្ខភាព/i.test(h)) || headers[1] || headers[0];
    const ticketHeader = headers.find(h => /លេខ|ឆ្នោត|ស្លាក|រៀង|no|number|code|id|ល\.រ/i.test(h)) || headers[0];

    // Make sure nameHeader and ticketHeader are different
    let finalTicketHeader = ticketHeader;
    if (nameHeader === ticketHeader && headers.length > 1) {
      finalTicketHeader = headers.find(h => h !== nameHeader && /លេខ|ឆ្នោត|ស្លាក|រៀង|no|number|code|id|ល\.រ/i.test(h)) || headers[0];
      if (finalTicketHeader === nameHeader) {
        finalTicketHeader = headers[0] !== nameHeader ? headers[0] : headers[1];
      }
    }

    console.log('[GroupDuplicates] nameHeader:', nameHeader, '| ticketHeader:', finalTicketHeader);
    console.log('[GroupDuplicates] Total rows before grouping:', rows.length);

    const groupedMap = new Map();

    rows.forEach((row, idx) => {
      const rawName = normalizeName(row[nameHeader]);
      if (!rawName) return;

      const key = rawName.toLowerCase();
      const rawTicket = String(row[finalTicketHeader] ?? '').trim();

      if (idx < 3) {
        console.log(`[GroupDuplicates] Row ${idx}: name="${rawName}" key="${key}" ticket="${rawTicket}"`);
      }

      if (!groupedMap.has(key)) {
        groupedMap.set(key, {
          name: rawName,
          tickets: rawTicket ? [rawTicket] : [],
          rowCopy: { ...row }
        });
      } else {
        const item = groupedMap.get(key);
        if (rawTicket && !item.tickets.includes(rawTicket)) {
          item.tickets.push(rawTicket);
        }
        headers.forEach(h => {
          if (h !== nameHeader && h !== finalTicketHeader && row[h]) {
            const existingVal = String(item.rowCopy[h] ?? '').trim();
            const newVal = String(row[h]).trim();
            if (!existingVal) {
              item.rowCopy[h] = newVal;
            } else if (newVal && !existingVal.includes(newVal)) {
              item.rowCopy[h] = existingVal + ', ' + newVal;
            }
          }
        });
      }
    });

    console.log('[GroupDuplicates] Unique names after grouping:', groupedMap.size);

    const result = [];
    groupedMap.forEach(item => {
      const row = { ...item.rowCopy };
      row[nameHeader] = item.name;
      row[finalTicketHeader] = formatTicketNumbers(item.tickets);

      headers.forEach(h => {
        row[h] = toKhmerDigits(row[h]);
      });

      result.push(row);
    });

    return result;
  }

  /** Build state.excelHeaders and state.excelRows from raw array-of-arrays */
  function applyExcelData(rows, headerRowIndex, autoGroup = true) {
    let headers = [];
    let startDataRow = 0;

    if (headerRowIndex === -1) {
      // No header row — generate Column 1, Column 2, Column 3...
      const maxCols = Math.max(...rows.map(r => r.length));
      for (let c = 0; c < maxCols; c++) {
        headers.push(`Column ${c + 1}`);
      }
      startDataRow = 0; // Row 1 is already data!
    } else {
      const headerArr = rows[headerRowIndex];
      headers = headerArr.map((h, i) => {
        const s = String(h).trim();
        if (!s || s.startsWith('__EMPTY')) {
          return 'Column ' + (i + 1);
        }
        return s;
      });
      startDataRow = headerRowIndex + 1;
    }

    // Build data rows (skip completely empty rows)
    let dataRows = [];
    for (let r = startDataRow; r < rows.length; r++) {
      const arr = rows[r];
      if (!arr || arr.every(c => c === null || c === undefined || String(c).trim() === '')) continue;
      const obj = {};
      headers.forEach((h, i) => {
        const val = arr[i] !== undefined ? arr[i] : '';
        obj[h] = val;
      });
      dataRows.push(obj);
    }

    if (autoGroup) {
      dataRows = groupDuplicateRecords(dataRows, headers);
    } else {
      dataRows = dataRows.map(obj => {
        const newObj = {};
        headers.forEach(h => { newObj[h] = toKhmerDigits(obj[h]); });
        return newObj;
      });
    }

    if (!dataRows.length) {
      alert('រកមិនឃើញទិន្នន័យ ក្នុងឯកសារ Excel ទេ!');
      return;
    }

    state.excelHeaders       = headers;
    state.excelRows          = dataRows;
    state.currentRecordIndex = 0;

    // Auto-map column headers to template text placeholders if needed
    autoMapPlaceholders(headers);

    dom.excelStatusBadge.textContent = `${dataRows.length} Records`;
    dom.excelStatusBadge.classList.add('active');

    updateMergeFieldsList();
    updateRecordDropdown();
    renderExcelTableModal();

    // Automatically switch to Live Preview Mode (Record 1) so user immediately sees filled poster!
    setMode('preview');

    let singleCount = 0;
    let multiCount = 0;
    const countsMap = buildPersonTicketCountsMap();
    const nameHeader = state.excelHeaders.find(h => /ឈ្មោះ|name|owner|ម្ចាស់|គោត្តនាម|បេក្ខភាព/i.test(h)) || state.excelHeaders[1] || state.excelHeaders[0];

    state.excelRows.forEach((r, i) => {
      const rawName = normalizeName(r[nameHeader]);
      const key = rawName ? rawName.toLowerCase() : `row_${i}`;
      const tc = countsMap.get(key) || getRecordTicketCount(r, state.excelHeaders);
      if (tc > 1) multiCount++;
      else singleCount++;
    });

    alert(`✔ Upload ឯកសារ Excel ជោគជ័យ!
បញ្ចូលបានសរុប ${state.excelRows.length} ឈ្មោះ។ ប្រព័ន្ធបានបែងចែកជា ២ ក្រុមស្វ័យប្រវត្តិ៖
• 🟢 ឈ្មោះ ១ មានលេខរៀង ១ ៖ ${toKhmerDigits(singleCount)} នាក់
• 🔵 ឈ្មោះ តែ ១ មានលេខរៀងច្រើន ៖ ${toKhmerDigits(multiCount)} នាក់

លោកអ្នកអាចចុចលើ Filter ឬ Tab ដើម្បីជ្រើសរើសមើល និង Export Posters តាមក្រុមបានភ្លាមៗ!`);
  }

  /** Smart placeholder auto-mapper */
  function autoMapPlaceholders(headers) {
    if (!headers || !headers.length) return;

    const nameHeader = headers.find(h => /ឈ្មោះ|name|owner|ម្ចាស់|គោត្តនាម|បេក្ខភាព/i.test(h)) || headers[1] || headers[0];
    let ticketHeader = headers.find(h => h !== nameHeader && /លេខ|ឆ្នោត|ស្លាក|រៀង|no|number|code|id|ល\.រ/i.test(h));
    if (!ticketHeader) {
      ticketHeader = headers.find(h => h !== nameHeader) || (headers[0] !== nameHeader ? headers[0] : headers[1]);
    }

    state.elements.forEach((el) => {
      if (el.type !== 'text') return;
      const currentText = el.text || '';

      if (/ឈ្មោះ|គោត្តនាម|owner|name|ម្ចាស់|បេក្ខភាព/i.test(currentText)) {
        el.text = `{{ ${nameHeader} }}`;
      } else if (/លេខ|ឆ្នោត|ស្លាក|រៀង|no|number|code|id|ល\.រ/i.test(currentText)) {
        el.text = `{{ ${ticketHeader} }}`;
      }
    });

    updateInspectorForm();
    renderLayersList();
  }

  // ──────────────────────────────────────────────────────────────────────────
  // Canvas & Zoom
  // ──────────────────────────────────────────────────────────────────────────
  function resizeCanvas(w, h) {
    state.canvas.width  = w;
    state.canvas.height = h;
    // The wrapper's CSS size follows the canvas naturally
    dom.canvasWrapper.style.width  = w + 'px';
    dom.canvasWrapper.style.height = h + 'px';
  }

  function changeZoom(delta) {
    state.zoomLevel = Math.max(0.15, Math.min(3.0, state.zoomLevel + delta));
    applyZoom();
  }

  function fitZoomToScreen() {
    const vw = dom.canvasViewport.clientWidth  - 80;
    const vh = dom.canvasViewport.clientHeight - 80;
    const sx = vw / state.templateWidth;
    const sy = vh / state.templateHeight;
    state.zoomLevel = parseFloat(Math.min(sx, sy, 1.0).toFixed(2));
    applyZoom();
  }

  function applyZoom() {
    const z  = state.zoomLevel;
    const sw = state.templateWidth  * z;
    const sh = state.templateHeight * z;
    const vw = dom.canvasViewport.clientWidth;
    const vh = dom.canvasViewport.clientHeight;
    // Center the scaled canvas using margins on the wrapper
    const ml = Math.max(40, (vw - sw) / 2);
    const mt = Math.max(40, (vh - sh) / 2);
    dom.canvasWrapper.style.transform  = `scale(${z})`;
    dom.canvasWrapper.style.marginLeft = ml + 'px';
    dom.canvasWrapper.style.marginTop  = mt + 'px';
    dom.zoomLevelDisplay.textContent   = Math.round(z * 100) + '%';
    renderOverlay();
  }

  // ──────────────────────────────────────────────────────────────────────────
  // Element Management
  // ──────────────────────────────────────────────────────────────────────────
  function addTextElement(defaultText) {
    const el = {
      id:          'el_' + Date.now(),
      type:        'text',
      text:        defaultText || 'អត្ថបទថ្មី',
      x:           Math.round(state.templateWidth  / 2),
      y:           Math.round(state.templateHeight / 2),
      width:       400,
      height:      60,
      fontFamily:  "'Battambang', sans-serif",
      fontSize:    36,
      color:       '#5b360f',
      strokeColor: '#ffffff',
      strokeWidth: 0,
      shadowBlur:  0,
      align:       'center'
    };
    state.elements.push(el);
    selectElement(el.id);
  }

  function addPhotoFrameElement() {
    const el = {
      id:     'el_' + Date.now(),
      type:   'photoFrame',
      x:      Math.round(state.templateWidth  / 2),
      y:      Math.round(state.templateHeight / 2),
      width:  160,
      height: 160,
      shape:  'circle'
    };
    state.elements.push(el);
    selectElement(el.id);
  }

  function selectElement(id) {
    state.selectedElementId = id;
    updateInspectorForm();
    renderLayersList();
    renderCanvas();   // renders overlay too
  }

  function deleteSelectedElement() {
    if (!state.selectedElementId) return;
    state.elements = state.elements.filter(el => el.id !== state.selectedElementId);
    state.selectedElementId = null;
    updateInspectorForm();
    renderLayersList();
    renderCanvas();
  }

  function moveLayer(dir) {
    if (!state.selectedElementId) return;
    const i = state.elements.findIndex(el => el.id === state.selectedElementId);
    if (i < 0) return;
    const j = i + dir;
    if (j >= 0 && j < state.elements.length) {
      [state.elements[i], state.elements[j]] = [state.elements[j], state.elements[i]];
      renderLayersList();
      renderCanvas();
    }
  }

  function getActiveElement() {
    return state.elements.find(el => el.id === state.selectedElementId) || null;
  }

  // ──────────────────────────────────────────────────────────────────────────
  // Inspector (Right Sidebar)
  // ──────────────────────────────────────────────────────────────────────────
  function updateInspectorForm() {
    const el = getActiveElement();
    if (!el) {
      dom.elementProperties.classList.add('hidden');
      dom.noSelectionText.classList.remove('hidden');
      return;
    }
    dom.noSelectionText.classList.add('hidden');
    dom.elementProperties.classList.remove('hidden');

    if (el.type === 'text') {
      dom.textGroup.style.display = 'block';
      dom.propText.value              = el.text          || '';
      dom.propFontFamily.value        = el.fontFamily    || "'Battambang', sans-serif";
      dom.propFontSize.value          = el.fontSize      || 32;
      dom.propAlign.value             = el.align         || 'center';
      dom.propColor.value             = el.color         || '#5b360f';
      dom.propColorText.value         = el.color         || '#5b360f';
      dom.propStrokeColor.value       = el.strokeColor   || '#ffffff';
      dom.propStrokeColorText.value   = el.strokeColor   || '#ffffff';
      dom.propStrokeWidth.value       = el.strokeWidth   || 0;
      dom.propShadowBlur.value        = el.shadowBlur    || 0;
    } else {
      dom.textGroup.style.display = 'none';
    }

    dom.propX.value = Math.round(el.x);
    dom.propY.value = Math.round(el.y);
  }

  function onFormChange() {
    const el = getActiveElement();
    if (!el) return;

    if (el.type === 'text') {
      el.text        = dom.propText.value;
      el.fontFamily  = dom.propFontFamily.value;
      el.fontSize    = parseInt(dom.propFontSize.value,   10) || 32;
      el.align       = dom.propAlign.value;
      el.color       = dom.propColor.value;
      el.strokeColor = dom.propStrokeColor.value;
      el.strokeWidth = parseInt(dom.propStrokeWidth.value, 10) || 0;
      el.shadowBlur  = parseInt(dom.propShadowBlur.value,  10) || 0;
    }

    const nx = parseInt(dom.propX.value, 10);
    const ny = parseInt(dom.propY.value, 10);
    if (!isNaN(nx)) el.x = nx;
    if (!isNaN(ny)) el.y = ny;

    renderCanvas();
  }

  // ──────────────────────────────────────────────────────────────────────────
  // Layer List
  // ──────────────────────────────────────────────────────────────────────────
  function renderLayersList() {
    dom.layersList.innerHTML = '';
    state.elements.forEach(el => {
      const item  = document.createElement('div');
      item.className = `layer-item ${el.id === state.selectedElementId ? 'active' : ''}`;
      const title = el.type === 'text' ? (el.text || 'Text Layer') : 'Photo Frame';
      const icon  = el.type === 'text' ? 'type' : 'image';

      item.innerHTML = `
        <div class="layer-info">
          <i data-lucide="${icon}" style="width:14px;height:14px;"></i>
          <span>${escapeHtml(title.substring(0, 30))}</span>
        </div>`;

      item.addEventListener('click', () => selectElement(el.id));
      dom.layersList.appendChild(item);
    });
    lucide.createIcons();
  }

  // ──────────────────────────────────────────────────────────────────────────
  // Merge Fields
  // ──────────────────────────────────────────────────────────────────────────
  function updateMergeFieldsList() {
    dom.mergeFieldsContainer.innerHTML = '';
    if (!state.excelHeaders.length) {
      dom.mergeFieldsContainer.innerHTML = '<p class="empty-hint">Upload Excel ដើម្បីរៀបចំ Merge Fields</p>';
      return;
    }
    state.excelHeaders.forEach(header => {
      const tag = document.createElement('div');
      tag.className = 'tag-field';
      tag.innerHTML = `<i data-lucide="plus" style="width:12px;height:12px;"></i> {{ ${escapeHtml(header)} }}`;
      tag.addEventListener('click', () => insertMergeField(header));
      dom.mergeFieldsContainer.appendChild(tag);
    });
    lucide.createIcons();
  }

  function insertMergeField(header) {
    const tag = `{{ ${header} }}`;
    const el  = getActiveElement();
    if (el && el.type === 'text') {
      el.text += ' ' + tag;
      updateInspectorForm();
      renderCanvas();
    } else {
      addTextElement(tag);
    }
  }

  function evaluateMerge(pattern, row) {
    if (!pattern || !row) return pattern || '';
    return pattern.replace(/\{\{\s*([^}]+?)\s*\}\}/g, (_, col) => {
      const v = row[col.trim()];
      return v !== undefined ? toKhmerDigits(v) : `{{ ${col.trim()} }}`;
    });
  }

  // ──────────────────────────────────────────────────────────────────────────
  // Record Navigation & Filtering (1 Ticket vs Multiple Tickets)
  // ──────────────────────────────────────────────────────────────────────────

  /** Map of person name -> total ticket count across the whole dataset */
  function buildPersonTicketCountsMap() {
    const countsMap = new Map();
    if (!state.excelRows || !state.excelRows.length) return countsMap;

    const currentHeaders = state.excelHeaders;
    const nameHeader = currentHeaders.find(h => /ឈ្មោះ|name|owner|ម្ចាស់|គោត្តនាម|បេក្ខភាព/i.test(h)) || currentHeaders[1] || currentHeaders[0];
    let ticketHeader = currentHeaders.find(h => h !== nameHeader && /លេខ|ឆ្នោត|ស្លាក|រៀង|no|number|code|id|ល\.រ/i.test(h));
    if (!ticketHeader) {
      ticketHeader = currentHeaders.find(h => h !== nameHeader) || (currentHeaders[0] !== nameHeader ? currentHeaders[0] : currentHeaders[1]);
    }

    state.excelRows.forEach(row => {
      const rawName = normalizeName(row[nameHeader]);
      if (!rawName) return;
      const key = rawName.toLowerCase();
      
      const val = ticketHeader ? String(row[ticketHeader] ?? '').trim() : '';
      const rowTickets = parseTicketCountFromValue(val);
      
      countsMap.set(key, (countsMap.get(key) || 0) + rowTickets);
    });

    return countsMap;
  }

  function parseTicketCountFromValue(val) {
    if (!val) return 1;

    const latin = toLatinDigits(val)
      .replace(/[\u200B-\u200D\uFEFF\u00AD]/g, '')
      .replace(/[\s\u00A0\u1680\u2000-\u200A\u3000]+/g, ' ')
      .trim();

    const numbers = latin.match(/\d+/g);
    if (!numbers || numbers.length === 0) return 1;

    const hasRangeSymbol = /[\-\u2010-\u2015\u2212]|ដល់|\bto\b/i.test(latin);

    if (hasRangeSymbol && numbers.length >= 2) {
      const n1 = parseInt(numbers[0], 10);
      const n2 = parseInt(numbers[numbers.length - 1], 10);
      if (!isNaN(n1) && !isNaN(n2) && n2 > n1) {
        return (n2 - n1) + 1;
      }
    }

    if (numbers.length > 1) {
      return numbers.length;
    }

    return 1;
  }

  /** Count total ticket numbers in a record row */
  function getRecordTicketCount(row, headers) {
    if (!row) return 1;
    if (row._ticketCount && typeof row._ticketCount === 'number' && row._ticketCount > 0) {
      return row._ticketCount;
    }

    const currentHeaders = (headers && headers.length) ? headers : state.excelHeaders;
    const nameHeader = currentHeaders.find(h => /ឈ្មោះ|name|owner|ម្ចាស់|គោត្តនាម|បេក្ខភាព/i.test(h)) || currentHeaders[1] || currentHeaders[0];
    let ticketHeader = currentHeaders.find(h => h !== nameHeader && /លេខ|ឆ្នោត|ស្លាក|រៀង|no|number|code|id|ល\.រ/i.test(h));
    if (!ticketHeader) {
      ticketHeader = currentHeaders.find(h => h !== nameHeader) || (currentHeaders[0] !== nameHeader ? currentHeaders[0] : currentHeaders[1]);
    }

    const val = ticketHeader ? String(row[ticketHeader] ?? '').trim() : '';
    return parseTicketCountFromValue(val);
  }

  /** Filter records according to active filter type ('all' | 'single' | 'multiple') */
  function getFilteredRecords(filterType) {
    const fType = filterType || state.ticketFilter || 'all';
    const countsMap = buildPersonTicketCountsMap();
    const nameHeader = state.excelHeaders.find(h => /ឈ្មោះ|name|owner|ម្ចាស់|គោត្តនាម|បេក្ខភាព/i.test(h)) || state.excelHeaders[1] || state.excelHeaders[0];

    const list = [];
    state.excelRows.forEach((row, originalIndex) => {
      const rawName = normalizeName(row[nameHeader]);
      const key = rawName ? rawName.toLowerCase() : `row_${originalIndex}`;
      
      const personTotalTickets = countsMap.get(key) || getRecordTicketCount(row, state.excelHeaders);
      const isSingle = personTotalTickets === 1;
      const isMultiple = personTotalTickets > 1;

      if (fType === 'single' && !isSingle) return;
      if (fType === 'multiple' && !isMultiple) return;

      list.push({
        row,
        originalIndex,
        personTotalTickets,
        ticketCount: getRecordTicketCount(row, state.excelHeaders),
        isSingle,
        isMultiple
      });
    });
    return list;
  }

  /** Sync active ticket filter across sidebar buttons, toolbar dropdown, and modal tabs */
  function setTicketFilter(filterType) {
    state.ticketFilter = filterType;
    state.modalTabFilter = filterType;

    if (dom.filterTicketType) dom.filterTicketType.value = filterType;

    // Update Side Filter Buttons
    document.querySelectorAll('.btn-side-filter').forEach(btn => {
      if (btn.getAttribute('data-filter') === filterType) {
        btn.classList.add('active');
      } else {
        btn.classList.remove('active');
      }
    });

    // Update Modal Filter Tabs
    document.querySelectorAll('.filter-tab').forEach(tab => {
      if (tab.getAttribute('data-filter') === filterType) {
        tab.classList.add('active');
      } else {
        tab.classList.remove('active');
      }
    });

    updateRecordDropdown();
    renderExcelTableModal();

    const filtered = getFilteredRecords(state.ticketFilter);
    if (filtered.length > 0) {
      const hasCurrent = filtered.some(f => f.originalIndex === state.currentRecordIndex);
      if (!hasCurrent) {
        state.currentRecordIndex = filtered[0].originalIndex;
      }
      setMode('preview');
    } else {
      renderCanvas();
    }
  }

  function updateRecordDropdown() {
    dom.recordDropdown.innerHTML = '<option value="template">-- Mode: រៀបចំ Template (ទម្រង់ដើម) --</option>';
    const filteredList = getFilteredRecords(state.ticketFilter);

    const nameHeader = state.excelHeaders.find(h => /ឈ្មោះ|name|owner|ម្ចាស់/i.test(h)) || state.excelHeaders[1] || state.excelHeaders[0];
    const ticketHeader = state.excelHeaders.find(h => /លេខ|ឆ្នោត|ស្លាក|រៀង|no|number|code|id/i.test(h)) || state.excelHeaders[0];

    filteredList.forEach((item, i) => {
      const opt = document.createElement('option');
      opt.value = item.originalIndex;
      
      const nameVal = String(item.row[nameHeader] || Object.values(item.row)[0] || `Record #${item.originalIndex + 1}`);
      const ticketVal = ticketHeader ? String(item.row[ticketHeader] || '') : '';
      
      const typeLabel = item.personTotalTickets > 1 ? `[👤🔢 ${toKhmerDigits(item.personTotalTickets)} លេខ]` : '[👤1️⃣ 1 លេខ]';
      opt.textContent = `${toKhmerDigits(i + 1)}. ${nameVal} (${ticketVal || 'គ្មានលេខ'}) ${typeLabel}`;
      
      if (item.originalIndex === state.currentRecordIndex) {
        opt.selected = true;
      }
      dom.recordDropdown.appendChild(opt);
    });

    // Update filter badge text
    if (dom.filteredCountBadge) {
      dom.filteredCountBadge.textContent = `${toKhmerDigits(filteredList.length)}/${toKhmerDigits(state.excelRows.length)}`;
    }
  }

  function setMode(mode) {
    state.mode = mode;
    if (mode === 'template') {
      dom.modeTemplateBtn.classList.add('active');
      dom.modePreviewBtn.classList.remove('active');
      dom.recordDropdown.value = 'template';
    } else {
      dom.modeTemplateBtn.classList.remove('active');
      dom.modePreviewBtn.classList.add('active');
      dom.recordDropdown.value = state.currentRecordIndex;
    }
    renderCanvas();
  }

  function navigateRecord(dir) {
    const filteredList = getFilteredRecords(state.ticketFilter);
    if (!filteredList.length) return;

    let currentPos = filteredList.findIndex(item => item.originalIndex === state.currentRecordIndex);
    if (currentPos === -1) currentPos = 0;

    const nextPos = (currentPos + dir + filteredList.length) % filteredList.length;
    state.currentRecordIndex = filteredList[nextPos].originalIndex;
    setMode('preview');
  }

  // ──────────────────────────────────────────────────────────────────────────
  // Canvas Rendering
  // ──────────────────────────────────────────────────────────────────────────
  function renderCanvas(targetCtx, targetRow) {
    const ctx      = targetCtx || state.ctx;
    const isExport = !!targetCtx;
    const row      = targetRow !== undefined
      ? targetRow
      : (state.mode === 'preview' ? state.excelRows[state.currentRecordIndex] : null);

    ctx.clearRect(0, 0, state.templateWidth, state.templateHeight);

    if (state.bgImage) {
      ctx.drawImage(state.bgImage, 0, 0, state.templateWidth, state.templateHeight);
    } else {
      ctx.fillStyle = '#1e293b';
      ctx.fillRect(0, 0, state.templateWidth, state.templateHeight);
    }

    state.elements.forEach(el => {
      if (el.type === 'text')       renderTextElement(ctx, el, row);
      else if (el.type === 'photoFrame') renderPhotoFrame(ctx, el);
    });

    if (!isExport) renderOverlay();
  }

  function renderTextElement(ctx, el, row) {
    ctx.save();
    const text  = row ? evaluateMerge(el.text, row) : (el.text || '');
    ctx.font         = `${el.fontSize || 32}px ${el.fontFamily || "'Battambang', sans-serif"}`;
    ctx.textAlign    = el.align        || 'center';
    ctx.textBaseline = 'middle';

    if ((el.shadowBlur || 0) > 0) {
      ctx.shadowColor   = 'rgba(0,0,0,0.7)';
      ctx.shadowBlur    = el.shadowBlur;
      ctx.shadowOffsetX = 2;
      ctx.shadowOffsetY = 2;
    }

    if ((el.strokeWidth || 0) > 0) {
      ctx.lineWidth    = el.strokeWidth;
      ctx.strokeStyle  = el.strokeColor || '#000000';
      ctx.lineJoin     = 'round';
      ctx.strokeText(text, el.x, el.y);
    }

    ctx.fillStyle = el.color || '#ffffff';
    ctx.fillText(text, el.x, el.y);
    ctx.restore();
  }

  function renderPhotoFrame(ctx, el) {
    ctx.save();
    ctx.beginPath();
    ctx.arc(el.x, el.y, (el.width || 160) / 2, 0, Math.PI * 2);
    ctx.closePath();
    ctx.fillStyle   = 'rgba(255,255,255,0.2)';
    ctx.fill();
    ctx.lineWidth   = 4;
    ctx.strokeStyle = '#d97706';
    ctx.stroke();
    ctx.fillStyle    = '#d97706';
    ctx.font         = '18px sans-serif';
    ctx.textAlign    = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText('Photo', el.x, el.y);
    ctx.restore();
  }

  // ──────────────────────────────────────────────────────────────────────────
  // Overlay (Drag/Resize Handles) — scaled with CSS transform
  // ──────────────────────────────────────────────────────────────────────────
  function renderOverlay() {
    dom.canvasOverlay.innerHTML = '';

    // In template mode show boxes; in preview mode hide them (read-only)
    if (state.mode === 'preview') return;

    state.elements.forEach(el => {
      const box  = document.createElement('div');
      const isSelected = el.id === state.selectedElementId;
      box.className = 'element-box' + (isSelected ? ' active' : '');

      const hw = (el.width  || 200) / 2;
      const hh = (el.height || 60)  / 2;

      box.style.cssText = `
        left:${el.x - hw}px; top:${el.y - hh}px;
        width:${el.width || 200}px; height:${el.height || 60}px;
      `;

      // Click to select
      box.addEventListener('mousedown', e => {
        if (e.target.classList.contains('resize-handle')) return;
        e.stopPropagation();
        e.preventDefault();
        beginDrag(e, el.id);
      });

      if (isSelected) {
        ['nw','ne','sw','se'].forEach(pos => {
          const h = document.createElement('div');
          h.className = `resize-handle handle-${pos}`;
          h.addEventListener('mousedown', e => {
            e.stopPropagation();
            e.preventDefault();
            beginResize(e, el.id);
          });
          box.appendChild(h);
        });
      }

      dom.canvasOverlay.appendChild(box);
    });
  }

  // ── Drag ──────────────────────────────────────────────────────────────────
  function beginDrag(e, id) {
    if (state.selectedElementId !== id) {
      state.selectedElementId = id;
      updateInspectorForm();
      renderLayersList();
      renderOverlay();
    }

    const el = getActiveElement();
    if (!el) return;

    state.dragState.isDragging      = true;
    state.dragState.startX          = e.clientX;
    state.dragState.startY          = e.clientY;
    state.dragState.elementStartX   = el.x;
    state.dragState.elementStartY   = el.y;

    const move = ev => {
      if (!state.dragState.isDragging) return;
      const dx = (ev.clientX - state.dragState.startX) / state.zoomLevel;
      const dy = (ev.clientY - state.dragState.startY) / state.zoomLevel;
      el.x = Math.round(state.dragState.elementStartX + dx);
      el.y = Math.round(state.dragState.elementStartY + dy);
      dom.propX.value = el.x;
      dom.propY.value = el.y;
      renderCanvas();
    };

    const up = () => {
      state.dragState.isDragging = false;
      document.removeEventListener('mousemove', move);
      document.removeEventListener('mouseup',   up);
    };

    document.addEventListener('mousemove', move);
    document.addEventListener('mouseup',   up);
  }

  // ── Resize ────────────────────────────────────────────────────────────────
  function beginResize(e, id) {
    state.selectedElementId = id;
    const el = getActiveElement();
    if (!el) return;

    state.dragState.isResizing      = true;
    state.dragState.startX          = e.clientX;
    state.dragState.startY          = e.clientY;
    state.dragState.elementStartW   = el.width  || 200;
    state.dragState.elementStartH   = el.height || 60;

    const move = ev => {
      if (!state.dragState.isResizing) return;
      const dx = (ev.clientX - state.dragState.startX) / state.zoomLevel;
      const dy = (ev.clientY - state.dragState.startY) / state.zoomLevel;
      el.width  = Math.max(40,  Math.round(state.dragState.elementStartW + dx));
      el.height = Math.max(20,  Math.round(state.dragState.elementStartH + dy));
      renderCanvas();
    };

    const up = () => {
      state.dragState.isResizing = false;
      document.removeEventListener('mousemove', move);
      document.removeEventListener('mouseup',   up);
    };

    document.addEventListener('mousemove', move);
    document.addEventListener('mouseup',   up);
  }

  // ──────────────────────────────────────────────────────────────────────────
  // Excel Data Table Modal
  // ──────────────────────────────────────────────────────────────────────────
  // ──────────────────────────────────────────────────────────────────────────
  // Excel Data Table Modal
  // ──────────────────────────────────────────────────────────────────────────
  function renderExcelTableModal() {
    if (!state.excelRows.length) {
      dom.excelTableContainer.innerHTML = '<p class="empty-hint">មិនទាន់មានទិន្នន័យ Excel</p>';
      if (dom.modalCountAll) dom.modalCountAll.textContent = '0';
      if (dom.modalCountSingle) dom.modalCountSingle.textContent = '0';
      if (dom.modalCountMultiple) dom.modalCountMultiple.textContent = '0';
      if (dom.sideCountAll) dom.sideCountAll.textContent = '0';
      if (dom.sideCountSingle) dom.sideCountSingle.textContent = '0';
      if (dom.sideCountMultiple) dom.sideCountMultiple.textContent = '0';
      return;
    }

    // Calculate total counts for filter tabs & sidebar badges using person total tickets
    let singleCount = 0;
    let multiCount = 0;
    const countsMap = buildPersonTicketCountsMap();
    const nameHeader = state.excelHeaders.find(h => /ឈ្មោះ|name|owner|ម្ចាស់/i.test(h)) || state.excelHeaders[1] || state.excelHeaders[0];

    state.excelRows.forEach((r, i) => {
      const rawName = normalizeName(r[nameHeader]);
      const key = rawName ? rawName.toLowerCase() : `row_${i}`;
      const tc = countsMap.get(key) || getRecordTicketCount(r, state.excelHeaders);
      if (tc > 1) multiCount++;
      else singleCount++;
    });

    if (dom.modalCountAll) dom.modalCountAll.textContent = toKhmerDigits(state.excelRows.length);
    if (dom.modalCountSingle) dom.modalCountSingle.textContent = toKhmerDigits(singleCount);
    if (dom.modalCountMultiple) dom.modalCountMultiple.textContent = toKhmerDigits(multiCount);

    if (dom.sideCountAll) dom.sideCountAll.textContent = toKhmerDigits(state.excelRows.length);
    if (dom.sideCountSingle) dom.sideCountSingle.textContent = toKhmerDigits(singleCount);
    if (dom.sideCountMultiple) dom.sideCountMultiple.textContent = toKhmerDigits(multiCount);

    const query = (state.modalSearchQuery || '').toLowerCase().trim();

    // Filter rows for modal view based on tab and search
    const visibleRows = [];
    state.excelRows.forEach((row, i) => {
      const rawName = normalizeName(row[nameHeader]);
      const key = rawName ? rawName.toLowerCase() : `row_${i}`;
      const personTotalTickets = countsMap.get(key) || getRecordTicketCount(row, state.excelHeaders);
      const isMulti = personTotalTickets > 1;

      // Tab filter
      if (state.modalTabFilter === 'single' && isMulti) return;
      if (state.modalTabFilter === 'multiple' && !isMulti) return;

      // Search query filter
      if (query) {
        const rowText = Object.values(row).join(' ').toLowerCase();
        if (!rowText.includes(query)) return;
      }

      visibleRows.push({ row, originalIndex: i, personTotalTickets, isMulti });
    });

    if (!visibleRows.length) {
      dom.excelTableContainer.innerHTML = '<p class="empty-hint" style="text-align:center;padding:24px;">មិនមានទិន្នន័យត្រូវនឹងការ Filter/ស្វែងរកនេះទេ</p>';
      return;
    }

    let html = '<table class="data-table"><thead><tr><th>#</th><th>ប្រភេទ</th>';
    state.excelHeaders.forEach(h => { html += `<th>${escapeHtml(h)}</th>`; });
    html += '<th style="text-align:center;">សកម្មភាព</th></tr></thead><tbody>';

    visibleRows.forEach((item) => {
      const typeBadge = item.isMulti
        ? `<span class="badge-type badge-multiple"><i data-lucide="layers" style="width:12px;height:12px"></i> ${toKhmerDigits(item.personTotalTickets)} លេខ</span>`
        : `<span class="badge-type badge-single"><i data-lucide="check-circle-2" style="width:12px;height:12px"></i> 1 លេខ</span>`;

      const isCurrent = item.originalIndex === state.currentRecordIndex;
      html += `<tr style="${isCurrent ? 'background:rgba(217,119,6,0.18);' : ''}">
        <td>${toKhmerDigits(item.originalIndex + 1)}</td>
        <td>${typeBadge}</td>`;

      state.excelHeaders.forEach(h => {
        html += `<td>${escapeHtml(String(item.row[h] ?? ''))}</td>`;
      });

      html += `<td style="text-align:center;">
        <button class="btn btn-secondary btn-sm-preview" data-index="${item.originalIndex}" style="padding:3px 8px;font-size:0.75rem;">
          <i data-lucide="eye" style="width:12px;height:12px"></i> Preview
        </button>
      </td></tr>`;
    });

    html += '</tbody></table>';
    dom.excelTableContainer.innerHTML = html;
    lucide.createIcons();

    // Click handler for Preview buttons in modal table
    dom.excelTableContainer.querySelectorAll('.btn-sm-preview').forEach(btn => {
      btn.addEventListener('click', e => {
        const idx = parseInt(e.currentTarget.getAttribute('data-index'), 10);
        if (!isNaN(idx)) {
          state.currentRecordIndex = idx;
          setMode('preview');
          dom.dataModal.classList.add('hidden');
        }
      });
    });
  }

  // ──────────────────────────────────────────────────────────────────────────
  // Batch Export → ZIP (Option to export all or filtered rows)
  // ──────────────────────────────────────────────────────────────────────────
  async function handleBatchExport(filteredOnly = false) {
    if (!state.excelRows.length) {
      alert('Upload ឯកសារ Excel ជាមុនសិន!');
      return;
    }

    let exportTargetRows = state.excelRows;
    let zipFilename = 'Posters_Bulk_Export.zip';

    if (filteredOnly) {
      const activeFilter = state.modalTabFilter !== 'all' ? state.modalTabFilter : state.ticketFilter;
      const filteredItems = getFilteredRecords(activeFilter);
      
      // Apply search query if present
      const query = (state.modalSearchQuery || '').toLowerCase().trim();
      const finalItems = query
        ? filteredItems.filter(item => Object.values(item.row).join(' ').toLowerCase().includes(query))
        : filteredItems;

      exportTargetRows = finalItems.map(item => item.row);
      const tag = activeFilter === 'single' ? 'Single_Tickets' : (activeFilter === 'multiple' ? 'Multiple_Tickets' : 'Filtered');
      zipFilename = `Posters_${tag}_Export.zip`;
    }

    if (!exportTargetRows.length) {
      alert('គ្មានទិន្នន័យសម្រាប់ Export តាម Filter នេះទេ!');
      return;
    }

    state.cancelExport = false;
    dom.exportModal.classList.remove('hidden');
    dom.exportProgressBar.style.width = '0%';

    const zip     = new JSZip();
    const off     = document.createElement('canvas');
    off.width     = state.templateWidth;
    off.height    = state.templateHeight;
    const offCtx  = off.getContext('2d');
    const total   = exportTargetRows.length;

    for (let i = 0; i < total; i++) {
      if (state.cancelExport) break;

      const row = exportTargetRows[i];
      renderCanvas(offCtx, row);

      const blob = await new Promise(res => off.toBlob(res, 'image/png', 0.95));
      const nameHeader = state.excelHeaders.find(h => /ឈ្មោះ|name|owner|ម្ចាស់/i.test(h)) || state.excelHeaders[0];
      const name = String(row[nameHeader] || Object.values(row)[0] || `poster_${i + 1}`).replace(/[\\/:*?"<>|]/g, '_');
      zip.file(`poster_${String(i + 1).padStart(3, '0')}_${name}.png`, blob);

      const pct = Math.round(((i + 1) / total) * 100);
      dom.exportProgressBar.style.width  = pct + '%';
      dom.exportStatusText.textContent   = `Export ${i + 1} / ${total} (${name}) ...`;

      await new Promise(r => setTimeout(r, 10));
    }

    if (!state.cancelExport) {
      dom.exportStatusText.textContent = 'Compressing ZIP...';
      const content = await zip.generateAsync({ type: 'blob' });
      saveAs(content, zipFilename);
    }

    dom.exportModal.classList.add('hidden');
  }

  // ──────────────────────────────────────────────────────────────────────────
  // Utilities
  // ──────────────────────────────────────────────────────────────────────────
  function escapeHtml(str) {
    return String(str)
      .replace(/&/g,  '&amp;')
      .replace(/</g,  '&lt;')
      .replace(/>/g,  '&gt;')
      .replace(/"/g,  '&quot;');
  }

  // ──────────────────────────────────────────────────────────────────────────
  // Boot
  // ──────────────────────────────────────────────────────────────────────────
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
