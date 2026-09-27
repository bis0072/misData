document.addEventListener('DOMContentLoaded', () => {
  const form = document.getElementById('uploadForm');
  const fileBase = document.getElementById('fileBase');
  const dzBase = document.getElementById('dzBase');
  const nameBase = document.getElementById('nameBase');
  const btnAddFile = document.getElementById('btnAddFile');
  const comparisonList = document.getElementById('comparisonList');
  const submitBtn = document.getElementById('submitBtn');
  const btnSpinner = document.getElementById('btnSpinner');
  const btnText = document.getElementById('btnText');
  const btnIcon = document.getElementById('btnIcon');
  const errorBox = document.getElementById('errorBox');
  const resultsCard = document.getElementById('resultsCard');
  const breakdownList = document.getElementById('breakdownList');

  // Option Mode Elements
  const modeAuto = document.getElementById('modeAuto');
  const modeCustom = document.getElementById('modeCustom');
  const nopColIndex = document.getElementById('nopColIndex');
  const premColIndex = document.getElementById('premColIndex');
  const nopSubText = document.getElementById('nopSubText');
  const premSubText = document.getElementById('premSubText');
  const indexBaseRow = document.getElementById('indexBaseRow');
  const base1Radio = document.getElementById('base1Radio');
  const base0Radio = document.getElementById('base0Radio');
  const optionsTip = document.getElementById('optionsTip');
  const outputFilenameInput = document.getElementById('outputFilename');
  const dateTagPreview = document.getElementById('dateTagPreview');
  const downloadedFileName = document.getElementById('downloadedFileName');

  // Dynamic Footer Year
  const currentYearSpan = document.getElementById('currentYear');
  if (currentYearSpan) {
    currentYearSpan.textContent = new Date().getFullYear();
  }

  // Progress Bar & Stage Tracker Elements
  const progressContainer = document.getElementById('progressContainer');
  const progressBar = document.getElementById('progressBar');
  const progressStatusText = document.getElementById('progressStatusText');
  const progressPercent = document.getElementById('progressPercent');
  const stageSteps = [
    document.getElementById('stageStep1'),
    document.getElementById('stageStep2'),
    document.getElementById('stageStep3'),
    document.getElementById('stageStep4'),
    document.getElementById('stageStep5')
  ];

  class ProgressTracker {
    constructor() {
      this.timer = null;
      this.currentPercent = 0;
    }

    start(totalFiles) {
      if (!progressContainer) return;
      this.reset();
      progressContainer.style.display = 'block';
      this.setStage(1, `Uploading Base + ${totalFiles} Comparison report(s) to cloud...`, 15);

      let elapsed = 0;
      this.timer = setInterval(() => {
        elapsed += 200;

        if (elapsed >= 300 && elapsed < 900) {
          this.setStage(2, `Parsing HTML tables & finding column headers...`, Math.min(25 + (elapsed / 25), 45));
        } else if (elapsed >= 900 && elapsed < 2000) {
          this.setStage(3, `Cross-matching Agent codes across all files (VLOOKUP in RAM)...`, Math.min(48 + (elapsed / 45), 72));
        } else if (elapsed >= 2000 && elapsed < 3500) {
          this.setStage(4, `Generating Excel workbook with styled headers & merged banner...`, Math.min(74 + (elapsed / 90), 92));
        } else if (elapsed >= 3500) {
          this.setStage(4, `Finalizing in-memory Excel file streaming from cloud server...`, Math.min(this.currentPercent + 0.3, 95));
        }
      }, 200);
    }

    setStage(stageNum, text, targetPercent) {
      if (text && progressStatusText) progressStatusText.textContent = text;
      this.currentPercent = Math.max(this.currentPercent, Math.round(targetPercent));
      if (progressBar) progressBar.style.width = `${this.currentPercent}%`;
      if (progressPercent) progressPercent.textContent = `${this.currentPercent}%`;

      stageSteps.forEach((step, idx) => {
        if (!step) return;
        const currentStageIndex = stageNum - 1;
        if (idx < currentStageIndex) {
          step.className = 'stage-step completed';
        } else if (idx === currentStageIndex) {
          step.className = 'stage-step active';
        } else {
          step.className = 'stage-step';
        }
      });
    }

    complete(callback) {
      clearInterval(this.timer);
      this.setStage(5, `✨ Excel workbook generated! Triggering browser download...`, 100);
      stageSteps.forEach(step => { if (step) step.className = 'stage-step completed'; });
      if (progressBar) {
        progressBar.style.background = 'linear-gradient(90deg, #10b981, #34d399)';
      }

      setTimeout(() => {
        if (progressContainer) {
          progressContainer.style.display = 'none';
        }
        if (callback) callback();
      }, 600);
    }

    error(errMessage) {
      clearInterval(this.timer);
      if (progressStatusText) progressStatusText.textContent = `Processing failed: ${errMessage}`;
      if (progressBar) {
        progressBar.style.background = '#ef4444';
        progressBar.style.width = '100%';
      }
      if (progressPercent) {
        progressPercent.textContent = 'Failed';
        progressPercent.style.color = '#ef4444';
      }
    }

    reset() {
      clearInterval(this.timer);
      this.currentPercent = 0;
      if (progressBar) {
        progressBar.style.width = '0%';
        progressBar.style.background = 'linear-gradient(90deg, #3b82f6, #06b6d4, #10b981)';
      }
      if (progressPercent) {
        progressPercent.textContent = '0%';
        progressPercent.style.color = '#60a5fa';
      }
      stageSteps.forEach(step => { if (step) step.className = 'stage-step'; });
    }
  }

  const tracker = new ProgressTracker();

  // Format today's date (YYYY-MM-DD)
  const now = new Date();
  const yyyy = now.getFullYear();
  const mm = String(now.getMonth() + 1).padStart(2, '0');
  const dd = String(now.getDate()).padStart(2, '0');
  const todayDateStr = `${yyyy}-${mm}-${dd}`;

  if (dateTagPreview) {
    dateTagPreview.textContent = `_${todayDateStr}.xlsx`;
  }

  function formatBytes(bytes) {
    if (!bytes) return '0 B';
    const k = 1024;
    const sizes = ['B', 'KB', 'MB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
  }

  // --- Column Index Mode Handling ---
  function updateColumnSettingsMode() {
    const isCustom = modeCustom.checked;
    if (isCustom) {
      indexBaseRow.style.display = 'flex';
      nopColIndex.disabled = false;
      premColIndex.disabled = false;
      nopColIndex.classList.remove('auto-mode');
      premColIndex.classList.remove('auto-mode');
      applyIndexDefaults();
    } else {
      indexBaseRow.style.display = 'none';
      nopColIndex.disabled = true;
      premColIndex.disabled = true;
      nopColIndex.value = '';
      premColIndex.value = '';
      nopColIndex.placeholder = 'Auto';
      premColIndex.placeholder = 'Auto';
      nopColIndex.classList.add('auto-mode');
      premColIndex.classList.add('auto-mode');
      nopSubText.textContent = "Auto-detects 'Total POL' / 'NOP'";
      premSubText.textContent = "Auto-detects 'TOTAL PREM'";
      optionsTip.innerHTML = `✨ <strong>No index guessing required!</strong> In <strong>Auto-Detect</strong> mode, the system automatically finds <code>Total POL</code> (NOP) and <code>TOTAL PREM(SP+NSP)</code> by header name across any 2nd, 3rd, or Nth HTML file.<br>If you ever need to manually set columns, switch to <strong>Custom Column Position</strong>.`;
    }
  }

  function applyIndexDefaults() {
    const is1Based = base1Radio.checked;
    if (is1Based) {
      // Excel 1-based indexing (Col 1 = A, Col 8 = H, Col 11 = K)
      nopColIndex.value = 8;
      premColIndex.value = 11;
      nopSubText.textContent = 'Excel Column 8 (Column H / Total POL)';
      premSubText.textContent = 'Excel Column 11 (Column K / Total Prem)';
      optionsTip.innerHTML = `📊 <strong>Excel Style (1-based):</strong> 1st column is <code>1</code>. Standard MIS column for Total POL is <code>8</code> (Col H), and Total Prem is <code>11</code> (Col K, or <code>10</code> for non-single).`;
    } else {
      // Python 0-based indexing (Col 0 = A, Col 7 = H, Col 10 = K)
      nopColIndex.value = 7;
      premColIndex.value = 10;
      nopSubText.textContent = 'Python Index 7 (Total POL)';
      premSubText.textContent = 'Python Index 10 (Total Prem)';
      optionsTip.innerHTML = `🐍 <strong>Python Style (0-based):</strong> 1st column is <code>0</code>. Standard MIS column for Total POL is <code>7</code>, and Total Prem is <code>10</code> (or <code>9</code> for non-single).`;
    }
  }

  modeAuto.addEventListener('change', updateColumnSettingsMode);
  modeCustom.addEventListener('change', updateColumnSettingsMode);

  base1Radio.addEventListener('change', () => {
    if (modeCustom.checked) {
      // Switch from 0-based to 1-based: add 1
      if (nopColIndex.value !== '') nopColIndex.value = parseInt(nopColIndex.value, 10) + 1;
      if (premColIndex.value !== '') premColIndex.value = parseInt(premColIndex.value, 10) + 1;
      nopSubText.textContent = `Excel Column ${nopColIndex.value}`;
      premSubText.textContent = `Excel Column ${premColIndex.value}`;
    }
  });

  base0Radio.addEventListener('change', () => {
    if (modeCustom.checked) {
      // Switch from 1-based to 0-based: subtract 1
      if (nopColIndex.value !== '') nopColIndex.value = Math.max(0, parseInt(nopColIndex.value, 10) - 1);
      if (premColIndex.value !== '') premColIndex.value = Math.max(0, parseInt(premColIndex.value, 10) - 1);
      nopSubText.textContent = `Python Index ${nopColIndex.value}`;
      premSubText.textContent = `Python Index ${premColIndex.value}`;
    }
  });

  btnApplyDefaults.addEventListener('click', applyIndexDefaults);

  // --- Live Table Header Inspector on Selected Files ---
  async function inspectFileHeaders(file, targetContainerId) {
    if (!file) return;
    try {
      const slice = file.slice(0, 80000);
      const text = await slice.text();
      const parser = new DOMParser();
      const doc = parser.parseFromString(text, 'text/html');
      const tableIndex = parseInt(document.getElementById('tableIndex').value || '4', 10);
      const tables = doc.querySelectorAll('table');
      if (tables.length <= tableIndex) return;

      const table = tables[tableIndex];
      const firstRow = table.querySelector('tr');
      if (!firstRow) return;

      const headers = Array.from(firstRow.querySelectorAll('th, td')).map(c => c.textContent.trim());
      if (headers.length === 0) return;

      let box = document.getElementById(targetContainerId);
      if (!box) {
        box = document.createElement('div');
        box.id = targetContainerId;
        box.className = 'live-columns-box';
      }

      box.innerHTML = `
        <div class="live-columns-header">
          <span>Detected Columns in Table #${tableIndex} (${headers.length} cols)</span>
          <span style="font-size:0.7rem; color:#60a5fa;">[Col # = Excel 1-based | Idx # = Python 0-based]</span>
        </div>
        <div class="live-columns-chips"></div>
      `;

      const chipsContainer = box.querySelector('.live-columns-chips');
      headers.forEach((h, i) => {
        const chip = document.createElement('span');
        chip.className = 'live-col-chip';
        const hLow = h.toLowerCase();
        let badgeType = '';

        if (hLow === 'agent') {
          chip.classList.add('active-key');
          badgeType = ' (Key)';
        } else if (hLow.includes('total pol') || hLow === 'tot pol' || hLow === 'nop') {
          chip.classList.add('active-pol');
          badgeType = ' (NOP)';
        } else if (hLow.includes('total prem') || hLow.includes('prem(sp+nsp)')) {
          chip.classList.add('active-prem');
          badgeType = ' (Prem)';
        }

        chip.textContent = `Col ${i + 1} [Idx ${i}]: ${h}${badgeType}`;
        chip.title = `Click to set as custom column in settings`;
        chip.style.cursor = 'pointer';

        chip.addEventListener('click', () => {
          modeCustom.checked = true;
          updateColumnSettingsMode();
          const is1Based = base1Radio.checked;
          const chosenVal = is1Based ? (i + 1) : i;
          if (hLow.includes('prem')) {
            premColIndex.value = chosenVal;
          } else {
            nopColIndex.value = chosenVal;
          }
        });

        chipsContainer.appendChild(chip);
      });

      return box;
    } catch (_) {
      return null;
    }
  }

  // --- Base File Dropzone ---
  function setupBaseDropzone() {
    fileBase.addEventListener('change', async () => {
      const existingPreview = document.getElementById('preview_base');
      if (existingPreview) existingPreview.remove();

      if (fileBase.files.length > 0) {
        const file = fileBase.files[0];
        dzBase.classList.add('has-file');
        nameBase.textContent = `${file.name} (${formatBytes(file.size)})`;

        const previewEl = await inspectFileHeaders(file, 'preview_base');
        if (previewEl) {
          dzBase.parentNode.appendChild(previewEl);
        }
      } else {
        dzBase.classList.remove('has-file');
        nameBase.textContent = '';
      }
      updateSubmitButtonText();
    });

    ['dragover', 'dragenter'].forEach(e => {
      dzBase.addEventListener(e, (evt) => {
        evt.preventDefault();
        dzBase.classList.add('dragover');
      });
    });

    ['dragleave', 'drop'].forEach(e => {
      dzBase.addEventListener(e, () => {
        dzBase.classList.remove('dragover');
      });
    });

    dzBase.addEventListener('drop', (e) => {
      e.preventDefault();
      if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
        fileBase.files = e.dataTransfer.files;
        const event = new Event('change');
        fileBase.dispatchEvent(event);
      }
    });
  }

  // --- Dynamic Comparison Cards ---
  let comparisonCards = [];
  let nextCardId = 1;

  function updateSubmitButtonText() {
    const validCompFiles = comparisonCards.filter(c => c.file !== null);
    const count = validCompFiles.length;
    if (count <= 1) {
      btnText.textContent = 'Convert & Download Excel';
    } else {
      btnText.textContent = `Convert & Merge (1 Base + ${count} Comparison Files)`;
    }
  }

  function renumberCards() {
    const cards = comparisonList.querySelectorAll('.comp-card');
    cards.forEach((cardEl, index) => {
      const numBadge = cardEl.querySelector('.comp-card-num');
      if (numBadge) {
        numBadge.textContent = `#${index + 1}`;
      }
    });
  }

  function addComparisonCard(initialFile = null) {
    const cardId = nextCardId++;
    const cardData = {
      id: cardId,
      file: initialFile
    };
    comparisonCards.push(cardData);

    const card = document.createElement('div');
    card.className = 'comp-card';
    card.id = `compCard_${cardId}`;

    card.innerHTML = `
      <div class="comp-card-header">
        <div class="comp-card-title">
          <span>Comparison File</span>
          <span class="comp-card-num">#${comparisonCards.length}</span>
        </div>
        <button type="button" class="btn-remove-file" title="Remove this comparison file">
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"></line><line x1="6" y1="6" x2="18" y2="18"></line></svg>
        </button>
      </div>

      <div class="dropzone ${initialFile ? 'has-file' : ''}" id="dzComp_${cardId}">
        <input type="file" id="inputComp_${cardId}" accept=".htm,.html">
        <div class="drop-icon">
          <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="12" y1="18" x2="12" y2="12"></line><line x1="9" y1="15" x2="15" y2="15"></line></svg>
        </div>
        <div class="drop-title">Drop Comparison HTML Report</div>
        <div class="drop-desc">e.g. <code>temadata.htm</code>, <code>dvnDor.html</code></div>
        <div class="file-badge" id="badgeComp_${cardId}">${initialFile ? `${initialFile.name} (${formatBytes(initialFile.size)})` : ''}</div>
      </div>
      <div id="previewContainer_${cardId}"></div>
    `;

    comparisonList.appendChild(card);
    renumberCards();

    const fileInput = card.querySelector(`#inputComp_${cardId}`);
    const dz = card.querySelector(`#dzComp_${cardId}`);
    const badge = card.querySelector(`#badgeComp_${cardId}`);
    const removeBtn = card.querySelector('.btn-remove-file');
    const previewContainer = card.querySelector(`#previewContainer_${cardId}`);

    async function applySelectedFile(file) {
      cardData.file = file;
      previewContainer.innerHTML = '';

      if (file) {
        dz.classList.add('has-file');
        badge.textContent = `${file.name} (${formatBytes(file.size)})`;
        badge.style.display = 'inline-block';

        const previewEl = await inspectFileHeaders(file, `preview_comp_${cardId}`);
        if (previewEl) {
          previewContainer.appendChild(previewEl);
        }
      } else {
        dz.classList.remove('has-file');
        badge.textContent = '';
        badge.style.display = 'none';
      }
      updateSubmitButtonText();
    }

    if (initialFile) {
      applySelectedFile(initialFile);
    }

    fileInput.addEventListener('change', () => {
      if (fileInput.files.length > 0) {
        applySelectedFile(fileInput.files[0]);
      } else {
        applySelectedFile(null);
      }
    });

    ['dragover', 'dragenter'].forEach(e => {
      dz.addEventListener(e, (evt) => {
        evt.preventDefault();
        dz.classList.add('dragover');
      });
    });

    ['dragleave', 'drop'].forEach(e => {
      dz.addEventListener(e, () => {
        dz.classList.remove('dragover');
      });
    });

    dz.addEventListener('drop', (e) => {
      e.preventDefault();
      const files = e.dataTransfer.files;
      if (files && files.length > 0) {
        applySelectedFile(files[0]);
        for (let i = 1; i < files.length; i++) {
          addComparisonCard(files[i]);
        }
      }
    });

    removeBtn.addEventListener('click', () => {
      if (comparisonCards.length <= 1) {
        applySelectedFile(null);
        fileInput.value = '';
        return;
      }
      comparisonCards = comparisonCards.filter(c => c.id !== cardId);
      card.remove();
      renumberCards();
      updateSubmitButtonText();
    });

    updateSubmitButtonText();
  }

  // Initialize
  setupBaseDropzone();
  addComparisonCard();

  btnAddFile.addEventListener('click', () => {
    addComparisonCard();
  });

  // API accordion toggle
  const apiHeader = document.getElementById('apiHeader');
  const apiBody = document.getElementById('apiBody');
  if (apiHeader && apiBody) {
    apiHeader.addEventListener('click', () => {
      const isVisible = apiBody.style.display === 'block';
      apiBody.style.display = isVisible ? 'none' : 'block';
    });
  }

  // --- Form Submit Handler ---
  form.addEventListener('submit', async (e) => {
    e.preventDefault();
    errorBox.style.display = 'none';
    resultsCard.style.display = 'none';
    breakdownList.innerHTML = '';

    const baseFile = fileBase.files[0];
    if (!baseFile) {
      errorBox.textContent = 'Please select the Constant Base Report (HTML 1) before proceeding.';
      errorBox.style.display = 'block';
      return;
    }

    const activeComparisonFiles = comparisonCards
      .map(c => c.file)
      .filter(f => f !== null && f !== undefined);

    if (activeComparisonFiles.length === 0) {
      errorBox.textContent = 'Please upload at least one comparison HTML file (e.g. temadata.htm).';
      errorBox.style.display = 'block';
      return;
    }

    // UI Loading state & Progress Tracker
    submitBtn.disabled = true;
    btnSpinner.style.display = 'block';
    btnIcon.style.display = 'none';
    btnText.textContent = `Processing 1 Base + ${activeComparisonFiles.length} Comparison files in RAM...`;

    errorBox.style.display = 'none';
    resultsCard.style.display = 'none';
    breakdownList.innerHTML = '';
    tracker.start(activeComparisonFiles.length);

    const formData = new FormData();
    formData.append('file1', baseFile);
    activeComparisonFiles.forEach(file => {
      formData.append('comparison_files', file);
    });

    const tableIndex = document.getElementById('tableIndex').value || '4';
    formData.append('table_index', tableIndex);

    if (outputFilenameInput && outputFilenameInput.value.trim() !== '') {
      formData.append('output_filename', outputFilenameInput.value.trim());
    }

    const reportHeadingInput = document.getElementById('reportHeading');
    if (reportHeadingInput && reportHeadingInput.value.trim() !== '') {
      formData.append('report_heading', reportHeadingInput.value.trim());
    }

    // Only send column indices if in Custom Column mode
    if (modeCustom.checked) {
      const is1Based = base1Radio.checked;
      formData.append('index_base', is1Based ? '1' : '0');

      if (nopColIndex && nopColIndex.value.trim() !== '') {
        formData.append('nop_col_index', nopColIndex.value.trim());
      }
      if (premColIndex && premColIndex.value.trim() !== '') {
        formData.append('prem_col_index', premColIndex.value.trim());
      }
    }
    // If in Auto-detect mode, nop_col_index & prem_col_index are omitted:
    // backend automatically finds 'Total POL' and 'TOTAL PREM(SP+NSP)'!

    try {
      const response = await fetch('/api/convert', {
        method: 'POST',
        body: formData
      });

      if (!response.ok) {
        let errDetail = 'An error occurred during conversion.';
        try {
          const errJson = await response.json();
          errDetail = errJson.detail || errDetail;
        } catch (_) {
          errDetail = await response.text();
        }
        throw new Error(errDetail);
      }

      const matched = response.headers.get('X-Stats-Matched') || '0';
      const unmatched = response.headers.get('X-Stats-Unmatched') || '0';
      const totalAgents = response.headers.get('X-Stats-Total-Agents') || '0';
      const filesCount = response.headers.get('X-Stats-Files-Count') || String(activeComparisonFiles.length);

      // Determine dynamic date-stamped filename
      let downloadName = response.headers.get('X-Output-Filename');
      if (!downloadName) {
        const disposition = response.headers.get('Content-Disposition');
        if (disposition && disposition.includes('filename=')) {
          const match = disposition.match(/filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/);
          if (match && match[1]) {
            downloadName = match[1].replace(/['"]/g, '').trim();
          }
        }
      }
      if (!downloadName) {
        const prefix = (outputFilenameInput && outputFilenameInput.value.trim()) || 'report_';
        downloadName = `${prefix}_${todayDateStr}.xlsx`;
      }

      // Download binary Excel blob
      const blob = await response.blob();
      const downloadUrl = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = downloadUrl;
      a.download = downloadName;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(downloadUrl);

      if (downloadedFileName) {
        downloadedFileName.textContent = downloadName;
      }

      // Populate results
      document.getElementById('resTotalAgents').textContent = totalAgents;
      document.getElementById('resMatched').textContent = matched;
      document.getElementById('resUnmatched').textContent = unmatched;
      document.getElementById('resFilesCount').textContent = filesCount;

      // Fetch summary breakdown
      try {
        const summaryRes = await fetch('/api/summary', {
          method: 'POST',
          body: formData
        });
        if (summaryRes.ok) {
          const summaryJson = await summaryRes.json();
          const filesMeta = summaryJson.summary.files || [];
          filesMeta.forEach((f) => {
            const item = document.createElement('div');
            item.className = 'breakdown-item';
            item.innerHTML = `
              <div class="breakdown-file-name">
                <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline></svg>
                <span>${f.filename}</span>
                <small style="color:#64748b; margin-left:6px;">[POL: ${f.pol_col} | Prem: ${f.prem_col}]</small>
              </div>
              <div class="breakdown-badges">
                <span class="badge-stat matched">${f.matched} Matched</span>
                <span class="badge-stat">${f.total_policies.toLocaleString()} POL</span>
                <span class="badge-stat">₹ ${f.total_premium.toLocaleString()}</span>
              </div>
            `;
            breakdownList.appendChild(item);
          });
        }
      } catch (_) {
        // Non-critical
      }

      tracker.complete(() => {
        resultsCard.style.display = 'block';
      });

    } catch (err) {
      tracker.error(err.message);
      errorBox.textContent = err.message;
      errorBox.style.display = 'block';
    } finally {
      submitBtn.disabled = false;
      btnSpinner.style.display = 'none';
      btnIcon.style.display = 'block';
      updateSubmitButtonText();
    }
  });
});
