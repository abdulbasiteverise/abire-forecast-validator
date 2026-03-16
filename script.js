/* ═══════════════════════════════════════════════════
   ABIRE — Workforce Intelligence Platform
   script.js  — Pure-browser, no framework
   ═══════════════════════════════════════════════════ */

'use strict';

/* ───────────────────────────────────────────────────
   STATE
   ─────────────────────────────────────────────────── */
const S = {
  startTime:      '08:00',
  endTime:        '23:00',
  intervalMins:   15,
  weekCount:      5,
  intervals:      [],   // ['08:00','08:15',...]
  weeks:          [],   // [[], [], ...] per week per interval
  today:          [],   // today actuals per interval
  undoStack:      [],   // [{weeks, today}]
  charts:         {},
  pendingPaste:   null, // {cols:[[]], rowCount:int}
  dataMode:       'manual', // 'manual' | 'auto'
};

/* ───────────────────────────────────────────────────
   CONSTANTS
   ─────────────────────────────────────────────────── */
const EXAMPLE_INTERVALS_30 = [
  '08:00','08:30','09:00','09:30','10:00','10:30','11:00','11:30',
  '12:00','12:30','13:00','13:30','14:00','14:30','15:00','15:30',
  '16:00','16:30','17:00','17:30','18:00','18:30','19:00','19:30'
];
const EXAMPLE_WEEKS_30 = [
  [12,15,22,28,32,35,38,34,30,28,26,24,22,20,18,16,15,14,12,10,9,8,7,6],
  [11,14,21,27,31,34,37,33,29,27,25,23,21,19,17,15,14,13,11,9,8,7,6,5],
  [13,16,23,29,33,36,40,36,32,30,28,26,24,22,20,18,17,15,13,11,10,9,8,7],
  [10,13,20,26,30,33,36,32,28,26,24,22,20,18,16,14,13,12,10,8,7,6,5,4],
  [14,17,24,30,34,38,42,38,34,32,30,27,25,23,21,19,17,16,14,11,10,9,8,7],
];
const EXAMPLE_TODAY_30  = [13,16,25,32,38,44,48,42,38,34,31,28,26,23,21,19,18,16,14,12,10,9,8,7];

/* ───────────────────────────────────────────────────
   INTERVAL GENERATION
   ─────────────────────────────────────────────────── */
function timeToMins(t) {
  const [h, m] = t.split(':').map(Number);
  return h * 60 + m;
}
function minsToTime(m) {
  const h = Math.floor(m / 60) % 24;
  const mm = m % 60;
  return `${String(h).padStart(2,'0')}:${String(mm).padStart(2,'0')}`;
}
function generateIntervals(start, end, stepMins) {
  const list = [];
  let cur = timeToMins(start);
  const endM = timeToMins(end);
  while (cur <= endM) {
    list.push(minsToTime(cur));
    cur += stepMins;
  }
  return list;
}
function rebuildIntervalState() {
  const newIntervals = generateIntervals(S.startTime, S.endTime, S.intervalMins);
  const n = newIntervals.length;
  // Resize weeks data
  for (let w = 0; w < S.weekCount; w++) {
    if (!S.weeks[w]) S.weeks[w] = [];
    S.weeks[w].length = n;
    for (let i = 0; i < n; i++) {
      if (S.weeks[w][i] === undefined || S.weeks[w][i] === null) S.weeks[w][i] = 0;
    }
  }
  S.today.length = n;
  for (let i = 0; i < n; i++) {
    if (S.today[i] === undefined || S.today[i] === null) S.today[i] = 0;
  }
  S.intervals = newIntervals;
}

/* ───────────────────────────────────────────────────
   TABLE RENDERING
   ─────────────────────────────────────────────────── */
function buildTableHeaders() {
  let h = '<tr><th>Interval</th>';
  for (let w = 1; w <= S.weekCount; w++) h += `<th>WK${w}</th>`;
  h += '<th>Avg</th><th>Today</th><th>Dev%</th></tr>';
  document.getElementById('tbl-head').innerHTML = h;
}

function buildTableRows() {
  const tbody = document.getElementById('tbl-body');
  const n = S.intervals.length;
  let rows = '';
  for (let i = 0; i < n; i++) {
    const wkVals = [];
    for (let w = 0; w < S.weekCount; w++) wkVals.push(S.weeks[w] ? (S.weeks[w][i] || 0) : 0);
    const avg = wkVals.reduce((a,b)=>a+b,0) / (wkVals.length || 1);
    const tod = S.today[i] || 0;
    const dev = avg ? ((tod - avg) / avg * 100) : 0;
    const devStr = (dev >= 0 ? '+' : '') + dev.toFixed(1) + '%';
    const devCls = Math.abs(dev) > 15 ? 'dev-critical' : Math.abs(dev) > 10 ? 'dev-major' : Math.abs(dev) > 5 ? 'dev-mild' : 'dev-normal';

    let wkCells = '';
    for (let w = 0; w < S.weekCount; w++) {
      wkCells += `<td><input class="tc" type="number" min="0" value="${S.weeks[w][i]||0}" data-row="${i}" data-wk="${w}" tabindex="0"></td>`;
    }
    rows += `<tr>
      <td><span class="int-lbl">${S.intervals[i]}</span></td>
      ${wkCells}
      <td class="td-avg">${avg.toFixed(1)}</td>
      <td><input class="tc tc-today" type="number" min="0" value="${tod}" data-row="${i}" data-today="1" tabindex="0"></td>
      <td class="td-dev ${devCls}">${devStr}</td>
    </tr>`;
  }
  tbody.innerHTML = rows;
  attachTableListeners();
}

function buildTable() {
  buildTableHeaders();
  buildTableRows();
  updateWeekCountLabel();
  updateBulkHeaders();
}

function updateWeekCountLabel() {
  const el = document.getElementById('week-count-label');
  if (el) el.textContent = `${S.intervals.length} intervals · ${S.weekCount} historical week${S.weekCount !== 1 ? 's' : ''}`;
}

/* ───────────────────────────────────────────────────
   TABLE CELL LISTENERS
   ─────────────────────────────────────────────────── */
function attachTableListeners() {
  document.querySelectorAll('.tc').forEach(inp => {
    inp.addEventListener('change', onCellChange);
    inp.addEventListener('keydown', onCellKeydown);
    inp.addEventListener('paste', onCellPaste);
  });
}

function onCellChange(e) {
  const el = e.target;
  const row = parseInt(el.dataset.row);
  const val = Math.max(0, parseFloat(el.value) || 0);
  el.value = val;
  if (el.dataset.today) {
    S.today[row] = val;
  } else {
    const wk = parseInt(el.dataset.wk);
    S.weeks[wk][row] = val;
  }
  updateRowStats(row);
  clearValidation();
}

function updateRowStats(row) {
  const tr = document.querySelectorAll('#tbl-body tr')[row];
  if (!tr) return;
  const wkVals = [];
  for (let w = 0; w < S.weekCount; w++) wkVals.push(S.weeks[w][row] || 0);
  const avg = wkVals.reduce((a,b)=>a+b,0) / (wkVals.length || 1);
  const tod = S.today[row] || 0;
  const dev = avg ? ((tod - avg) / avg * 100) : 0;
  const devStr = (dev >= 0 ? '+' : '') + dev.toFixed(1) + '%';
  const devCls = Math.abs(dev) > 15 ? 'dev-critical' : Math.abs(dev) > 10 ? 'dev-major' : Math.abs(dev) > 5 ? 'dev-mild' : 'dev-normal';
  const avgCell = tr.querySelector('.td-avg');
  const devCell = tr.querySelector('.td-dev');
  if (avgCell) avgCell.textContent = avg.toFixed(1);
  if (devCell) { devCell.textContent = devStr; devCell.className = `td-dev ${devCls}`; }
}

/* ───────────────────────────────────────────────────
   KEYBOARD NAVIGATION IN TABLE
   ─────────────────────────────────────────────────── */
function onCellKeydown(e) {
  if (e.key === 'Tab') return; // natural tab
  const cells = Array.from(document.querySelectorAll('.tc'));
  const idx = cells.indexOf(e.target);
  if (idx < 0) return;
  if (e.key === 'ArrowDown' || e.key === 'Enter') {
    e.preventDefault();
    // move to same column, next row
    const totalCols = S.weekCount + 1; // wk cells + today cell
    const nextIdx = idx + totalCols;
    if (cells[nextIdx]) cells[nextIdx].focus();
  } else if (e.key === 'ArrowUp') {
    e.preventDefault();
    const totalCols = S.weekCount + 1;
    const prevIdx = idx - totalCols;
    if (prevIdx >= 0 && cells[prevIdx]) cells[prevIdx].focus();
  } else if (e.key === 'ArrowRight') {
    if (cells[idx + 1]) cells[idx + 1].focus();
  } else if (e.key === 'ArrowLeft') {
    if (idx > 0 && cells[idx - 1]) cells[idx - 1].focus();
  }
}

/* ───────────────────────────────────────────────────
   EXCEL PASTE INTO CELL
   ─────────────────────────────────────────────────── */
function onCellPaste(e) {
  e.preventDefault();
  const text = (e.clipboardData || window.clipboardData).getData('text');
  const el = e.target;
  const startRow = parseInt(el.dataset.row);
  const startWk  = el.dataset.today ? null : parseInt(el.dataset.wk);
  const isToday  = !!el.dataset.today;

  const parsed = parseClipboard(text);
  if (!parsed.cols.length) return;

  // Save undo
  pushUndo();

  // Apply
  const n = S.intervals.length;
  if (isToday || parsed.cols.length === 1) {
    // Single column → fill down Today or the target week column
    const col = parsed.cols[0];
    for (let i = 0; i < col.length; i++) {
      const row = startRow + i;
      if (row >= n) break;
      if (isToday) S.today[row] = col[i];
      else if (startWk !== null) S.weeks[startWk][row] = col[i];
    }
  } else {
    // Multi-column → fill down+right
    for (let c = 0; c < parsed.cols.length; c++) {
      const wk = isToday ? null : (startWk !== null ? startWk + c : c);
      if (wk !== null && wk >= S.weekCount) continue;
      for (let r = 0; r < parsed.cols[c].length; r++) {
        const row = startRow + r;
        if (row >= n) break;
        if (wk !== null) S.weeks[wk][row] = parsed.cols[c][r];
      }
    }
  }

  buildTableRows();
  flashPasteCells(startRow, Math.min(startRow + parsed.rowCount, n));
  showValidation(`✓ Pasted ${parsed.rowCount} row(s) × ${parsed.cols.length} col(s)`, 'ok');
  showToast(`Pasted ${parsed.rowCount} rows`);
}

function flashPasteCells(startRow, endRow) {
  const cells = document.querySelectorAll('.tc');
  cells.forEach(c => {
    const r = parseInt(c.dataset.row);
    if (r >= startRow && r < endRow) c.classList.add('paste-highlight');
  });
  setTimeout(() => document.querySelectorAll('.paste-highlight').forEach(c => c.classList.remove('paste-highlight')), 600);
}

/* ───────────────────────────────────────────────────
   CLIPBOARD PARSER — TSV / CSV / newline
   ─────────────────────────────────────────────────── */
function parseClipboard(text) {
  const lines = text.trim().split(/\r?\n/);
  // Detect delimiter
  const firstLine = lines[0];
  let delim = '\t';
  if (!firstLine.includes('\t')) {
    if (firstLine.includes(',')) delim = ',';
    else delim = ' ';
  }
  const rows = lines.map(l => l.split(delim).map(v => parseFloat(v.trim())).filter(v => !isNaN(v)));
  const maxCols = Math.max(...rows.map(r => r.length));
  const cols = [];
  for (let c = 0; c < maxCols; c++) {
    cols.push(rows.map(r => r[c] !== undefined ? r[c] : 0));
  }
  return { cols, rowCount: rows.length };
}

/* ───────────────────────────────────────────────────
   BULK PASTE PANELS
   ─────────────────────────────────────────────────── */
function updateBulkHeaders() {
  const cont = document.getElementById('bulk-cols-container');
  if (!cont) return;
  let html = '';
  for (let w = 1; w <= S.weekCount; w++) {
    html += bulkColBox(`WK-${w}`, w - 1, false);
  }
  html += bulkColBox('TODAY', -1, true);
  cont.innerHTML = html;
}

function bulkColBox(label, wkIdx, isToday) {
  const id = isToday ? 'bulk-today' : `bulk-wk-${wkIdx}`;
  return `<div class="bulk-col-box">
    <div class="bulk-col-lbl">
      <span>${label}</span>
      <button class="btn btn-sm btn-neutral" onclick="clearBulkCol('${id}')">✕</button>
    </div>
    <textarea class="bulk-textarea" id="${id}" placeholder="Paste values here\none per line\nor tab-separated…"></textarea>
    <div class="bulk-col-actions">
      <button class="btn btn-sm btn-primary" onclick="applyBulkCol('${id}', ${wkIdx}, ${isToday})">Apply</button>
      <button class="btn btn-sm btn-neutral" onclick="previewBulkCol('${id}', '${label}')">Preview</button>
    </div>
  </div>`;
}

function clearBulkCol(id) {
  const el = document.getElementById(id);
  if (el) el.value = '';
}

function applyBulkCol(id, wkIdx, isToday) {
  const el = document.getElementById(id);
  if (!el || !el.value.trim()) return;
  const text = el.value.trim();
  const parsed = parseClipboard(text);
  const col = parsed.cols[0] || [];
  const n = S.intervals.length;
  if (col.length > n) {
    showValidation(`⚠ Pasted ${col.length} values but only ${n} intervals exist. Extra values ignored.`, 'warn');
  }
  pushUndo();
  for (let i = 0; i < Math.min(col.length, n); i++) {
    if (isToday) S.today[i] = col[i];
    else if (wkIdx >= 0 && wkIdx < S.weekCount) S.weeks[wkIdx][i] = col[i];
  }
  buildTableRows();
  showToast(`${isToday ? 'Today' : 'WK' + (wkIdx + 1)} updated — ${Math.min(col.length, n)} intervals`);
  clearValidation();
  el.value = '';
}

function previewBulkCol(id, label) {
  const el = document.getElementById(id);
  if (!el || !el.value.trim()) { showToast('Nothing to preview', 'warn'); return; }
  const parsed = parseClipboard(el.value.trim());
  const col = parsed.cols[0] || [];
  const preview = col.slice(0, 10).join(', ') + (col.length > 10 ? ` … (+${col.length - 10} more)` : '');
  document.getElementById('paste-preview').style.display = 'block';
  document.getElementById('paste-preview-title').textContent = `Preview: ${label} — ${col.length} values detected`;
  document.getElementById('paste-preview-content').textContent = preview;
}

function dismissPreview() {
  document.getElementById('paste-preview').style.display = 'none';
}

/* ───────────────────────────────────────────────────
   UNDO
   ─────────────────────────────────────────────────── */
function pushUndo() {
  S.undoStack.push({
    weeks: S.weeks.map(w => [...w]),
    today: [...S.today],
  });
  if (S.undoStack.length > 20) S.undoStack.shift();
  showUndoBar();
}

function undoLast() {
  if (!S.undoStack.length) { showToast('Nothing to undo', 'warn'); return; }
  const snap = S.undoStack.pop();
  S.weeks = snap.weeks.map(w => [...w]);
  S.today = [...snap.today];
  buildTableRows();
  showToast('Undo applied');
  if (!S.undoStack.length) hideUndoBar();
}

function showUndoBar() {
  const b = document.getElementById('undo-bar');
  if (b) { b.style.display = 'block'; b.textContent = `↩ Undo available (${S.undoStack.length} step${S.undoStack.length !== 1 ? 's' : ''})`; }
}
function hideUndoBar() {
  const b = document.getElementById('undo-bar');
  if (b) b.style.display = 'none';
}

/* ───────────────────────────────────────────────────
   INTERVAL CONFIGURATION CONTROLS
   ─────────────────────────────────────────────────── */
function onConfigChange() {
  const startEl  = document.getElementById('cfg-start');
  const endEl    = document.getElementById('cfg-end');
  const stepEl   = document.getElementById('cfg-step');

  S.startTime    = startEl.value || '08:00';
  S.endTime      = endEl.value   || '23:00';
  S.intervalMins = parseInt(stepEl.value) || 15;

  // Validate
  if (timeToMins(S.startTime) >= timeToMins(S.endTime)) {
    showValidation('⚠ End time must be after start time.', 'error');
    return;
  }
  clearValidation();
  rebuildIntervalState();
  buildTable();
  showToast(`Table rebuilt: ${S.intervals.length} intervals`);
}

function addWeek() {
  S.weekCount++;
  const n = S.intervals.length;
  if (!S.weeks[S.weekCount - 1]) {
    S.weeks.push(new Array(n).fill(0));
  }
  buildTable();
  showToast(`WK${S.weekCount} added`);
}

function removeWeek() {
  if (S.weekCount <= 1) { showToast('Minimum 1 week required', 'warn'); return; }
  S.weekCount--;
  buildTable();
  showToast(`WK${S.weekCount + 1} removed`);
}

/* ───────────────────────────────────────────────────
   DATA SOURCE TOGGLE
   ─────────────────────────────────────────────────── */
function setDataMode(mode) {
  S.dataMode = mode;
  document.querySelectorAll('.ds-toggle-btn').forEach(b => {
    b.classList.toggle('active', b.dataset.mode === mode);
  });
  document.getElementById('api-panel').style.display = mode === 'auto' ? 'block' : 'none';
}

/* ───────────────────────────────────────────────────
   API INGESTION
   ─────────────────────────────────────────────────── */
function ingestAPI() {
  const raw = document.getElementById('api-json').value.trim();
  if (!raw) { showToast('Paste JSON first', 'warn'); return; }
  try {
    const data = JSON.parse(raw);
    const arr = data.intervals || data;
    if (!Array.isArray(arr)) throw new Error('Expected array of {time, volume}');
    pushUndo();
    const n = S.intervals.length;
    arr.forEach((item, i) => {
      // Match by time or by index
      const timeIdx = S.intervals.indexOf(item.time);
      const idx = timeIdx >= 0 ? timeIdx : i;
      if (idx < n) S.today[idx] = parseFloat(item.volume) || 0;
    });
    buildTableRows();
    showToast(`API ingested: ${arr.length} intervals → Today column`);
    clearValidation();
  } catch (err) {
    showValidation('JSON parse error: ' + err.message, 'error');
  }
}

function loadAPIExample() {
  const example = S.intervals.slice(0, 10).map((t, i) => ({ time: t, volume: 15 + i * 3 }));
  document.getElementById('api-json').value = JSON.stringify({ intervals: example }, null, 2);
}

/* ───────────────────────────────────────────────────
   CSV UPLOAD
   ─────────────────────────────────────────────────── */
function handleCSVUpload(event) {
  const file = event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = e => parseCSV(e.target.result);
  reader.readAsText(file);
  event.target.value = '';
}

function parseCSV(text) {
  try {
    const lines = text.trim().split(/\r?\n/);
    const headers = lines[0].split(',').map(h => h.trim().toLowerCase());
    const intervalIdx = headers.indexOf('interval');
    const todayIdx    = headers.indexOf('today');
    const weekIndices = headers.map((h, i) => h.startsWith('wk') ? i : -1).filter(i => i >= 0);

    if (intervalIdx < 0) { showValidation('CSV must include an "interval" column.', 'error'); return; }

    pushUndo();
    S.weekCount = Math.max(weekIndices.length, 1);
    const newIntervals = [], newWeeks = weekIndices.map(() => []), newToday = [];

    for (let l = 1; l < lines.length; l++) {
      const cols = lines[l].split(',').map(c => c.trim());
      if (!cols[intervalIdx]) continue;
      newIntervals.push(cols[intervalIdx]);
      weekIndices.forEach((wi, j) => newWeeks[j].push(parseFloat(cols[wi]) || 0));
      newToday.push(todayIdx >= 0 ? (parseFloat(cols[todayIdx]) || 0) : 0);
    }

    S.intervals = newIntervals;
    S.weeks = newWeeks;
    S.today = newToday;

    // Sync config UI
    if (newIntervals.length > 1) {
      const step = timeToMins(newIntervals[1]) - timeToMins(newIntervals[0]);
      S.intervalMins = step > 0 ? step : 15;
      S.startTime = newIntervals[0];
      S.endTime   = newIntervals[newIntervals.length - 1];
      syncConfigUI();
    }
    buildTable();
    showToast(`CSV loaded: ${newIntervals.length} intervals, ${S.weekCount} weeks`);
  } catch (err) {
    showValidation('CSV error: ' + err.message, 'error');
  }
}

function syncConfigUI() {
  const s = document.getElementById('cfg-start');
  const e = document.getElementById('cfg-end');
  const st = document.getElementById('cfg-step');
  if (s) s.value = S.startTime;
  if (e) e.value = S.endTime;
  if (st) st.value = String(S.intervalMins);
}

/* ───────────────────────────────────────────────────
   LOAD EXAMPLE
   ─────────────────────────────────────────────────── */
function loadExample() {
  S.startTime    = '08:00';
  S.endTime      = '19:30';
  S.intervalMins = 30;
  S.weekCount    = 5;
  S.intervals    = [...EXAMPLE_INTERVALS_30];
  S.weeks        = EXAMPLE_WEEKS_30.map(w => [...w]);
  S.today        = [...EXAMPLE_TODAY_30];
  S.undoStack    = [];

  syncConfigUI();
  buildTable();
  hideUndoBar();

  document.getElementById('daily-forecast').value  = 2200;
  document.getElementById('actual-calls').value    = 1340;
  document.getElementById('hist-completion').value = 60;
  document.getElementById('current-interval').value= '13:00';
  document.getElementById('agents').value          = 45;
  document.getElementById('aht').value             = 280;
  document.getElementById('shrinkage').value       = 30;
  document.getElementById('interval-len').value    = 30;

  showToast('Example WFM data loaded');
}

/* ───────────────────────────────────────────────────
   VALIDATION DISPLAY
   ─────────────────────────────────────────────────── */
function showValidation(msg, type) {
  const bar = document.getElementById('validation-bar');
  if (!bar) return;
  bar.textContent = msg;
  bar.className = type;
  bar.style.display = 'block';
}
function clearValidation() {
  const bar = document.getElementById('validation-bar');
  if (bar) bar.style.display = 'none';
}

/* ───────────────────────────────────────────────────
   MAIN ANALYSIS ENGINE
   ─────────────────────────────────────────────────── */
function runAnalysis() {
  setStatus('running');

  // Short delay for UI to update
  requestAnimationFrame(() => setTimeout(() => {
    try {
      const dailyFc   = parseFloat(document.getElementById('daily-forecast').value) || 0;
      const actualCalls = parseFloat(document.getElementById('actual-calls').value) || 0;
      const histPct   = parseFloat(document.getElementById('hist-completion').value) || 60;

      /* — Reforecast — */
      const expected  = dailyFc * (histPct / 100);
      const variance  = expected ? ((actualCalls - expected) / expected * 100) : 0;
      const corrected = expected ? (dailyFc * (actualCalls / expected)) : dailyFc;

      /* — Forecast Bias — */
      const bias = dailyFc ? ((actualCalls - dailyFc) / dailyFc * 100) : 0;

      /* — CV (variability) — */
      const dailyTotals = [];
      for (let w = 0; w < S.weekCount; w++) {
        if (S.weeks[w]) dailyTotals.push(S.weeks[w].reduce((a,b)=>a+b,0));
      }
      const mean = dailyTotals.length ? dailyTotals.reduce((a,b)=>a+b,0)/dailyTotals.length : 0;
      const std  = dailyTotals.length > 1
        ? Math.sqrt(dailyTotals.map(v=>(v-mean)**2).reduce((a,b)=>a+b,0)/dailyTotals.length) : 0;
      const cv   = mean ? (std/mean*100) : 0;

      /* — Spike — */
      const todayTotal = S.today.reduce((a,b)=>a+b,0);
      const deviation  = mean ? ((todayTotal - mean) / mean * 100) : 0;

      /* — Capacity — */
      const agents       = parseFloat(document.getElementById('agents').value) || 0;
      const aht          = parseFloat(document.getElementById('aht').value) || 280;
      const shrinkage    = parseFloat(document.getElementById('shrinkage').value) || 30;
      const intLen       = parseFloat(document.getElementById('interval-len').value) || 30;
      const effAgents    = agents * (1 - shrinkage / 100);
      const callsPerAgt  = (intLen * 60) / aht;
      const totalCap     = effAgents * callsPerAgt;
      const remaining    = corrected - actualCalls;
      const utilization  = totalCap ? (remaining / totalCap * 100) : 0;
      const riskLevel    = utilization <= 90 ? 'safe' : utilization <= 110 ? 'moderate' : 'critical';

      /* — Render KPIs — */
      document.getElementById('kpi-row').style.display = 'grid';
      renderKPIs(bias, corrected, variance, cv, riskLevel, utilization);

      /* — Analysis panels — */
      document.getElementById('analysis-section').style.display = 'block';
      renderReforecast(dailyFc, expected, actualCalls, variance, corrected);
      detectAndRenderPattern(deviation, variance, histPct);
      renderSpike(mean, todayTotal, deviation);

      /* — Capacity — */
      document.getElementById('capacity-results').style.display = 'grid';
      renderCapacity(effAgents, callsPerAgt, totalCap, riskLevel, utilization);

      /* — Charts — */
      document.getElementById('charts-section').style.display = 'block';
      renderCharts(dailyFc, corrected, expected, actualCalls, totalCap, remaining);

      /* — Recs + Summary — */
      document.getElementById('recommendations-section').style.display = 'block';
      generateRecommendations(bias, deviation, utilization, cv, riskLevel);
      generateExecSummary(bias, corrected, dailyFc, cv, variance, deviation, riskLevel, effAgents, totalCap, remaining);

      setStatus('done');
    } catch(err) {
      setStatus('idle');
      showToast('Analysis error: ' + err.message, 'err');
      console.error(err);
    }
  }, 150));
}

/* ───────────────────────────────────────────────────
   RENDER HELPERS
   ─────────────────────────────────────────────────── */
function renderKPIs(bias, corrected, variance, cv, riskLevel, utilization) {
  // Bias
  const biasEl = document.getElementById('kpi-bias');
  biasEl.textContent = fmtPct(bias);
  biasEl.className   = 'kpi-val ' + (Math.abs(bias)<=5?'c-green':Math.abs(bias)<=10?'c-yellow':'c-red');
  const biasBadge = Math.abs(bias)<=5
    ? '<span class="badge b-safe">Accurate</span>'
    : Math.abs(bias)<=10
      ? '<span class="badge b-warn">Moderate Bias</span>'
      : '<span class="badge b-danger">High Bias</span>';
  document.getElementById('bias-badge').innerHTML = biasBadge;

  // Reforecast
  document.getElementById('kpi-reforecast').textContent = Math.round(corrected).toLocaleString();

  // Variance
  const varEl = document.getElementById('kpi-variance');
  varEl.textContent = fmtPct(variance);
  varEl.className   = 'kpi-val ' + (Math.abs(variance)<=5?'c-green':Math.abs(variance)<=10?'c-yellow':'c-red');
  document.getElementById('kpi-variance-sub').textContent = variance>=0?'Above Expected':'Below Expected';

  // CV
  const cvEl = document.getElementById('kpi-cv');
  cvEl.textContent = cv.toFixed(1)+'%';
  cvEl.className   = 'kpi-val ' + (cv<10?'c-green':cv<20?'c-yellow':'c-red');
  document.getElementById('cv-badge').innerHTML = cv<10
    ? '<span class="badge b-safe">Stable</span>'
    : cv<20
      ? '<span class="badge b-warn">Moderate</span>'
      : '<span class="badge b-danger">High Variability</span>';

  // Risk
  document.getElementById('kpi-risk-label').textContent  = riskLevel.toUpperCase();
  document.getElementById('kpi-risk-label').className    = 'risk-lbl risk-'+riskLevel;
}

function renderReforecast(dailyFc, expected, actualCalls, variance, corrected) {
  setText('rf-daily',    dailyFc.toLocaleString());
  setText('rf-expected', Math.round(expected).toLocaleString());
  setText('rf-actual',   actualCalls.toLocaleString());
  setText('rf-variance', fmtPct(variance));
  setText('rf-corrected',Math.round(corrected).toLocaleString());
}

function renderCapacity(effAgents, callsPerAgt, totalCap, riskLevel, utilization) {
  setText('cap-effective',  effAgents.toFixed(1));
  setText('cap-per-agent',  callsPerAgt.toFixed(2));
  setText('cap-total',      Math.round(totalCap).toLocaleString());
  const el = document.getElementById('cap-risk-label');
  el.textContent = riskLevel.toUpperCase();
  el.className   = 'risk-lbl risk-'+riskLevel;
  setText('cap-risk-sub', `Utilization: ${utilization.toFixed(0)}%`);
}

function renderSpike(histAvg, todayTotal, deviation) {
  setText('spike-hist-avg', Math.round(histAvg).toLocaleString());
  setText('spike-today',    todayTotal.toLocaleString());
  setText('spike-dev',      fmtPct(deviation));
  const abs = Math.abs(deviation);
  const badge = abs<5
    ? '<span class="badge b-safe">Normal</span>'
    : abs<10
      ? '<span class="badge b-info">Mild Spike</span>'
      : abs<15
        ? '<span class="badge b-warn">Major Spike</span>'
        : '<span class="badge b-danger">Critical Spike</span>';
  document.getElementById('spike-badge').innerHTML = badge;
  setText('spike-drift', deviation>=0?'▲ Above Baseline':'▼ Below Baseline');
}

/* ───────────────────────────────────────────────────
   PATTERN DETECTION
   ─────────────────────────────────────────────────── */
function detectAndRenderPattern(deviation, variance, histPct) {
  const today = S.today;
  if (!today || !today.length) return;
  const total = today.reduce((a,b)=>a+b,0);
  if (!total) return;

  let cum = 0;
  const cumPcts = today.map(v => { cum += v; return cum / total; });
  const n = today.length;
  const q1 = cumPcts[Math.floor(n*0.25)] || 0;
  const q3 = 1 - (cumPcts[Math.floor(n*0.75)] || 0);
  const peakIdx = today.indexOf(Math.max(...today));
  const peakPct = peakIdx / n;

  let icon, name, desc;
  if (q1 > 0.38) {
    icon='⚡'; name='Front-Loaded Traffic';
    desc='Heavy early-day volume. Ensure max staffing at open. Protect morning intervals and delay lunch breaks.';
  } else if (q3 > 0.30) {
    icon='🌆'; name='Evening Surge';
    desc='Volume peaks late in the day. Late shift coverage is critical. Monitor evening SLA closely.';
  } else if (peakPct > 0.35 && peakPct < 0.65) {
    icon='🔥'; name='Midday Spike';
    desc='Volume concentrates around midday. Reduce breaks 11:00–14:00. Protect peak coverage.';
  } else if (peakPct < 0.22) {
    icon='⏳'; name='Delayed Arrival';
    desc='Slow early start. Volume arrives later than expected. Flexible staffing and flexible breaks recommended.';
  } else if (Math.abs(deviation) > 12 && peakPct > 0.3) {
    icon='📈'; name='Double Peak';
    desc='Two distinct volume peaks detected. Apply split-shift staffing patterns and interval-level planning.';
  } else {
    icon='✅'; name='Normal Pattern';
    desc='Volume aligns with historical baseline. Standard staffing plan is appropriate.';
  }

  setText('pattern-icon', icon);
  setText('pattern-name', name);
  setText('pattern-desc', desc);
}

/* ───────────────────────────────────────────────────
   RECOMMENDATIONS
   ─────────────────────────────────────────────────── */
function generateRecommendations(bias, deviation, utilization, cv, riskLevel) {
  const recs = [];
  if (Math.abs(bias) > 10) {
    if (bias > 0) recs.push({t:'critical',ico:'🚨',m:`High positive forecast bias (+${bias.toFixed(1)}%). Actual demand is significantly exceeding forecast. Recommend immediate staffing escalation.`});
    else recs.push({t:'warn',ico:'⚠️',m:`High negative forecast bias (${bias.toFixed(1)}%). Demand is running below forecast. Consider redeploying agents or approving voluntary time off.`});
  } else if (Math.abs(bias) > 5) {
    recs.push({t:'warn',ico:'📊',m:`Moderate forecast bias (${bias.toFixed(1)}%). Monitor closely and adjust next-day planning assumptions.`});
  } else {
    recs.push({t:'safe',ico:'✅',m:`Forecast bias within acceptable range (${bias.toFixed(1)}%). No immediate correction needed.`});
  }
  if (Math.abs(deviation) > 15) recs.push({t:'critical',ico:'🔥',m:`Critical spike detected (${deviation.toFixed(1)}% above baseline). Activate contingency staffing plan immediately.`});
  else if (Math.abs(deviation) > 10) recs.push({t:'warn',ico:'⚡',m:`Major spike in volume (${deviation.toFixed(1)}% deviation). Increase coverage for remaining intervals.`});
  if (riskLevel==='critical') recs.push({t:'critical',ico:'💥',m:`Backlog risk: Demand exceeds capacity (${utilization.toFixed(0)}% utilization). Authorize overtime or pull agents from back-office now.`});
  else if (riskLevel==='moderate') recs.push({t:'warn',ico:'⚠️',m:`Moderate capacity pressure (${utilization.toFixed(0)}% utilization). Review interval staffing for next 2 hours.`});
  else recs.push({t:'safe',ico:'✅',m:`Capacity appears adequate (${utilization.toFixed(0)}% utilization). Continue monitoring variance through peak windows.`});
  if (cv > 20) recs.push({t:'info',ico:'📉',m:`High demand variability (CV: ${cv.toFixed(1)}%). Historical data inconsistent — use wider staffing buffers.`});
  else if (cv > 10) recs.push({t:'info',ico:'📌',m:`Moderate variability (CV: ${cv.toFixed(1)}%). Consider safety staffing of +5–8% above modelled requirement.`});

  document.getElementById('recommendations-panel').innerHTML =
    recs.map(r=>`<div class="rec-item ${r.t}"><span>${r.ico}</span><span>${r.m}</span></div>`).join('');
}

/* ───────────────────────────────────────────────────
   EXEC SUMMARY
   ─────────────────────────────────────────────────── */
function generateExecSummary(bias, corrected, dailyFc, cv, variance, deviation, riskLevel, effAgents, totalCap, remaining) {
  const biasLbl = Math.abs(bias)<=5?'Accurate':Math.abs(bias)<=10?'Moderate Bias':'High Bias';
  const cvLbl   = cv<10?'Stable':cv<20?'Moderate':'High';
  const conf    = (Math.abs(bias)<=5&&cv<10)?'High':(Math.abs(bias)<=10&&cv<20)?'Medium':'Low';
  const trend   = bias>5?'Traffic trending ABOVE historical baseline.':bias<-5?'Traffic trending BELOW historical baseline.':'Traffic aligning with historical baseline.';
  const rec     = riskLevel==='critical'?'CRITICAL: Activate contingency staffing. Authorize overtime immediately.':riskLevel==='moderate'?'MODERATE RISK: Review staffing for upcoming intervals.':'Capacity appears sufficient. Continue monitoring through peak windows.';
  const ts      = new Date().toLocaleString('en-GB',{day:'2-digit',month:'short',year:'numeric',hour:'2-digit',minute:'2-digit'});
  const absDevLbl = Math.abs(deviation)<5?'Normal':Math.abs(deviation)<10?'Mild Spike':Math.abs(deviation)<15?'Major Spike':'Critical Spike';

  const s = `══════════════════════════════════════
  ABIRE FORECAST ANALYSIS REPORT
  Generated: ${ts}
  Intervals: ${S.intervals.length} · Step: ${S.intervalMins}min
══════════════════════════════════════

  FORECAST INTELLIGENCE
  ─────────────────────────────────────
  Original Forecast   : ${Number(dailyFc).toLocaleString()}
  Corrected (EOD)     : ${Math.round(corrected).toLocaleString()}
  Forecast Bias       : ${fmtPct(bias)}  [${biasLbl}]
  Intraday Variance   : ${fmtPct(variance)}

  DEMAND VARIABILITY
  ─────────────────────────────────────
  CV Index            : ${cv.toFixed(1)}%  [${cvLbl} Variability]
  Confidence Level    : ${conf}
  Spike Severity      : ${absDevLbl}  (${fmtPct(deviation)})

  CAPACITY ASSESSMENT
  ─────────────────────────────────────
  Effective Agents    : ${effAgents.toFixed(1)}
  Capacity (interval) : ${Math.round(totalCap)}
  Remaining Demand    : ${Math.round(remaining)}
  Risk Level          : ${riskLevel.toUpperCase()}

  OPERATIONAL INSIGHT
  ─────────────────────────────────────
  ${trend}

  RECOMMENDATION
  ─────────────────────────────────────
  ${rec}

══════════════════════════════════════
  ABIRE | Workforce Intelligence Suite
  Built by Basit
══════════════════════════════════════`;

  setText('exec-summary-box', s);
}

/* ───────────────────────────────────────────────────
   CHARTS (Chart.js)
   ─────────────────────────────────────────────────── */
function renderCharts(dailyFc, corrected, expected, actualCalls, totalCap, remaining) {
  const labels   = S.intervals;
  const todayV   = S.today;
  const histAvgs = labels.map((_,i) => {
    let sum = 0, cnt = 0;
    for (let w = 0; w < S.weekCount; w++) { if(S.weeks[w]){sum+=S.weeks[w][i]||0;cnt++;} }
    return cnt ? sum/cnt : 0;
  });

  const base = {
    responsive:true, maintainAspectRatio:false,
    plugins:{legend:{labels:{color:'#7ab8d4',font:{family:'Exo 2',size:10}}}},
    scales:{
      x:{ticks:{color:'#2e607a',font:{size:9},maxTicksLimit:12},grid:{color:'rgba(12,37,64,0.4)'}},
      y:{ticks:{color:'#2e607a',font:{size:9}},grid:{color:'rgba(12,37,64,0.4)'}}
    }
  };

  // 1. Historical trend
  const totals = [];
  for(let w=0;w<S.weekCount;w++) totals.push(S.weeks[w]?S.weeks[w].reduce((a,b)=>a+b,0):0);
  mkChart('chart-trend','line',
    {labels:totals.map((_,i)=>`WK${i+1}`),
     datasets:[{label:'Daily Volume',data:totals,borderColor:'#00d4ff',backgroundColor:'rgba(0,212,255,0.07)',fill:true,tension:0.4,pointBackgroundColor:'#00d4ff',pointRadius:4}]},
    base);

  // 2. Forecast vs Actual
  mkChart('chart-fa','bar',
    {labels,
     datasets:[
       {label:'Hist Avg',data:histAvgs,backgroundColor:'rgba(122,184,212,0.25)',borderColor:'#7ab8d4',borderWidth:1},
       {label:'Today',   data:todayV,  backgroundColor:'rgba(0,255,157,0.22)',  borderColor:'#00ff9d',borderWidth:1}
     ]},
    base);

  // 3. Cumulative
  let ch=0,ct=0;
  const cumH = histAvgs.map(v=>{ch+=v;return ch;});
  const cumT = todayV.map(v=>{ct+=v;return ct;});
  mkChart('chart-cumulative','line',
    {labels,
     datasets:[
       {label:'Expected (Hist)',data:cumH,borderColor:'#f7c948',fill:false,tension:0.4,borderDash:[5,4]},
       {label:'Actual Today',  data:cumT,borderColor:'#00ff9d',fill:false,tension:0.4}
     ]},
    base);

  // 4. Demand vs Capacity
  mkChart('chart-dc','bar',
    {labels:['Remaining Demand','Total Capacity'],
     datasets:[{
       label:'Volume',
       data:[Math.max(0,remaining),totalCap],
       backgroundColor:[remaining>totalCap?'rgba(255,59,92,0.35)':'rgba(247,201,72,0.28)','rgba(0,255,157,0.28)'],
       borderColor:[remaining>totalCap?'#ff3b5c':'#f7c948','#00ff9d'],
       borderWidth:2
     }]},
    {...base, plugins:{...base.plugins,legend:{display:false}}});
}

function mkChart(id, type, data, options) {
  const canvas = document.getElementById(id);
  if (!canvas) return;
  if (S.charts[id]) S.charts[id].destroy();
  S.charts[id] = new Chart(canvas, {type,data,options});
}

/* ───────────────────────────────────────────────────
   EMAIL / COPY
   ─────────────────────────────────────────────────── */
function sendEmail() {
  const s = document.getElementById('exec-summary-box').textContent;
  window.location.href = `mailto:?subject=${encodeURIComponent('ABIRE Forecast Analysis')}&body=${encodeURIComponent(s)}`;
}
function copySummary() {
  const s = document.getElementById('exec-summary-box').textContent;
  navigator.clipboard.writeText(s).then(() => showToast('Summary copied to clipboard'));
}

/* ───────────────────────────────────────────────────
   RESET
   ─────────────────────────────────────────────────── */
function resetAll() {
  S.weekCount = 5;
  S.startTime = '08:00'; S.endTime = '23:00'; S.intervalMins = 15;
  S.undoStack = [];
  rebuildIntervalState();
  buildTable();
  hideUndoBar();
  ['kpi-row','analysis-section','charts-section','recommendations-section'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.style.display = 'none';
  });
  const cr = document.getElementById('capacity-results');
  if (cr) cr.style.display = 'none';
  setText('exec-summary-box', 'Run analysis to generate executive summary…');
  const rp = document.getElementById('recommendations-panel');
  if (rp) rp.innerHTML = '';
  Object.values(S.charts).forEach(c => c.destroy());
  S.charts = {};
  syncConfigUI();
  setStatus('idle');
  showToast('Dashboard reset');
}

/* ───────────────────────────────────────────────────
   STATUS
   ─────────────────────────────────────────────────── */
function setStatus(state) {
  const dot  = document.getElementById('status-dot');
  const text = document.getElementById('status-text');
  if (state==='running') { dot.className='status-dot running'; text.textContent='ANALYSIS RUNNING'; }
  else if (state==='done') { dot.className='status-dot done'; text.textContent='FORECAST GENERATED'; }
  else { dot.className='status-dot idle'; text.textContent='IDLE'; }
}

/* ───────────────────────────────────────────────────
   TOAST
   ─────────────────────────────────────────────────── */
function showToast(msg, type) {
  const t = document.createElement('div');
  t.className = 'toast ' + (type||'');
  t.textContent = msg;
  document.body.appendChild(t);
  setTimeout(() => {
    t.style.animation = 'toastOut 0.3s ease forwards';
    setTimeout(() => t.remove(), 350);
  }, 2800);
}

/* ───────────────────────────────────────────────────
   FAQ ACCORDION
   ─────────────────────────────────────────────────── */
function initFAQ() {
  document.querySelectorAll('.faq-item').forEach(item => {
    item.querySelector('.faq-q').addEventListener('click', () => {
      item.classList.toggle('open');
    });
  });
}

/* ───────────────────────────────────────────────────
   UTILS
   ─────────────────────────────────────────────────── */
function fmtPct(v) { return (v>=0?'+':'') + v.toFixed(1) + '%'; }
function setText(id, val) { const el = document.getElementById(id); if(el) el.textContent = val; }

/* ───────────────────────────────────────────────────
   INIT
   ─────────────────────────────────────────────────── */
document.addEventListener('DOMContentLoaded', () => {
  // Build initial table with 15-min intervals
  rebuildIntervalState();
  buildTable();
  initFAQ();
  // Wire up config change events
  ['cfg-start','cfg-end','cfg-step'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.addEventListener('change', onConfigChange);
  });
});
