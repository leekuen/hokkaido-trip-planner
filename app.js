/* ---------- utils ---------- */
const uid = (p) => `${p}-${Math.random().toString(36).slice(2, 9)}`;
const $ = (sel, root = document) => root.querySelector(sel);
const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));
const isMapsUrl = (v) => typeof v === 'string' && /maps\.app\.goo\.gl|google\.[a-z.]+\/maps|goo\.gl\/maps/i.test(v);
const isUrl = (v) => typeof v === 'string' && /^https?:\/\//i.test(v.trim());
function extractMapsUrl(raw) {
  if (!raw) return null;
  const candidates = [raw];
  try { candidates.push(decodeURIComponent(raw)); } catch (e) { /* not encoded, ignore */ }
  for (const c of candidates) {
    const m = c.match(/https?:\/\/maps\.app\.goo\.gl\/[A-Za-z0-9_-]+/);
    if (m) return m[0];
  }
  return isMapsUrl(raw) ? raw : null;
}
const buildMapsSearch = (q) => `https://www.google.com/maps/search/${encodeURIComponent(q)}`;
const pad2 = (n) => String(n).padStart(2, '0');
const toDateKey = (d) => `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
const isDateObj = (v) => Object.prototype.toString.call(v) === '[object Date]';
const cleanArrow = (s) => (s || '').replace(/\$\\rightarrow\$/g, '→').trim();

function parseAmount(v) {
  if (v == null || v === '') return null;
  if (typeof v === 'number') return v;
  const n = parseFloat(String(v).replace(/[¥$,\s]/g, ''));
  return isNaN(n) ? null : n;
}

/* ---------- state ---------- */
const STORAGE_KEY = 'hokkaidoTripData_v1';

const DEFAULT_SHEET_URL = 'https://docs.google.com/spreadsheets/d/1IabCOL39Uzbl8ZxspX5u-GRz4WGOpJ461Ebn-K2njdQ/edit';

let state = {
  tripTitle: '北海道行程規劃',
  days: [],
  expenses: [],
  candidates: [],
  shopping: [],
  sheetUrl: DEFAULT_SHEET_URL,
  dirty: false,        // true = in-app itinerary/candidate/shopping edits exist; blocks auto-sync
  lastSyncAt: null,
};

// Shopping progress (checked-off items) is precious, so re-import always merges
// by (location, item) rather than overwriting: keep "bought" if it was already true,
// and drop items no longer present in the sheet.
function mergeShoppingList(existing, incoming) {
  const key = (it) => `${it.location}|${it.item}`;
  const existingMap = new Map(existing.map((it) => [key(it), it]));
  return incoming.map((it) => {
    const prev = existingMap.get(key(it));
    return {
      id: prev ? prev.id : it.id,
      location: it.location,
      item: it.item,
      desc: it.desc || '',
      searchTerm: it.searchTerm || '',
      bought: (prev && prev.bought) || it.bought,
    };
  });
}

function markDirty() { state.dirty = true; }

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) state = JSON.parse(raw);
  } catch (e) { console.error('load failed', e); }
}
function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

/* ---------- Excel: grid reader ---------- */
function sheetToGrid(ws) {
  if (!ws || !ws['!ref']) return [];
  const range = XLSX.utils.decode_range(ws['!ref']);
  const grid = [];
  for (let R = range.s.r; R <= range.e.r; R++) {
    const row = [];
    for (let C = range.s.c; C <= range.e.c; C++) {
      const addr = XLSX.utils.encode_cell({ r: R, c: C });
      const cell = ws[addr];
      if (!cell) { row.push({ v: '', url: null }); continue; }
      let v = cell.v;
      if (v == null) v = '';
      row.push({ v: isDateObj(v) ? v : (v === '' ? '' : String(v).trim()), url: cell.l ? cell.l.Target : null, isDate: isDateObj(v) });
    }
    grid.push(row);
  }
  return grid;
}
const cv = (row, c) => (row && row[c] ? row[c].v : '');
const cu = (row, c) => (row && row[c] ? row[c].url : null);

/* ---------- Excel: 行程表 (detailed day blocks) ---------- */
function parseItineraryDetailSheet(ws) {
  const grid = sheetToGrid(ws);
  const dayLabels = new Set(['日期', '地點', '住宿', '其他參考']);
  const days = [];
  let current = null;
  let expectDescription = false;

  for (let r = 0; r < grid.length; r++) {
    const row = grid[r];
    const a = cv(row, 0);

    if (a === '日期') {
      if (current) days.push(current);
      current = { dateKey: null, place: '', accommodation: null, items: [], lastPeriod: '' };
      for (let c = 0; c < row.length; c++) {
        const label = cv(row, c);
        if (dayLabels.has(label)) {
          const valCell = row[c + 1];
          const val = cv(row, c + 1);
          if (label === '日期' && valCell && valCell.isDate) current.dateKey = toDateKey(val);
          if (label === '地點') current.place = val;
          if (label === '住宿' && val) current.accommodation = { name: val, url: cu(row, c + 1) };
        }
      }
      expectDescription = true;
      continue;
    }

    if (!current) continue;

    if (expectDescription) {
      expectDescription = false;
      const restBlank = row.slice(1).every((cell) => !cell || cell.v === '');
      if (a && restBlank) { current.description = a; continue; }
    }

    if (a === '行程架構') continue;

    const b = cv(row, 1), c2 = cv(row, 2), d = cv(row, 3), e = cv(row, 4);
    const rowEmpty = !a && !b && !c2 && !d && !e;
    if (rowEmpty) continue;
    if (!b && !c2 && !d && !e) continue;

    if (a) current.lastPeriod = a;
    let content = [c2, d].filter(Boolean).join('\n');
    if (!content && e) content = e;
    let mapUrl = null;
    for (let c = 0; c < row.length; c++) { if (cu(row, c)) { mapUrl = cu(row, c); break; } }
    current.items.push({
      id: uid('item'),
      period: current.lastPeriod,
      time: b,
      content,
      note: (c2 && d) ? '' : (d && !content.includes(d) ? d : ''),
      mapUrl,
    });
  }
  if (current) days.push(current);
  return days;
}

/* ---------- date label formatting (行程表 is the sole source of day info) ---------- */
const WEEKDAYS = ['日', '一', '二', '三', '四', '五', '六'];
function formatDateLabel(dateKey) {
  const [y, m, d] = dateKey.split('-').map(Number);
  const dt = new Date(y, m - 1, d);
  return `${m}/${d}（${WEEKDAYS[dt.getDay()]}）`;
}
// The block's free-text description usually starts by repeating "M/D（週X）地點",
// so the highlight shown on the collapsed card is whatever comes after that first line.
function deriveHighlight(description) {
  if (!description) return '';
  const lines = description.split('\n').filter(Boolean);
  const rest = lines.slice(1).join(' ').trim();
  if (!rest) return '';
  return rest.length > 60 ? rest.slice(0, 60) + '…' : rest;
}
// Fallback when the day-header's 地點 cell is blank: the description's first
// line repeats "M/D（週X）地點" (sometimes typo'd with a backslash), so strip that prefix.
function derivePlaceFromDescription(description) {
  if (!description) return '';
  const firstLine = description.split('\n')[0];
  return firstLine.replace(/^\d{1,2}[/\\]\d{1,2}（.）/, '').trim();
}

/* ---------- Excel: 待買清單 (shopping checklist) ---------- */
// Column order isn't stable (the sheet has already been reshuffled once), so
// map by header text instead of fixed position.
function parseShoppingSheet(ws) {
  const grid = sheetToGrid(ws);
  if (!grid.length) return [];
  const header = grid[0].map((c) => (c.v || '').replace(/[✅\s]/g, ''));
  const findCol = (...keywords) => header.findIndex((h) => keywords.some((k) => h.includes(k)));
  const idx = {
    location: findCol('地點'),
    item: findCol('品項'),
    desc: findCol('說明', '推薦對象'),
    bought: findCol('已買', '購買'),
    searchTerm: findCol('搜尋', '部落客'),
  };
  const items = [];
  for (let r = 1; r < grid.length; r++) {
    const row = grid[r];
    const item = idx.item >= 0 ? cv(row, idx.item) : '';
    if (!item) continue;
    items.push({
      id: uid('shop'),
      location: idx.location >= 0 ? cv(row, idx.location) : '',
      item,
      desc: idx.desc >= 0 ? cv(row, idx.desc) : '',
      searchTerm: idx.searchTerm >= 0 ? cv(row, idx.searchTerm) : '',
      bought: idx.bought >= 0 ? /^true$/i.test(cv(row, idx.bought)) : false,
    });
  }
  return items;
}

/* ---------- Excel: candidate lists (stacked name/desc/link) ---------- */
function parseStackedCandidateSheet(ws, sourceName) {
  const grid = sheetToGrid(ws);
  const items = [];
  let category = '';
  let buffer = [];
  const flush = (urlVal) => {
    let name = '(未命名地點)', desc = '';
    if (buffer.length >= 2) { name = buffer[buffer.length - 2]; desc = buffer[buffer.length - 1]; }
    else if (buffer.length === 1) { name = buffer[0]; }
    items.push({ id: uid('cand'), name, desc, category, url: urlVal, status: 'idle', source: sourceName });
    buffer = [];
  };
  for (const row of grid) {
    const a = cv(row, 0);
    const bVal = cv(row, 1);
    const bUrl = cu(row, 1);
    if (a) category = a;
    const linkVal = extractMapsUrl(bUrl) || extractMapsUrl(bVal);
    if (linkVal) { flush(linkVal); }
    else if (bVal) { buffer.push(bVal); }
    else { buffer = []; }
  }
  return items;
}

/* ---------- Excel: reference article sheet (link, note) pairs ---------- */
function parseLinkNoteSheet(ws, sourceName) {
  const grid = sheetToGrid(ws);
  const items = [];
  for (const row of grid) {
    const a = cv(row, 0);
    const aUrl = cu(row, 0);
    const link = extractMapsUrl(aUrl) || extractMapsUrl(a);
    if (!link) continue;
    const note = cv(row, 1) || '';
    items.push({ id: uid('cand'), name: note || '(未命名地點)', desc: '', category: sourceName, url: link, status: 'idle', source: sourceName });
  }
  return items;
}

/* ---------- Excel: 住宿安排 (accommodation options table) ---------- */
function parseAccommodationSheet(ws) {
  const grid = sheetToGrid(ws);
  const items = [];
  for (let r = 1; r < grid.length; r++) {
    const row = grid[r];
    const name = cv(row, 3);
    if (!name) continue;
    items.push({
      dateRange: cv(row, 0),
      price: parseAmount(cv(row, 1)),
      perPerson: parseAmount(cv(row, 2)),
      name,
      url: cu(row, 4) || (isUrl(cv(row, 4)) ? cv(row, 4) : null),
      note: cv(row, 5),
    });
  }
  return items;
}

/* ---------- Excel: 記帳 sheet ---------- */
function parseExpenseSheet(ws) {
  const grid = sheetToGrid(ws);
  const items = [];
  for (let r = 1; r < grid.length; r++) {
    const row = grid[r];
    const name = cv(row, 3);
    if (!name) continue;
    const dateVal = row[1] && row[1].isDate ? toDateKey(cv(row, 1)) : (cv(row, 1) || '');
    items.push({
      id: uid('exp'),
      date: dateVal,
      category: cv(row, 2),
      name,
      ledger: cv(row, 4),
      payer: cv(row, 5),
      method: cv(row, 6),
      amountJPY: parseAmount(cv(row, 7)),
      amountTWD: parseAmount(cv(row, 8)),
      note: cv(row, 9),
      actualTWD: parseAmount(cv(row, 10)),
    });
  }
  return items;
}

/* ---------- combine into day list (行程表 is the sole source) ---------- */
function buildDayList(detailDays) {
  return detailDays.map((detail, idx) => ({
    id: uid('day'),
    mmdd: detail.dateKey ? `${parseInt(detail.dateKey.split('-')[1])}/${parseInt(detail.dateKey.split('-')[2])}` : String(idx),
    dateKey: detail.dateKey,
    dateLabel: detail.dateKey ? formatDateLabel(detail.dateKey) : `第 ${idx + 1} 天`,
    place: detail.place || derivePlaceFromDescription(detail.description),
    highlight: deriveHighlight(detail.description),
    transport: '',
    description: detail.description || '',
    accommodation: detail.accommodation,
    items: detail.items,
    open: false,
  }));
}

/* ---------- full workbook import ---------- */
function importWorkbook(wb) {
  const sheet = (name) => wb.Sheets[name];
  const detailSheetName = wb.SheetNames.find((n) => n.includes('行程表'));
  const foodSheetName = wb.SheetNames.find((n) => n.includes('食物清單'));
  const refSheetName = wb.SheetNames.find((n) => n.includes('參考文章'));
  const accSheetName = wb.SheetNames.find((n) => n.includes('住宿安排'));
  const expSheetName = wb.SheetNames.find((n) => n.includes('記帳'));
  const shoppingSheetName = wb.SheetNames.find((n) => n.includes('待買清單'));

  const result = { days: [], expenses: [], candidates: [], accommodationOptions: [], shopping: [] };

  if (detailSheetName) {
    const detail = parseItineraryDetailSheet(sheet(detailSheetName));
    result.days = buildDayList(detail);
  }
  if (foodSheetName) result.candidates.push(...parseStackedCandidateSheet(sheet(foodSheetName), foodSheetName));
  if (refSheetName) result.candidates.push(...parseLinkNoteSheet(sheet(refSheetName), refSheetName));
  if (accSheetName) result.accommodationOptions = parseAccommodationSheet(sheet(accSheetName));
  if (expSheetName) result.expenses = parseExpenseSheet(sheet(expSheetName));
  if (shoppingSheetName) result.shopping = parseShoppingSheet(sheet(shoppingSheetName));

  return result;
}

/* ---------- Google Maps saved-places import ---------- */
function parseTakeoutJSON(text) {
  const data = JSON.parse(text);
  const features = data.features || data.Features || [];
  return features.map((f) => {
    const props = f.properties || {};
    const loc = props.Location || props.location || {};
    const name = props.Title || loc.Name || loc.name || loc['Business Name'] || props.name || '(未命名地點)';
    const address = loc.Address || loc.address || props.Address || '';
    let url = props.google_maps_url || props['Google Maps URL'] || null;
    if (!url && f.geometry && f.geometry.coordinates) {
      const [lng, lat] = f.geometry.coordinates;
      url = `https://www.google.com/maps?q=${lat},${lng}`;
    }
    if (!url) url = buildMapsSearch(`${name} ${address}`);
    return { id: uid('cand'), name, desc: address, category: 'Google Maps 已存地點', url, status: 'idle', source: 'google_maps_import' };
  });
}

function parseKML(text) {
  const xml = new DOMParser().parseFromString(text, 'text/xml');
  const placemarks = Array.from(xml.getElementsByTagName('Placemark'));
  return placemarks.map((pm) => {
    const name = pm.getElementsByTagName('name')[0]?.textContent?.trim() || '(未命名地點)';
    const desc = pm.getElementsByTagName('description')[0]?.textContent?.trim() || '';
    const coordText = pm.getElementsByTagName('coordinates')[0]?.textContent?.trim();
    let url = null;
    if (coordText) {
      const [lng, lat] = coordText.split(',').map(parseFloat);
      if (!isNaN(lat) && !isNaN(lng)) url = `https://www.google.com/maps?q=${lat},${lng}`;
    }
    if (!url) url = buildMapsSearch(name);
    return { id: uid('cand'), name, desc, category: 'Google Maps 已存地點', url, status: 'idle', source: 'google_maps_import' };
  });
}

function parseCSVSimple(text) {
  const lines = text.split(/\r?\n/).filter((l) => l.trim());
  if (!lines.length) return [];
  const splitLine = (l) => l.split(',').map((s) => s.trim().replace(/^"|"$/g, ''));
  const header = splitLine(lines[0]).map((h) => h.toLowerCase());
  const nameIdx = header.findIndex((h) => /title|name|地點|名稱/.test(h));
  const urlIdx = header.findIndex((h) => /url|link|連結/.test(h));
  const out = [];
  for (let i = 1; i < lines.length; i++) {
    const cols = splitLine(lines[i]);
    const name = nameIdx >= 0 ? cols[nameIdx] : cols[0];
    const url = urlIdx >= 0 ? cols[urlIdx] : cols.find((c) => isUrl(c));
    if (!name && !url) continue;
    out.push({ id: uid('cand'), name: name || '(未命名地點)', desc: '', category: 'Google Maps 已存地點', url: url || buildMapsSearch(name), status: 'idle', source: 'google_maps_import' });
  }
  return out;
}

/* ---------- rendering: itinerary ---------- */
function itemMapUrl(day, item) {
  if (item.mapUrl) return item.mapUrl;
  const firstLine = (item.content || '').split('\n')[0].replace(/^\d{1,2}:\d{2}\s*/, '').slice(0, 40);
  return buildMapsSearch(`${day.place} ${firstLine}`.trim());
}

function renderItinerary() {
  const list = $('#itineraryList');
  const empty = $('#itineraryEmpty');
  if (!state.days.length) { list.innerHTML = ''; empty.hidden = false; return; }
  empty.hidden = true;
  list.innerHTML = state.days.map((day, idx) => `
    <div class="day-card ${day.open ? 'open' : ''}" data-day-id="${day.id}">
      <div class="day-header" data-action="toggle-day">
        <div>
          <div class="day-title">${escapeHtml(day.dateLabel)} ${escapeHtml(day.place)}</div>
          <div class="day-sub">${escapeHtml(day.highlight || '')}</div>
        </div>
        <div class="day-transport">${escapeHtml(day.transport || '')}</div>
      </div>
      <div class="day-body">
        ${day.description ? `<p class="hint">${escapeHtml(day.description)}</p>` : ''}
        <a class="map-btn" target="_blank" rel="noopener" href="${buildMapsSearch(day.place)}">📍 ${escapeHtml(day.place)}</a>
        ${day.accommodation ? `
          <div class="accommodation-row">
            🏨 ${escapeHtml(day.accommodation.name)}
            <a class="map-btn" target="_blank" rel="noopener" href="${day.accommodation.url || buildMapsSearch(day.accommodation.name)}">地圖</a>
          </div>` : ''}
        <div class="items">
          ${day.items.map((item, i) => `
            <div class="item-row" data-item-id="${item.id}">
              <div class="item-move">
                <button data-action="move-up" ${i === 0 ? 'disabled' : ''}>▲</button>
                <button data-action="move-down" ${i === day.items.length - 1 ? 'disabled' : ''}>▼</button>
              </div>
              <div class="item-main">
                <div class="item-period">${escapeHtml(item.period || '')} <span class="item-time">${escapeHtml(item.time || '')}</span></div>
                <div class="item-content">${escapeHtml(item.content || '')}</div>
                <div class="item-actions">
                  <a class="map-btn" target="_blank" rel="noopener" href="${itemMapUrl(day, item)}">📍地圖</a>
                  <button class="icon-btn" data-action="edit-item">✏️ 編輯</button>
                  <button class="icon-btn" data-action="delete-item">🗑️ 刪除</button>
                </div>
              </div>
            </div>
          `).join('')}
        </div>
        <button class="add-item-btn" data-action="add-item">＋ 新增行程項目</button>
        <button class="add-item-btn" data-action="edit-day">✏️ 編輯這一天（地點／住宿／重點）</button>
      </div>
    </div>
  `).join('');
}

function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

/* ---------- rendering: expenses ---------- */
function renderExpenseSummary() {
  const totals = { '公帳_TWD': 0, '私帳_TWD': 0, '公帳_JPY': 0, '私帳_JPY': 0 };
  for (const e of state.expenses) {
    const ledger = e.ledger && e.ledger.includes('私') ? '私帳' : '公帳';
    totals[`${ledger}_TWD`] += (e.actualTWD ?? e.amountTWD ?? 0) || 0;
    totals[`${ledger}_JPY`] += e.amountJPY || 0;
  }
  const fmt = (n) => Math.round(n).toLocaleString();
  $('#expenseSummary').innerHTML = `
    <div><div class="total-label">公帳 (TWD)</div><div class="total-value">$${fmt(totals['公帳_TWD'])}</div></div>
    <div><div class="total-label">私帳 (TWD)</div><div class="total-value">$${fmt(totals['私帳_TWD'])}</div></div>
    <div><div class="total-label">公帳 (¥)</div><div class="total-value">¥${fmt(totals['公帳_JPY'])}</div></div>
    <div><div class="total-label">私帳 (¥)</div><div class="total-value">¥${fmt(totals['私帳_JPY'])}</div></div>
  `;
}

function renderExpenseList() {
  const list = $('#expenseList');
  const sorted = [...state.expenses].sort((a, b) => (b.date || '').localeCompare(a.date || ''));
  list.innerHTML = sorted.map((e) => `
    <div class="expense-item" data-exp-id="${e.id}" data-action="edit-expense">
      <div class="exp-left">
        <div class="exp-title">${escapeHtml(e.name)}</div>
        <div class="exp-meta">
          <span class="exp-tag">${escapeHtml(e.ledger || '')}</span>
          ${escapeHtml(e.date || '')} · ${escapeHtml(e.category || '')} · ${escapeHtml(e.payer || '')}
        </div>
      </div>
      <div class="exp-amount">
        ${e.amountTWD != null ? `$${Math.round(e.amountTWD).toLocaleString()}` : ''}
        <small>${e.amountJPY != null ? `¥${Math.round(e.amountJPY).toLocaleString()}` : ''}</small>
      </div>
    </div>
  `).join('');
  renderExpenseSummary();
}

/* ---------- rendering: candidates ---------- */
function populateCandidateCategoryFilter() {
  const sel = $('#candidateCategoryFilter');
  const cats = Array.from(new Set(state.candidates.map((c) => c.category).filter(Boolean)));
  const current = sel.value;
  sel.innerHTML = `<option value="">全部分類</option>` + cats.map((c) => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join('');
  sel.value = current;
}

function renderCandidates() {
  populateCandidateCategoryFilter();
  const search = $('#candidateSearch').value.trim().toLowerCase();
  const cat = $('#candidateCategoryFilter').value;
  const filtered = state.candidates.filter((c) => {
    if (cat && c.category !== cat) return false;
    if (search && !(`${c.name} ${c.desc}`.toLowerCase().includes(search))) return false;
    return true;
  });
  const empty = $('#candidateEmpty');
  const list = $('#candidateList');
  if (!filtered.length) { list.innerHTML = ''; empty.hidden = state.candidates.length !== 0; return; }
  empty.hidden = true;
  list.innerHTML = filtered.map((c) => `
    <div class="candidate-card status-${c.status}" data-cand-id="${c.id}">
      <div class="cand-cat">${escapeHtml(c.category || '')}</div>
      <div class="cand-title">${escapeHtml(c.name)}</div>
      ${c.desc ? `<div class="cand-desc">${escapeHtml(c.desc)}</div>` : ''}
      <div class="cand-actions">
        <a target="_blank" rel="noopener" href="${c.url || buildMapsSearch(c.name)}">📍 地圖</a>
        <button data-action="schedule-candidate">📅 排入行程</button>
        <button data-action="edit-candidate">✏️ 編輯</button>
        ${c.status === 'idle' ? `<button data-action="skip-candidate">🙈 略過</button>` : `<button data-action="unskip-candidate">↩️ 還原</button>`}
        <button data-action="delete-candidate">🗑️</button>
      </div>
    </div>
  `).join('');
}

/* ---------- rendering: shopping list ---------- */
function renderShopping() {
  const empty = $('#shoppingEmpty');
  const list = $('#shoppingList');
  const summary = $('#shoppingSummary');
  if (!state.shopping.length) { list.innerHTML = ''; summary.textContent = ''; empty.hidden = false; return; }
  empty.hidden = true;
  const bought = state.shopping.filter((s) => s.bought).length;
  summary.textContent = `已買 ${bought} / ${state.shopping.length}`;

  const groups = new Map();
  for (const s of state.shopping) {
    const key = s.location || '未分類';
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(s);
  }
  list.innerHTML = Array.from(groups.entries()).map(([location, items]) => `
    <div class="shop-group">
      <div class="shop-group-title">📍 ${escapeHtml(location)}</div>
      ${items.map((s) => `
        <div class="shop-item-row ${s.bought ? 'bought' : ''}">
          <label class="shop-item" data-shop-id="${s.id}">
            <input type="checkbox" data-action="toggle-shop" ${s.bought ? 'checked' : ''}>
            <span class="shop-item-text">
              <span class="shop-item-name">${escapeHtml(s.item)}</span>
              ${s.desc ? `<span class="shop-item-desc">${escapeHtml(s.desc)}</span>` : ''}
            </span>
          </label>
          ${s.searchTerm ? `<a class="shop-search-btn" target="_blank" rel="noopener" href="https://www.google.com/search?q=${encodeURIComponent(s.searchTerm)}" title="搜尋「${escapeHtml(s.searchTerm)}」">🔍</a>` : ''}
        </div>
      `).join('')}
    </div>
  `).join('');
}

/* ---------- modal helper ---------- */
function openModal(html, onMount) {
  const root = $('#modalRoot');
  root.innerHTML = `<div class="modal-backdrop" id="modalBackdrop"><div class="modal-sheet">${html}</div></div>`;
  $('#modalBackdrop').addEventListener('click', (e) => { if (e.target.id === 'modalBackdrop') closeModal(); });
  if (onMount) onMount(root);
}
function closeModal() { $('#modalRoot').innerHTML = ''; }

/* ---------- item modal ---------- */
function openItemModal(dayId, itemId) {
  const day = state.days.find((d) => d.id === dayId);
  const item = itemId ? day.items.find((i) => i.id === itemId) : null;
  openModal(`
    <h3>${item ? '編輯' : '新增'}行程項目</h3>
    <div class="form-row"><label>時段（如：早餐/上午）</label><input id="f-period" value="${escapeHtml(item?.period || '')}"></div>
    <div class="form-row"><label>時間</label><input id="f-time" value="${escapeHtml(item?.time || '')}"></div>
    <div class="form-row"><label>內容</label><textarea id="f-content">${escapeHtml(item?.content || '')}</textarea></div>
    <div class="form-row"><label>Google Maps 連結（選填）</label><input id="f-mapurl" value="${escapeHtml(item?.mapUrl || '')}"></div>
    <div class="modal-actions">
      ${item ? `<button class="delete-btn" id="btnDelete">刪除</button>` : ''}
      <button class="cancel-btn" id="btnCancel">取消</button>
      <button class="save-btn" id="btnSave">儲存</button>
    </div>
  `, () => {
    $('#btnCancel').onclick = closeModal;
    $('#btnSave').onclick = () => {
      const data = {
        period: $('#f-period').value.trim(),
        time: $('#f-time').value.trim(),
        content: $('#f-content').value.trim(),
        mapUrl: $('#f-mapurl').value.trim() || null,
      };
      if (item) Object.assign(item, data);
      else day.items.push({ id: uid('item'), ...data });
      markDirty(); saveState(); renderItinerary(); closeModal();
    };
    if (item) $('#btnDelete').onclick = () => {
      day.items = day.items.filter((i) => i.id !== itemId);
      markDirty(); saveState(); renderItinerary(); closeModal();
    };
  });
}

function openDayEditModal(dayId) {
  const day = state.days.find((d) => d.id === dayId);
  openModal(`
    <h3>編輯 ${escapeHtml(day.dateLabel)}</h3>
    <div class="form-row"><label>地點</label><input id="f-place" value="${escapeHtml(day.place)}"></div>
    <div class="form-row"><label>行程重點</label><input id="f-highlight" value="${escapeHtml(day.highlight || '')}"></div>
    <div class="form-row"><label>交通工具</label><input id="f-transport" value="${escapeHtml(day.transport || '')}"></div>
    <div class="form-row"><label>住宿名稱</label><input id="f-acc-name" value="${escapeHtml(day.accommodation?.name || '')}"></div>
    <div class="form-row"><label>住宿地圖連結</label><input id="f-acc-url" value="${escapeHtml(day.accommodation?.url || '')}"></div>
    <div class="modal-actions">
      <button class="cancel-btn" id="btnCancel">取消</button>
      <button class="save-btn" id="btnSave">儲存</button>
    </div>
  `, () => {
    $('#btnCancel').onclick = closeModal;
    $('#btnSave').onclick = () => {
      day.place = $('#f-place').value.trim();
      day.highlight = $('#f-highlight').value.trim();
      day.transport = $('#f-transport').value.trim();
      const accName = $('#f-acc-name').value.trim();
      day.accommodation = accName ? { name: accName, url: $('#f-acc-url').value.trim() || null } : null;
      markDirty(); saveState(); renderItinerary(); closeModal();
    };
  });
}

/* ---------- expense modal ---------- */
const CUSTOM_OPT = '__custom__';

function expenseFieldOptions(field, seeds) {
  const seen = new Set();
  const opts = [];
  for (const v of [...seeds, ...state.expenses.map((e) => e[field])]) {
    const val = (v || '').trim();
    if (val && !seen.has(val)) { seen.add(val); opts.push(val); }
  }
  return opts;
}

// A <select> of known values plus a 自訂 option that reveals a free-text input.
function selectWithCustom(id, options, currentVal) {
  const current = (currentVal || '').trim();
  const list = current && !options.includes(current) ? [current, ...options] : options;
  const selected = current || list[0] || '';
  return `
    <select id="${id}">
      ${list.map((o) => `<option value="${escapeHtml(o)}" ${o === selected ? 'selected' : ''}>${escapeHtml(o)}</option>`).join('')}
      <option value="${CUSTOM_OPT}">＋ 自訂...</option>
    </select>
    <input id="${id}-custom" placeholder="輸入新的選項" style="display:none; margin-top:6px;">
  `;
}

function wireSelectWithCustom(id) {
  const sel = $(`#${id}`);
  const custom = $(`#${id}-custom`);
  sel.addEventListener('change', () => {
    custom.style.display = sel.value === CUSTOM_OPT ? 'block' : 'none';
    if (sel.value === CUSTOM_OPT) custom.focus();
  });
}

function readSelectWithCustom(id) {
  const sel = $(`#${id}`);
  return sel.value === CUSTOM_OPT ? $(`#${id}-custom`).value.trim() : sel.value;
}

function openExpenseModal(expId) {
  const exp = expId ? state.expenses.find((e) => e.id === expId) : null;
  const categoryOpts = expenseFieldOptions('category', ['🚌 交通', '🏨 住宿', '🍜 飲食', '🎟️ 門票體驗', '🛍️ 購物', '📦 其他']);
  const payerOpts = expenseFieldOptions('payer', ['婍瑞', '力堃']);
  const methodOpts = expenseFieldOptions('method', ['CUBE', '現金']);
  const today = toDateKey(new Date());
  openModal(`
    <h3>${exp ? '編輯' : '新增'}記帳</h3>
    <div class="form-row"><label>日期</label><input id="f-date" type="date" value="${exp?.date || today}"></div>
    <div class="form-row"><label>大項目</label>${selectWithCustom('f-category', categoryOpts, exp?.category)}</div>
    <div class="form-row"><label>項目名稱</label><input id="f-name" value="${escapeHtml(exp?.name || '')}"></div>
    <div class="form-row"><label>公／私帳</label>
      <select id="f-ledger"><option ${exp?.ledger === '公帳' ? 'selected' : ''}>公帳</option><option ${exp?.ledger === '私帳' ? 'selected' : ''}>私帳</option></select>
    </div>
    <div class="form-row"><label>支付人</label>${selectWithCustom('f-payer', payerOpts, exp?.payer)}</div>
    <div class="form-row"><label>支付方式</label>${selectWithCustom('f-method', methodOpts, exp?.method)}</div>
    <div class="form-row"><label>金額（日幣）</label><input id="f-jpy" type="number" value="${exp?.amountJPY ?? ''}"></div>
    <div class="form-row"><label>金額（台幣）</label><input id="f-twd" type="number" value="${exp?.amountTWD ?? ''}"></div>
    <div class="form-row"><label>備註</label><input id="f-note" value="${escapeHtml(exp?.note || '')}"></div>
    <div class="modal-actions">
      ${exp ? `<button class="delete-btn" id="btnDelete">刪除</button>` : ''}
      <button class="cancel-btn" id="btnCancel">取消</button>
      <button class="save-btn" id="btnSave">儲存</button>
    </div>
  `, () => {
    wireSelectWithCustom('f-category');
    wireSelectWithCustom('f-payer');
    wireSelectWithCustom('f-method');
    $('#btnCancel').onclick = closeModal;
    $('#btnSave').onclick = () => {
      const data = {
        date: $('#f-date').value,
        category: readSelectWithCustom('f-category'),
        name: $('#f-name').value.trim(),
        ledger: $('#f-ledger').value,
        payer: readSelectWithCustom('f-payer'),
        method: readSelectWithCustom('f-method'),
        amountJPY: parseAmount($('#f-jpy').value),
        amountTWD: parseAmount($('#f-twd').value),
        note: $('#f-note').value.trim(),
      };
      if (!data.name) return;
      if (exp) Object.assign(exp, data);
      else state.expenses.push({ id: uid('exp'), actualTWD: null, ...data });
      saveState(); renderExpenseList(); closeModal();
    };
    if (exp) $('#btnDelete').onclick = () => {
      state.expenses = state.expenses.filter((e) => e.id !== expId);
      saveState(); renderExpenseList(); closeModal();
    };
  });
}

/* ---------- candidate modal ---------- */
function openCandidateModal(candId) {
  const cand = candId ? state.candidates.find((c) => c.id === candId) : null;
  openModal(`
    <h3>${cand ? '編輯' : '新增'}候選地點</h3>
    <div class="form-row"><label>名稱</label><input id="f-name" value="${escapeHtml(cand?.name || '')}"></div>
    <div class="form-row"><label>分類</label><input id="f-category" value="${escapeHtml(cand?.category || '')}"></div>
    <div class="form-row"><label>說明</label><textarea id="f-desc">${escapeHtml(cand?.desc || '')}</textarea></div>
    <div class="form-row"><label>Google Maps 連結</label><input id="f-url" value="${escapeHtml(cand?.url || '')}"></div>
    <div class="modal-actions">
      ${cand ? `<button class="delete-btn" id="btnDelete">刪除</button>` : ''}
      <button class="cancel-btn" id="btnCancel">取消</button>
      <button class="save-btn" id="btnSave">儲存</button>
    </div>
  `, () => {
    $('#btnCancel').onclick = closeModal;
    $('#btnSave').onclick = () => {
      const name = $('#f-name').value.trim();
      if (!name) return;
      const data = { name, category: $('#f-category').value.trim(), desc: $('#f-desc').value.trim(), url: $('#f-url').value.trim() };
      if (cand) Object.assign(cand, data);
      else state.candidates.push({ id: uid('cand'), status: 'idle', source: 'manual', ...data });
      markDirty(); saveState(); renderCandidates(); closeModal();
    };
    if (cand) $('#btnDelete').onclick = () => {
      state.candidates = state.candidates.filter((c) => c.id !== candId);
      markDirty(); saveState(); renderCandidates(); closeModal();
    };
  });
}

function openScheduleModal(candId) {
  const cand = state.candidates.find((c) => c.id === candId);
  if (!state.days.length) { alert('請先匯入或建立行程天數'); return; }
  openModal(`
    <h3>把「${escapeHtml(cand.name)}」排入行程</h3>
    <div class="form-row"><label>選擇日期</label>
      <select id="f-day">${state.days.map((d) => `<option value="${d.id}">${escapeHtml(d.dateLabel)} ${escapeHtml(d.place)}</option>`).join('')}</select>
    </div>
    <div class="form-row"><label>時段（選填）</label><input id="f-period" placeholder="如：午餐"></div>
    <div class="modal-actions">
      <button class="cancel-btn" id="btnCancel">取消</button>
      <button class="save-btn" id="btnSave">加入</button>
    </div>
  `, () => {
    $('#btnCancel').onclick = closeModal;
    $('#btnSave').onclick = () => {
      const day = state.days.find((d) => d.id === $('#f-day').value);
      day.items.push({ id: uid('item'), period: $('#f-period').value.trim(), time: '', content: cand.name + (cand.desc ? `\n${cand.desc}` : ''), mapUrl: cand.url, fromCandidate: candId });
      cand.status = 'scheduled';
      markDirty(); saveState(); renderItinerary(); renderCandidates(); closeModal();
    };
  });
}

/* ---------- shared import (Google Sheet sync & Excel upload) ---------- */
function applyImport(wb, resultEl) {
  const parsed = importWorkbook(wb);
  const overwriteItin = $('#overwriteItinerary').checked;
  const overwriteExp = $('#overwriteExpense').checked;

  if (overwriteItin) {
    if (state.days.length && !confirm('這會覆蓋目前工具內的行程與候選清單，確定要繼續嗎？')) { resultEl.textContent = '已取消。'; return; }
    state.days = parsed.days;
    state.candidates = parsed.candidates;
  }
  state.shopping = mergeShoppingList(state.shopping, parsed.shopping);
  if (overwriteExp) {
    if (state.expenses.length && !confirm('這會覆蓋目前工具內的所有記帳資料，確定要繼續嗎？')) { resultEl.textContent = '已取消。'; return; }
    state.expenses = parsed.expenses;
  } else {
    const existingKeys = new Set(state.expenses.map((x) => `${x.date}|${x.name}`));
    for (const exp of parsed.expenses) {
      if (!existingKeys.has(`${exp.date}|${exp.name}`)) state.expenses.push(exp);
    }
  }
  if (overwriteItin) state.dirty = false;
  state.lastSyncAt = new Date().toISOString();
  saveState();
  renderItinerary(); renderExpenseList(); renderCandidates(); renderShopping();
  resultEl.textContent = `匯入完成：${parsed.days.length} 天行程、${parsed.candidates.length} 個候選地點、${parsed.expenses.length} 筆記帳、${parsed.shopping.length} 項待買。`;
}

/* ---------- auto-sync on open ---------- */
function showToast(msg, ms = 3500) {
  let el = $('#toast');
  if (!el) {
    el = document.createElement('div');
    el.id = 'toast';
    document.body.appendChild(el);
  }
  el.textContent = msg;
  el.classList.add('show');
  clearTimeout(el._t);
  el._t = setTimeout(() => el.classList.remove('show'), ms);
}

async function fetchSheetWorkbook(sheetUrl) {
  const idMatch = (sheetUrl || '').match(/\/spreadsheets\/d\/([A-Za-z0-9_-]+)/);
  if (!idMatch) throw new Error('試算表網址格式不對');
  const resp = await fetch(`https://docs.google.com/spreadsheets/d/${idMatch[1]}/export?format=xlsx`);
  if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
  const buf = new Uint8Array(await resp.arrayBuffer());
  return XLSX.read(buf, { type: 'array', cellDates: true });
}

async function autoSyncOnOpen() {
  if (!state.sheetUrl) return;
  if (state.dirty) {
    showToast('⚠️ 你在工具內調整過行程，已暫停自動同步。要改用試算表版本請到「設定」按立即同步。', 6000);
    return;
  }
  if (typeof navigator !== 'undefined' && navigator.onLine === false) return;
  showToast('🔄 正在同步試算表...');
  try {
    const wb = await fetchSheetWorkbook(state.sheetUrl);
    const parsed = importWorkbook(wb);
    state.days = parsed.days;
    state.candidates = parsed.candidates;
    state.shopping = mergeShoppingList(state.shopping, parsed.shopping);
    const existingKeys = new Set(state.expenses.map((x) => `${x.date}|${x.name}`));
    for (const exp of parsed.expenses) {
      if (!existingKeys.has(`${exp.date}|${exp.name}`)) state.expenses.push(exp);
    }
    state.dirty = false;
    state.lastSyncAt = new Date().toISOString();
    saveState();
    renderItinerary(); renderExpenseList(); renderCandidates(); renderShopping();
    showToast(`✅ 已同步：${parsed.days.length} 天行程、${parsed.candidates.length} 個候選地點`);
  } catch (err) {
    showToast('同步失敗（' + err.message + '），顯示上次的資料。', 5000);
  }
}

/* ---------- tab switching ---------- */
function switchTab(tabId) {
  $$('.tab-panel').forEach((p) => p.classList.toggle('active', p.id === tabId));
  $$('.nav-btn').forEach((b) => b.classList.toggle('active', b.dataset.tab === tabId));
}

/* ---------- event wiring ---------- */
function wireEvents() {
  $$('.nav-btn').forEach((btn) => btn.addEventListener('click', () => switchTab(btn.dataset.tab)));

  $('#itineraryList').addEventListener('click', (e) => {
    const dayEl = e.target.closest('.day-card');
    if (!dayEl) return;
    const dayId = dayEl.dataset.dayId;
    const day = state.days.find((d) => d.id === dayId);
    const action = e.target.closest('[data-action]')?.dataset.action;

    if (e.target.closest('.day-header') && (!action || action === 'toggle-day')) {
      day.open = !day.open; saveState(); renderItinerary();
      const el = $(`.day-card[data-day-id="${dayId}"]`);
      if (el && el.scrollIntoView) el.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
      return;
    }
    if (!action) return;
    const itemRow = e.target.closest('.item-row');
    const itemId = itemRow?.dataset.itemId;

    if (action === 'add-item') openItemModal(dayId, null);
    if (action === 'edit-item') openItemModal(dayId, itemId);
    if (action === 'delete-item') { day.items = day.items.filter((i) => i.id !== itemId); markDirty(); saveState(); renderItinerary(); }
    if (action === 'edit-day') openDayEditModal(dayId);
    if (action === 'move-up' || action === 'move-down') {
      const idx = day.items.findIndex((i) => i.id === itemId);
      const swapWith = action === 'move-up' ? idx - 1 : idx + 1;
      if (swapWith < 0 || swapWith >= day.items.length) return;
      [day.items[idx], day.items[swapWith]] = [day.items[swapWith], day.items[idx]];
      markDirty(); saveState(); renderItinerary();
    }
  });

  $('#addExpenseBtn').addEventListener('click', () => openExpenseModal(null));
  $('#expenseList').addEventListener('click', (e) => {
    const row = e.target.closest('[data-exp-id]');
    if (row) openExpenseModal(row.dataset.expId);
  });

  $('#addCandidateBtn').addEventListener('click', () => openCandidateModal(null));
  $('#candidateSearch').addEventListener('input', renderCandidates);
  $('#candidateCategoryFilter').addEventListener('change', renderCandidates);
  $('#candidateList').addEventListener('click', (e) => {
    const card = e.target.closest('.candidate-card');
    if (!card) return;
    const id = card.dataset.candId;
    const action = e.target.closest('[data-action]')?.dataset.action;
    if (!action) return;
    const cand = state.candidates.find((c) => c.id === id);
    if (action === 'edit-candidate') openCandidateModal(id);
    if (action === 'schedule-candidate') openScheduleModal(id);
    if (action === 'skip-candidate') { cand.status = 'skip'; markDirty(); saveState(); renderCandidates(); }
    if (action === 'unskip-candidate') { cand.status = 'idle'; markDirty(); saveState(); renderCandidates(); }
    if (action === 'delete-candidate') { state.candidates = state.candidates.filter((c) => c.id !== id); markDirty(); saveState(); renderCandidates(); }
  });

  $('#shoppingList').addEventListener('change', (e) => {
    if (e.target.dataset.action !== 'toggle-shop') return;
    const row = e.target.closest('[data-shop-id]');
    const item = state.shopping.find((s) => s.id === row.dataset.shopId);
    item.bought = e.target.checked;
    markDirty(); saveState(); renderShopping();
  });

  $('#syncSheetBtn').addEventListener('click', async () => {
    const resultEl = $('#sheetSyncResult');
    const url = $('#sheetUrlInput').value.trim();
    if (!/\/spreadsheets\/d\/[A-Za-z0-9_-]+/.test(url)) { resultEl.textContent = '網址格式不對，請貼上 Google 試算表的完整連結。'; return; }
    state.sheetUrl = url;
    resultEl.textContent = '下載中...';
    try {
      const wb = await fetchSheetWorkbook(url);
      applyImport(wb, resultEl);
    } catch (err) {
      resultEl.textContent = '同步失敗：' + err.message + '（請確認試算表已設為「知道連結的使用者可檢視」）';
    }
  });

  $('#excelFileInput').addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const buf = new Uint8Array(await file.arrayBuffer());
    const wb = XLSX.read(buf, { type: 'array', cellDates: true });
    applyImport(wb, $('#excelImportResult'));
  });

  $('#mapsFileInput').addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const text = await file.text();
    let items = [];
    try {
      if (file.name.endsWith('.json')) items = parseTakeoutJSON(text);
      else if (file.name.endsWith('.kml')) items = parseKML(text);
      else if (file.name.endsWith('.csv')) items = parseCSVSimple(text);
    } catch (err) {
      $('#mapsImportResult').textContent = '解析失敗，請確認檔案格式。' + err.message;
      return;
    }
    const existingUrls = new Set(state.candidates.map((c) => c.url));
    let added = 0;
    for (const it of items) { if (!existingUrls.has(it.url)) { state.candidates.push(it); added++; } }
    saveState(); renderCandidates();
    $('#mapsImportResult').textContent = `匯入完成：新增 ${added} 個地點（共讀到 ${items.length} 筆，略過重複）。`;
  });

  $('#exportBackupBtn').addEventListener('click', () => {
    const blob = new Blob([JSON.stringify(state, null, 2)], { type: 'application/json' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `trip-backup-${new Date().toISOString().slice(0, 10)}.json`;
    a.click();
  });
  $('#restoreBackupBtn').addEventListener('click', () => $('#restoreBackupInput').click());
  $('#restoreBackupInput').addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    if (!confirm('這會覆蓋目前所有資料，確定要從備份還原嗎？')) return;
    state = JSON.parse(await file.text());
    saveState(); renderAll();
  });
  $('#clearAllBtn').addEventListener('click', () => {
    if (!confirm('確定要清除本機所有資料嗎？此動作無法復原（建議先匯出備份）。')) return;
    state = { tripTitle: '北海道行程規劃', days: [], expenses: [], candidates: [] };
    saveState(); renderAll();
  });
}

function renderAll() {
  $('#tripTitle').textContent = state.tripTitle || '行程規劃';
  $('#sheetUrlInput').value = state.sheetUrl || DEFAULT_SHEET_URL;
  renderItinerary();
  renderExpenseList();
  renderCandidates();
  renderShopping();
}

loadState();
if (!Array.isArray(state.shopping)) state.shopping = []; // migrate pre-shopping-list saves
wireEvents();
renderAll();
autoSyncOnOpen();
