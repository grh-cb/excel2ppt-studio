/* Excel2PPT Studio 前端：四步向导（无框架，直接操作 DOM） */
'use strict';

const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => Array.from(document.querySelectorAll(sel));
const NS = 'http://www.w3.org/2000/svg';
const STORE_KEY = 'e2p-studio-v1';

const COLORS = { text: '#3b82f6', table: '#10b981', image: '#f59e0b', chart: '#8b5cf6' };
const KIND_NAMES = { text: '文本', table: '表格', image: '图片', chart: '图表', group: '组合' };

const S = {
  template: '', excel: '', slide: 0, sheet: '', headerRow: 1,
  width: 0, height: 0, fields: [], columns: [], rows: [], total: 0,
  bindings: [], selected: null, built: false, lastOutput: '',
  pageIndex: 0, pageTotal: 1,
  titles: null, titleKey: '序号', titleColumn: '',
  slideCount: 1, slidesData: [],
  cover: { enabled: false, slide: 0, title: '', subtitle: '' },
  end: { enabled: false, slide: 0, text: '' },
};

/* ── 通用 ─────────────────────────────────────────────── */
function setMsg(el, text, kind) {
  if (!el) return;
  el.textContent = text || '';
  el.className = 'msg' + (kind ? ' ' + kind : '');
}

async function api(path, body, allowBusinessFail) {
  const res = await fetch(path, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  let data = {};
  try { data = await res.json(); } catch (_) { /* 非 JSON 响应 */ }
  // 有些接口用 ok:false 表达业务结论（如校验不通过），不应视为请求失败
  if (!res.ok || (data.ok === false && !allowBusinessFail)) {
    throw new Error(data.detail || data.error || ('请求失败：HTTP ' + res.status));
  }
  return data;
}

function gotoStep(n) {
  $$('.panel').forEach((p) => { p.hidden = p.id !== 'step' + n; });
  $$('.step').forEach((b) => b.classList.toggle('active', Number(b.dataset.step) === n));
  hidePopover();
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

function persist() {
  const data = {
    template: S.template, excel: S.excel, slide: S.slide, sheet: S.sheet,
    headerRow: S.headerRow, bindings: S.bindings,
    cover: S.cover, end: S.end,
  };
  try { localStorage.setItem(STORE_KEY, JSON.stringify(data)); } catch (_) { /* 隐私模式忽略 */ }
}

function restore() {
  let data = null;
  try { data = JSON.parse(localStorage.getItem(STORE_KEY) || 'null'); } catch (_) { return; }
  if (!data) return;
  $('#templatePath').value = data.template || '';
  $('#excelPath').value = data.excel || '';
  $('#slideIndex').value = data.slide ?? 0;
  $('#sheet').value = data.sheet || '';
  $('#headerRow').value = data.headerRow ?? 1;
  S.bindings = Array.isArray(data.bindings) ? data.bindings : [];
  if (data.cover) Object.assign(S.cover, data.cover);
  if (data.end) Object.assign(S.end, data.end);
  const withMap = S.bindings.find((b) => b.source && b.source.value_map);
  if (withMap) {
    S.titles = withMap.source.value_map;
    S.titleKey = withMap.source.key_column || '序号';
    S.titleColumn = withMap.source.column || '';
  }
}

/* ── 第 1 步：读取 ─────────────────────────────────────── */
async function upload(kind, file) {
  const fd = new FormData();
  fd.append('file', file);
  const res = await fetch('/api/upload?kind=' + encodeURIComponent(kind), { method: 'POST', body: fd });
  const data = await res.json().catch(() => ({}));
  if (!res.ok || data.ok === false) {
    throw new Error(data.detail || data.error || '上传失败');
  }
  return data.path;
}

async function doScan() {
  const template = $('#templatePath').value.trim();
  const excel = $('#excelPath').value.trim();
  const slide = Number($('#slideIndex').value || 0);
  const sheet = $('#sheet').value.trim();
  const headerRow = Number($('#headerRow').value || 1);
  const msg = $('#scanMsg');
  if (!template || !excel) { setMsg(msg, '请先填写模板与 Excel 路径。', 'error'); return; }

  const btn = $('#btnScan');
  btn.disabled = true;
  setMsg(msg, '正在读取模板与 Excel…');
  try {
    const scan = await api('/api/scan', { template, slide });
    const book = await api('/api/excel', { path: excel, sheet: sheet || null, header_row: headerRow });
    // 智能识别正文页：用户没手动改索引（仍为 0）时采用推荐页
    let usedSlide = slide;
    let autoNote = '';
    if (slide === 0 && typeof scan.auto_slide === 'number' && scan.auto_slide !== 0) {
      usedSlide = scan.auto_slide;
      $('#slideIndex').value = String(usedSlide);
      const detail = await api('/api/scan', { template, slide: usedSlide });
      Object.assign(scan, detail);
      autoNote = `（已自动识别正文页：第 ${usedSlide + 1} 页）`;
    }
    Object.assign(S, {
      template: scan.template, excel: book.path, slide: usedSlide, sheet: book.sheet, headerRow,
      width: scan.width, height: scan.height, fields: scan.fields || [],
      slideCount: scan.slide_count || 1,
      columns: book.columns || [], rows: book.rows || [], total: book.total || 0,
      selected: null, built: false, pageIndex: 0, pageTotal: 1,
    });
    setMsg(msg, `已读取 ${S.fields.length} 个元素、${S.total} 行数据。${autoNote}`, 'ok');
    await runDoctor(template, usedSlide, excel);
    try {
      const slideInfo = await api('/api/slides', { template });
      S.slidesData = slideInfo.slides || [];
    } catch (_) {
      S.slidesData = [];
    }
    renderLayoutPanel();
    renderFieldList();
    renderCanvas();
    renderDetail();
    renderDataPreview();
    await autoMap(true);
    persist();
    gotoStep(2);
  } catch (err) {
    setMsg(msg, err.message, 'error');
  } finally {
    btn.disabled = false;
  }
}

/* ── 模板体检 ─────────────────────────────────────────── */
async function runDoctor(template, slide, excel) {
  const card = $('#healthCard');
  if (!card) return;
  try {
    const report = await api('/api/doctor', { template, slide, excel: excel || null });
    const level = report.score >= 80 ? 'ok' : report.score >= 60 ? 'warn' : 'bad';
    card.hidden = false;
    const score = $('#healthScore');
    score.textContent = `${report.score} 分`;
    score.className = `health-score ${level}`;
    const bar = $('#healthBar').querySelector('i');
    bar.style.width = `${report.score}%`;
    bar.className = level;
    $('#healthIssues').innerHTML = (report.issues || [])
      .map((i) => {
        const tag = { error: '错误', warning: '警告', info: '信息' }[i.level] || i.level;
        const ref = i.field_id ? ` <code>${escapeHtml(i.field_id)}</code>` : '';
        return `<li class="lv-${escapeAttr(i.level)}"><b>${tag}</b> ${escapeHtml(i.message)}${ref}</li>`;
      })
      .join('');
  } catch (err) {
    card.hidden = true;
  }
}

/* ── 第 2 步：元素、画布、映射 ─────────────────────────── */
function fieldPreview(f) {
  if (f.kind === 'table' && f.table) return `表格 ${f.table.rows}×${f.table.cols}`;
  if (f.kind === 'image') return '图片';
  if (f.kind === 'chart') return '图表';
  return (f.text || '').replace(/\s*\n\s*/g, ' / ') || '空文本框';
}

function posLabel(rect) {
  if (!S.width || !S.height || !rect) return '';
  const cx = (rect.x + rect.w / 2) / S.width;
  const cy = (rect.y + rect.h / 2) / S.height;
  const h = cx < 0.34 ? '左' : cx > 0.66 ? '右' : '中';
  const v = cy < 0.34 ? '上' : cy > 0.66 ? '下' : '中';
  return `${v}${h}`;
}

function bindingForField(fieldId, kind, slotName) {
  return S.bindings.find((b) => {
    if (b.field_id !== fieldId) return false;
    if (kind && b.kind !== kind) return false;
    if (b.kind === 'text' && slotName !== undefined) {
      return (b.source.slot || null) === (slotName || null);
    }
    return true;
  });
}

function bindingLabel(b) {
  if (!b) return '';
  if (b.kind === 'table') return '表格';
  if (b.kind === 'chart') return '图表';
  if (b.kind === 'image') return b.source.column || '图片';
  return b.source.column || '';
}

function renderFieldList() {
  const box = $('#fieldList');
  box.innerHTML = '';
  $('#fieldCount').textContent = String(S.fields.length);
  S.fields.forEach((f) => {
    const mapped = S.bindings.some((b) => b.field_id === f.id);
    const item = document.createElement('div');
    item.className = 'field-item' + (f.id === S.selected ? ' active' : '') + (mapped ? ' mapped' : '');
    item.dataset.id = f.id;
    item.innerHTML = `
      <div class="row1">
        <span class="tag ${f.kind}">${KIND_NAMES[f.kind] || f.kind}</span>
        <span class="name">${escapeHtml(f.name || f.id)}</span>
        <span class="pos">${posLabel(f.rect)}</span>
        ${mapped ? '<span class="tick">✓</span>' : ''}
      </div>
      <div class="preview">${escapeHtml(fieldPreview(f))}</div>`;
    item.addEventListener('click', () => select(f.id));
    box.appendChild(item);
  });
}

function renderCanvas() {
  const svg = $('#canvas');
  svg.innerHTML = '';
  if (!S.width || !S.height) return;
  // 用 1280 宽的像素坐标系承载：EMU 直接缩放，文字按像素渲染才清晰
  const VW = 1280;
  const scale = VW / S.width;
  const VH = Math.round(S.height * scale);
  svg.setAttribute('viewBox', `0 0 ${VW} ${VH}`);
  svg.style.aspectRatio = `${VW} / ${VH}`;
  const font = 13;
  const stroke = 1.5;

  S.fields.forEach((f) => {
    const color = COLORS[f.kind] || '#94a3b8';
    const binding = S.bindings.find((b) => b.field_id === f.id);
    const empty = f.kind === 'text' && !(f.text || '').trim() && !(f.slots || []).length;
    const rect = document.createElementNS(NS, 'rect');
    rect.setAttribute('x', f.rect.x * scale);
    rect.setAttribute('y', f.rect.y * scale);
    rect.setAttribute('width', f.rect.w * scale);
    rect.setAttribute('height', f.rect.h * scale);
    rect.setAttribute('fill', binding ? color + '2b' : color + '12');
    rect.setAttribute('stroke', empty ? '#cbd5e1' : color);
    rect.setAttribute('stroke-width', binding ? stroke + 0.6 : stroke);
    if (!binding) rect.setAttribute('stroke-dasharray', '5 4');
    rect.setAttribute('rx', 3);
    rect.setAttribute('class', 'box' + (f.id === S.selected ? ' selected' : ''));
    rect.addEventListener('click', (e) => {
      e.stopPropagation();
      select(f.id);
      showPopover(f, e.clientX, e.clientY);
    });
    svg.appendChild(rect);

    // 空文本框不写标签，避免把画布糊满；槽位名比形状名更有信息量
    if (empty) return;
    const label = document.createElementNS(NS, 'text');
    label.setAttribute('x', f.rect.x * scale + 6);
    label.setAttribute('y', f.rect.y * scale + font + 3);
    label.setAttribute('font-size', font);
    label.setAttribute('fill', color);
    label.setAttribute('class', 'label');
    const slotName = (f.slots || []).map((s) => s.name).filter(Boolean).slice(0, 2).join('/');
    const suffix = binding ? ` → ${bindingLabel(binding)}` : '';
    label.textContent = (f.id + ' ' + (slotName || fieldPreview(f)) + suffix).slice(0, 32);
    svg.appendChild(label);
  });
}

function select(id) {
  S.selected = id;
  renderFieldList();
  renderCanvas();
  renderDetail();
  const el = document.querySelector(`.field-item[data-id="${id}"]`);
  if (el) el.scrollIntoView({ block: 'nearest' });
}

function renderDetail() {
  const box = $('#detailBody');
  const f = S.fields.find((x) => x.id === S.selected);
  if (!f) {
    box.innerHTML = '<p class="muted">尚未选中元素。</p>';
    return;
  }
  const cm = f.rect.cm || [0, 0, 0, 0];
  const slots = (f.slots || []).map((s) => s.name).filter(Boolean);
  box.innerHTML = `
    <div class="kv"><span>类型</span><span>${KIND_NAMES[f.kind] || f.kind}</span></div>
    <div class="kv"><span>名称</span><span>${escapeHtml(f.name || f.id)}</span></div>
    <div class="kv"><span>位置 cm</span><span>${cm[0]}, ${cm[1]}</span></div>
    <div class="kv"><span>尺寸 cm</span><span>${cm[2]} × ${cm[3]}</span></div>
    <div class="kv"><span>槽位</span><span>${slots.length ? escapeHtml(slots.join('、')) : '—'}</span></div>
    <div class="kv" style="display:block"><span>文本</span><div class="preview" style="white-space:normal">${escapeHtml((f.text || '').slice(0, 160) || '—')}</div></div>`;
  const mapTitle = document.createElement('h3');
  mapTitle.className = 'detail-map-title';
  mapTitle.textContent = '映射到 Excel 列';
  box.appendChild(mapTitle);
  box.appendChild(buildFieldMapping(f));
}

function renderDataPreview() {
  const box = $('#dataTable');
  $('#dataMeta').textContent = `${S.sheet || '默认表'} · 共 ${S.total} 行 · 预览前 ${Math.min(S.rows.length, 10)} 行`;
  if (!S.columns.length) {
    box.innerHTML = '<p class="muted" style="padding:10px">没有数据。</p>';
    return;
  }
  const head = S.columns.map((c) =>
    `<th>${escapeHtml(c.name)}<span class="muted"> · ${c.dtype}</span></th>`
  ).join('');
  const body = S.rows.slice(0, 10).map((row) => '<tr>' + S.columns.map((c) => {
    const v = row[c.name];
    if (v === null || v === undefined || v === '') return '<td class="empty">—</td>';
    return `<td>${escapeHtml(String(v))}</td>`;
  }).join('') + '</tr>').join('');
  box.innerHTML = `<table class="data"><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table>`;
}

/* ── 映射控件（详情与浮层共用） ───────────────────────── */
function columnSelect(selected) {
  const sel = document.createElement('select');
  const vars = [
    ['@page', '页码'],
    ['@total', '总页数'],
    ['@date', '生成日期'],
  ];
  sel.innerHTML =
    '<option value="">（不填充）</option>' +
    S.columns.map((c) =>
      `<option value="${escapeAttr(c.name)}"${c.name === selected ? ' selected' : ''}>${escapeHtml(c.name)}</option>`
    ).join('') +
    (S.titles
      ? `<optgroup label="AI"><option value="@titles"${selected === '@titles' ? ' selected' : ''}>AI 短标题（${Object.keys(S.titles).length} 条）</option></optgroup>`
      : '') +
    '<optgroup label="变量">' +
    vars.map(([v, label]) =>
      `<option value="${v}"${v === selected ? ' selected' : ''}>${label}</option>`
    ).join('') +
    '</optgroup>' +
    `<option value="@constant"${selected === '@constant' ? ' selected' : ''}>固定文本…（封面标题/结尾致辞）</option>`;
  return sel;
}

function buildFieldMapping(field) {
  const wrap = document.createElement('div');
  wrap.className = 'field-mapping';
  const specs = [];
  if (field.kind === 'text') {
    const slots = (field.slots || []).filter((s) => s.name);
    if (slots.length) slots.forEach((slot) => specs.push({ slot }));
    else specs.push({ slot: null });
  } else if (field.kind === 'image') {
    specs.push({ image: true });
  } else if (field.kind === 'table' && field.table && field.table.first_row_header) {
    (field.table.cells[0] || []).forEach((name, i) => specs.push({ tableCol: i, tableName: name }));
  } else if (field.kind === 'chart') {
    specs.push({ chart: true });
  }
  if (!specs.length) {
    wrap.innerHTML = '<p class="muted">该元素没有可映射位置。</p>';
    return wrap;
  }
  specs.forEach((spec) => wrap.appendChild(buildMappingRow(field, spec)));
  return wrap;
}

function currentColumnFor(field, spec) {
  if (spec.tableCol !== undefined) {
    const b = S.bindings.find((x) => x.field_id === field.id && x.kind === 'table');
    return (b && b.source.columns && b.source.columns[spec.tableCol]) || '';
  }
  const b = bindingForField(field.id, field.kind, spec.slot ? spec.slot.name : null);
  if (b && b.source.var) return '@' + b.source.var;
  if (b && b.source.value_map) return '@titles';
  if (b && b.source.value !== undefined && !b.source.column) return '@constant';
  return (b && b.source.column) || '';
}

function currentFitFor(field) {
  const b = S.bindings.find((x) => x.field_id === field.id && x.kind === 'image');
  return (b && b.options && b.options.fit) || 'contain';
}

function buildMappingRow(field, spec) {
  if (spec.chart) return buildChartControl(field);

  const line = document.createElement('label');
  line.className = 'map-line';
  const name = document.createElement('span');
  name.className = 'map-name';
  if (spec.slot) name.textContent = `「${spec.slot.name}」`;
  else if (spec.tableCol !== undefined) name.textContent = `表头「${spec.tableName}」`;
  else if (spec.image) name.textContent = '图片位';
  else name.textContent = '整段文本';

  const select = columnSelect(currentColumnFor(field, spec));
  select.addEventListener('change', () => {
    setFieldBinding(field, spec, select.value);
    afterMappingChange();
  });
  line.append(name, select);

  if (spec.image) {
    const fit = document.createElement('select');
    fit.className = 'fit-select';
    fit.innerHTML = ['contain', 'cover', 'stretch'].map((v) =>
      `<option value="${v}"${v === currentFitFor(field) ? ' selected' : ''}>${v}</option>`
    ).join('');
    fit.addEventListener('change', () => {
      setFieldBinding(field, spec, select.value, fit.value);
      afterMappingChange();
    });
    line.append(fit);
  }
  return line;
}

function buildChartControl(field) {
  const b = S.bindings.find((x) => x.field_id === field.id && x.kind === 'chart');
  const cat = (b && b.source.categories) || '';
  const val = (b && b.source.series && b.source.series[0] && b.source.series[0].value) || '';

  const line = document.createElement('label');
  line.className = 'map-line chart-line';
  const name = document.createElement('span');
  name.className = 'map-name';
  name.textContent = '类目 / 数值';

  const catSel = columnSelect(cat);
  const valSel = columnSelect(val);
  catSel.classList.add('mini');
  valSel.classList.add('mini');

  const update = () => {
    const idx = S.bindings.findIndex((x) => x.field_id === field.id && x.kind === 'chart');
    if (!catSel.value || !valSel.value) {
      if (idx >= 0) S.bindings.splice(idx, 1);
      afterMappingChange();
      return;
    }
    const nb = {
      field_id: field.id, kind: 'chart',
      source: { categories: catSel.value, series: [{ name: valSel.value, value: valSel.value }] },
      options: {},
    };
    if (idx >= 0) S.bindings[idx] = nb; else S.bindings.push(nb);
    afterMappingChange();
  };
  catSel.addEventListener('change', update);
  valSel.addEventListener('change', update);

  const pair = document.createElement('span');
  pair.className = 'chart-pair';
  pair.append(catSel, document.createTextNode(' / '), valSel);
  line.append(name, pair);
  return line;
}

function setFieldBinding(field, spec, column, fit) {
  if (spec.tableCol !== undefined) {
    let b = S.bindings.find((x) => x.field_id === field.id && x.kind === 'table');
    if (!b) {
      b = { field_id: field.id, kind: 'table', source: { mode: 'rows', columns: [] }, options: { header: true } };
      S.bindings.push(b);
    }
    const cols = b.source.columns || (b.source.columns = []);
    cols[spec.tableCol] = column;
    if (!cols.some(Boolean)) S.bindings = S.bindings.filter((x) => x !== b);
    return;
  }
  const kind = field.kind;
  const slotName = spec.slot ? spec.slot.name : null;
  const idx = S.bindings.findIndex((b) => {
    if (b.field_id !== field.id || b.kind !== kind) return false;
    if (kind === 'text') return (b.source.slot || null) === slotName;
    return true;
  });
  if (!column) {
    if (idx >= 0) S.bindings.splice(idx, 1);
    return;
  }
  let source;
  if (column === '@titles' && S.titles) {
    source = {
      column: S.titleColumn,
      key_column: S.titleKey,
      value_map: S.titles,
    };
  } else if (column === '@constant') {
    const text = window.prompt('输入固定文本（例如封面标题、结尾致辞）', '');
    if (text === null) return;
    if (!text.trim()) return;
    source = { value: text.trim() };
  } else if (column.startsWith('@')) {
    source = { var: column.slice(1) };
  } else {
    source = { column };
  }
  if (kind === 'text' && slotName) source.slot = slotName;
  const options = kind === 'image' ? { fit: fit || currentFitFor(field) } : {};
  const binding = { field_id: field.id, kind, source, options };
  if (idx >= 0) S.bindings[idx] = binding; else S.bindings.push(binding);
}

function afterMappingChange() {
  renderCanvas();
  renderFieldList();
  renderMappingSummary();
  persist();
}

/* ── 画布浮层 ─────────────────────────────────────────── */
function showPopover(field, clientX, clientY) {
  const pop = $('#canvasPopover');
  pop.innerHTML = '';
  const head = document.createElement('div');
  head.className = 'pop-head';
  head.innerHTML = `<span class="tag ${field.kind}">${KIND_NAMES[field.kind] || field.kind}</span>
    <b>${escapeHtml(field.name || field.id)}</b>`;
  pop.appendChild(head);
  pop.appendChild(buildFieldMapping(field));
  pop.hidden = false;
  pop.style.left = '0px';
  pop.style.top = '0px';
  const box = pop.getBoundingClientRect();
  let left = clientX + 14;
  let top = clientY + 14;
  if (left + box.width > window.innerWidth - 12) left = Math.max(12, clientX - box.width - 14);
  if (top + box.height > window.innerHeight - 12) {
    top = Math.max(12, window.innerHeight - box.height - 12);
  }
  pop.style.left = left + 'px';
  pop.style.top = top + 'px';
}

function hidePopover() {
  const pop = $('#canvasPopover');
  if (pop) pop.hidden = true;
}

/* ── 第 3 步：预览 ─────────────────────────────────────── */
function currentPagination() {
  const mode = $('#pageMode').value;
  if (mode === 'n_per_page') {
    return { mode, size: Math.max(1, Number($('#pageSize').value || 5)) };
  }
  return { mode: 'one_per_row' };
}

function renderMappingSummary() {
  const box = $('#mappingSummary');
  if (!box) return;
  if (!S.bindings.length) {
    box.innerHTML = '<p class="muted">还没有映射。回到第 2 步点元素选列。</p>';
    return;
  }
  const lines = S.bindings.map((b) => {
    const f = S.fields.find((x) => x.id === b.field_id);
    const label = f ? (f.name || f.id) : b.field_id;
    if (b.kind === 'text') {
      const slot = b.source.slot ? `「${b.source.slot}」` : '整段';
      if (b.source.var) return `${label} · ${slot} → 变量 ${b.source.var}`;
      return `${label} · ${slot} → ${b.source.column}`;
    }
    if (b.kind === 'table') return `${label} · 表格 → ${(b.source.columns || []).join(' / ')}`;
    if (b.kind === 'image') return `${label} · 图片 → ${b.source.column}`;
    if (b.kind === 'chart') {
      const s = (b.source.series || [])[0] || {};
      return `${label} · 图表 → ${b.source.categories} / ${s.value || ''}`;
    }
    return `${label} · ${b.kind}`;
  });
  box.innerHTML = lines.map((t) => `<div class="sum-item">${escapeHtml(t)}</div>`).join('');
}

function updatePageNav() {
  $('#pageLabel').textContent = `第 ${S.pageIndex + 1} / ${Math.max(1, S.pageTotal)} 页`;
  $('#btnPrevPage').disabled = S.pageIndex <= 0;
  $('#btnNextPage').disabled = S.pageIndex >= S.pageTotal - 1;
}

function renderPreviewWarnings(warnings) {
  const box = $('#previewWarnings');
  if (!warnings || !warnings.length) { box.innerHTML = ''; return; }
  box.innerHTML = '<div class="warn-title">提示</div>' + warnings.map((w) =>
    `<div class="warn-item">${escapeHtml(w)}</div>`
  ).join('');
}

async function refreshPreview(pageIndex) {
  if (!S.fields.length) return;
  const img = $('#previewImage');
  const loading = $('#previewLoading');
  const fallback = $('#previewFallback');
  loading.hidden = false;
  img.hidden = true;
  fallback.hidden = true;
  const profile = buildProfile();
  S.previewDirty = false;
  try {
    const hasLayout = Object.keys(profile.layout || {}).length > 0;
    const res = await api('/api/preview-page', {
      profile,
      page_index: Math.max(0, pageIndex),
      full: hasLayout,
    });
    S.pageIndex = res.page_index;
    S.pageTotal = res.page_total;
    updatePageNav();
    renderPreviewWarnings(res.warnings || []);
    if (res.image) {
      img.src = res.image;
      img.hidden = false;
      $('#previewMeta').textContent = `引擎 ${res.engine} · 本页 ${res.page_rows} 行 · ${res.cached ? '缓存命中' : '刚渲染'}`;
    } else {
      renderFallbackPreview(res);
      $('#previewMeta').textContent = `引擎不可用 · 本页 ${res.page_rows} 行`;
    }
  } catch (err) {
    fallback.hidden = false;
    fallback.innerHTML = `<p class="msg error">预览失败：${escapeHtml(err.message)}</p>`;
    $('#previewMeta').textContent = '';
  } finally {
    loading.hidden = true;
  }
}

function previewFieldText(field, row) {
  let text = field.text || '';
  (field.slots || []).forEach((s) => {
    const b = S.bindings.find((x) =>
      x.field_id === field.id && x.kind === 'text' && (x.source.slot || null) === (s.name || null)
    );
    if (!b) return;
    const v = row[b.source.column];
    if (v !== undefined && v !== null) text = text.split(s.raw).join(String(v));
  });
  return text || fieldPreview(field);
}

function renderFallbackPreview(res) {
  const box = $('#previewFallback');
  box.hidden = false;
  box.innerHTML = '';
  const row = (res.rows && res.rows[0]) || {};
  const VW = 960;
  const scale = VW / S.width;
  const VH = Math.round(S.height * scale);
  const svg = document.createElementNS(NS, 'svg');
  svg.setAttribute('viewBox', `0 0 ${VW} ${VH}`);
  svg.setAttribute('class', 'fallback-svg');
  S.fields.forEach((f) => {
    const rect = document.createElementNS(NS, 'rect');
    rect.setAttribute('x', f.rect.x * scale);
    rect.setAttribute('y', f.rect.y * scale);
    rect.setAttribute('width', f.rect.w * scale);
    rect.setAttribute('height', f.rect.h * scale);
    rect.setAttribute('fill', '#ffffff');
    rect.setAttribute('stroke', '#cbd5e1');
    svg.appendChild(rect);
    if (f.kind === 'text') {
      const t = document.createElementNS(NS, 'text');
      t.setAttribute('x', f.rect.x * scale + 6);
      t.setAttribute('y', f.rect.y * scale + 18);
      t.setAttribute('font-size', 13);
      t.setAttribute('fill', '#334155');
      t.textContent = previewFieldText(f, row).slice(0, 46);
      svg.appendChild(t);
    }
  });
  box.appendChild(svg);
  const note = document.createElement('p');
  note.className = 'muted';
  note.textContent = 'WPS 不可用，已降级为几何预览（仅文字与位置，样式不准确）。';
  box.appendChild(note);
}

function enterPreview() {
  if (!S.fields.length) return;
  renderMappingSummary();
  gotoStep(3);
  // 有封面时从第 2 页（第一个正文页）开始，避免一进来看到封面以为是正文
  const hasCover = S.cover.enabled && S.cover.title.trim();
  refreshPreview(hasCover ? 1 : 0);
}

/* ── 第 4 步：生成 ─────────────────────────────────────── */
function cleanBindings(list) {
  return list.filter((b) => {
    if (b.kind === 'table') return (b.source.columns || []).some(Boolean);
    if (b.kind === 'chart') return Boolean(b.source.categories);
    if (b.kind === 'constant') return Boolean(b.source.value);
    return Boolean(b.source.column || b.source.var || b.source.value_map);
  });
}

function renderLayoutPanel() {
  const coverSlide = $('#coverSlide');
  const endSlide = $('#endSlide');
  if (!coverSlide || !endSlide) return;
  const opts = (S.slidesData || [])
    .map((s) => {
      const label = s.kind === 'cover' ? '（封面）' : s.kind === 'end' ? '（结尾）' : '（正文）';
      return `<option value="${s.index}">第 ${s.index + 1} 页${label}</option>`;
    })
    .join('');
  coverSlide.innerHTML = opts;
  endSlide.innerHTML = opts;
  if (!S.cover.slide) S.cover.slide = (S.slidesData[0] || {}).index || 0;
  if (!S.end.slide) S.end.slide = (S.slidesData[S.slidesData.length - 1] || {}).index || 0;
  coverSlide.value = String(S.cover.slide);
  endSlide.value = String(S.end.slide);
  $('#coverEnabled').checked = S.cover.enabled;
  $('#coverTitle').value = S.cover.title;
  $('#coverSubtitle').value = S.cover.subtitle;
  $('#endEnabled').checked = S.end.enabled;
  $('#endText').value = S.end.text;
}

function bindLayoutPanel() {
  const bind = (sel, handler) => {
    const el = $(sel);
    if (el) el.addEventListener('change', handler);
  };
  bind('#coverEnabled', (e) => { S.cover.enabled = e.target.checked; persist(); });
  bind('#coverSlide', (e) => { S.cover.slide = Number(e.target.value); persist(); });
  bind('#endEnabled', (e) => { S.end.enabled = e.target.checked; persist(); });
  bind('#endSlide', (e) => { S.end.slide = Number(e.target.value); persist(); });
  ['#coverTitle', '#coverSubtitle', '#endText'].forEach((sel) => {
    const el = $(sel);
    if (!el) return;
    el.addEventListener('input', (e) => {
      if (sel === '#coverTitle') S.cover.title = e.target.value;
      else if (sel === '#coverSubtitle') S.cover.subtitle = e.target.value;
      else S.end.text = e.target.value;
      persist();
    });
  });
}

function pageTextField(slideIndex, order) {
  const page = (S.slidesData || []).find((s) => s.index === slideIndex);
  if (!page) return null;
  // 只挑「有文字」的文本框，避免选中装饰性空框；再按面积从大到小
  const texts = (page.fields || [])
    .filter((f) => f.kind === 'text' && (f.text || '').trim())
    .sort((a, b) => b.rect.w * b.rect.h - a.rect.w * a.rect.h);
  return texts[order] || null;
}

function buildProfile() {
  const layout = {};
  if (S.cover.enabled && S.cover.title.trim()) {
    const bindings = [];
    const main = pageTextField(S.cover.slide, 0);
    if (main) {
      bindings.push({ field_id: main.id, kind: 'constant', source: { value: S.cover.title.trim() }, options: {} });
    }
    if (S.cover.subtitle.trim()) {
      const sub = pageTextField(S.cover.slide, 1);
      if (sub) {
        bindings.push({ field_id: sub.id, kind: 'constant', source: { value: S.cover.subtitle.trim() }, options: {} });
      }
    }
    if (bindings.length) layout.cover = { slide_index: S.cover.slide, bindings };
  }
  if (S.end.enabled && S.end.text.trim()) {
    const main = pageTextField(S.end.slide, 0);
    if (main) {
      layout.end = {
        slide_index: S.end.slide,
        bindings: [{ field_id: main.id, kind: 'constant', source: { value: S.end.text.trim() }, options: {} }],
      };
    }
  }
  return {
    version: 1,
    template: S.template,
    slide_index: S.slide,
    excel: S.excel,
    sheet: S.sheet || '',
    header_row: S.headerRow,
    bindings: cleanBindings(S.bindings),
    pagination: currentPagination(),
    output: {},
    layout,
  };
}

async function doBuild() {
  const btn = $('#btnBuild');
  const msg = $('#mapMsg');
  const profile = buildProfile();
  if (!profile.bindings.length) { setMsg(msg, '至少建立一条映射再生成。', 'error'); gotoStep(2); return; }
  btn.disabled = true;
  try {
    const res = await api('/api/build', { profile });
    S.built = true;
    renderResult(res);
    gotoStep(4);
  } catch (err) {
    setMsg(msg, err.message, 'error');
    gotoStep(2);
  } finally {
    btn.disabled = false;
  }
}

function renderResult(res) {
  S.lastOutput = res.output;
  const box = $('#buildResult');
  box.innerHTML = `
    <div class="result-line"><b>完成</b><span>${res.rows} 行数据 → ${res.slides} 页幻灯片，替换 ${res.replaced} 处</span></div>
    <div class="result-path">${escapeHtml(res.output)}</div>
    ${res.warnings && res.warnings.length ? `<ul class="warn-list">${res.warnings.map((w) => `<li>${escapeHtml(w)}</li>`).join('')}</ul>` : '<p class="muted" style="margin-top:10px">没有警告。</p>'}`;
  $('#verifyBox').innerHTML = '';
  $('#previewBox').innerHTML = '';
}

async function doVerify() {
  const box = $('#verifyBox');
  if (!S.lastOutput) { box.innerHTML = '<p class="msg error">请先生成。</p>'; return; }
  const layout = buildProfile().layout || {};
  if (Object.keys(layout).length) {
    box.innerHTML =
      '<p class="muted">当前输出包含封面/结尾（多模板编排），逐页零漂移校验暂只支持单页模板。' +
      '可以先用单页模板验证映射，再回到这里生成整册。</p>';
    return;
  }
  box.innerHTML = '<p class="muted">校验中…</p>';
  try {
    const res = await api('/api/verify', {
      template: S.template, output: S.lastOutput, slide: S.slide,
    }, true);
    const issues = res.issues || [];
    const warns = issues.filter((i) => i.level === 'warning');
    box.innerHTML = `
      <div class="verify-box">
        <p class="${res.ok ? 'verify-ok' : 'verify-fail'}">
          ${res.ok ? '✓ 零漂移校验通过' : '✗ 校验发现结构性差异'}
          <span class="muted">（${res.pages_checked} 页 / ${res.shapes_checked} 个元素）</span>
        </p>
        ${issues.length ? `<ul class="issue-list">${issues.slice(0, 30).map((i) =>
          `<li class="${i.level === 'error' ? 'err' : 'warn'}">${escapeHtml(i.message)}${i.page !== null && i.page !== undefined ? `（第 ${i.page + 1} 页）` : ''}</li>`).join('')}</ul>` : ''}
        ${warns.length > 30 ? `<p class="muted">其余 ${warns.length - 30} 条已省略。</p>` : ''}
      </div>`;
  } catch (err) {
    box.innerHTML = `<p class="msg error">${escapeHtml(err.message)}</p>`;
  }
}

async function doPreview() {
  const box = $('#previewBox');
  if (!S.lastOutput) { box.innerHTML = '<p class="msg error">请先生成。</p>'; return; }
  box.innerHTML = '<p class="muted">正在渲染预览…</p>';
  try {
    const res = await api('/api/preview', { path: S.lastOutput, page: 1 });
    box.innerHTML = `<div class="preview-box"><span class="muted">第 1 页 · ${res.backend}</span><img alt="第 1 页预览" src="${res.image}"></div>`;
  } catch (err) {
    box.innerHTML = `<p class="msg error">预览不可用：${escapeHtml(err.message)}</p>`;
  }
}

/* ── 自动匹配 ─────────────────────────────────────────── */
async function autoMap(silent) {
  const msg = $('#mapMsg') || $('#scanMsg');
  if (!S.template || !S.excel) return;
  try {
    const res = await api('/api/suggest', {
      template: S.template, slide: S.slide, excel: S.excel,
      sheet: S.sheet || null, header_row: S.headerRow, threshold: 60,
    });
    S.bindings = res.profile.bindings || [];
    renderCanvas();
    renderFieldList();
    renderMappingSummary();
    persist();
    if (!silent) setMsg(msg, `自动匹配 ${res.matched} 项（槽位共 ${res.slots_total} 个）。`, res.notes.length ? 'error' : 'ok');
  } catch (err) {
    setMsg(msg, '自动建议失败：' + err.message, 'error');
  }
}

/* ── AI 短标题 ────────────────────────────────────────── */
function openTitlesModal() {
  if (!S.columns.length) {
    setMsg($('#mapMsg'), '请先读取 Excel。', 'error');
    return;
  }
  const col = $('#titleColumn');
  const key = $('#titleKey');
  const opts = S.columns.map((c) =>
    `<option value="${escapeAttr(c.name)}">${escapeHtml(c.name)}</option>`
  ).join('');
  col.innerHTML = opts;
  key.innerHTML = opts;
  const guessCol = S.columns.find((c) => /发现问题|问题描述|描述/.test(c.name));
  if (guessCol) col.value = guessCol.name;
  const guessKey = S.columns.find((c) => /序号|编号/.test(c.name));
  if (guessKey) key.value = guessKey.name;
  $('#titlePreview').innerHTML = '';
  setMsg($('#titleMsg'), S.titles ? `已有 ${Object.keys(S.titles).length} 条，重新生成会覆盖。` : '');
  $('#titlesModal').hidden = false;
}

async function runTitles() {
  const btn = $('#btnTitleRun');
  btn.disabled = true;
  setMsg($('#titleMsg'), '正在生成，约 30–60 秒…');
  try {
    const res = await api('/api/titles', {
      excel: S.excel,
      sheet: S.sheet || null,
      column: $('#titleColumn').value,
      key: $('#titleKey').value,
      header_row: S.headerRow,
      api_key: $('#titleApiKey').value.trim() || null,
    });
    S.titles = res.titles;
    S.titleKey = $('#titleKey').value;
    S.titleColumn = $('#titleColumn').value;
    const keys = Object.keys(res.titles).slice(0, 6);
    $('#titlePreview').innerHTML = keys.map((k) =>
      `<div class="row"><b>${escapeHtml(k)}</b><span>${escapeHtml(res.titles[k])}</span></div>`
    ).join('');
    setMsg(
      $('#titleMsg'),
      `已生成 ${res.count} 条${res.over16 ? `（${res.over16} 条超 16 字）` : '，全部 ≤16 字'}。关闭后在标题槽位选「AI 短标题」。`,
      'ok'
    );
    renderCanvas();
    renderFieldList();
    renderDetail();
    persist();
  } catch (err) {
    setMsg($('#titleMsg'), err.message, 'error');
  } finally {
    btn.disabled = false;
  }
}

/* ── 工具 ─────────────────────────────────────────────── */
function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}
function escapeAttr(s) { return escapeHtml(s); }

/* ── 事件绑定 ─────────────────────────────────────────── */
$('#btnScan').addEventListener('click', doScan);
bindLayoutPanel();
$('#btnToPreview').addEventListener('click', enterPreview);
$('#btnAutoMap').addEventListener('click', () => autoMap(false));
$('#btnTitles').addEventListener('click', openTitlesModal);
$('#btnTitleCancel').addEventListener('click', () => { $('#titlesModal').hidden = true; });
$('#btnTitleRun').addEventListener('click', runTitles);
$('#btnBuild').addEventListener('click', doBuild);
$('#btnBuildAgain').addEventListener('click', doBuild);
$('#btnVerify').addEventListener('click', doVerify);
$('#btnPreview').addEventListener('click', doPreview);
$('#btnRefreshPreview').addEventListener('click', () => refreshPreview(S.pageIndex));
$('#btnPrevPage').addEventListener('click', () => refreshPreview(S.pageIndex - 1));
$('#btnNextPage').addEventListener('click', () => refreshPreview(S.pageIndex + 1));
$('#btnEditMapping').addEventListener('click', () => gotoStep(2));

$('#btnExportProfile').addEventListener('click', () => {
  const profile = buildProfile();
  const blob = new Blob([JSON.stringify(profile, null, 2)], { type: 'application/json' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'excel2ppt-profile.json';
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 2000);
});

$('#importProfile').addEventListener('change', async (e) => {
  const file = e.target.files[0];
  if (!file) return;
  try {
    const data = JSON.parse(await file.text());
    if (!data || !Array.isArray(data.bindings)) throw new Error('文件里没有 bindings 数组');
    if (data.template) $('#templatePath').value = data.template;
    if (data.excel) $('#excelPath').value = data.excel;
    if (typeof data.slide_index === 'number') $('#slideIndex').value = data.slide_index;
    if (data.sheet) $('#sheet').value = data.sheet;
    if (typeof data.header_row === 'number') $('#headerRow').value = data.header_row;
    S.bindings = data.bindings;
    persist();
    renderCanvas();
    renderFieldList();
    renderMappingSummary();
    setMsg($('#mapMsg'), `已导入 ${data.bindings.length} 条映射。`, 'ok');
  } catch (err) {
    setMsg($('#mapMsg'), '导入失败：' + err.message, 'error');
  }
});

$('#pageMode').addEventListener('change', () => {
  $('#pageSizeWrap').hidden = $('#pageMode').value !== 'n_per_page';
  if (!$('#step3').hidden) refreshPreview(0);
});
$('#pageSize').addEventListener('change', () => {
  if (!$('#step3').hidden) refreshPreview(0);
});

$$('.step').forEach((b) => b.addEventListener('click', () => {
  const n = Number(b.dataset.step);
  if (n === 2 && !S.fields.length) return;
  if (n === 3) { renderMappingSummary(); gotoStep(3); refreshPreview(S.pageIndex); return; }
  if (n === 4 && !S.built) return;
  gotoStep(n);
}));
$$('[data-goto]').forEach((b) => b.addEventListener('click', () => {
  const n = Number(b.dataset.goto);
  if (n === 3) { renderMappingSummary(); gotoStep(3); refreshPreview(S.pageIndex); return; }
  gotoStep(n);
}));

document.addEventListener('click', (e) => {
  const pop = $('#canvasPopover');
  if (pop.hidden) return;
  if (pop.contains(e.target) || (e.target.closest && e.target.closest('#canvas'))) return;
  hidePopover();
});
document.addEventListener('keydown', (e) => { if (e.key === 'Escape') hidePopover(); });

$('#templateFile').addEventListener('change', async (e) => {
  const file = e.target.files[0];
  if (!file) return;
  try {
    const path = await upload('template', file);
    $('#templatePath').value = path;
    setMsg($('#scanMsg'), '模板已上传：' + file.name, 'ok');
  } catch (err) { setMsg($('#scanMsg'), err.message, 'error'); }
});

$('#excelFile').addEventListener('change', async (e) => {
  const file = e.target.files[0];
  if (!file) return;
  try {
    const path = await upload('excel', file);
    $('#excelPath').value = path;
    setMsg($('#scanMsg'), 'Excel 已上传：' + file.name, 'ok');
  } catch (err) { setMsg($('#scanMsg'), err.message, 'error'); }
});

/* 初始化 */
restore();
$('#footStatus').textContent = 'v0.4.0';
fetch('/api/health').then((r) => r.json()).then((d) => {
  if (d.ok) $('#footStatus').textContent = 'v' + d.version + ' · 本地服务已就绪';
}).catch(() => {});
