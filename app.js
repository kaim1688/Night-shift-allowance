import {buildReports, pageChunks, NIGHT_AMOUNT, NIGHT_TIMES, NIGHT_LABELS, FORM_WORK_TYPES, FORM_NOTES} from './model.mjs';
import {readSchedule, readEmployees, validateEmployees, createExportWorkbook, serializeExportWorkbook} from './spreadsheet.mjs';

const $ = id => document.getElementById(id);
const escape = text => String(text).replace(/[&<>"']/g, char => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[char]));
const money = value => value.toLocaleString('zh-TW');
const statusLabel = {ready:'可列印', none:'無夜班', missing:'無班表資料', review:'待核對', unknown:'未建檔'};
let catTimer;
let employees = [], schedule = null, reports = [], selected = new Set(), activeName = '', importVersion = 0;

function setScheduleEnabled(enabled) {
  $('schedule-file').disabled = !enabled;
  $('choose-schedule').disabled = !enabled;
}
$('choose-roster').addEventListener('click', () => $('roster-file').click());
$('choose-schedule').addEventListener('click', () => $('schedule-file').click());

let catInView = false;
function syncCatMotion() {
  $('cat-station').classList.toggle('cat-awake', catInView && !document.hidden);
}
new IntersectionObserver(([entry]) => {
  catInView = entry.isIntersecting;
  syncCatMotion();
}, {threshold:.05}).observe($('cat-station'));
document.addEventListener('visibilitychange', syncCatMotion);

function setCatConfused(confused) {
  clearTimeout(catTimer);
  $('cat-station').classList.remove('eating', 'patting');
  $('cat-station').classList.toggle('confused', confused);
  $('cat-greeting').textContent = confused ? '這個看不懂~' : '喵~幫你填表!';
  if (confused) $('empty').hidden = false;
}

function applyTheme(theme) {
  document.documentElement.dataset.theme = theme;
  $('theme-toggle').textContent = theme === 'dark' ? '淺色模式' : '深色模式';
  $('theme-toggle').setAttribute('aria-pressed', String(theme === 'dark'));
}
let savedTheme;
try { savedTheme = localStorage.getItem('theme'); } catch {}
applyTheme(savedTheme === 'dark' ? 'dark' : 'light');
$('theme-toggle').addEventListener('click', () => {
  const theme = document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark';
  applyTheme(theme);
  try { localStorage.setItem('theme', theme); } catch {}
});

function reset() {
  schedule = null; reports = []; selected.clear(); activeName = '';
  setCatConfused(false);
  $('workspace').hidden = true; $('empty').hidden = false;
  $('print-area').replaceChildren(); $('preview').replaceChildren();
  $('print-button').disabled = true; $('print-person').disabled = true;
  $('export-button').disabled = true; $('export-status').textContent = '';
}

function formHtml(report, days, page, pageCount) {
  const blank = report === null;
  const {year, month} = schedule ?? {};
  report ??= {title: '', code: ''};
  const rows = Array.from({length:19}, (_, index) => {
    const night = days[index];
    return `<tr>${index === 0 ? `<td class="work-type" rowspan="19">${FORM_WORK_TYPES.map((label,i) => `<span class="work-option">${!blank && i === 0 ? '■' : '□'} ${label}</span>`).join('')}</td>` : ''}<td>${night ? `${month}月${night.day}日` : '月　日'}</td><td class="time">${night ? NIGHT_TIMES[night.role].start : ''}</td><td class="time end-time">次日<span class="end-clock">${night ? NIGHT_TIMES[night.role].end : ''}</span></td><td>${night ? '1' : ''}</td><td>${night ? '1' : ''}</td><td>${night ? '400' : ''}</td><td></td></tr>`;
  }).join('');
  return `<article class="sheet" aria-label="個人夜點費請領表"><h2 class="form-title">${blank ? '____' : year-1911} 年 ${blank ? '____' : month} 月個人夜點費請領表</h2><div class="form-unit">單位：<span class="unit-value">竹南站</span></div><div class="form-identity"><span>姓名：<span class="signature"></span></span><span>（請簽名及蓋章）</span><span>職稱：<span class="title-value">${escape(report.title)}</span></span><span>員工代號：<span class="form-code">${escape(report.code)}</span></span></div><table class="claim-table"><colgroup><col style="width:30mm"><col style="width:22mm"><col style="width:16mm"><col style="width:22mm"><col style="width:23mm"><col style="width:25mm"><col style="width:18mm"><col style="width:34mm"></colgroup><thead><tr><th scope="col">夜間工作別</th><th scope="col">日期</th><th scope="col">開始<br>時間</th><th scope="col">結束<br>時間</th><th scope="col">基本時段<br>250 元<small>（22時至次日6時）</small></th><th scope="col">加發時段<br>加發150元<small>（24時至次日4時）</small></th><th scope="col">每晚<br>金額</th><th scope="col">直接主管<br>審核簽章</th></tr></thead><tbody>${rows}<tr class="total-row"><td></td><td class="total-label" colspan="3">${pageCount > 1 ? '本頁合計' : '合計'}</td><td>${blank ? '' : days.length + ' 次'}</td><td>${blank ? '' : days.length + ' 次'}</td><td>${blank ? '' : money(days.length*NIGHT_AMOUNT) + '元'}</td><td></td></tr></tbody></table><div class="form-notes"><p>註：</p><ol>${FORM_NOTES.map(note => `<li>${escape(note)}</li>`).join('')}</ol></div>${pageCount > 1 ? `<div class="continuation">第 ${page+1}／${pageCount} 頁 · 全月 ${report.nights.length} 次，共 ${money(report.total)} 元</div>` : ''}</article>`;
}

function sizePreview() {
  const width = $('paper-viewport').clientWidth - (matchMedia('(max-width:760px)').matches ? 16 : 36);
  const scale = Math.min(1, Math.max(.1, width / (210 * 96 / 25.4)));
  for (const box of $('preview').children) {
    box.style.height = `${297 * 96 / 25.4 * scale}px`;
    const sheet = box.querySelector('.sheet');
    if (sheet) sheet.style.transform = `scale(${scale})`;
  }
}

function showPerson(name) {
  activeName = name;
  const person = reports.find(report => report.name === name);
  if (!person) return;
  $('preview-name').textContent = `${person.name}的請領表`;
  $('preview-note').textContent = person.status === 'ready' ? `${person.nights.length} 次夜班 · ${money(person.total)} 元 · 姓名簽章欄留白` : statusLabel[person.status];
  $('print-person').disabled = person.status !== 'ready';
  $('person-message').hidden = person.status === 'ready';
  const messages = {none:'這位人員在本月班表沒有夜班，不產生請領表。', missing:'本月班表未列出這位人員，無法判斷夜班，不會產生請領表。', unknown:'名冊沒有這位人員的職稱與代號。請先更新 employees.json，再重新開啟並匯入班表。', review:'以下內容無法確定實際夜間出勤，請修正來源班表後重新匯入：'};
  $('person-message').innerHTML = escape(messages[person.status] ?? '') + (person.issues.length ? `<ul class="issue-list">${person.issues.map(item => `<li>${schedule.month}月${item.day}日：${escape(item.shift)}</li>`).join('')}</ul>` : '');
  $('preview').innerHTML = person.status === 'ready' ? pageChunks(person.nights).map((days,index,array) => `<div class="preview-sheet-box">${formHtml(person,days,index,array.length)}</div>`).join('') : '';
  renderNightEditor(person);
  sizePreview();
  for (const row of document.querySelectorAll('.person-row')) row.classList.toggle('active', row.dataset.name === name);
}

function updateSelection() {
  $('print-button').disabled = !selected.size;
  $('print-button').textContent = `列印已勾選 ${selected.size} 人`;
  $('export-button').disabled = !selected.size;
  $('export-button').textContent = `匯出 Excel（已勾選 ${selected.size} 人）`;
  $('export-status').textContent = '';
  const picked = reports.filter(person => selected.has(person.name) && person.status === 'ready');
  const nights = picked.reduce((sum,person) => sum+person.nights.length,0);
  $('selection-summary').textContent = `已勾選 ${picked.length} 人 · ${nights} 次夜班 · ${money(nights*NIGHT_AMOUNT)} 元`;
}

function renderNightEditor(person) {
  $('night-editor').hidden = person.status !== 'ready';
  $('night-roles').innerHTML = person.status === 'ready' ? person.nights.map(night =>
    `<label>${schedule.month}月${night.day}日<select data-day="${night.day}">${Object.entries(NIGHT_TIMES).map(([role,time]) => `<option value="${role}" ${night.role===role?'selected':''}>${NIGHT_LABELS[role]} ${time.start}～${time.end}</option>`).join('')}</select></label>`).join('') : '';
}
$('night-roles').addEventListener('change', event => {
  const select = event.target;
  if (!select.matches('select') || !Object.hasOwn(NIGHT_TIMES, select.value)) return;
  const person = reports.find(person => person.name === activeName);
  const night = person?.nights.find(night => night.day === Number(select.dataset.day));
  if (!night) return;
  night.role = select.value;
  $('export-status').textContent = '';
  $('preview').innerHTML = pageChunks(person.nights).map((days,index,array) => `<div class="preview-sheet-box">${formHtml(person,days,index,array.length)}</div>`).join('');
  sizePreview();
});

function renderPeople() {
  const query = $('search').value.trim();
  const visible = reports.filter(person => !['none','missing'].includes(person.status) && (person.name.includes(query) || person.code.includes(query)));
  const excluded = reports.filter(person => ['none','missing'].includes(person.status));
  $('excluded-summary').textContent = `未列入請領的人員（${excluded.length} 人）`;
  $('excluded-people').hidden = !excluded.length;
  $('excluded-list').innerHTML = excluded.map(person => `<p>${escape(person.name)}<span>${statusLabel[person.status]}</span></p>`).join('');
  $('people-list').innerHTML = visible.length ? visible.map(person => `<div class="person-row${person.name===activeName?' active':''}" data-name="${escape(person.name)}"><input type="checkbox" aria-label="選取${escape(person.name)}" ${selected.has(person.name)?'checked':''} ${person.status!=='ready'?'disabled':''}><button type="button" class="person-open"><strong>${escape(person.name)}</strong><small>${escape(person.title || '職稱待補')} · ${escape(person.code || '代號待補')}</small></button><span class="person-right">${person.status==='ready'?`<b>${person.nights.length} 次</b>${money(person.total)} 元`:`<span class="${['review','unknown'].includes(person.status)?'pending':''}">${statusLabel[person.status]}</span>`}</span></div>`).join('') : '<p class="no-results">找不到相符的人員。</p>';
  for (const row of $('people-list').querySelectorAll('.person-row')) {
    row.querySelector('input').addEventListener('change', event => {
      event.target.checked ? selected.add(row.dataset.name) : selected.delete(row.dataset.name);
      updateSelection();
    });
    row.querySelector('button').addEventListener('click', () => showPerson(row.dataset.name));
  }
  updateSelection();
}

function printReports(names) {
  if (!schedule) return;
  const printable = reports.filter(person => names.has(person.name) && person.status === 'ready');
  if (!printable.length) return;
  $('print-area').innerHTML = printable.map(person => pageChunks(person.nights).map((days,index,array) => formHtml(person,days,index,array.length)).join('')).join('');
  document.title = `${schedule.year-1911}年${schedule.month}月夜點費請領表`;
  window.print();
}

$('schedule-file').addEventListener('change', async event => {
  const version = ++importVersion;
  const file = event.target.files[0];
  reset(); $('error').hidden = true;
  if (!file) { $('file-state').textContent = '尚未匯入'; return; }
  $('file-state').textContent = '正在讀取班表…';
  try {
    if (!file.name.toLowerCase().endsWith('.xlsx')) throw new Error('請使用 .xlsx 檔案，舊版 .xls 請先另存新檔。');
    if (file.size > 10 * 1024 * 1024) throw new Error('班表超過 10 MB，請使用較小的檔案。');
    if (!employees.length) throw new Error('人員名冊尚未載入。請重新整理頁面後再試。');
    const data = readSchedule(await file.arrayBuffer());
    if (version !== importVersion) return;
    schedule = data;
    reports = buildReports(schedule, employees);
    selected = new Set(reports.filter(person => person.status === 'ready').map(person => person.name));
    $('empty').hidden = true; $('workspace').hidden = false; $('search').value = '';
    $('period-title').textContent = `${schedule.year-1911} 年 ${schedule.month} 月夜點費`;
    const ready = reports.filter(person => person.status === 'ready');
    const nights = ready.reduce((sum,person) => sum+person.nights.length,0);
    const incomplete = reports.filter(person => ['review','unknown'].includes(person.status));
    $('summary').textContent = `${schedule.people.length} 人排班 · ${ready.length} 人可列印 · ${nights} 次夜班 · ${money(nights*NIGHT_AMOUNT)} 元${incomplete.length?'（另有待核對資料）':''}`;
    const warnings = [...schedule.warnings];

    if (incomplete.length) warnings.unshift(`待核對／未建檔：${incomplete.map(person => person.name).join('、')}。這些人員尚未納入金額及列印。`);
    $('warnings').hidden = !warnings.length;
    $('warnings').textContent = warnings.join(' ');
    $('file-state').textContent = `${file.name} · ${schedule.year-1911}年${schedule.month}月 · 已匯入`;
    $('excluded-people').open = false;
    $('night-editor').open = false;
    renderPeople();
    showPerson(ready[0]?.name ?? reports[0].name);
  } catch (error) {
    if (version !== importVersion) return;
    reset(); $('error').hidden = false;
    $('error').textContent = error.message || '無法讀取班表，請重新選擇檔案。';
    $('file-state').textContent = '匯入未完成，請重新選擇檔案。';
    setCatConfused(true);
  } finally {
    if (version === importVersion) event.target.value = '';
  }
});
$('search').addEventListener('input', renderPeople);
$('select-all').addEventListener('click', () => {selected = new Set(reports.filter(person => person.status==='ready').map(person => person.name)); renderPeople();});
$('clear-all').addEventListener('click', () => {selected.clear();renderPeople();});
$('print-button').addEventListener('click', () => printReports(selected));
$('print-person').addEventListener('click', () => printReports(new Set([activeName])));
$('export-button').addEventListener('click', () => {
  if (!schedule || !selected.size) return;
  try {
    const workbook = createExportWorkbook(schedule, reports, selected);
    const url=URL.createObjectURL(new Blob([serializeExportWorkbook(workbook)],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
    const link=document.createElement('a');
    link.href=url;link.download=`${schedule.year-1911}年${schedule.month}月夜點費_已勾選${selected.size}人.xlsx`;
    document.body.append(link);link.click();link.remove();
    setTimeout(()=>URL.revokeObjectURL(url),1000);
    $('export-status').textContent = `已匯出 ${selected.size} 人的請領表、彙總與逐日明細（含目前替班設定）。`;
  } catch {
    $('export-status').textContent = '匯出未完成，請再試一次，或使用列印另存 PDF。';
  }
});
$('print-blank').addEventListener('click', () => {
  $('print-area').innerHTML = formHtml(null, [], 0, 1);
  document.title = '夜點費空白請領表';
  window.print();
});
function playCat(state, message, duration) {
  if ($('cat-station').classList.contains('confused')) return;
  clearTimeout(catTimer);
  $('cat-station').classList.remove('eating', 'patting');
  $('cat-station').classList.add(state);
  $('cat-greeting').textContent = message;
  catTimer = setTimeout(() => {
    $('cat-station').classList.remove(state);
    $('cat-greeting').textContent = '喵~幫你填表!';
  }, duration);
}
$('pet-cat').addEventListener('click', () => playCat('eating', '喵，夜點開動！你也記得休息。', 3600));
$('pet-head').addEventListener('click', () => playCat('patting', '呼嚕~謝謝你，慢慢核對就好。', 1600));

window.addEventListener('afterprint', () => {$('print-area').replaceChildren();document.title='夜點費填表器';});
new ResizeObserver(sizePreview).observe($('paper-viewport'));

try {
  const response = await fetch('./employees.json');
  if (!response.ok) throw new Error('名冊載入失敗');
  employees = await response.json();
  if (Array.isArray(employees) && !employees.length) {
    $('roster-import').hidden = false;
    setScheduleEnabled(false);
    $('file-state').textContent = '請先載入人員名冊，再選擇班表。';
  } else {
    validateEmployees(employees);
    $('roster-state').textContent = `已使用本機內建 ${employees.length} 人名冊，可直接進行步驟 2；也可選擇檔案更換。`;
  }
} catch {
  $('error').hidden=false;
  $('error').textContent='無法載入人員名冊，請確認 employees.json 存在且格式正確，再重新整理頁面。';
  setScheduleEnabled(false);
}
$('roster-file').addEventListener('change', async event => {
  const file = event.target.files[0];
  if (!file) return;
  try {
    if (file.size > 10 * 1024 * 1024) throw new Error('名冊請使用10 MB以內的檔案。');
    const next = file.name.toLowerCase().endsWith('.json') ? validateEmployees(JSON.parse(await file.text())) : readEmployees(await file.arrayBuffer());
    reset();
    employees = next;
    setScheduleEnabled(true);
    $('roster-state').textContent = `已載入 ${employees.length} 人；名冊只保留在本頁，重新整理後需重選。`;
    $('file-state').textContent = '人員名冊已載入，請選擇班表。';
    $('error').hidden = true;
  } catch (error) {
    $('error').hidden = false;
    $('error').textContent = (error instanceof SyntaxError ? '名冊格式不正確，請選擇正確的 .json 或 .xlsx 名冊。' : error.message) + (employees.length ? ' 原名冊仍保留，請重新選擇。' : '');
    setCatConfused(true);
  } finally {event.target.value = '';}
});
