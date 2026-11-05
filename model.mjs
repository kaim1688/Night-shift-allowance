export const NIGHT_AMOUNT = 400;
export const NIGHT_START = '18:50';
export const NIGHT_END = '06:50';
export const FORM_WORK_TYPES = ['三班制（2輪1例隔週休班制）','夜班制','電務輪勤制','工務夜間現場實際工作','乘務隨車制','車勤服務制'];
export const FORM_NOTES = [
  '依排定班表，於 22:00～06:00 間實際執行工作者，每次發給 250 元；如於 00:00～04:00 有實際工作事實者，再加發 150 元。每人每晚最高發給 400 元。',
  '於 22:00～06:00 間有請假情形者，不予發給。',
  '依其他規定支領與夜點費同性質給與者，僅得擇一支領（同時段只能擇一領取）。',
  '各單位對所屬夜間工作，應確實督導考核。如有審核不實或虛報情事，除追繳已領之夜點費，失職人員視情節輕重依章議處。',
];
export const NIGHT_TIMES = {
  regular: {start: NIGHT_START, end: NIGHT_END},
  ticket: {start: '18:40', end: '06:40'},
  deputy: {start: '18:56', end: '06:56'},
};
export const NIGHT_LABELS = {regular:'一般夜班', ticket:'售票領班夜班', deputy:'副站長夜班'};

export function classifyShift(raw) {
  const shift = String(raw).replace(/\s/g, '').replace(/^[*$＄＊]+/, '');
  if (shift === '夜') return 'night';
  if (shift.includes('夜') || shift.includes('公式無有效結果')) return 'review';
  return 'other';
}

export function buildReports(schedule, employees) {
  const source = new Map(schedule.people.map(person => [person.name, person]));
  const known = new Set(employees.map(person => person.name));
  const reports = employees.map(employee => {
    const person = source.get(employee.name);
    const role = employee.title === '副站長' ? 'deputy' : Object.hasOwn(NIGHT_TIMES, employee.nightRole) ? employee.nightRole : 'regular';
    const nights = person?.days.filter(item => classifyShift(item.shift) === 'night')
      .map(item => ({...item, role})) ?? [];
    const issues = person?.days.filter(item => classifyShift(item.shift) === 'review') ?? [];
    return {...employee, nights, issues, present: !!person, total: nights.length * NIGHT_AMOUNT,
      status: !person ? 'missing' : issues.length ? 'review' : nights.length ? 'ready' : 'none'};
  });
  for (const person of schedule.people) {
    if (!known.has(person.name)) {
      reports.push({name: person.name, title: '', code: '', nights: person.days.filter(item => classifyShift(item.shift) === 'night'),
        issues: [], present: true, total: 0, status: 'unknown'});
    }
  }
  return reports;
}

export function pageChunks(nights) {
  const pages = [];
  for (let i = 0; i < nights.length; i += 19) pages.push(nights.slice(i, i + 19));
  return pages;
}
