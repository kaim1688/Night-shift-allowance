export const NIGHT_AMOUNT = 400;
export const NIGHT_START = '18:50';
export const NIGHT_END = '06:50';
export const NIGHT_TIMES = {
  regular: {start: NIGHT_START, end: NIGHT_END},
  ticket: {start: '18:40', end: '06:40'},
};

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
    const role = employee.nightRole === 'ticket' ? 'ticket' : 'regular';
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
