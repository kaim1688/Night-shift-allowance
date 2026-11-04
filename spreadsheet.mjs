function readWorkbook(data) {
  const bytes = new Uint8Array(data);
  if (!bytes.length || bytes.length > 10 * 1024 * 1024) throw new Error('請選擇10 MB以內的.xlsx檔案。');
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let end = -1;
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--) {
    if (view.getUint32(i, true) === 0x06054b50 && i + 22 + view.getUint16(i + 20, true) === bytes.length) {end = i; break;}
  }
  if (end < 0) throw new Error('無法讀取Excel，請使用未加密的.xlsx檔案。');
  const count = view.getUint16(end + 10, true);
  let offset = view.getUint32(end + 16, true), total = 0;
  if (!count || count > 2000) throw new Error('Excel內部檔案數量不正確。');
  for (let i = 0; i < count; i++) {
    if (offset + 46 > end || view.getUint32(offset, true) !== 0x02014b50) throw new Error('Excel格式損壞。');
    if (view.getUint16(offset + 8, true) & 1) throw new Error('不支援加密的Excel。');
    total += view.getUint32(offset + 24, true);
    if (total > 50 * 1024 * 1024) throw new Error('Excel解壓縮後過大，請另存精簡班表。');
    offset += 46 + view.getUint16(offset + 28, true) + view.getUint16(offset + 30, true) + view.getUint16(offset + 32, true);
  }
  if (!globalThis.XLSX) throw new Error('Excel讀取元件尚未載入，請重新整理。');
  try {return XLSX.read(bytes, {type:'array', cellFormula:true});}
  catch {throw new Error('Excel格式損壞或不受支援，請另存.xlsx後重試。');}
}

function sheetRows(sheet) {
  const rows = new Map();
  const addresses = Object.keys(sheet).filter(key => !key.startsWith('!'));
  if (addresses.length > 200000) throw new Error('Excel儲存格過多。');
  for (const address of addresses) {
    const {r,c} = XLSX.utils.decode_cell(address);
    const cell = sheet[address];
    const value = cell.t === 'e' || (cell.f && cell.v == null) ? '【公式無有效結果】' : String(cell.v ?? '').trim();
    if (!value) continue;
    if (!rows.has(r)) rows.set(r, new Map());
    rows.get(r).set(c, value);
  }
  return [...rows.entries()].sort((a,b) => a[0]-b[0]);
}

export function readSchedule(data) {
  const workbook = readWorkbook(data), results = [];
  for (const name of workbook.SheetNames) {
    const rows = sheetRows(workbook.Sheets[name]), headers = [], periods = new Set();
    rows.forEach(([row,cells],index) => {
      const days = [...cells].filter(([,v]) => /^\d{1,2}$/.test(v) && +v >= 1 && +v <= 31)
        .map(([col,v]) => [col,+v]).sort((a,b)=>a[0]-b[0]);
      if (days.length >= 28 && days.every(([,v],i)=>v===i+1)) headers.push({index,days});
      for (const text of cells.values()) for (const match of text.matchAll(/(\d{3,4})\s*年\s*(\d{1,2})\s*月/g)) {
        const year = +match[1] < 1911 ? +match[1] + 1911 : +match[1];
        periods.add(`${year}-${+match[2]}`);
      }
    });
    if (!headers.length) continue;
    if (periods.size !== 1) throw new Error('請在班表標題提供唯一的年月，例如「115年11月份」。');
    const [year,month] = [...periods][0].split('-').map(Number);
    if (year < 1912 || year > 2200 || month < 1 || month > 12) throw new Error('班表年月不正確。');
    const count = new Date(Date.UTC(year,month,0)).getUTCDate(), people = [], warnings = [], seen = new Set();
    headers.forEach(({index,days},block) => {
      if (days.length !== count) throw new Error('日期欄與該月天數不一致，請確認沒有缺日或多日。');
      const end = headers[block+1]?.index ?? rows.length;
      for (const [row,cells] of rows.slice(index+1,end)) {
        const person = (cells.get(days[0][0]-1) ?? '').replace(/\s/g,'');
        if (!person || ['星期','姓名','日期'].includes(person)) continue;
        const shifts = days.map(([col,day])=>({day,shift:cells.get(col) ?? ''}));
        if (person === '段替班') {
          if (shifts.some(d=>d.shift.includes('夜'))) throw new Error(`第 ${row+1} 列段替班有夜班，請先補上實際人員姓名。`);
          continue;
        }
        if (seen.has(person)) throw new Error(`「${person}」重複出現在班表，請確認後重新匯入。`);
        seen.add(person);
        people.push({name:person,sourceRow:row+1,days:shifts});
      }
    });
    if (!people.length) throw new Error('找到日期欄，但未找到人員姓名。');
    results.push({year,month,sheet:name,people,warnings});
  }
  if (results.length !== 1) throw new Error('請使用只有一張有效班表的Excel；目前找不到班表或有多張有效班表。');
  return results[0];
}

export function validateEmployees(people) {
  if (!Array.isArray(people) || !people.length || people.some(p=>!p.name || !p.title || !/^\d{7}$/.test(p.code) || (p.nightRole != null && !['regular','ticket'].includes(p.nightRole))) ||
      new Set(people.map(p=>p.name)).size !== people.length || new Set(people.map(p=>p.code)).size !== people.length)
    throw new Error('名冊需有姓名、職稱、7碼員工代號，且姓名及代號不能重複。');
  return people;
}

export function readEmployees(data) {
  const workbook = readWorkbook(data), people = [];
  for (const sheet of Object.values(workbook.Sheets)) {
    const rows = sheetRows(sheet);
    const header = rows.findIndex(([,cells])=>cells.get(0)==='姓名' && cells.get(1)==='職稱' && cells.get(2)==='員工代號');
    if (header < 0) continue;
    for (const [,cells] of rows.slice(header+1)) {
      if (!cells.get(0)) continue;
      const person = {name:cells.get(0),title:cells.get(1),code:cells.get(2)};
      if (cells.get(3)) person.nightRole = cells.get(3);
      people.push(person);
    }
  }
  return validateEmployees(people);
}
