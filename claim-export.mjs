import {pageChunks, NIGHT_AMOUNT, NIGHT_TIMES, FORM_NOTES, FORM_WORK_TYPES} from './model.mjs';

// Millimetre boundaries shared by the table and the 42mm signature / 32mm title fields.
const EDGES = [0,12,30,40,52,54,68,90,102,113,134,138,156,160,190];
const TABLE_COLS = [[0,1],[2,3],[4,5],[6,6],[7,8],[9,10],[11,11],[12,13]];

export function appendClaimSheet(workbook, schedule, person) {
  const sheet = {'!merges':[], '!rows':[], '!cols':EDGES.slice(1).map((end,i)=>({width:(end-EDGES[i])*96/25.4/7})),
    '!margins':{left:10/25.4,right:10/25.4,top:10/25.4,bottom:10/25.4,header:0,footer:0}, '!claimStyles':{}};
  const pages = pageChunks(person.nights);
  function box(row, first, last, text, style, bottom=row) {
    for (let r=row; r<=bottom; r++) for (let c=first; c<=last; c++) {
      const address=XLSX.utils.encode_cell({r,c});
      sheet[address]={t:typeof text==='number' && r===row && c===first?'n':'s',v:r===row && c===first?text:''};
      sheet['!claimStyles'][address]=style;
    }
    if (first!==last || row!==bottom) sheet['!merges'].push({s:{r:row,c:first},e:{r:bottom,c:last}});
  }
  pages.forEach((days,index)=>{
    const start=index*34;
    const heights=[25,7,20,5,22,6,62,...Array(19).fill(22),25,6,14,32,18,28,36,14];
    heights.forEach((hpt,row)=>sheet['!rows'][start+row]={hpt});
    const line=(row,first,last,text,style,bottom=row)=>box(start+row,first,last,text,style,start+bottom);
    line(0,0,13,`${schedule.year-1911} 年 ${schedule.month} 月個人夜點費請領表`,'title');
    line(2,0,0,'單位：','unitLabel');line(2,1,2,'竹南站','unitValue');
    line(4,0,0,'姓名：','identity');line(4,1,4,'','signature');
    line(4,5,6,'（請簽名及蓋章）','identity');line(4,7,7,'職稱：','identity');
    line(4,8,9,person.title,'identityValue');line(4,10,12,'員工代號：','identity');line(4,13,13,person.code,'code');
    const headers=['夜間工作別','日期','開始\n時間','結束\n時間','基本時段\n250 元\n（22時至次日6時）','加發時段\n加發150元\n（24時至次日4時）','每晚\n金額','直接主管\n審核簽章'];
    headers.forEach((text,col)=>line(6,...TABLE_COLS[col],text,col===4||col===5?'smallHeader':'header'));
    const labels=FORM_WORK_TYPES.map((label,i)=>(i===0?'■ ':'□ ')+(i===0?'三班制（2\n　輪1例隔週\n　休班制）':i===3?'工務夜間現\n　場實際工作':label));
    line(7,0,1,labels.join('\n'),'workType',25);
    for (let row=0; row<19; row++) {
      const night=days[row], time=night?NIGHT_TIMES[night.role]:null;
      const values=[night?`${schedule.month}月${night.day}日`:'月　日',time?.start??'',`次日${time?' '+time.end:''}`,night?1:'',night?1:'',night?NIGHT_AMOUNT:'',''];
      values.forEach((text,col)=>line(7+row,...TABLE_COLS[col+1],text,col===2?'endTime':col===1?'time':'cell'));
    }
    line(26,0,1,'','cell');line(26,2,6,pages.length>1?'本頁合計':'合計','totalLabel');
    line(26,7,8,`${days.length} 次`,'cell');line(26,9,10,`${days.length} 次`,'cell');
    line(26,11,11,days.length*NIGHT_AMOUNT,'totalMoney');line(26,12,13,'','cell');
    line(28,0,13,'註：','notes');
    FORM_NOTES.forEach((text,i)=>line(29+i,0,13,`${i+1}. ${text}`,'notes'));
    if (pages.length>1) line(33,0,13,`第 ${index+1}／${pages.length} 頁 · 全月 ${person.nights.length} 次，共 ${person.total.toLocaleString('zh-TW')} 元`,'continuation');
  });
  sheet['!ref']=`A1:N${pages.length*34}`;
  sheet['!claimPages']=pages.length;
  const base=person.name.replace(/[\\/:?*\[\]]/g,'_').replace(/^'+|'+$/g,'').slice(0,31)||'請領表';
  let name=base, suffix=1;
  while (name.toLowerCase()==='history' || workbook.SheetNames.some(n=>n.toLowerCase()===name.toLowerCase())) {
    const tail=` (${suffix++})`;name=base.slice(0,31-tail.length)+tail;
  }
  XLSX.utils.book_append_sheet(workbook,sheet,name);
  workbook.Workbook??={};workbook.Workbook.Names??=[];
  workbook.Workbook.Names.push({Name:'_xlnm.Print_Area',Sheet:workbook.SheetNames.length-1,Ref:`'${name.replace(/'/g,"''")}'!$A$1:$N$${pages.length*34}`});
}

// SheetJS CE builds the workbook; its bundled ZIP API adds the native Excel form styles and page setup.
// See https://github.com/SheetJS/js-cfb for the supported ZIP read/write interface.
export function serializeExportWorkbook(workbook) {
  const zip=XLSX.CFB.read(new Uint8Array(XLSX.write(workbook,{bookType:'xlsx',type:'array'})),{type:'buffer'});
  const read=path=>new TextDecoder().decode(XLSX.CFB.find(zip,path).content);
  const write=(path,content)=>XLSX.CFB.utils.cfb_add(zip,path,new TextEncoder().encode(content));
  let styles=read('/xl/styles.xml');
  // Fix the workbook's Normal font so column widths do not inherit the user's Excel default.
  styles=styles.replace(/(<fonts[^>]*>)<font>[\s\S]*?<\/font>/,'$1<font><sz val="11"/><name val="Calibri"/><family val="2"/><charset val="0"/></font>')
    .replace(/(<cellStyleXfs[^>]*>)<xf[^>]*\/>/,'$1<xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="1"/>');
  function extend(tag,items) {
    const match=styles.match(new RegExp(`<${tag} count="(\\d+)"[^>]*>([\\s\\S]*?)</${tag}>`));
    if (!match) throw new Error('無法建立Excel請領表樣式。');
    const first=Number(match[1]);
    styles=styles.replace(match[0],`<${tag} count="${first+items.length}">${match[2]}${items.join('')}</${tag}>`);
    return first;
  }
  const font=extend('fonts',[10,11,12,17].map(size=>`<font><sz val="${size}"/><name val="標楷體"/><family val="3"/><charset val="136"/></font>`).concat('<font><sz val="10"/><name val="Microsoft JhengHei"/><family val="2"/></font>'));
  const edge='<color rgb="FF000000"/>';
  const border=extend('borders',[
    `<border><left style="thin">${edge}</left><right style="thin">${edge}</right><top style="thin">${edge}</top><bottom style="thin">${edge}</bottom><diagonal/></border>`,
    `<border><left/><right/><top/><bottom style="thin">${edge}</bottom><diagonal/></border>`,
  ]);
  const numFormats=styles.match(/<numFmts count="(\d+)"[^>]*>([\s\S]*?)<\/numFmts>/);
  const moneyFormat=Math.max(164,...[...styles.matchAll(/numFmtId="(\d+)"/g)].map(m=>Number(m[1])))+1;
  const format=`<numFmt numFmtId="${moneyFormat}" formatCode="#,##0&quot;元&quot;"/>`;
  if (numFormats) styles=styles.replace(numFormats[0],`<numFmts count="${Number(numFormats[1])+1}">${numFormats[2]}${format}</numFmts>`);
  else styles=styles.replace('<fonts',`<numFmts count="1">${format}</numFmts><fonts`);
  const definitions={title:[3,0,'center','center'],unitLabel:[1,0,'left','center'],unitValue:[2,2,'center','center'],
    identity:[1,0,'left','center'],signature:[1,2,'left','center'],identityValue:[1,2,'center','center'],code:[4,2,'center','center'],
    header:[1,1,'center','center'],smallHeader:[0,1,'center','center'],cell:[1,1,'center','center'],
    time:[4,1,'center','center'],endTime:[4,1,'left','center'],workType:[0,1,'left','top'],
    totalLabel:[1,1,'right','center'],totalMoney:[1,1,'right','center',true],notes:[0,0,'left','top'],continuation:[0,0,'right','center']};
  const names=Object.keys(definitions);
  const firstStyle=extend('cellXfs',Object.values(definitions).map(([f,b,h,v,money])=>
    `<xf numFmtId="${money?moneyFormat:0}" fontId="${font+f}" fillId="0" borderId="${b?border+b-1:0}" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1" applyNumberFormat="1"><alignment horizontal="${h}" vertical="${v}" wrapText="1"/></xf>`));
  write('/xl/styles.xml',styles);
  workbook.SheetNames.forEach((name,index)=>{
    const sheet=workbook.Sheets[name];
    if (!sheet['!claimStyles']) return;
    const path=`/xl/worksheets/sheet${index+1}.xml`;
    let xml=read(path).replace(/<c r="([^"]+)"([^>]*)>/g,(tag,address,attributes)=>{
      const style=sheet['!claimStyles'][address];
      return style?`<c r="${address}"${attributes.replace(/ s="\d+"/,'')} s="${firstStyle+names.indexOf(style)}">`:tag;
    });
    xml=xml.replace(/(<worksheet[^>]*>)/,'$1<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>')
      .replace(/<sheetView /,'<sheetView showGridLines="0" ');
    const breaks=Array.from({length:sheet['!claimPages']-1},(_,i)=>`<brk id="${(i+1)*34}" min="0" max="16383" man="1"/>`).join('');
    const setup='<pageSetup paperSize="9" orientation="portrait" fitToWidth="1" fitToHeight="0"/>'+(breaks?`<rowBreaks count="${sheet['!claimPages']-1}" manualBreakCount="${sheet['!claimPages']-1}">${breaks}</rowBreaks>`:'');
    xml=xml.replace(/(<pageMargins[^>]*\/>)/,`$1${setup}`);
    write(path,xml);
  });
  return new Uint8Array(XLSX.CFB.write(zip,{fileType:'zip',compression:true}));
}
