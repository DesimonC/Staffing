const CONFIG = {
  spreadsheetId: '', // Leave blank when this script is bound to the Staffing Google Sheet.
  sheets: {
    colleagues: 'Colleague',
    rostering: 'Rostering',
    availability: 'Availability',
    regular: 'RegularShifts',
    absences: 'Absences',
    rotas: 'Rotas',
    assignments: 'RotaAssignments'
  }
};

const DAYS = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'];
const SHIFT_ROWS = [
  'Morning cook','Morning Till / Plate up','Lunch cover','Clear down helper',
  'Extra cover','Close down No.1','Close down No.2'
];
const SKILL_RULES = {
  'Morning cook': {all:['Morning Cook','Open up'], any:[]},
  'Morning Till / Plate up': {all:[], any:['Till','Plate up']},
  'Lunch cover': {all:[], any:['Till','Plate up','Wash up','Diningroom']},
  'Clear down helper': {all:[], any:['Wash up','Diningroom']},
  'Extra cover': {all:[], any:['Till','Plate up','Wash up','Diningroom']},
  'Close down No.1': {all:['Close down'], any:[]},
  'Close down No.2': {all:['Close down'], any:[]}
};

function doGet(e) {
  try {
    const action = String((e && e.parameter && e.parameter.action) || 'bootstrap');
    if (action === 'bootstrap') return json_(bootstrap_());
    if (action === 'health') return json_({ok:true, now:new Date().toISOString()});
    throw new Error('Unknown GET action: ' + action);
  } catch (err) {
    return json_({ok:false,error:String(err.message || err)});
  }
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.action === 'saveAbsence') return json_(saveAbsence_(body));
    if (body.action === 'saveRota') return json_(saveRota_(body));
    throw new Error('Unknown POST action: ' + body.action);
  } catch (err) {
    return json_({ok:false,error:String(err.message || err)});
  }
}

function bootstrap_() {
  ensureAppSheets_();
  const colleagues = readColleagues_();
  return {
    ok:true,
    meta:{version:Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Europe/London','yyyyMMddHHmmss'),source:'Google Sheets'},
    colleagues:colleagues,
    availability:readAvailability_(colleagues),
    shifts:readRostering_(),
    regular:readTable_(CONFIG.sheets.regular),
    absences:readTable_(CONFIG.sheets.absences)
  };
}

function ss_() {
  return CONFIG.spreadsheetId ? SpreadsheetApp.openById(CONFIG.spreadsheetId) : SpreadsheetApp.getActiveSpreadsheet();
}
function sheet_(name) {
  const s=ss_().getSheetByName(name);
  if(!s) throw new Error('Missing sheet: '+name);
  return s;
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function norm_(v){ return String(v == null ? '' : v).trim(); }
function bool_(v){ return v === true || /^(true|yes|y|1)$/i.test(norm_(v)); }
function time_(v) {
  if (v === '' || v == null) return '';
  if (Object.prototype.toString.call(v)==='[object Date]') return Utilities.formatDate(v, Session.getScriptTimeZone() || 'Europe/London','HH:mm');
  if (typeof v === 'number') {
    const mins=Math.round(v*24*60), h=Math.floor(mins/60)%24, m=mins%60;
    return String(h).padStart(2,'0')+':'+String(m).padStart(2,'0');
  }
  const s=norm_(v);
  if (/^\d{1,2}:\d{2}$/.test(s)) return s.padStart(5,'0');
  return s;
}
function date_(v) {
  if(!v) return '';
  if(Object.prototype.toString.call(v)==='[object Date]') return Utilities.formatDate(v, Session.getScriptTimeZone() || 'Europe/London','yyyy-MM-dd');
  return norm_(v);
}

function readColleagues_() {
  const s=sheet_(CONFIG.sheets.colleagues), vals=s.getDataRange().getValues(), h=vals[0].map(norm_);
  const ix=n=>h.indexOf(n);
  const skillNames=['Morning Cook','Evening Cook','Till','Plate up','Wash up','Diningroom','Open up','Close down'];
  return vals.slice(1).filter(r=>norm_(r[ix('Name')])).map(r=>{
    const till=norm_(r[ix('Till number')]);
    return {
      id:till,name:norm_(r[ix('Name')]),contract:Number(r[ix('Contract')]||0),
      supervisor:/^christina\s*$/i.test(norm_(r[ix('Name')])),
      skills:skillNames.filter(k=>ix(k)>=0 && bool_(r[ix(k)]))
    };
  }).filter(c=>c.id);
}

function readAvailability_(colleagues) {
  const s=sheet_(CONFIG.sheets.availability), v=s.getDataRange().getValues(), out={};
  const nameToId={}; colleagues.forEach(c=>nameToId[c.name.trim().toLowerCase()]=c.id);
  // Current sheet layout: row 1 day headings, row 2 From/Until, row 3+ colleague name in col A.
  const dayCols={Monday:[1,2],Tuesday:[3,4],Wednesday:[5,6],Thursday:[7,8],Friday:[9,10],Saturday:[11,12],Sunday:[13,14]};
  v.slice(2).forEach(r=>{
    const id=nameToId[norm_(r[0]).toLowerCase()]; if(!id) return;
    out[id]={};
    DAYS.forEach(d=>{const c=dayCols[d]; out[id][d]={from:time_(r[c[0]]),until:time_(r[c[1]])};});
  });
  return out;
}

function readRostering_() {
  const s=sheet_(CONFIG.sheets.rostering), v=s.getDataRange().getValues(), out=[];
  // Current Rostering layout: task names in col A; Monday B:F, Tue G:J, Wed K:N, Thu O:R, Fri S:V, Sat W:Z, Sun AA:AD.
  const starts={Monday:1,Tuesday:6,Wednesday:10,Thursday:14,Friday:18,Saturday:22,Sunday:26};
  for(let r=3;r<v.length;r++){
    const name=norm_(v[r][0]); if(!SHIFT_ROWS.includes(name)) continue;
    DAYS.forEach(day=>{
      const c=starts[day], start=time_(v[r][c]), finish=time_(v[r][c+1]);
      if(!start || !finish) return; // blank Extra cover etc. means not required.
      const rule=SKILL_RULES[name]||{all:[],any:[]};
      out.push({day:day,name:name,start:start,end:finish,skillsAll:rule.all,skillsAny:rule.any});
    });
  }
  return out;
}

function ensureAppSheets_() {
  ensureSheet_(CONFIG.sheets.regular,['RegularShiftID','TillNumber','Day','ShiftName','Active']);
  ensureSheet_(CONFIG.sheets.absences,['AbsenceID','TillNumber','Type','StartDate','EndDate','Status','Notes','Created']);
  ensureSheet_(CONFIG.sheets.rotas,['RotaID','WeekCommencing','Status','Created','CreatedBy','Published']);
  ensureSheet_(CONFIG.sheets.assignments,['RotaID','Date','Day','ShiftName','RegularTillNumber','AssignedTillNumber','BaseStart','BaseFinish','ActualStart','ActualFinish','ElapsedHours','BreakMinutes','PaidHours','HourlyRate','BudgetCost','AssignmentType','Notes']);
}
function ensureSheet_(name, headers) {
  let s=ss_().getSheetByName(name); if(!s) s=ss_().insertSheet(name);
  if(s.getLastRow()===0) s.getRange(1,1,1,headers.length).setValues([headers]);
}
function readTable_(name) {
  const s=sheet_(name), v=s.getDataRange().getValues(); if(v.length<2) return [];
  const h=v[0].map(norm_);
  return v.slice(1).filter(r=>r.some(x=>x!=='' && x!=null)).map(r=>{
    const o={}; h.forEach((k,i)=>o[k]=r[i]); return o;
  });
}

function saveAbsence_(b) {
  ensureAppSheets_();
  const id='ABS-'+Utilities.getUuid().slice(0,8).toUpperCase();
  const row=[id,norm_(b.tillNumber),norm_(b.type),date_(b.startDate),date_(b.endDate),norm_(b.status)||'Active',norm_(b.notes),new Date()];
  sheet_(CONFIG.sheets.absences).appendRow(row);
  return {ok:true,absenceId:id};
}

function saveRota_(b) {
  ensureAppSheets_();
  const rotaId=norm_(b.rotaId)||('R-'+norm_(b.weekCommencing).replace(/-/g,''));
  const rotas=sheet_(CONFIG.sheets.rotas), as=sheet_(CONFIG.sheets.assignments);
  rotas.appendRow([rotaId,date_(b.weekCommencing),norm_(b.status)||'DRAFT',new Date(),norm_(b.createdBy),'']);
  (b.assignments||[]).forEach(a=>{
    const elapsed=elapsedHours_(a.actualStart,a.actualFinish), breakMinutes=elapsed>6?15:0, paid=elapsed-(breakMinutes/60);
    const rate=Number(a.hourlyRate||0), cost=paid*rate;
    as.appendRow([rotaId,date_(a.date),norm_(a.day),norm_(a.shiftName),norm_(a.regularTillNumber),norm_(a.assignedTillNumber),time_(a.baseStart),time_(a.baseFinish),time_(a.actualStart),time_(a.actualFinish),elapsed,breakMinutes,paid,rate,cost,norm_(a.assignmentType)||'MANUAL',norm_(a.notes)]);
  });
  return {ok:true,rotaId:rotaId};
}
function elapsedHours_(a,b) {
  const m=t=>{const p=String(t).split(':').map(Number);return p[0]*60+p[1];};
  return (m(b)-m(a))/60;
}