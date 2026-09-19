/**
 * PrepOne V14.3.8 — Sheet-driven Exam Logos
 * Live Master Sheet: 1LC8YFDe89_yJc49BUkTAy9bq2eUVL3CgNk4pn4KAXyE
 *
 * Public: all active courses/content, no login required.
 * Student: password-protected personal dashboard + completion tracking + private topic notes.
 * Teacher: creates students, resets passwords, assigns courses.
 *
 * Teacher password is read from Settings → TeacherPassword in the verified master Sheet.
 * Change that Sheet cell any time; no Code.gs edit or redeploy is needed for password changes.
 */

let REQUEST_MASTER_SPREADSHEET_ID_ = '';

const CONFIG = Object.freeze({
  APP_NAME: 'PrepOne',
  VERSION: '14.3.8-sheet-driven-exam-logos',
  MASTER_SPREADSHEET_ID: '1LC8YFDe89_yJc49BUkTAy9bq2eUVL3CgNk4pn4KAXyE',
  CACHE_SECONDS: 90,
  SESSION_SECONDS: 21600,
  SHEETS: {
    EXAMS: 'Exams',
    SYLLABUS: 'Syllabus',
    STUDENTS: 'Students',
    STUDENT_PROGRESS: 'StudentProgress',
    STUDENT_NOTES: 'StudentNotes',
    SETTINGS: 'Settings',
    EXAM_UPDATES: 'ExamUpdates'
  }
});

function doGet(e) {
  const p = (e && e.parameter) ? e.parameter : {};
  const action = clean_(p.action || 'health').toLowerCase();
  // Public reads must stay fast. Schema maintenance is only needed for diagnostics/admin setup.
  if (action === 'health' || action === 'teachersetup') ensureSystemSheetsCached_();
  const callback = clean_(p.callback || '');
  try {
    let result;
    switch (action) {
      case 'health': result = health_(); break;
      case 'storageidentity': result = storageIdentity_(); break;
      case 'config': result = publicConfig_(); break;
      case 'schema': result = schema_(); break;
      case 'exams': result = getExams_(); break;
      case 'exam': result = getExam_(p.examId || p.exam || ''); break;
      case 'syllabus': result = getSyllabus_(p.examId || p.exam || ''); break;
      case 'course': result = getCourseBundle_(p.examId || p.exam || ''); break;
      case 'examupdates':
      case 'exam_updates':
      case 'updates': result = getExamUpdates_(p.examId || p.exam || '', p.limit || 4); break;

      case 'studentspublic': result = studentsPublic_(); break;
      case 'studentlogin': result = studentLogin_(p.studentId || '', p.passwordHash || ''); break;
      case 'studentsession': result = studentSession_(p.token || ''); break;
      case 'studentdashboard': result = studentDashboard_(p.token || ''); break;
      case 'studentcourseprogress': result = studentCourseProgress_(p.token || '', p.examId || p.exam || ''); break;
      case 'resourceprogress': result = resourceProgress_(p); break;
      case 'studentnotes': result = studentNotes_(p.token || ''); break;
      case 'studentsavenote': result = studentSaveNote_(p); break;
      case 'studentdeletenote': result = studentDeleteNote_(p); break;
      case 'studentlogout': result = logoutStudent_(p.token || ''); break;

      case 'teacherstatus': result = teacherStatus_(); break;
      case 'teachersetup': result = teacherSetup_(); break;
      case 'teacherlogin': result = teacherLogin_(p.passwordHash || ''); break;
      case 'teacherstudents': result = teacherStudents_(p.token || ''); break;
      case 'teachersavestudent': result = teacherSaveStudent_(p); break;
      case 'teacherdeletestudent': result = teacherDeleteStudent_(p); break;
      case 'teacherlogout': result = logoutTeacher_(p.token || ''); break;

      default:
        result = {ok:false,error:'Unknown action',action:action};
    }
    return output_(withBackendMeta_(result), callback);
  } catch (err) {
    return output_(withBackendMeta_({ok:false,error:err && err.message ? err.message : String(err)}), callback);
  }
}

function withBackendMeta_(payload){
  const out=(payload && typeof payload==='object' && !Array.isArray(payload)) ? payload : {ok:true,data:payload};
  const expected=getMasterSpreadsheetId_();
  const actual=clean_(REQUEST_MASTER_SPREADSHEET_ID_||expected);
  out.backendVersion=CONFIG.VERSION;
  out.configuredMasterSpreadsheetId=expected;
  out.masterSpreadsheetId=actual;
  out.actualSpreadsheetId=out.actualSpreadsheetId||actual;
  out.readSpreadsheetId=out.readSpreadsheetId||actual;
  out.writeSpreadsheetId=out.writeSpreadsheetId||actual;
  out.readWriteUnified=actual===expected;
  return out;
}

function getMasterSpreadsheetId_(){
  // ONE source of truth. Never fall back to Script Properties, active/bound Sheets,
  // or another spreadsheet ID.
  const id=clean_(CONFIG.MASTER_SPREADSHEET_ID);
  if(!id) throw new Error('No PrepOne master spreadsheet ID is configured.');
  return id;
}

function getSS_(){
  const expected=getMasterSpreadsheetId_();
  let ss;
  try{ss=SpreadsheetApp.openById(expected)}catch(err){
    throw new Error('Cannot open the configured PrepOne Master Sheet: '+expected+'. Check that the Apps Script execution account has access to it.');
  }
  const actual=clean_(ss.getId());
  REQUEST_MASTER_SPREADSHEET_ID_=actual;
  if(actual!==expected){
    throw new Error('Storage target mismatch. Expected '+expected+' but Apps Script opened '+actual+'.');
  }
  return ss;
}

function actualMasterSpreadsheetId_(){
  try{return clean_(getSS_().getId())}catch(_){return ''}
}

function assertUnifiedMaster_(){
  const expected=getMasterSpreadsheetId_(), ss=getSS_(), actual=clean_(ss.getId());
  if(actual!==expected) throw new Error('Master Sheet verification failed. Writes are blocked.');
  return actual;
}

function assertWriteSheet_(sh){
  if(!sh) throw new Error('Write sheet is unavailable.');
  const expected=getMasterSpreadsheetId_(), actual=clean_(sh.getParent().getId());
  if(actual!==expected) throw new Error('Write blocked: target spreadsheet is '+actual+', expected '+expected+'.');
  return actual;
}

function storageIdentity_(){
  const ss=getSS_(), actual=clean_(ss.getId()), expected=getMasterSpreadsheetId_();
  let serviceUrl='';
  try{serviceUrl=ScriptApp.getService().getUrl()||''}catch(_){}
  let scriptId='';
  try{scriptId=ScriptApp.getScriptId()||''}catch(_){}
  return {
    ok:actual===expected,
    version:CONFIG.VERSION,
    configuredSpreadsheetId:expected,
    actualSpreadsheetId:actual,
    readSpreadsheetId:actual,
    writeSpreadsheetId:actual,
    writeTargetVerified:actual===expected,
    serviceUrl:serviceUrl,
    scriptId:scriptId
  };
}

function ensureSystemSheetsCached_(){
  const cache=CacheService.getScriptCache(), key='schema-ready:v1426';
  if(cache.get(key)==='1')return;
  ensureSystemSheets_();
  try{cache.put(key,'1',21600)}catch(_){}
}
function ensureSystemSheets_(){
  const ss=getSS_();
  ensureSheetHeaders_(ss,CONFIG.SHEETS.STUDENTS,['StudentID','StudentName','PasswordHash','AssignedCourses','Active','CreatedAt','LastLogin']);
  ensureSheetHeaders_(ss,CONFIG.SHEETS.STUDENT_PROGRESS,['ProgressID','StudentID','ExamID','SyllabusID','Subject','Chapter','Topic','ResourceKey','ResourceLabel','ResourceType','Completed','CompletedAt','LastOpenedAt']);
  ensureSheetHeaders_(ss,CONFIG.SHEETS.STUDENT_NOTES,['NoteID','StudentID','ExamID','Subject','Chapter','Topic','NoteTitle','NoteText','UpdatedAt']);
  ensureSheetHeaders_(ss,CONFIG.SHEETS.SETTINGS,['Key','Value','Description']);
  ensureSheetHeaders_(ss,CONFIG.SHEETS.EXAM_UPDATES,['UpdateID','ExamID','Title','Summary','Content','PublishedAt','SourceName','SourceURL','Active']);
  ensureSettingRow_(ss,'TeacherPassword','', 'Teacher dashboard password in plain text. Change this cell any time; keep the Sheet private.');
  ensureSettingRow_(ss,'TeacherPasswordStorage','SETTINGS_SHEET','Teacher password is read only from Settings → TeacherPassword.');
  ensureSettingRow_(ss,'SchemaVersion','14.2.8','PrepOne V14.2.8 schema.');
  ensureExamDateColumns_(ss);
}
function ensureSettingRow_(ss,key,defaultValue,description){
  const sh=ss.getSheetByName(CONFIG.SHEETS.SETTINGS); if(!sh)return;
  const vals=sh.getDataRange().getValues(); if(!vals.length)return;
  const heads=vals[0].map(clean_),ix=indexMap_(heads); if(ix.Key===undefined||ix.Value===undefined)return;
  for(let i=1;i<vals.length;i++){
    if(clean_(vals[i][ix.Key])!==key)continue;
    let changed=false;
    if(defaultValue!=='' && clean_(vals[i][ix.Value])!==clean_(defaultValue)){vals[i][ix.Value]=defaultValue;changed=true;}
    if(ix.Description!==undefined && description && clean_(vals[i][ix.Description])!==clean_(description)){vals[i][ix.Description]=description;changed=true;}
    if(changed){sh.getRange(i+1,1,1,heads.length).setValues([vals[i]]);clearCache_();}
    return;
  }
  const row=heads.map(h=>h==='Key'?key:h==='Value'?defaultValue:h==='Description'?description:'');
  sh.appendRow(row); clearCache_();
}
function ensureExamDateColumns_(ss){
  const sh=ss.getSheetByName(CONFIG.SHEETS.EXAMS); if(!sh)return;
  const required=['ExamDateStart','ExamDateEnd','DateStatus','DateSource','AnnualApproxStart','AnnualApproxEnd'];
  let last=Math.max(1,sh.getLastColumn()), heads=sh.getRange(1,1,1,last).getDisplayValues()[0].map(clean_);
  const missing=required.filter(h=>heads.indexOf(h)<0);
  if(missing.length){sh.getRange(1,last+1,1,missing.length).setValues([missing]);last+=missing.length;heads=heads.concat(missing);}
  const defaults={
    AP_EAPCET_E:['2026-05-12','2026-05-18','Official 2026','https://cets.apsche.ap.gov.in/eapcet/EapcetHomepages/ImportantDates','05-12','05-18'],
    AP_EAPCET_AP:['2026-05-19','2026-05-20','Official 2026','https://cets.apsche.ap.gov.in/eapcet/EapcetHomepages/ImportantDates','05-19','05-20'],
    NEET_UG:['2026-05-03','2026-05-03','Official 2026','https://nta.ac.in/Download/Notice/Notice_20260502215114.pdf','05-03','05-03'],
    AP_POLYCET:['2026-04-25','2026-04-25','Official 2026','https://polycetap.nic.in/','04-25','04-25'],
    SSC_CHSL:['2026-07-01','2026-09-30','Tentative 2026 window','https://ssc.gov.in/api/attachment/uploads/masterData/ExamCalendar/Tentative_Calendar2026_27_08012026.pdf','07-01','09-30'],
    SSC_MTS:['2026-09-01','2026-11-30','Tentative 2026 window','https://ssc.gov.in/api/attachment/uploads/masterData/ExamCalendar/Tentative_Calendar2026_27_08012026.pdf','09-01','11-30']
  };
  const vals=sh.getDataRange().getDisplayValues(),ix=indexMap_(heads); let changed=false;
  for(let r=1;r<vals.length;r++){
    const id=clean_(vals[r][ix.ExamID]); if(!defaults[id])continue;
    required.forEach((h,j)=>{
      if(clean_(vals[r][ix[h]])===''){
        sh.getRange(r+1,ix[h]+1).setValue(defaults[id][j]);
        changed=true;
      }
    });
  }
  if(changed)clearCache_();
}
function ensureSheetHeaders_(ss,name,headers){
  let sh=ss.getSheetByName(name);
  if(!sh)sh=ss.insertSheet(name);
  if(sh.getLastRow()===0 || sh.getLastColumn()===0){sh.getRange(1,1,1,headers.length).setValues([headers]);sh.setFrozenRows(1);return;}
  const existing=sh.getRange(1,1,1,Math.max(sh.getLastColumn(),headers.length)).getDisplayValues()[0].map(clean_);
  if(existing.filter(Boolean).length===0){sh.getRange(1,1,1,headers.length).setValues([headers]);sh.setFrozenRows(1);}
}

function requireSheet_(name){ const sh=getSS_().getSheetByName(name); if(!sh) throw new Error('Missing sheet: '+name); return sh; }
function clean_(v){ return String(v === null || v === undefined ? '' : v).trim(); }
function num_(v,f){ const n=Number(v); return isFinite(n)?n:(f===undefined?0:f); }
function enabled_(v){ const s=clean_(v).toLowerCase(); return s==='' || !['false','0','no','off','inactive'].includes(s); }
function bool_(v){ return ['true','1','yes','on'].includes(clean_(v).toLowerCase()); }
function safeHash_(v){ const s=clean_(v).toLowerCase(); if(!/^[a-f0-9]{64}$/.test(s)) throw new Error('Invalid password hash'); return s; }
function sha256Hex_(text){ const bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(text), Utilities.Charset.UTF_8); return bytes.map(b=>(b<0?b+256:b).toString(16).padStart(2,'0')).join(''); }
function credentialPepper_(){ const p=PropertiesService.getScriptProperties(); let v=p.getProperty('STUDENT_CREDENTIAL_PEPPER'); if(!v){v=Utilities.getUuid()+Utilities.getUuid();p.setProperty('STUDENT_CREDENTIAL_PEPPER',v)} return v; }
function studentCredential_(studentId,clientHash){ return sha256Hex_(clean_(studentId)+'|'+safeHash_(clientHash)+'|'+credentialPepper_()); }
function activeSyllabusExamIds_(){
  const ids=new Set();
  rowsAsObjects_(CONFIG.SHEETS.SYLLABUS).filter(r=>enabled_(r.Active)).forEach(r=>{const id=clean_(r.ExamID);if(id)ids.add(id)});
  return ids;
}
function availableExamRows_(){
  const contentIds=activeSyllabusExamIds_();
  return rowsAsObjects_(CONFIG.SHEETS.EXAMS)
    .filter(r=>enabled_(r.Active))
    .filter(r=>{const id=clean_(r.ExamID);return id && contentIds.has(id)});
}
function availableExamIds_(){ return availableExamRows_().map(r=>clean_(r.ExamID)).filter(Boolean); }
function ensureExamAllowed_(id){
  id=clean_(id);
  if(!id) throw new Error('examId is required');
  if(availableExamIds_().indexOf(id)<0) throw new Error('Exam is inactive or has no active syllabus content: '+id);
  return id;
}
function indexMap_(headers){ const x={}; headers.forEach((h,i)=>x[h]=i); return x; }
function settingValue_(key){
  const sh=requireSheet_(CONFIG.SHEETS.SETTINGS), vals=sh.getDataRange().getDisplayValues(); if(!vals.length)return '';
  const heads=vals[0].map(clean_),ix=indexMap_(heads); if(ix.Key===undefined||ix.Value===undefined)return '';
  for(let i=1;i<vals.length;i++)if(clean_(vals[i][ix.Key])===clean_(key))return clean_(vals[i][ix.Value]);
  return '';
}
function parseDateOnly_(value){
  const s=clean_(value); if(!s)return null; let m=s.match(/^(\d{4})-(\d{2})-(\d{2})$/); if(m)return new Date(Number(m[1]),Number(m[2])-1,Number(m[3]),12,0,0,0);
  const d=new Date(s); return isNaN(d.getTime())?null:new Date(d.getFullYear(),d.getMonth(),d.getDate(),12,0,0,0);
}
function monthDayFromDate_(d){return d?String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0'):'';}
function dateFromMonthDay_(year,md){const m=clean_(md).match(/^(\d{2})-(\d{2})$/);if(!m)return null;return new Date(year,Number(m[1])-1,Number(m[2]),12,0,0,0);}
function indiaToday_(){const s=Utilities.formatDate(new Date(),'Asia/Kolkata','yyyy-MM-dd');return parseDateOnly_(s);}
function isoDateOnly_(d){return d?Utilities.formatDate(d,'Asia/Kolkata','yyyy-MM-dd'):'';}
function humanDate_(d){return d?Utilities.formatDate(d,'Asia/Kolkata','dd MMM yyyy'):'';}
function resolveExamDate_(exam){
  const today=indiaToday_(), dayMs=86400000; let start=parseDateOnly_(exam.ExamDateStart||exam.ExamDate), end=parseDateOnly_(exam.ExamDateEnd||exam.ExamDateStart||exam.ExamDate);
  const annualStart=clean_(exam.AnnualApproxStart)||(start?monthDayFromDate_(start):''), annualEnd=clean_(exam.AnnualApproxEnd)||(end?monthDayFromDate_(end):annualStart);
  let status=clean_(exam.DateStatus)||'Scheduled', source=normalizeUrl_(exam.DateSource||exam.OfficialURL||''), approximate=false;
  if(start&&!end)end=start;
  if(!start || (end && end<today)){
    let year=today.getFullYear(), aStart=dateFromMonthDay_(year,annualStart), aEnd=dateFromMonthDay_(year,annualEnd||annualStart);
    if(aStart&&aEnd&&aEnd<aStart)aEnd=dateFromMonthDay_(year+1,annualEnd||annualStart);
    if(aEnd&&aEnd<today){year++;aStart=dateFromMonthDay_(year,annualStart);aEnd=dateFromMonthDay_(year,annualEnd||annualStart);if(aStart&&aEnd&&aEnd<aStart)aEnd=dateFromMonthDay_(year+1,annualEnd||annualStart);}
    if(aStart){start=aStart;end=aEnd||aStart;approximate=true;status='Approximate next cycle';}
  }
  if(!start)return {ResolvedExamDateStart:'',ResolvedExamDateEnd:'',ResolvedDateStatus:status,ResolvedDateSource:source,ResolvedDateApproximate:false,DaysLeft:null,CountdownState:'unknown',CountdownLabel:'Date not available',ExamDateLabel:'Date not available'};
  if(!end)end=start;
  const same=isoDateOnly_(start)===isoDateOnly_(end); let days=0,state='',countdown='';
  if(today<start){days=Math.max(0,Math.ceil((start-today)/dayMs));state='upcoming';countdown=days===0?'Exam today':days+' day'+(days===1?'':'s')+' left';}
  else if(today<=end){days=Math.max(0,Math.ceil((end-today)/dayMs));state=same?'today':'active-window';countdown=same?'Exam today':days+' day'+(days===1?'':'s')+' left in window';}
  else {state='passed';countdown='Date passed';}
  const dateLabel=same?humanDate_(start):(humanDate_(start)+' – '+humanDate_(end));
  return {ResolvedExamDateStart:isoDateOnly_(start),ResolvedExamDateEnd:isoDateOnly_(end),ResolvedDateStatus:status,ResolvedDateSource:source,ResolvedDateApproximate:approximate,DaysLeft:days,CountdownState:state,CountdownLabel:countdown,ExamDateLabel:(approximate?'Approx. ':'')+dateLabel};
}
function decorateExam_(exam){return Object.assign({},exam,resolveExamDate_(exam));}

function rowsAsObjects_(sheetName){
  const cache=CacheService.getScriptCache(), key='rows:'+sheetName, cached=cache.get(key);
  if(cached){ try{return JSON.parse(cached)}catch(_){} }
  const sh=requireSheet_(sheetName), vals=sh.getDataRange().getDisplayValues();
  if(!vals.length) return [];
  const heads=vals[0].map(clean_), rows=[];
  for(let r=1;r<vals.length;r++){
    if(vals[r].every(v=>clean_(v)==='')) continue;
    const o={}; heads.forEach((h,i)=>{if(h)o[h]=vals[r][i]!==undefined?vals[r][i]:''}); rows.push(o);
  }
  try{cache.put(key,JSON.stringify(rows),CONFIG.CACHE_SECONDS)}catch(_){}
  return rows;
}
function clearCache_(){ const c=CacheService.getScriptCache(); Object.keys(CONFIG.SHEETS).forEach(k=>{try{c.remove('rows:'+CONFIG.SHEETS[k])}catch(_){}}); }

function health_(){
  const ss=getSS_(), actual=clean_(ss.getId()), expected=getMasterSpreadsheetId_();
  const existing=ss.getSheets().map(s=>s.getName()), required=Object.values(CONFIG.SHEETS), missing=required.filter(x=>existing.indexOf(x)<0);
  const examRows=existing.includes(CONFIG.SHEETS.EXAMS)?rowsAsObjects_(CONFIG.SHEETS.EXAMS):[];
  const syllabusRows=existing.includes(CONFIG.SHEETS.SYLLABUS)?rowsAsObjects_(CONFIG.SHEETS.SYLLABUS):[];
  const available=missing.length?[]:availableExamIds_();
  const students=existing.includes(CONFIG.SHEETS.STUDENTS)?rowsAsObjects_(CONFIG.SHEETS.STUDENTS).filter(r=>enabled_(r.Active)).length:0;
  let serviceUrl=''; try{serviceUrl=ScriptApp.getService().getUrl()||''}catch(_){}
  return {ok:missing.length===0 && actual===expected,app:CONFIG.APP_NAME,version:CONFIG.VERSION,spreadsheetConfigured:true,spreadsheetName:ss.getName(),spreadsheetId:actual,configuredSpreadsheetId:expected,actualSpreadsheetId:actual,storageMode:'single-master-sheet-strict',readWriteUnified:actual===expected,readSpreadsheetId:actual,writeSpreadsheetId:actual,writeTargetVerified:actual===expected,serviceUrl:serviceUrl,studentDataSheets:[CONFIG.SHEETS.STUDENTS,CONFIG.SHEETS.STUDENT_PROGRESS,CONFIG.SHEETS.STUDENT_NOTES],missingSheets:missing,counts:{exams:examRows.length,publicExams:available.length,syllabus:syllabusRows.length,students},rawExamIds:examRows.filter(r=>enabled_(r.Active)).map(r=>clean_(r.ExamID)).filter(Boolean),availableExamIds:available,teacherConfigured:teacherConfigured_(),serverTime:new Date().toISOString()};
}
function publicConfig_(){ return {ok:true,app:CONFIG.APP_NAME,version:CONFIG.VERSION,enabledExams:availableExamIds_(),publicCourses:true,studentTracking:true,dynamicResources:true,examUpdates:true,dynamicExams:true,teacherPasswordInSheet:true,examCountdown:true,themeToggle:true}; }
function schema_(){
  const sh=requireSheet_(CONFIG.SHEETS.SYLLABUS), heads=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0].map(clean_);
  return {ok:true,syllabusHeaders:heads,studentHeaders:requireSheet_(CONFIG.SHEETS.STUDENTS).getRange(1,1,1,requireSheet_(CONFIG.SHEETS.STUDENTS).getLastColumn()).getDisplayValues()[0]};
}

// ---------------- PUBLIC COURSE CONTENT ----------------
function getExams_(){
  scheduleAutoRowGrouping_();
  const rows=availableExamRows_();
  rows.sort((a,b)=>num_(a.DisplayOrder,999)-num_(b.DisplayOrder,999)||clean_(a.ExamName).localeCompare(clean_(b.ExamName)));
  return {ok:true,exams:rows.map(decorateExam_)};
}
function getExam_(examId){
  const id=ensureExamAllowed_(examId), exam=rowsAsObjects_(CONFIG.SHEETS.EXAMS).find(r=>clean_(r.ExamID)===id && enabled_(r.Active));
  return exam?{ok:true,exam:decorateExam_(exam)}:{ok:false,error:'Exam not found',examId:id};
}
function getSyllabus_(examId){
  const id=ensureExamAllowed_(examId);
  const rows=rowsAsObjects_(CONFIG.SHEETS.SYLLABUS).filter(r=>clean_(r.ExamID)===id && enabled_(r.Active)).sort((a,b)=>num_(a.UnitOrder,999)-num_(b.UnitOrder,999)||num_(a.ChapterOrder,999)-num_(b.ChapterOrder,999)||num_(a.TopicOrder,999)-num_(b.TopicOrder,999));
  const optionalResources=optionalResourceRows_();
  return {ok:true,examId:id,syllabus:rows.map(r=>normalizeTopic_(r,optionalResources)),count:rows.length};
}
function optionalResourceRows_(){
  try{
    const ss=getSS_(),sh=ss.getSheetByName('Resources');
    if(!sh)return [];
    return rowsAsObjects_('Resources').filter(r=>enabled_(r.Active));
  }catch(_){return []}
}
function getCourseBundle_(examId){
  const e=getExam_(examId); if(!e.ok)return e; const s=getSyllabus_(examId); return {ok:true,exam:e.exam,syllabus:s.syllabus,summary:buildCourseSummary_(s.syllabus)};
}
function normalizeTopic_(row,optionalResources=[]){
  const links=extractTopicResources_(row,optionalResources);
  return {...row,ResourceLinks:links,HasMCQ:links.some(x=>x.Type==='MCQ')};
}
function extractTopicResources_(row,optionalResources=[]){
  const metadata=new Set(['SyllabusID','ExamID','Subject','Chapter','Topic','UnitOrder','ChapterOrder','TopicOrder','Weightage','SourceURL','Active','Description','Notes']);
  const out=[],seen=new Set();
  Object.keys(row||{}).forEach((label,i)=>{
    if(metadata.has(label))return;
    const u=normalizeUrl_(row[label]);
    if(!u || !/^https?:\/\//i.test(u))return;
    const type=/^mcq(?: quiz|url)?$/i.test(label)||/quiz|mcq|test/i.test(label)?'MCQ':resourceType_(label,u);
    const key=/^mcqurl$/i.test(label)?'MCQ Quiz':label;
    const d=(key+'|'+u).toLowerCase(); if(seen.has(d))return; seen.add(d);
    out.push({Key:key,Label:key,URL:u,Type:type,Order:i});
  });
  (optionalResources||[]).filter(r=>clean_(r.SyllabusID)===clean_(row.SyllabusID) || (clean_(r.ExamID)===clean_(row.ExamID)&&clean_(r.Subject)===clean_(row.Subject)&&clean_(r.Chapter)===clean_(row.Chapter)&&clean_(r.Topic)===clean_(row.Topic))).forEach((r,i)=>{
    const u=normalizeUrl_(r.ResourceURL||r.URL||r.Link); if(!u || !/^https?:\/\//i.test(u))return;
    const label=clean_(r.ResourceTitle||r.Title||r.Label||r.ResourceType||'Resource '+(i+1));
    const key=clean_(r.ResourceID)||label, d=(key+'|'+u).toLowerCase(); if(seen.has(d))return; seen.add(d);
    out.push({Key:key,Label:label,URL:u,Type:clean_(r.ResourceType)||resourceType_(label,u),Order:1000+num_(r.ResourceOrder,i)});
  });
  out.sort((a,b)=>{const am=a.Type==='MCQ',bm=b.Type==='MCQ';if(am!==bm)return am?1:-1;return num_(a.Order,9999)-num_(b.Order,9999)});
  return out;
}
function normalizeUrl_(v){ let u=clean_(v); if(/^http:\/\/script\.google\.com\//i.test(u))u='https://'+u.slice(7); return u; }
function resourceType_(label,url){ const s=(clean_(label)+' '+clean_(url)).toLowerCase(); if(/mcq|quiz|test/.test(s))return'MCQ'; if(/youtube|youtu\.be/.test(s))return'YOUTUBE'; if(/video|\.mp4|\.webm|\.m4v/.test(s))return'VIDEO'; if(/pdf|\.pdf/.test(s))return'PDF'; if(/audio|\.mp3|\.wav|\.m4a|\.ogg|\.aac/.test(s))return'AUDIO'; if(/image|photo|diagram|\.png|\.jpe?g|\.webp|\.gif/.test(s))return'IMAGE'; return'WEBSITE'; }
function buildCourseSummary_(syllabus){
  const subjects={}; let resources=0;
  syllabus.forEach(t=>{const s=t.Subject||'General',c=t.Chapter||'General'; if(!subjects[s])subjects[s]={topicCount:0,resourceCount:0,chapters:{}}; if(!subjects[s].chapters[c])subjects[s].chapters[c]={topicCount:0,resourceCount:0}; const n=(t.ResourceLinks||[]).length; resources+=n; subjects[s].topicCount++;subjects[s].resourceCount+=n;subjects[s].chapters[c].topicCount++;subjects[s].chapters[c].resourceCount+=n;});
  return {subjectCount:Object.keys(subjects).length,topicCount:syllabus.length,resourceCount:resources,subjects};
}


// ---------------- EXAM UPDATES ----------------
function examUpdateQueries_(){
  return {
    AP_EAPCET_E:'"AP EAPCET" engineering APSCHE',
    AP_EAPCET_AP:'"AP EAPCET" agriculture pharmacy APSCHE',
    NEET_UG:'"NEET UG" NTA',
    AP_POLYCET:'"AP POLYCET" SBTET',
    SSC_CHSL:'"SSC CHSL"',
    SSC_MTS:'"SSC MTS" OR "SSC MTS Havaldar"'
  };
}
function examUpdateQuery_(examId){
  const exam=rowsAsObjects_(CONFIG.SHEETS.EXAMS).find(r=>clean_(r.ExamID)===clean_(examId));
  if(!exam)return '';
  const custom=clean_(exam.NewsQuery||exam.UpdateQuery||exam.NewsSearchQuery);
  if(custom)return custom;
  const override=examUpdateQueries_()[examId];
  if(override)return override;
  const name=clean_(exam.ExamName)||clean_(exam.ShortName)||clean_(examId).replace(/_/g,' ');
  const shortName=clean_(exam.ShortName);
  return shortName && name.toLowerCase().indexOf(shortName.toLowerCase())<0
    ? '"'+name+'" OR "'+shortName+'"'
    : '"'+name+'"';
}
function examUpdateQueryCandidates_(examId){
  const base=examUpdateQuery_(examId); if(!base)return [];
  return [
    base+' (notification OR application OR registration OR "exam date" OR schedule OR "admit card" OR result)',
    base+' (latest OR notification OR exam OR recruitment)',
    base
  ];
}
function getExamUpdates_(examId,limit){
  const wanted=clean_(examId); if(wanted)ensureExamAllowed_(wanted);
  const max=Math.max(1,Math.min(8,num_(limit,4)));
  const ids=wanted?[wanted]:availableExamIds_(), grouped={}; ids.forEach(id=>grouped[id]=[]);
  const cache=CacheService.getScriptCache(), key='examupdates:v1422:'+(wanted||ids.join(','))+':'+max, cached=cache.get(key);
  if(cached){try{return JSON.parse(cached)}catch(_){} }

  // Curated updates come ONLY from the configured V14.2 master Sheet.
  rowsAsObjects_(CONFIG.SHEETS.EXAM_UPDATES).filter(r=>enabled_(r.Active)).forEach(r=>{
    const id=clean_(r.ExamID); if(!Object.prototype.hasOwnProperty.call(grouped,id))return;
    grouped[id].push(normalizeExamUpdate_(r,'curated'));
  });

  // Add current/relevant news from Google News RSS. Multiple query levels improve coverage.
  let newsByExam={};
  try{newsByExam=fetchNewsUpdatesBatch_(ids)}catch(_){ }

  ids.forEach(id=>{
    (newsByExam[id]||[]).forEach(u=>grouped[id].push(u));
    const seen=new Set();
    grouped[id]=grouped[id]
      .filter(u=>u&&u.Title&&(u.URL||u.Content||u.Summary))
      .sort((a,b)=>new Date(b.PublishedAt||0)-new Date(a.PublishedAt||0))
      .filter(u=>{const k=(u.Title+'|'+(u.URL||u.Content||u.Summary||'')).toLowerCase();if(seen.has(k))return false;seen.add(k);return true})
      .slice(0,max);

    // Never leave a course card empty. This is intentionally phrased as a link to
    // the latest official information, not as a fabricated dated news event.
    if(!grouped[id].length){
      const fallback=officialLatestUpdate_(id);
      if(fallback)grouped[id]=[fallback];
    }
  });

  const result={ok:true,updates:grouped,generatedAt:new Date().toISOString(),sourceSpreadsheetId:assertUnifiedMaster_()};
  try{cache.put(key,JSON.stringify(result),300)}catch(_){ }
  return result;
}
function normalizeExamUpdate_(r,source){
  const published=r.PublishedAt?new Date(r.PublishedAt):new Date();
  return {ID:clean_(r.UpdateID)||Utilities.getUuid(),ExamID:clean_(r.ExamID),Title:clean_(r.Title),Summary:clean_(r.Summary),Content:clean_(r.Content)||clean_(r.Summary),PublishedAt:isNaN(published.getTime())?new Date().toISOString():published.toISOString(),SourceName:clean_(r.SourceName)||'Official update',URL:normalizeUrl_(r.SourceURL||r.URL),Kind:source||'curated'};
}
function officialLatestUpdate_(examId){
  const exam=rowsAsObjects_(CONFIG.SHEETS.EXAMS).find(r=>clean_(r.ExamID)===clean_(examId)&&enabled_(r.Active));
  if(!exam)return null;
  const name=clean_(exam.ExamName)||clean_(exam.ShortName)||clean_(examId).replace(/_/g,' ');
  const url=normalizeUrl_(exam.OfficialURL||exam.SourceURL||'');
  return {
    ID:'OFFICIAL-LATEST-'+clean_(examId),ExamID:clean_(examId),
    Title:'Latest official '+name+' updates',
    Summary:'Open the official '+name+' source for the latest notification, application, schedule, admit card and result information.',
    Content:'Official update source for '+name+'. Use the source link for the newest published information.',
    PublishedAt:new Date().toISOString(),SourceName:'Official website',URL:url,Kind:'official-latest'
  };
}
function parseNewsUpdatesResponse_(examId,res){
  if(!res||res.getResponseCode()<200||res.getResponseCode()>=300)return [];
  const doc=XmlService.parse(res.getContentText()), channel=doc.getRootElement().getChild('channel'); if(!channel)return [];
  return channel.getChildren('item').slice(0,10).map((item,i)=>{
    const text=n=>{const x=item.getChild(n);return x?clean_(x.getText()):''};
    const rawTitle=text('title'), parts=rawTitle.split(' - '), source=parts.length>1?parts.pop():'News';
    const title=parts.join(' - ')||rawTitle, link=text('link'), pub=new Date(text('pubDate')), desc=stripHtml_(text('description'));
    return {ID:'RSS-'+examId+'-'+i+'-'+Utilities.getUuid().slice(0,6),ExamID:examId,Title:title,Summary:desc.slice(0,280),Content:desc,PublishedAt:isNaN(pub.getTime())?'':pub.toISOString(),SourceName:source,URL:link,Kind:'news'};
  }).filter(x=>x.Title&&x.URL);
}
function fetchNewsUpdatesBatch_(examIds){
  const ids=(examIds||[]).map(clean_).filter(Boolean), out={}; ids.forEach(id=>out[id]=[]);
  const requests=[], requestMeta=[];
  ids.forEach(id=>{
    examUpdateQueryCandidates_(id).forEach((q,priority)=>{
      requests.push({url:'https://news.google.com/rss/search?q='+encodeURIComponent(q)+'&hl=en-IN&gl=IN&ceid=IN:en',muteHttpExceptions:true,followRedirects:true,headers:{'User-Agent':'Mozilla/5.0 PrepOne/14.2'}});
      requestMeta.push({id:id,priority:priority});
    });
  });
  if(!requests.length)return out;
  const responses=UrlFetchApp.fetchAll(requests);
  responses.forEach((res,i)=>{
    const meta=requestMeta[i];
    try{
      const rows=parseNewsUpdatesResponse_(meta.id,res);
      rows.forEach(r=>{r.SearchPriority=meta.priority;out[meta.id].push(r)});
    }catch(_){ }
  });
  ids.forEach(id=>{
    const seen=new Set();
    out[id]=out[id]
      .sort((a,b)=>(num_(a.SearchPriority,9)-num_(b.SearchPriority,9))||(new Date(b.PublishedAt||0)-new Date(a.PublishedAt||0)))
      .filter(x=>{const k=(clean_(x.Title)+'|'+clean_(x.URL)).toLowerCase();if(seen.has(k))return false;seen.add(k);return true})
      .slice(0,8);
  });
  return out;
}
function fetchNewsUpdates_(examId){ return (fetchNewsUpdatesBatch_([examId])[clean_(examId)]||[]); }
function stripHtml_(s){return clean_(String(s||'').replace(/<[^>]*>/g,' ').replace(/&nbsp;/g,' ').replace(/&amp;/g,'&').replace(/&quot;/g,'"').replace(/&#39;/g,"'").replace(/\s+/g,' '));}

// ---------------- STUDENTS ----------------
function studentsPublic_(){
  const rows=rowsAsObjects_(CONFIG.SHEETS.STUDENTS).filter(r=>enabled_(r.Active)).map(r=>({StudentID:clean_(r.StudentID),StudentName:clean_(r.StudentName)})).filter(r=>r.StudentID&&r.StudentName).sort((a,b)=>a.StudentName.localeCompare(b.StudentName));
  return {ok:true,students:rows};
}
function studentLogin_(studentId,passwordHash){
  assertUnifiedMaster_();
  const id=clean_(studentId), hash=safeHash_(passwordHash); if(!id)throw new Error('Select a student'); loginGuard_('student:'+id);
  const sh=requireSheet_(CONFIG.SHEETS.STUDENTS); assertWriteSheet_(sh); const vals=sh.getDataRange().getValues(), heads=vals[0].map(clean_), ix=indexMap_(heads);
  for(let i=1;i<vals.length;i++){
    if(clean_(vals[i][ix.StudentID])===id && enabled_(vals[i][ix.Active])){
      if(clean_(vals[i][ix.PasswordHash]).toLowerCase()!==studentCredential_(id,hash)){recordLoginFailure_('student:'+id);return {ok:false,error:'Incorrect password'}};
      clearLoginFailures_('student:'+id);
      if(ix.LastLogin!==undefined){vals[i][ix.LastLogin]=new Date();sh.getRange(i+1,1,1,heads.length).setValues([vals[i]])}
      clearCache_(); const token=createSession_('student',id); return {ok:true,token:token,student:studentSafe_(objectFromRow_(heads,vals[i]))};
    }
  }
  return {ok:false,error:'Student not found'};
}
function studentSession_(token){ const id=requireSession_('student',token); const s=findStudent_(id); return {ok:true,student:studentSafe_(s)}; }
function logoutStudent_(token){ deleteSession_('student',token); return {ok:true}; }
function findStudent_(studentId){ const s=rowsAsObjects_(CONFIG.SHEETS.STUDENTS).find(r=>clean_(r.StudentID)===clean_(studentId)&&enabled_(r.Active)); if(!s)throw new Error('Student account is inactive or missing'); return s; }
function studentSafe_(s){ return {StudentID:clean_(s.StudentID),StudentName:clean_(s.StudentName),AssignedCourses:courseIds_(s.AssignedCourses),Active:enabled_(s.Active),CreatedAt:s.CreatedAt||'',LastLogin:s.LastLogin||''}; }
function courseIds_(csv){ const allowed=new Set(availableExamIds_()); return clean_(csv).split(',').map(clean_).filter(x=>allowed.has(x)); }
function assertAssigned_(student,examId){ const id=ensureExamAllowed_(examId), ids=courseIds_(student.AssignedCourses); if(ids.indexOf(id)<0)throw new Error('This course is not assigned to the student'); return id; }
function studentCourseProgress_(token,examId){ const sid=requireSession_('student',token), student=findStudent_(sid), id=assertAssigned_(student,examId); return {ok:true,student:studentSafe_(student),examId:id,progress:progressRows_(sid,id)}; }
function studentDashboard_(token){
  const sid=requireSession_('student',token), student=findStudent_(sid), assigned=courseIds_(student.AssignedCourses), exams=getExams_().exams.filter(e=>assigned.includes(clean_(e.ExamID))), allProgress=progressRows_(sid,'');
  const courses=exams.map(e=>{
    const sy=getSyllabus_(e.ExamID).syllabus, progress=allProgress.filter(p=>clean_(p.ExamID)===e.ExamID), completedMap=completedMap_(progress); let totalRes=0,doneRes=0,totalTopics=0,doneTopics=0;
    sy.forEach(t=>{const links=t.ResourceLinks||[]; if(!links.length)return; totalTopics++; let td=0; links.forEach(r=>{totalRes++; if(completedMap[progressKey_(t.SyllabusID,r.Key)]){doneRes++;td++;}}); if(td===links.length)doneTopics++;});
    return {ExamID:e.ExamID,ExamName:e.ExamName,ShortName:e.ShortName,Category:e.Category,LogoURL:e.LogoURL||'',OfficialURL:e.OfficialURL||'',ResolvedExamDateStart:e.ResolvedExamDateStart||'',ResolvedExamDateEnd:e.ResolvedExamDateEnd||'',ResolvedDateStatus:e.ResolvedDateStatus||'',ResolvedDateSource:e.ResolvedDateSource||'',ResolvedDateApproximate:e.ResolvedDateApproximate===true,DaysLeft:e.DaysLeft,CountdownState:e.CountdownState||'',CountdownLabel:e.CountdownLabel||'',ExamDateLabel:e.ExamDateLabel||'',totalResources:totalRes,completedResources:doneRes,totalTopics:totalTopics,completedTopics:doneTopics,percent:totalRes?Math.round(doneRes/totalRes*100):0};
  });
  const recent=allProgress.filter(p=>clean_(p.LastOpenedAt)).sort((a,b)=>new Date(b.LastOpenedAt)-new Date(a.LastOpenedAt)).slice(0,8);
  const totals=courses.reduce((a,c)=>({resources:a.resources+c.totalResources,completed:a.completed+c.completedResources,topics:a.topics+c.totalTopics,topicsDone:a.topicsDone+c.completedTopics}),{resources:0,completed:0,topics:0,topicsDone:0});
  return {ok:true,student:studentSafe_(student),courses:courses,recent:recent,totals:{...totals,percent:totals.resources?Math.round(totals.completed/totals.resources*100):0}};
}
function progressRows_(studentId,examId){ let rows=rowsAsObjects_(CONFIG.SHEETS.STUDENT_PROGRESS).filter(r=>clean_(r.StudentID)===clean_(studentId)); if(examId)rows=rows.filter(r=>clean_(r.ExamID)===clean_(examId)); return rows; }
function completedMap_(rows){ const m={}; rows.forEach(r=>{if(bool_(r.Completed))m[progressKey_(r.SyllabusID,r.ResourceKey)]=true}); return m; }
function progressKey_(syllabusId,resourceKey){ return clean_(syllabusId)+'::'+clean_(resourceKey); }
function resourceProgress_(p){
  assertUnifiedMaster_();
  const sid=requireSession_('student',p.token||''), student=findStudent_(sid), examId=assertAssigned_(student,p.examId||p.exam||''), syllabusId=clean_(p.syllabusId), resourceKey=clean_(p.resourceKey), completed=clean_(p.completed); if(!syllabusId||!resourceKey)throw new Error('Missing resource identity');
  const topic=getSyllabus_(examId).syllabus.find(t=>clean_(t.SyllabusID)===syllabusId); if(!topic)throw new Error('Topic not found'); const resource=(topic.ResourceLinks||[]).find(r=>clean_(r.Key)===resourceKey); if(!resource)throw new Error('Resource not found');
  const lock=LockService.getScriptLock();lock.waitLock(5000);
  try{
    const sh=requireSheet_(CONFIG.SHEETS.STUDENT_PROGRESS); assertWriteSheet_(sh); const vals=sh.getDataRange().getValues(), heads=vals[0].map(clean_), ix=indexMap_(heads); let row=-1;
    for(let i=1;i<vals.length;i++){if(clean_(vals[i][ix.StudentID])===sid&&clean_(vals[i][ix.ExamID])===examId&&clean_(vals[i][ix.SyllabusID])===syllabusId&&clean_(vals[i][ix.ResourceKey])===resourceKey){row=i;break;}}
    const now=new Date(), existing=row>=0?objectFromRow_(heads,vals[row]):{}; const isComplete=completed===''?bool_(existing.Completed):bool_(completed);
    const obj={ProgressID:existing.ProgressID||Utilities.getUuid(),StudentID:sid,ExamID:examId,SyllabusID:syllabusId,Subject:topic.Subject,Chapter:topic.Chapter,Topic:topic.Topic,ResourceKey:resourceKey,ResourceLabel:resource.Label,ResourceType:resource.Type,Completed:isComplete,CompletedAt:isComplete?(existing.CompletedAt||now):'',LastOpenedAt:now};
    if(row>=0)writeObjectToRow_(sh,heads,row+1,obj);else appendObject_(sh,heads,obj); clearCache_(); return {ok:true,progress:obj,writeSpreadsheetId:assertWriteSheet_(sh)};
  } finally {try{lock.releaseLock()}catch(_) {}}
}

// ---------------- TEACHER ----------------
function teacherPassword_(){ return settingValue_('TeacherPassword'); }
function teacherConfigured_(){ return !!teacherPassword_(); }
function teacherStatus_(){ return {ok:true,configured:teacherConfigured_(),storage:'Settings',settingKey:'TeacherPassword'}; }
function teacherSetup_(){ return {ok:false,configured:teacherConfigured_(),error:'Set the teacher password in the master Sheet: Settings → TeacherPassword, then refresh this page.'}; }
function teacherLogin_(passwordHash){
  if(!teacherConfigured_())return {ok:false,error:'Teacher password is blank. Set Settings → TeacherPassword in the master Sheet, then refresh.'};
  loginGuard_('teacher'); const incoming=safeHash_(passwordHash), stored=sha256Hex_(teacherPassword_()).toLowerCase();
  if(incoming!==stored){recordLoginFailure_('teacher');return {ok:false,error:'Incorrect teacher password'}}
  clearLoginFailures_('teacher'); return {ok:true,token:createSession_('teacher','teacher')};
}
function logoutTeacher_(token){ deleteSession_('teacher',token); return {ok:true}; }
function requireTeacher_(token){ return requireSession_('teacher',token); }
function removeAllRowGroups_(sh){
  const last=sh.getLastRow();
  if(last<2)return;
  // One range operation removes existing native row-group depth quickly.
  // This avoids thousands of getRowGroup() calls that can time out in Apps Script.
  try{
    sh.getRange(2,1,last-1,Math.max(1,sh.getLastColumn())).shiftRowGroupDepth(-8);
  }catch(_){
    // Conservative fallback for Sheets that reject a large negative shift.
    for(let i=0;i<8;i++){
      try{sh.getRange(2,1,last-1,Math.max(1,sh.getLastColumn())).shiftRowGroupDepth(-1)}catch(__){break}
    }
  }
}
function groupRowsByExamId_(sheetName){
  const sh=requireSheet_(sheetName),last=sh.getLastRow(),lastCol=sh.getLastColumn();
  if(last<3||lastCol<1)return {sheet:sheetName,groups:0,rows:Math.max(0,last-1)};
  const vals=sh.getRange(1,1,last,lastCol).getDisplayValues(),heads=vals[0].map(clean_),ix=heads.indexOf('ExamID');
  if(ix<0)return {sheet:sheetName,groups:0,rows:last-1,skipped:true};

  removeAllRowGroups_(sh);
  try{sh.setRowGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.BEFORE);}catch(_){}

  let groups=0,start=1,current=clean_(vals[1][ix]);
  const closeBlock=end=>{
    // Leave the first row for an ExamID visible as a label/reference row and
    // put the remaining contiguous rows into a real Google Sheets row group.
    const groupStart=start+2, count=end-start;
    if(current&&count>0){
      sh.getRange(groupStart,1,count,lastCol).shiftRowGroupDepth(1);
      try{const g=sh.getRowGroup(groupStart,1);if(g)g.collapse();}catch(_){}
      groups++;
    }
  };
  for(let r=2;r<vals.length;r++){
    const id=clean_(vals[r][ix]);
    if(id!==current){closeBlock(r-1);start=r;current=id;}
  }
  closeBlock(vals.length-1);
  SpreadsheetApp.flush();
  return {sheet:sheetName,groups:groups,rows:last-1};
}
function rowGroupingTargets_(){
  return [CONFIG.SHEETS.SYLLABUS,CONFIG.SHEETS.EXAM_UPDATES,CONFIG.SHEETS.STUDENT_PROGRESS,CONFIG.SHEETS.STUDENT_NOTES];
}
function rowGroupingSignature_(){
  const ss=getSS_();
  return rowGroupingTargets_().map(name=>{
    const sh=ss.getSheetByName(name);
    return name+':'+(sh?sh.getLastRow():0)+':'+(sh?sh.getLastColumn():0);
  }).join('|');
}
function scheduleAutoRowGrouping_(){
  // V14.2.8 creates native +/- groups directly instead of relying on a
  // time-based trigger. It runs only when the relevant sheet dimensions change.
  try{
    assertUnifiedMaster_();
    const sig=rowGroupingSignature_(), cache=CacheService.getScriptCache();
    const key='native-row-groups:v1428:'+sig;
    if(cache.get(key)==='1')return {ok:true,cached:true};
    const lock=LockService.getScriptLock();
    if(!lock.tryLock(3000))return {ok:true,deferred:true};
    try{
      const results=rowGroupingTargets_().map(name=>{
        try{return groupRowsByExamId_(name)}catch(err){return {sheet:name,error:String(err&&err.message||err)}}
      });
      try{cache.put(key,'1',21600)}catch(_){}
      return {ok:true,results:results};
    }finally{try{lock.releaseLock()}catch(_){}}
  }catch(err){return {ok:false,error:String(err&&err.message||err)}}
}
function autoOrganizeSheetRows_(){
  // Compatibility alias. No installable trigger is required anymore.
  return scheduleAutoRowGrouping_();
}
function teacherStudents_(token){ requireTeacher_(token); const students=rowsAsObjects_(CONFIG.SHEETS.STUDENTS).map(studentSafe_).sort((a,b)=>a.StudentName.localeCompare(b.StudentName)); return {ok:true,students:students,exams:getExams_().exams}; }
function teacherSaveStudent_(p){
  assertUnifiedMaster_();
  requireTeacher_(p.token||'');
  const requestedId=clean_(p.studentId), originalId=clean_(p.originalStudentId), name=clean_(p.studentName), pass=clean_(p.passwordHash), courses=clean_(p.assignedCourses), active=clean_(p.active)===''?true:bool_(p.active);
  if(!name)throw new Error('Student name is required'); if(pass)safeHash_(pass);
  const lock=LockService.getScriptLock();lock.waitLock(5000);
  try{
    const courseIds=courseIds_(courses), sh=requireSheet_(CONFIG.SHEETS.STUDENTS); assertWriteSheet_(sh); const vals=sh.getDataRange().getValues(), heads=vals[0].map(clean_), ix=indexMap_(heads); let row=-1;
    const lookupId=originalId||requestedId;
    if(lookupId){for(let i=1;i<vals.length;i++){if(clean_(vals[i][ix.StudentID])===lookupId){row=i;break;}}}
    const isNew=row<0;
    if(!isNew && originalId && requestedId && requestedId!==originalId)throw new Error('Student ID cannot be changed after creation. Create a new student instead.');
    if(isNew && !pass)throw new Error('Password is required for a new student');
    const studentId=isNew?(requestedId||('STU-'+Utilities.getUuid().slice(0,8).toUpperCase())):clean_(vals[row][ix.StudentID]);
    if(isNew && requestedId){for(let i=1;i<vals.length;i++){if(clean_(vals[i][ix.StudentID])===requestedId)throw new Error('Student ID already exists');}}
    const existing=row>=0?objectFromRow_(heads,vals[row]):{};
    const storedPass=pass?studentCredential_(studentId,pass):(existing.PasswordHash||'');
    const obj={StudentID:studentId,StudentName:name,PasswordHash:storedPass,AssignedCourses:courseIds.join(','),Active:active,CreatedAt:existing.CreatedAt||new Date(),LastLogin:existing.LastLogin||''};
    if(row>=0)writeObjectToRow_(sh,heads,row+1,obj);else appendObject_(sh,heads,obj); clearCache_(); return {ok:true,student:studentSafe_(obj),writeSpreadsheetId:assertWriteSheet_(sh)};
  } finally {try{lock.releaseLock()}catch(_) {}}
}
function teacherDeleteStudent_(p){
  assertUnifiedMaster_();
  requireTeacher_(p.token||''); const id=clean_(p.studentId); if(!id)throw new Error('studentId is required');
  const lock=LockService.getScriptLock();lock.waitLock(5000);
  try{
    const sh=requireSheet_(CONFIG.SHEETS.STUDENTS); assertWriteSheet_(sh); const vals=sh.getDataRange().getValues(), heads=vals[0].map(clean_), ix=heads.indexOf('StudentID'); let deleted=false;
    for(let i=vals.length-1;i>=1;i--){if(clean_(vals[i][ix])===id){sh.deleteRow(i+1);deleted=true;}}
    if(deleted){deleteRowsForStudent_(CONFIG.SHEETS.STUDENT_PROGRESS,id);deleteRowsForStudent_(CONFIG.SHEETS.STUDENT_NOTES,id);clearCache_();}
    return {ok:true,deleted:deleted,writeSpreadsheetId:assertWriteSheet_(sh)};
  } finally {try{lock.releaseLock()}catch(_) {}}
}
function deleteRowsForStudent_(sheetName,studentId){
  const sh=requireSheet_(sheetName);assertWriteSheet_(sh);const vals=sh.getDataRange().getValues();if(vals.length<2)return;const heads=vals[0].map(clean_),ix=heads.indexOf('StudentID');if(ix<0)return;
  for(let i=vals.length-1;i>=1;i--){if(clean_(vals[i][ix])===studentId)sh.deleteRow(i+1);}
}

function loginGuard_(key){const c=CacheService.getScriptCache(),n=num_(c.get('loginfail:'+key),0);if(n>=8)throw new Error('Too many failed login attempts. Try again in about 15 minutes.');}
function recordLoginFailure_(key){const c=CacheService.getScriptCache(),k='loginfail:'+key,n=num_(c.get(k),0)+1;try{c.put(k,String(n),900)}catch(_){} }
function clearLoginFailures_(key){try{CacheService.getScriptCache().remove('loginfail:'+key)}catch(_){} }

// ---------------- SESSION / ROW HELPERS ----------------
function createSession_(kind,value){ const token=Utilities.getUuid().replace(/-/g,'')+Utilities.getUuid().replace(/-/g,'').slice(0,12); CacheService.getScriptCache().put('session:'+kind+':'+token,clean_(value),CONFIG.SESSION_SECONDS); return token; }
function requireSession_(kind,token){ token=clean_(token); if(!token)throw new Error('Login required'); const v=CacheService.getScriptCache().get('session:'+kind+':'+token); if(!v)throw new Error('Session expired. Please log in again.'); return v; }
function deleteSession_(kind,token){ try{CacheService.getScriptCache().remove('session:'+kind+':'+clean_(token))}catch(_){} }
function objectFromRow_(heads,row){ const o={}; heads.forEach((h,i)=>{if(h)o[h]=row[i]}); return o; }
function appendObject_(sheet,heads,obj){ sheet.appendRow(heads.map(h=>Object.prototype.hasOwnProperty.call(obj,h)?obj[h]:'')); }
function writeObjectToRow_(sheet,heads,rowNumber,obj){ const current=sheet.getRange(rowNumber,1,1,heads.length).getValues()[0]; heads.forEach((h,i)=>{if(Object.prototype.hasOwnProperty.call(obj,h))current[i]=obj[h]}); sheet.getRange(rowNumber,1,1,heads.length).setValues([current]); }
function output_(payload,callback){ const json=JSON.stringify(payload); if(callback){ if(!/^[A-Za-z_$][A-Za-z0-9_$.]*$/.test(callback))return ContentService.createTextOutput(JSON.stringify({ok:false,error:'Invalid callback'})).setMimeType(ContentService.MimeType.JSON); return ContentService.createTextOutput(callback+'('+json+');').setMimeType(ContentService.MimeType.JAVASCRIPT);} return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON); }


// ---------------- STUDENT NOTES ----------------
function studentNotes_(token){
  const sid=requireSession_('student',token);
  const rows=rowsAsObjects_(CONFIG.SHEETS.STUDENT_NOTES).filter(r=>clean_(r.StudentID)===sid);
  rows.sort((a,b)=>String(b.UpdatedAt||'').localeCompare(String(a.UpdatedAt||'')));
  return {ok:true,notes:rows};
}
function assertAssignedTopic_(student,examId,subject,chapter,topic){
  const id=assertAssigned_(student,examId), wantedSubject=clean_(subject), wantedChapter=clean_(chapter), wantedTopic=clean_(topic);
  if(!wantedSubject||!wantedChapter||!wantedTopic)throw new Error('Choose subject, chapter and topic');
  const match=getSyllabus_(id).syllabus.find(r=>clean_(r.Subject)===wantedSubject&&clean_(r.Chapter)===wantedChapter&&clean_(r.Topic)===wantedTopic);
  if(!match)throw new Error('Selected topic is not available in this course');
  return match;
}
function studentSaveNote_(p){
  assertUnifiedMaster_();
  const sid=requireSession_('student',p.token||''), id=clean_(p.noteId||''), examId=ensureExamAllowed_(p.examId||p.exam||'');
  const subject=clean_(p.subject),chapter=clean_(p.chapter),topic=clean_(p.topic),title=clean_(p.noteTitle).slice(0,120),text=clean_(p.noteText).slice(0,2000);
  if(!text)throw new Error('Note text is required');
  const student=findStudent_(sid), topicRow=assertAssignedTopic_(student,examId,subject,chapter,topic);
  const lock=LockService.getScriptLock();lock.waitLock(5000);
  try{
    const sh=requireSheet_(CONFIG.SHEETS.STUDENT_NOTES); assertWriteSheet_(sh); const vals=sh.getDataRange().getValues(),heads=vals[0].map(clean_),ix=indexMap_(heads);let row=-1,existing={};
    if(id){for(let i=1;i<vals.length;i++){if(clean_(vals[i][ix.NoteID])===id&&clean_(vals[i][ix.StudentID])===sid){row=i;existing=objectFromRow_(heads,vals[i]);break}}}
    const note={NoteID:id||('NOTE-'+Utilities.getUuid().slice(0,12).toUpperCase()),StudentID:sid,ExamID:examId,Subject:topicRow.Subject,Chapter:topicRow.Chapter,Topic:topicRow.Topic,NoteTitle:title,NoteText:text,UpdatedAt:new Date()};
    if(row>=0)writeObjectToRow_(sh,heads,row+1,note);else appendObject_(sh,heads,note);clearCache_();
    return {ok:true,note:{...note,UpdatedAt:new Date(note.UpdatedAt).toISOString()},writeSpreadsheetId:assertWriteSheet_(sh)};
  } finally {try{lock.releaseLock()}catch(_) {}}
}
function studentDeleteNote_(p){
  assertUnifiedMaster_();
  const sid=requireSession_('student',p.token||''),id=clean_(p.noteId||'');if(!id)throw new Error('noteId is required');
  const lock=LockService.getScriptLock();lock.waitLock(5000);
  try{
    const sh=requireSheet_(CONFIG.SHEETS.STUDENT_NOTES); assertWriteSheet_(sh); const vals=sh.getDataRange().getValues(),heads=vals[0].map(clean_),ix=indexMap_(heads);
    for(let i=1;i<vals.length;i++){if(clean_(vals[i][ix.NoteID])===id&&clean_(vals[i][ix.StudentID])===sid){sh.deleteRow(i+1);clearCache_();return {ok:true,deleted:true,writeSpreadsheetId:assertWriteSheet_(sh)}}}
    return {ok:true,deleted:false,writeSpreadsheetId:assertWriteSheet_(sh)};
  } finally {try{lock.releaseLock()}catch(_) {}}
}
