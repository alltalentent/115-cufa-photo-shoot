/* CUFA PHOTO CONTROL — 2026-10-31. No credentials belong in this file. */
const DB_ID = '1EhDDUVQ-Jvn9v9LXLLYYemBBlZ7dcEZAmfJQrJ6XJpw';
const VERSION = 'photo-control-20261101-v6';
const CLASSES = ['一甲','一乙','一丙','二甲','二乙','二丙','三甲','三乙','三丙','四甲','四乙','四丙','老師'];
const SHOTS = {'試鏡照':['全身正面','半身正面','半身側面','個人發揮'],'形象照':['個人發揮','固定姿勢']};
const TABS = {people:'拍攝學生名單',tasks:'ShootTasks',checkins:'CheckIns',live:'LiveControl',logs:'ActivityLogs',inbox:'RegistrationInbox',schedule:'拍攝流程'};
const HEADERS = {
 people:['personId','studentId','name','class','year','shootTypes','source','reviewFlag','phone','heightCm','weightKg','checkInStatus','queueNo','shootSlot','shootStatus','shootNotes','consentReceived'],
 tasks:['taskId','personId','shootType','checklist','status','checkedItems','notes','updatedAt'],
 checkins:['personId','queueNo','checkedAt','operator','consentStatus'],
 live:['key','value','updatedAt','operator'],logs:['timestamp','action','personId','operator','details'],
 inbox:['submissionId','studentId','name','class','heightCm','weightKg','phone','submittedAt','status']
};
const ROSTER_HEADERS={'學號':'studentId','班級':'class','姓名':'name','拍攝類型':'shootTypes','拍攝時段':'shootSlot','號碼':'queueNo','報到狀態':'checkInStatus','拍攝狀態':'shootStatus','拍攝備註':'shootNotes','行動電話':'phone','身高 cm':'heightCm','體重 kg':'weightKg','肖像權收件':'consentReceived','系統識別碼':'personId','年級':'year','資料來源':'source','待確認事項':'reviewFlag'};
function queueNo_(value){const raw=text_(value);if(!raw)return '';const n=raw.replace(/^A/i,'');if(!/^\d+$/.test(n))throw Error('叫號號碼格式不正確，請核對名單');return n.padStart(3,'0');}
function slotMinute_(value){const s=text_(value).replace(':','');if(!/^\d{3,4}$/.test(s))return null;const n=s.padStart(4,'0'),h=+n.slice(0,2),m=+n.slice(2);return h<24&&m<60?h*60+m:null;}
function slotTime_(n){return String(Math.floor(n/60)).padStart(2,'0')+':'+String(n%60).padStart(2,'0');}
function slotLabel_(p,schedule){const found=schedule.filter(s=>{if(s.kind==='老師')return p.class==='老師';if(s.kind==='個人')return p.shootTypes==='形象照'&&p.name===s.title;if(s.kind==='班級'){const c=text_(s.title).match(/[一二三四][甲乙丙]/);return !!c&&p.class===c[0]&&text_(s.title).includes(p.shootTypes==='試鏡照'?'試鏡':'形象');}return false;});return Array.from(new Set(found.map(s=>{const a=slotMinute_(s.start),b=slotMinute_(s.end);return a===null||b===null?'隨到隨拍':slotTime_(a)+'–'+slotTime_(b);}))).join('、')||(p.class==='老師'?'隨到隨拍':'時段待安排');}
let summaryTargets_=null;
let cachedBook_, tableCache_ = {}, stableSnapshot_ = null, pendingSheetWrites_ = false;
// Advanced Sheets keeps the same fresh safety checks, using one read and one atomic write.
let fastMetadata_ = null, pendingLogs_ = [];
function fastSheets_(){return typeof Sheets!=='undefined'&&!!Sheets.Spreadsheets;}
function metadata_(){if(fastMetadata_)return fastMetadata_;const cache=CacheService.getScriptCache(),saved=cache.get('photo-sheet-metadata');if(saved)return fastMetadata_=JSON.parse(saved);const result=Sheets.Spreadsheets.get(DB_ID,{fields:'sheets(properties(sheetId,title,gridProperties(rowCount,columnCount)))'});fastMetadata_=Object.fromEntries(result.sheets.map(s=>[s.properties.title,{id:s.properties.sheetId,rows:s.properties.gridProperties.rowCount,columns:s.properties.gridProperties.columnCount}]));cache.put('photo-sheet-metadata',JSON.stringify(fastMetadata_),21600);return fastMetadata_;}
function columnLetter_(n){let s='';while(n){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26);}return s;}
function hasTable_(key){return fastSheets_()?!!metadata_()[TABS[key]]:!!book_().getSheetByName(TABS[key]);}
function preload_(keys){if(!fastSheets_())return;const meta=metadata_(),needed=Array.from(new Set(keys)).filter(k=>!tableCache_[k]&&meta[TABS[k]]);if(!needed.length)return;const ranges=needed.map(k=>"'"+TABS[k].replace(/'/g,"''")+"'!A1:"+columnLetter_(meta[TABS[k]].columns)+meta[TABS[k]].rows),result=Sheets.Spreadsheets.Values.batchGet(DB_ID,{ranges,valueRenderOption:'FORMATTED_VALUE',dateTimeRenderOption:'FORMATTED_STRING'});needed.forEach((k,i)=>{const values=(result.valueRanges[i].values||[]).map(row=>row.map(v=>String(v==null?'':v))),headers=(values.shift()||[]).map(h=>k==='people'?(ROSTER_HEADERS[h]||h):h);tableCache_[k]={sheet:null,headers,values,initialLength:values.length};});}
function fastFlush_(){const meta=metadata_(),requests=[],dirty=[];let resized=false;const cells=row=>({values:row.map(value=>({userEnteredValue:{stringValue:String(value==null?'':value)},userEnteredFormat:{numberFormat:{type:'TEXT'}}}))});for(const k of Object.keys(tableCache_)){const t=tableCache_[k];if(!t.dirty?.size)continue;const first=Math.min(...t.dirty),last=Math.max(...t.dirty),m=meta[TABS[k]];if(last+2>m.rows){requests.push({appendDimension:{sheetId:m.id,dimension:'ROWS',length:last+2-m.rows}});resized=true;}for(const [index,cols]of t.dirtyCells||[]){const sorted=[...cols].sort((a,b)=>a-b);for(let i=0;i<sorted.length;){const start=sorted[i];let end=start;while(i+1<sorted.length&&sorted[i+1]===end+1)end=sorted[++i];requests.push({updateCells:{range:{sheetId:m.id,startRowIndex:index+1,endRowIndex:index+2,startColumnIndex:start,endColumnIndex:end+1},rows:[cells(t.values[index].slice(start,end+1))],fields:'userEnteredValue,userEnteredFormat.numberFormat'}});i++;}}dirty.push({t,m,last});}if(pendingLogs_.length)requests.push({appendCells:{sheetId:meta[TABS.logs].id,rows:pendingLogs_.map(cells),fields:'userEnteredValue,userEnteredFormat.numberFormat'}});if(!requests.length)return;Sheets.Spreadsheets.batchUpdate({requests},DB_ID);for(const {t,m,last}of dirty){m.rows=Math.max(m.rows,last+2);t.dirty.clear();t.dirtyCells?.clear();t.initialLength=t.values.length;}pendingLogs_=[];pendingSheetWrites_=false;if(resized)CacheService.getScriptCache().put('photo-sheet-metadata',JSON.stringify(meta),21600);}
function testFastPhotoConnection(){if(!fastSheets_())throw Error('請先啟用 Google Sheets API v4 服務');tableCache_={};cachedBook_=null;fastMetadata_=null;const started=Date.now();preload_(['people','tasks','checkins','live']);console.log(JSON.stringify({backend:'sheets-batch',people:rows_('people').length,tasks:rows_('tasks').length,readMs:Date.now()-started}));}
function book_(){return cachedBook_||(cachedBook_=SpreadsheetApp.openById(DB_ID));}
function sheet_(key){if(tableCache_[key]?.sheet)return tableCache_[key].sheet;const s=book_().getSheetByName(TABS[key]);if(!s)throw Error('缺少資料表');return s;}
function now_(){return Utilities.formatDate(new Date(),'Asia/Taipei','yyyy-MM-dd HH:mm:ss');}
function text_(x){return String(x==null?'':x).trim();}
function safe_(x){const s=String(x==null?'':x);return /^[=+@\-]/.test(s)?"'"+s:s;}
function table_(k){if(!tableCache_[k]){const sheet=sheet_(k),values=sheet.getDataRange().getDisplayValues(),headers=values.shift().map(h=>k==='people'?(ROSTER_HEADERS[h]||h):h);tableCache_[k]={sheet,headers,values,initialLength:values.length};}return tableCache_[k];}
function rows_(k){const t=table_(k);return t.values.map((r,i)=>Object.assign({_row:i+2},...t.headers.map((n,j)=>({[n]:r[j]||''})))).filter(r=>t.headers.some(n=>r[n]));}
function write_(k,obj,row){const t=table_(k),index=row?row-2:t.values.length,values=t.headers.map((n,i)=>Object.prototype.hasOwnProperty.call(obj,n)?text_(obj[n]):(t.values[index]||[])[i]||'');const before=t.values[index]||[],cols=(t.dirtyCells||(t.dirtyCells=new Map())).get(index)||new Set();t.headers.forEach((n,i)=>{if((!row||Object.prototype.hasOwnProperty.call(obj,n))&&String(before[i]||'')!==values[i])cols.add(i);});t.values[index]=values;if(cols.size){t.dirtyCells.set(index,cols);(t.dirty||(t.dirty=new Set())).add(index);}}
function flush_(){if(fastSheets_()){fastFlush_();return;}for(const k of Object.keys(tableCache_)){const t=tableCache_[k];if(!t.dirty?.size)continue;const first=Math.min(...t.dirty),last=Math.max(...t.dirty);for(const [index,cols]of t.dirtyCells||[]){const sorted=[...cols].sort((a,b)=>a-b);for(let i=0;i<sorted.length;){const start=sorted[i];let end=start;while(i+1<sorted.length&&sorted[i+1]===end+1)end=sorted[++i];const range=t.sheet.getRange(index+2,start+1,1,end-start+1);if(index>=t.initialLength)range.setNumberFormat('@');range.setValues([t.values[index].slice(start,end+1).map(safe_)]);i++;}}pendingSheetWrites_=true;t.initialLength=t.values.length;t.dirty.clear();t.dirtyCells?.clear();}}
function commitSheets_(){flush_();if(pendingSheetWrites_){SpreadsheetApp.flush();pendingSheetWrites_=false;}}
function cacheRead_(){try{const c=CacheService.getScriptCache(),meta=JSON.parse(c.get('photo-v6-meta')||'null');if(!meta)return null;const keys=Array.from({length:meta.count},(_,i)=>meta.id+'-'+i),parts=c.getAll(keys);let raw='';for(const key of keys){if(typeof parts[key]!=='string')return null;raw+=parts[key];}return JSON.parse(raw);}catch(_){return null;}}
function cacheWrite_(snapshot){try{const c=CacheService.getScriptCache(),raw=JSON.stringify(snapshot),id='photo-v6-'+Utilities.getUuid(),count=Math.ceil(raw.length/18000),parts={};for(let i=0;i<count;i++)parts[id+'-'+i]=raw.slice(i*18000,(i+1)*18000);c.putAll(parts,20);c.put('photo-v6-meta',JSON.stringify({id,count}),20);}catch(_){}}
function cacheClear_(){try{CacheService.getScriptCache().remove('photo-v6-meta');}catch(_){}}
function json_(data){return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);}
function doGet(){return json_({ok:true,version:VERSION,status:'online'});}
function doPost(e){const started=Date.now();try{const request=JSON.parse(e.postData.contents),result=dispatch_(request);return json_({ok:true,data:compactReply_(request,result),serverMs:Date.now()-started});}catch(err){return json_({ok:false,error:err.message||'操作失敗'});}}
function auth_(p){const expected=PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');if(!expected)throw Error('尚未設定工作人員密碼');if(typeof p!=='string'||p!==expected)throw Error('工作人員密碼錯誤');}
function setupPhotoControl(){const b=book_();Object.keys(HEADERS).forEach(k=>{let s=b.getSheetByName(TABS[k]);if(!s)s=b.insertSheet(TABS[k]);const h=s.getLastColumn()?s.getRange(1,1,1,s.getLastColumn()).getDisplayValues()[0].map(n=>k==='people'?(ROSTER_HEADERS[n]||n):n):[];HEADERS[k].filter(n=>!h.includes(n)).forEach((n,i)=>s.getRange(1,h.length+i+1).setValue(n));s.setFrozenRows(1);s.getRange(1,1,1,HEADERS[k].length).setBackground('#eeeeee').setFontWeight('bold');s.setColumnWidths(1,HEADERS[k].length,145);});b.setSpreadsheetTimeZone('Asia/Taipei');console.log('拍攝系統初始化完成；名單 '+rows_('people').length+' 人。保留原始名單與密碼。');}
function clean_(o){const c=Object.assign({},o);delete c._row;return c;}
function live_(){return Object.fromEntries(rows_('live').map(r=>[r.key,r.value]));}
function setLive_(key,value,op){const r=rows_('live').find(r=>r.key===key);if(r&&r.value===String(value))return;write_('live',{key,value,updatedAt:now_(),operator:op},r&&r._row);}
function taskItems_(t){return SHOTS[t.shootType]||text_(t.checklist).split('／').filter(Boolean);}
function checked_(t){try{const v=JSON.parse(t.checkedItems||'[]');return Array.isArray(v)?v:[];}catch(_){return text_(t.checkedItems).split('／').filter(Boolean);}}
function snapshot_(persist=true){const reuseReferences=stableSnapshot_&&Date.now()-Number(stableSnapshot_.referencesReadAt||0)<20000;preload_(['people','tasks','checkins','live',...(!reuseReferences?['schedule','inbox']:[])]);const people=rows_('people'),tasks=rows_('tasks'),checkins=rows_('checkins'),live=live_(),schedule=reuseReferences?stableSnapshot_.schedule:(hasTable_('schedule')?rows_('schedule').map(clean_):[]);
 for(const ci of checkins){const no=queueNo_(ci.queueNo);if(persist&&no!==ci.queueNo){write_('checkins',{queueNo:no},ci._row);ci.queueNo=no;}}
 const targets=summaryTargets_===null?null:new Set(summaryTargets_);if(targets)for(const key of ['people','tasks']){const table=tableCache_[key];for(const index of table?.dirty||[]){const col=table.headers.indexOf('personId');if(col>=0)targets.add(table.values[index][col]);}}
 for(const p of people){const mine=tasks.filter(t=>t.personId===p.personId&&t.status!=='取消'),ci=checkins.find(c=>c.personId===p.personId),status=p.checkInStatus==='請假'?'請假':mine.length&&mine.every(t=>t.status==='已完成')?'已完成':live.activePersonId===p.personId?'拍攝中':calledQueue_(live).includes(p.personId)?'正在叫號':mine.some(t=>t.status==='待補拍')?'待補拍':p.checkInStatus==='已報到'?'待拍攝':'未報到';
 const summary={queueNo:queueNo_(p.queueNo),shootSlot:p.shootSlot||slotLabel_(p,schedule),shootStatus:status,shootNotes:mine.filter(t=>t.notes).map(t=>mine.length>1?t.shootType+'：'+t.notes:t.notes).join('；'),consentReceived:p.class!=='老師'&&p.shootTypes==='形象照'?(ci&&ci.consentStatus==='已完成'?'已收件':'未收件'):'不適用'};
 const changes={};for(const key of Object.keys(summary))if(p[key]!==summary[key])changes[key]=summary[key];if(persist&&(!targets||targets.has(p.personId))&&Object.keys(changes).length)write_('people',changes,p._row);Object.assign(p,summary);}
 commitSheets_();const snapshot={version:VERSION,revision:Utilities.getUuid(),referencesReadAt:reuseReferences?stableSnapshot_.referencesReadAt:Date.now(),people:people.map(clean_),tasks:tasks.map(t=>Object.assign(clean_(t),{items:taskItems_(t),checked:checked_(t)})),checkins:checkins.map(clean_),live,schedule,inbox:!tableCache_.inbox&&reuseReferences?stableSnapshot_.inbox:rows_('inbox').filter(r=>r.status==='待核對').map(clean_),updatedAt:now_()};cacheWrite_(snapshot);return snapshot;}
function compactReply_(request,result){const previous=stableSnapshot_;if(!request.compact||!previous?.revision||request.knownRevision!==previous.revision||!result?.people||!result.tasks)return result;const changed=(key,id)=>{const old=new Map((previous[key]||[]).map(v=>[v[id],JSON.stringify(v)]));return result[key].filter(v=>old.get(v[id])!==JSON.stringify(v));};return {patch:true,baseRevision:previous.revision,revision:result.revision,updatedAt:result.updatedAt,version:result.version,live:result.live,people:changed('people','personId'),tasks:changed('tasks','taskId'),checkins:changed('checkins','personId'),inbox:result.inbox,receipt:result.receipt};}
function profile_(d){const p={studentId:text_(d.studentId),name:text_(d.name),class:text_(d.class),heightCm:text_(d.heightCm),weightKg:text_(d.weightKg),phone:text_(d.phone)};if(!CLASSES.includes(p.class)||!p.name||p.name.length>50)throw Error('請填寫正確班級及姓名');if(p.class==='老師')return p;if(!/^[A-Za-z0-9]{4,20}$/.test(p.studentId))throw Error('請確認學號');if(p.class==='老師'&&p.studentId&&!/^[A-Za-z0-9]{1,20}$/.test(p.studentId))throw Error('識別碼格式不正確');if(!(Number(p.heightCm)>=80&&Number(p.heightCm)<=250))throw Error('身高請填 80–250 cm');if(!(Number(p.weightKg)>=20&&Number(p.weightKg)<=250))throw Error('體重請填 20–250 kg');if(!/^\+?[0-9 ()-]{8,20}$/.test(p.phone))throw Error('請確認行動電話');return p;}
function logAction_(action,id,op,detail){if(fastSheets_()){pendingLogs_.push([now_(),action,id,op,JSON.stringify(detail||{})]);return;}sheet_('logs').appendRow([now_(),action,id,op,JSON.stringify(detail||{})].map(safe_));pendingSheetWrites_=true;}
function ready_(p,tasks,checkins){if(p.checkInStatus!=='已報到')throw Error('此人尚未報到或已請假');const ts=tasks.filter(t=>t.personId===p.personId);if(!ts.length)throw Error('尚未設定拍攝任務');if(p.class!=='老師'&&ts.some(t=>t.shootType==='形象照')&&!checkins.some(c=>c.personId===p.personId&&c.consentStatus==='已完成'))throw Error('請先在報到區檢核肖像權同意書');}
function calledQueue_(live){let q=[];try{q=JSON.parse(live.callQueue||'[]');}catch(_){}if(!Array.isArray(q))q=[];return Array.from(new Set([live.calledPersonId,...q].filter(Boolean)));}
function setCalledQueue_(q,op){setLive_('callQueue',JSON.stringify(q),op);setLive_('calledPersonId',q[0]||'',op);}
function saveCheckin_(d,op,people,tasks,checkins,p){
 const v=profile_(p?.class==='老師'?p:(d.profile||p||{})),type=p?p.shootTypes:text_(d.shootType);
 if(!SHOTS[type])throw Error('請選擇拍攝類型');
 if(p&&d.shootType&&d.shootType!==type)throw Error('報到資料不能變更拍攝類型');
 if(v.class!=='老師'&&type==='形象照'&&d.consent!==true)throw Error('形象照請先確認肖像權同意書已收件');
 if(people.some(x=>v.studentId&&x.studentId===v.studentId&&x.shootTypes===type&&x.personId!==(p&&p.personId)))throw Error('此學號已在此拍攝名單中');
 if(d.personId&&!p)throw Error('找不到人員');
 const id=p?p.personId:'P'+Utilities.getUuid(),ci=checkins.find(c=>c.personId===id);
 let no=(p&&p.queueNo)||(ci&&ci.queueNo);if(!no)no=String(Math.max(0,...people.map(x=>Number(text_(x.queueNo).replace(/^A/,''))||0),...checkins.map(x=>Number(text_(x.queueNo).replace(/^A/,''))||0))+1).padStart(3,'0');
 write_('people',Object.assign({},v,{personId:id,shootTypes:type,source:p?p.source:'現場新增',checkInStatus:'已報到',queueNo:no}),p&&p._row);
 if(v.studentId)people.filter(x=>x.personId!==id&&x.studentId===v.studentId).forEach(x=>write_('people',v,x._row));
 if(!p)write_('tasks',{taskId:'T'+Utilities.getUuid(),personId:id,shootType:type,checklist:SHOTS[type].join('／'),status:'待拍攝',checkedItems:'[]',updatedAt:now_()});
 tasks.filter(t=>t.personId===id&&['待報到','請假'].includes(t.status)).forEach(t=>write_('tasks',{status:'待拍攝',updatedAt:now_()},t._row));
 write_('checkins',{personId:id,queueNo:no,checkedAt:ci?ci.checkedAt:now_(),operator:op,consentStatus:v.class!=='老師'&&type==='形象照'?'已完成':'不適用'},ci&&ci._row);
 if(d.submissionId){const row=rows_('inbox').find(x=>x.submissionId===d.submissionId);if(row)write_('inbox',{status:'已核對'},row._row);}
 logAction_('saveCheckIn',id,op);const result=snapshot_();result.receipt={personId:id,name:v.name,queueNo:queueNo_(no),shootType:type,shootSlot:result.people.find(x=>x.personId===id).shootSlot};return result;
}
function taskUpdates_(mine,updates,complete){if(!Array.isArray(updates))throw Error('缺少拍攝項目');const seen=new Set();const clean=updates.map(u=>{const t=mine.find(t=>t.taskId===u.taskId);if(!t||seen.has(u.taskId))throw Error('拍攝任務不正確');seen.add(u.taskId);const items=taskItems_(t);if(!Array.isArray(u.checked)||u.checked.some(x=>!items.includes(x)))throw Error('拍攝項目不正確');const checked=Array.from(new Set(u.checked));if(complete&&checked.length!==items.length)throw Error('請先勾選所有拍攝項目');return {t,checked,notes:text_(u.notes).slice(0,2000)};});if(complete&&seen.size!==mine.length)throw Error('請確認所有拍攝任務');return clean;}
function dispatch_(d){
 if(PropertiesService.getScriptProperties().getProperty('SYNC_PAUSED')!=='false')throw Error('同步已暫停：正在編輯名單與流程，請等待重新啟用');
 tableCache_={};cachedBook_=null;stableSnapshot_=null;pendingSheetWrites_=false;fastMetadata_=null;pendingLogs_=[];
 if(d.action==='submitProfile'){
  const p=profile_(d),lock=LockService.getScriptLock();lock.waitLock(15000);
  try{preload_(['inbox']);const requestId=text_(d.requestId);if(!/^[a-zA-Z0-9-]{16,80}$/.test(requestId))throw Error('請重新整理頁面後再送出');if(rows_('inbox').some(r=>r.submissionId===requestId))return {received:true};const pending=rows_('inbox').filter(r=>r.status==='待核對'&&r.studentId===p.studentId&&r.name===p.name);if(pending.length>=3)throw Error('已收到資料，請至報到區核對');cacheClear_();write_('inbox',Object.assign(p,{submissionId:requestId,submittedAt:now_(),status:'待核對'}));commitSheets_();return {received:true};}finally{lock.releaseLock();}
 }
 auth_(d.password);
 if(d.action==='snapshot'){const cached=cacheRead_();if(cached&&!d.fresh)return cached;const readLock=LockService.getScriptLock();readLock.waitLock(20000);try{if(d.fresh){cacheClear_();return snapshot_(false);}return cacheRead_()||snapshot_(false);}finally{readLock.releaseLock();}}
 if(d.action==='display'){const cached=cacheRead_();if(!cached)preload_(['people','live','schedule']);const live=cached?cached.live:live_(),ps=cached?cached.people:rows_('people'),person=ps.find(p=>p.personId===(d.mode==='slate'?live.activePersonId:live.calledPersonId));return {live,upcoming:d.mode==='call'?calledQueue_(live).slice(1).map(id=>ps.find(p=>p.personId===id)).filter(Boolean).map(p=>({personId:p.personId,name:p.name,class:p.class,shootTypes:p.shootTypes,queueNo:queueNo_(p.queueNo),shootSlot:p.shootSlot||slotLabel_(p,hasTable_('schedule')?rows_('schedule'):[])})):[],person:person?{personId:person.personId,studentId:person.studentId,name:person.name,class:person.class,heightCm:person.heightCm,weightKg:person.weightKg,phone:person.phone,shootTypes:person.shootTypes,queueNo:queueNo_(person.queueNo),shootSlot:person.shootSlot||slotLabel_(person,hasTable_('schedule')?rows_('schedule'):[])}:null,updatedAt:cached?cached.updatedAt:now_()};}
 const lock=LockService.getScriptLock();lock.waitLock(20000);
 try{
 stableSnapshot_=cacheRead_();cacheClear_();preload_(['people','tasks','checkins','live']);
 const op=text_(d.operator).slice(0,40)||'工作人員',people=rows_('people'),tasks=rows_('tasks'),checkins=rows_('checkins'),live=live_();let p=people.find(p=>p.personId===d.personId);summaryTargets_=new Set([d.personId,live.activePersonId,...calledQueue_(live)].filter(Boolean));
 if(d.action==='saveCheckIn')return saveCheckin_(d,op,people,tasks,checkins,p);
 if(d.action==='savePerson'){
  const v=profile_(d.profile||{}),type=p?p.shootTypes:text_(d.shootType);if(!SHOTS[type])throw Error('請從試鏡照或形象照工作區新增人員');
  if(p&&((d.shootType&&d.shootType!==type)||(Array.isArray(d.shootTypes)&&d.shootTypes.join('／')!==type)))throw Error('報到資料不能變更拍攝類型');
  const duplicate=people.find(x=>v.studentId&&x.studentId===v.studentId&&x.shootTypes===type&&x.personId!==d.personId);if(duplicate)throw Error('此學號已在這個拍攝名單中，請搜尋後編輯');
  if(d.personId&&!p)throw Error('找不到人員');
  const id=p?p.personId:'P'+Utilities.getUuid();write_('people',Object.assign({},v,{personId:id,shootTypes:type,source:p?p.source:'現場新增',checkInStatus:p?p.checkInStatus:'未報到'}),p&&p._row);
  if(v.studentId)people.filter(x=>x.personId!==id&&x.studentId===v.studentId).forEach(x=>write_('people',v,x._row));
  if(!p)write_('tasks',{taskId:'T'+Utilities.getUuid(),personId:id,shootType:type,checklist:SHOTS[type].join('／'),status:'待報到',checkedItems:'[]',updatedAt:now_()});
  if(d.submissionId){const s=rows_('inbox').find(r=>r.submissionId===d.submissionId);if(s)write_('inbox',{status:'已核對'},s._row);}logAction_('savePerson',id,op);return snapshot_();
 }
 if(d.action==='dismissSubmission'){const s=rows_('inbox').find(r=>r.submissionId===d.submissionId);if(!s)throw Error('找不到填寫資料');write_('inbox',{status:'略過'},s._row);logAction_('dismissSubmission','',op);return snapshot_();}
 if(d.action==='clearCall'){setCalledQueue_(calledQueue_(live).slice(1),op);logAction_('clearCall','',op);return snapshot_();}
 if(!p)throw Error('找不到人員');const mine=tasks.filter(t=>t.personId===p.personId&&t.status!=='取消');const ci=checkins.find(c=>c.personId===p.personId);
 if(d.action==='checkIn'){
  const isImage=p.class!=='老師'&&mine.some(t=>t.shootType==='形象照'),consent=isImage?(d.consent===true?'已完成':'未完成'):'不適用';
  if(isImage&&consent!=='已完成')throw Error('形象照請先確認肖像權同意書');
  profile_(p);let no=p.queueNo||(ci&&ci.queueNo);if(!no){const max=Math.max(0,...people.map(x=>Number(text_(x.queueNo).replace(/^A/,''))||0),...checkins.map(x=>Number(text_(x.queueNo).replace(/^A/,''))||0));no=String(max+1).padStart(3,'0');}
  write_('people',{queueNo:no,checkInStatus:'已報到'},p._row);write_('checkins',{personId:p.personId,queueNo:no,checkedAt:ci?ci.checkedAt:now_(),operator:op,consentStatus:consent},ci&&ci._row);mine.filter(t=>['待報到','請假'].includes(t.status)).forEach(t=>write_('tasks',{status:'待拍攝',updatedAt:now_()},t._row));
 }else if(d.action==='absent'){
  if(live.activePersonId===p.personId)throw Error('請先結束棚內拍攝');write_('people',{checkInStatus:'請假'},p._row);mine.filter(t=>t.status!=='已完成').forEach(t=>write_('tasks',{status:'請假',updatedAt:now_()},t._row));if(calledQueue_(live).includes(p.personId))setCalledQueue_(calledQueue_(live).filter(id=>id!==p.personId),op);
 }else if(d.action==='call'||d.action==='start'){
  ready_(p,mine,checkins);if(!mine.some(t=>t.status!=='已完成'))throw Error('拍攝已完成，需要補拍請先標記');
  if(d.action==='call'){if(live.activePersonId===p.personId)throw Error('此人已在棚內拍攝');setCalledQueue_(Array.from(new Set([...calledQueue_(live),p.personId])),op);}
  else{if(live.calledPersonId!==p.personId)throw Error('請先叫號，再由正在叫號卡片按進棚');if(live.activePersonId&&live.activePersonId!==p.personId)throw Error('棚內仍有拍攝者，請先完成或暫緩');setLive_('activePersonId',p.personId,op);setCalledQueue_(calledQueue_(live).slice(1),op);mine.filter(t=>t.status!=='已完成').forEach(t=>write_('tasks',{status:'拍攝中',updatedAt:now_()},t._row));}
 }else if(d.action==='saveTask'){
  const t=mine.find(t=>t.taskId===d.taskId);if(!t)throw Error('找不到任務');const items=taskItems_(t);if(!Array.isArray(d.checked)||d.checked.some(i=>!items.includes(i)))throw Error('拍攝項目不正確');const checked=Array.from(new Set(d.checked));if(t.status==='已完成'&&checked.length!==items.length)throw Error('請先標記補拍再修改已完成任務');write_('tasks',{checkedItems:JSON.stringify(checked),notes:text_(d.notes).slice(0,2000),updatedAt:now_()},t._row);
 }else if(d.action==='finish'){
  if(live.activePersonId!==p.personId)throw Error('請先將此人設為棚內拍攝者');if(!mine.length)throw Error('尚未設定拍攝任務');const updates=d.tasks?taskUpdates_(mine,d.tasks,true):taskUpdates_(mine,mine.map(t=>({taskId:t.taskId,checked:checked_(t),notes:t.notes})),true);updates.forEach(u=>write_('tasks',{status:'已完成',checkedItems:JSON.stringify(u.checked),notes:u.notes,updatedAt:now_()},u.t._row));setLive_('activePersonId','',op);if(calledQueue_(live).includes(p.personId))setCalledQueue_(calledQueue_(live).filter(id=>id!==p.personId),op);
 }else if(d.action==='hold'||d.action==='retake'){
  ready_(p,mine,checkins);if(d.action==='hold'&&d.tasks)taskUpdates_(mine,d.tasks,false).forEach(u=>{u.t.checkedItems=JSON.stringify(u.checked);u.t.notes=u.notes;});const targets=d.action==='retake'?mine:mine.filter(t=>t.status!=='已完成');targets.forEach(t=>write_('tasks',{status:d.action==='retake'?'待補拍':'待拍攝',checkedItems:d.action==='retake'?'[]':t.checkedItems,notes:t.notes,updatedAt:now_()},t._row));if(live.activePersonId===p.personId)setLive_('activePersonId','',op);if(calledQueue_(live).includes(p.personId))setCalledQueue_(calledQueue_(live).filter(id=>id!==p.personId),op);
 }else if(d.action==='requeue'){
  ready_(p,mine,checkins);mine.filter(t=>['暫緩','待補拍'].includes(t.status)).forEach(t=>write_('tasks',{status:'待拍攝',updatedAt:now_()},t._row));
 }else throw Error('不支援的操作');
 logAction_(d.action,p.personId,op,{taskId:text_(d.taskId)});return snapshot_();
 }finally{lock.releaseLock();}
}
function testPhotoConnection(){const b=book_();console.log(JSON.stringify({version:VERSION,title:b.getName(),people:rows_('people').length,tasks:rows_('tasks').length,configured:!!PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD')}));}

function splitShootWorkflows(){
 stableSnapshot_=null;tableCache_={};cachedBook_=null;const lock=LockService.getScriptLock();lock.waitLock(20000);
 try{cacheClear_();const people=rows_('people'),tasks=rows_('tasks'),checkins=rows_('checkins'),live=live_();let next=Math.max(0,...people.map(p=>Number(text_(p.queueNo).replace(/^A/,''))||0));let count=0;
 for(const p of people){const mine=tasks.filter(t=>t.personId===p.personId&&t.status!=='取消'),types=Array.from(new Set(mine.map(t=>t.shootType))).filter(t=>SHOTS[t]);if(types.length<2){if(types.length===1&&p.shootTypes!==types[0])write_('people',{shootTypes:types[0]},p._row);continue;}
 const primary=types.includes('試鏡照')?'試鏡照':types[0],ci=checkins.find(c=>c.personId===p.personId);write_('people',{shootTypes:primary},p._row);if(ci&&primary==='試鏡照')write_('checkins',{consentStatus:'不適用'},ci._row);
 for(const type of types.filter(t=>t!==primary)){const id=p.personId+'__'+(type==='形象照'?'image':'audition');if(people.some(x=>x.personId===id))throw Error('存在拆分紀錄，請先核對');const own=mine.filter(t=>t.shootType===type),copy=clean_(p);copy.personId=id;copy.shootTypes=type;copy.queueNo=p.queueNo?String(++next).padStart(3,'0'):'';write_('people',copy);
 if(ci)write_('checkins',{personId:id,queueNo:copy.queueNo,checkedAt:ci.checkedAt,operator:ci.operator,consentStatus:ci.consentStatus});
 for(const t of own){const change={personId:id};if(live.activePersonId===p.personId&&t.status==='拍攝中')change.status='待拍攝';write_('tasks',change,t._row);}count++;}
 }
 if(live.calledPersonId&&live.calledPersonId===live.activePersonId)setLive_('calledPersonId','','系統更新');
 flush_();logAction_('splitShootWorkflows','','系統更新',{splitPeople:count});commitSheets_();console.log(JSON.stringify({splitPeople:count,records:rows_('people').length,tasks:rows_('tasks').length}));return {splitPeople:count};
 }finally{lock.releaseLock();}
}

// Run from the owner editor: hold the same lock as check-in/shooting writes while moving columns.
function organizePhotoRoster(){stableSnapshot_=null;const lock=LockService.getScriptLock();lock.waitLock(20000);try{tableCache_={};cachedBook_=null;cacheClear_();const sh=sheet_('people'),labels=Object.keys(ROSTER_HEADERS),keys=labels.map(x=>ROSTER_HEADERS[x]);let headers=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0].map(x=>ROSTER_HEADERS[x]||x);if(headers.some(x=>x&&!keys.includes(x)))throw Error('名單包含未知欄位，已停止整理');const before=rows_('people').map(p=>p.personId).sort().join('|');for(const key of keys)if(!headers.includes(key)){sh.getRange(1,headers.length+1).setValue(key);headers.push(key);}for(let i=0;i<keys.length;i++){const j=headers.indexOf(keys[i]);if(j!==i){sh.moveColumns(sh.getRange(1,j+1),i+1);headers.splice(i,0,headers.splice(j,1)[0]);}}sh.getRange(1,1,1,labels.length).setValues([labels]);SpreadsheetApp.flush();tableCache_={};const after=rows_('people').map(p=>p.personId).sort().join('|');if(before!==after)throw Error('識別碼核對不一致');const snap=snapshot_();console.log(JSON.stringify({organized:true,people:snap.people.length,tasks:snap.tasks.length,idsPreserved:true}));return true;}finally{lock.releaseLock();}}

// Owner-only editor controls; no public API route can resume synchronization.
function resumePhotoSync(){const lock=LockService.getScriptLock();lock.waitLock(20000);try{cacheClear_();PropertiesService.getScriptProperties().setProperty('SYNC_PAUSED','false');console.log('同步已恢復');}finally{lock.releaseLock();}}
function pausePhotoSync(){const lock=LockService.getScriptLock();lock.waitLock(20000);try{PropertiesService.getScriptProperties().setProperty('SYNC_PAUSED','true');cacheClear_();console.log('同步已暫停');}finally{lock.releaseLock();}}
