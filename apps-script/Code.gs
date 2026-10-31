/* CUFA PHOTO CONTROL — 2026-10-31. No credentials belong in this file. */
const DB_ID = '1EhDDUVQ-Jvn9v9LXLLYYemBBlZ7dcEZAmfJQrJ6XJpw';
const VERSION = 'photo-control-20261031';
const CLASSES = ['一甲','一乙','一丙','二甲','二乙','二丙','三甲','三乙','三丙','四甲','四乙','四丙','老師'];
const SHOTS = {'試鏡照':['全身正面','半身正面','半身側面','個人發揮'],'形象照':['個人發揮','固定姿勢']};
const TABS = {people:'拍攝學生名單',tasks:'ShootTasks',checkins:'CheckIns',live:'LiveControl',logs:'ActivityLogs',inbox:'RegistrationInbox'};
const HEADERS = {
 people:['personId','studentId','name','class','year','shootTypes','source','reviewFlag','phone','heightCm','weightKg','checkInStatus','queueNo'],
 tasks:['taskId','personId','shootType','checklist','status','checkedItems','notes','updatedAt'],
 checkins:['personId','queueNo','checkedAt','operator','consentStatus'],
 live:['key','value','updatedAt','operator'],logs:['timestamp','action','personId','operator','details'],
 inbox:['submissionId','studentId','name','class','heightCm','weightKg','phone','submittedAt','status']
};
function book_(){return SpreadsheetApp.openById(DB_ID);}
function sheet_(key){const s=book_().getSheetByName(TABS[key]);if(!s)throw Error('缺少資料表，請先執行 setupPhotoControl');return s;}
function now_(){return Utilities.formatDate(new Date(),'Asia/Taipei','yyyy-MM-dd HH:mm:ss');}
function text_(x){return String(x==null?'':x).trim();}
function safe_(x){const s=String(x==null?'':x);return /^[=+@\-]/.test(s)?"'"+s:s;}
function rows_(k){const s=sheet_(k),v=s.getDataRange().getDisplayValues(),h=v.shift();return v.map((r,i)=>Object.assign({_row:i+2},...h.map((n,j)=>({[n]:r[j]})))).filter(r=>h.some(n=>r[n]));}
function write_(k,obj,row){const s=sheet_(k),h=s.getRange(1,1,1,s.getLastColumn()).getDisplayValues()[0];if(row){Object.keys(obj).forEach(n=>{const c=h.indexOf(n);if(c>=0)s.getRange(row,c+1).setValue(safe_(obj[n]));});}else{s.appendRow(h.map(n=>safe_(obj[n])));}}
function json_(data){return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);}
function doGet(){return json_({ok:true,version:VERSION,status:'online'});}
function doPost(e){try{return json_({ok:true,data:dispatch_(JSON.parse(e.postData.contents))});}catch(err){return json_({ok:false,error:err.message||'操作失敗'});}}
function auth_(p){const expected=PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');if(!expected)throw Error('尚未設定工作人員密碼');if(typeof p!=='string'||p!==expected)throw Error('工作人員密碼錯誤');}
function setupPhotoControl(){const b=book_();Object.keys(HEADERS).forEach(k=>{let s=b.getSheetByName(TABS[k]);if(!s)s=b.insertSheet(TABS[k]);const h=s.getLastColumn()?s.getRange(1,1,1,s.getLastColumn()).getDisplayValues()[0]:[];HEADERS[k].filter(n=>!h.includes(n)).forEach((n,i)=>s.getRange(1,h.length+i+1).setValue(n));s.setFrozenRows(1);s.getRange(1,1,1,HEADERS[k].length).setBackground('#eeeeee').setFontWeight('bold');s.setColumnWidths(1,HEADERS[k].length,145);});b.setSpreadsheetTimeZone('Asia/Taipei');console.log('拍攝系統初始化完成；名單 '+rows_('people').length+' 人。保留原始名單與密碼。');}
function clean_(o){const c=Object.assign({},o);delete c._row;return c;}
function live_(){return Object.fromEntries(rows_('live').map(r=>[r.key,r.value]));}
function setLive_(key,value,op){const r=rows_('live').find(r=>r.key===key);write_('live',{key,value,updatedAt:now_(),operator:op},r&&r._row);}
function taskItems_(t){return SHOTS[t.shootType]||text_(t.checklist).split('／').filter(Boolean);}
function checked_(t){try{const v=JSON.parse(t.checkedItems||'[]');return Array.isArray(v)?v:[];}catch(_){return text_(t.checkedItems).split('／').filter(Boolean);}}
function snapshot_(){return {version:VERSION,people:rows_('people').map(clean_),tasks:rows_('tasks').map(t=>Object.assign(clean_(t),{items:taskItems_(t),checked:checked_(t)})),checkins:rows_('checkins').map(clean_),live:live_(),inbox:rows_('inbox').filter(r=>r.status==='待核對').map(clean_),updatedAt:now_()};}
function profile_(d){const p={studentId:text_(d.studentId),name:text_(d.name),class:text_(d.class),heightCm:text_(d.heightCm),weightKg:text_(d.weightKg),phone:text_(d.phone)};if(!CLASSES.includes(p.class)||!p.name||p.name.length>50)throw Error('請填寫正確班級及姓名');if(p.class!=='老師'&&!/^[A-Za-z0-9]{4,20}$/.test(p.studentId))throw Error('請確認學號');if(p.class==='老師'&&p.studentId&&!/^[A-Za-z0-9]{1,20}$/.test(p.studentId))throw Error('識別碼格式不正確');if(!(Number(p.heightCm)>=80&&Number(p.heightCm)<=250))throw Error('身高請填 80–250 cm');if(!(Number(p.weightKg)>=20&&Number(p.weightKg)<=250))throw Error('體重請填 20–250 kg');if(!/^\+?[0-9 ()-]{8,20}$/.test(p.phone))throw Error('請確認行動電話');return p;}
function logAction_(action,id,op,detail){write_('logs',{timestamp:now_(),action,personId:id,operator:op,details:JSON.stringify(detail||{})});}
function ready_(p,tasks,checkins){if(p.checkInStatus!=='已報到')throw Error('此人尚未報到或已請假');const ts=tasks.filter(t=>t.personId===p.personId);if(!ts.length)throw Error('尚未設定拍攝任務');if(ts.some(t=>t.shootType==='形象照')&&!checkins.some(c=>c.personId===p.personId&&c.consentStatus==='已完成'))throw Error('請先在報到區檢核肖像權同意書');}
function dispatch_(d){
 if(d.action==='submitProfile'){
  const p=profile_(d),lock=LockService.getScriptLock();lock.waitLock(15000);
  try{const requestId=text_(d.requestId);if(!/^[a-zA-Z0-9-]{16,80}$/.test(requestId))throw Error('請重新整理頁面後再送出');if(rows_('inbox').some(r=>r.submissionId===requestId))return {received:true};const pending=rows_('inbox').filter(r=>r.status==='待核對'&&r.studentId===p.studentId&&r.name===p.name);if(pending.length>=3)throw Error('已收到資料，請至報到區核對');write_('inbox',Object.assign(p,{submissionId:requestId,submittedAt:now_(),status:'待核對'}));return {received:true};}finally{lock.releaseLock();}
 }
 auth_(d.password);
 if(d.action==='snapshot')return snapshot_();
 if(d.action==='display'){const live=live_(),ps=rows_('people'),person=ps.find(p=>p.personId===(d.mode==='slate'?live.activePersonId:live.calledPersonId));return {live,person:person?{personId:person.personId,studentId:person.studentId,name:person.name,class:person.class,heightCm:person.heightCm,weightKg:person.weightKg,queueNo:person.queueNo}:null,updatedAt:now_()};}
 const lock=LockService.getScriptLock();lock.waitLock(20000);
 try{
 const op=text_(d.operator).slice(0,40)||'工作人員',people=rows_('people'),tasks=rows_('tasks'),checkins=rows_('checkins'),live=live_();let p=people.find(p=>p.personId===d.personId);
 if(d.action==='savePerson'){
  const v=profile_(d.profile||{}),types=d.shootTypes;if(!Array.isArray(types)||!types.length||types.some(t=>!SHOTS[t])||new Set(types).size!==types.length)throw Error('請選擇拍攝類型');
  const duplicate=people.find(x=>v.studentId&&x.studentId===v.studentId&&x.personId!==d.personId);if(duplicate)throw Error('學號已在名單中，請搜尋後編輯');
  if(d.personId&&!p)throw Error('找不到人員');if(p&&(p.checkInStatus==='已報到'||tasks.some(t=>t.personId===p.personId&&['拍攝中','已完成','待補拍'].includes(t.status)))){if(types.slice().sort().join('|')!==tasks.filter(t=>t.personId===p.personId&&t.status!=='取消').map(t=>t.shootType).sort().join('|'))throw Error('已報到的人員不可直接更改拍攝類型');}
  const id=p?p.personId:'P'+Utilities.getUuid();write_('people',Object.assign(v,{personId:id,shootTypes:types.join('／'),source:p?p.source:'現場新增',checkInStatus:p?p.checkInStatus:'未報到'}),p&&p._row);
  tasks.filter(t=>t.personId===id&&!types.includes(t.shootType)).forEach(t=>write_('tasks',{status:'取消'},t._row));
  types.forEach(type=>{let task=tasks.find(t=>t.personId===id&&t.shootType===type);if(!task)write_('tasks',{taskId:'T'+Utilities.getUuid(),personId:id,shootType:type,checklist:SHOTS[type].join('／'),status:'待報到',checkedItems:'[]',updatedAt:now_()});else if(task.status==='取消')write_('tasks',{status:'待報到'},task._row);});
  if(d.submissionId){const s=rows_('inbox').find(r=>r.submissionId===d.submissionId);if(s)write_('inbox',{status:'已核對'},s._row);}logAction_('savePerson',id,op);return snapshot_();
 }
 if(d.action==='dismissSubmission'){const s=rows_('inbox').find(r=>r.submissionId===d.submissionId);if(!s)throw Error('找不到填寫資料');write_('inbox',{status:'略過'},s._row);logAction_('dismissSubmission','',op);return snapshot_();}
 if(d.action==='clearCall'){setLive_('calledPersonId','',op);logAction_('clearCall','',op);return snapshot_();}
 if(!p)throw Error('找不到人員');const mine=tasks.filter(t=>t.personId===p.personId&&t.status!=='取消');const ci=checkins.find(c=>c.personId===p.personId);
 if(d.action==='checkIn'){
  const consent=d.consent===true?'已完成':ci?ci.consentStatus:'未完成';
  if(mine.some(t=>t.shootType==='形象照')&&consent!=='已完成')throw Error('形象照請先確認肖像權同意書');
  profile_(p);let no=p.queueNo||(ci&&ci.queueNo);if(!no){const max=Math.max(0,...people.map(x=>Number(text_(x.queueNo).replace(/^A/,''))||0),...checkins.map(x=>Number(text_(x.queueNo).replace(/^A/,''))||0));no='A'+String(max+1).padStart(2,'0');}
  write_('people',{queueNo:no,checkInStatus:'已報到'},p._row);write_('checkins',{personId:p.personId,queueNo:no,checkedAt:ci?ci.checkedAt:now_(),operator:op,consentStatus:consent},ci&&ci._row);mine.filter(t=>['待報到','請假'].includes(t.status)).forEach(t=>write_('tasks',{status:'待拍攝',updatedAt:now_()},t._row));
 }else if(d.action==='absent'){
  if(live.activePersonId===p.personId)throw Error('請先結束棚內拍攝');write_('people',{checkInStatus:'請假'},p._row);mine.filter(t=>t.status!=='已完成').forEach(t=>write_('tasks',{status:'請假',updatedAt:now_()},t._row));if(live.calledPersonId===p.personId)setLive_('calledPersonId','',op);
 }else if(d.action==='call'||d.action==='start'){
  ready_(p,mine,checkins);if(!mine.some(t=>t.status!=='已完成'))throw Error('拍攝已完成，需要補拍請先標記');
  if(d.action==='call')setLive_('calledPersonId',p.personId,op);
  else{if(live.activePersonId&&live.activePersonId!==p.personId)throw Error('棚內仍有拍攝者，請先完成或暫緩');setLive_('activePersonId',p.personId,op);mine.filter(t=>t.status!=='已完成').forEach(t=>write_('tasks',{status:'拍攝中',updatedAt:now_()},t._row));}
 }else if(d.action==='saveTask'){
  const t=mine.find(t=>t.taskId===d.taskId);if(!t)throw Error('找不到任務');const items=taskItems_(t);if(!Array.isArray(d.checked)||d.checked.some(i=>!items.includes(i)))throw Error('拍攝項目不正確');const checked=Array.from(new Set(d.checked));if(t.status==='已完成'&&checked.length!==items.length)throw Error('請先標記補拍再修改已完成任務');write_('tasks',{checkedItems:JSON.stringify(checked),notes:text_(d.notes).slice(0,2000),updatedAt:now_()},t._row);
 }else if(d.action==='finish'){
  if(live.activePersonId!==p.personId)throw Error('請先將此人設為棚內拍攝者');if(!mine.length||mine.some(t=>checked_(t).length!==taskItems_(t).length))throw Error('請先勾選所有拍攝項目');mine.forEach(t=>write_('tasks',{status:'已完成',updatedAt:now_()},t._row));setLive_('activePersonId','',op);if(live.calledPersonId===p.personId)setLive_('calledPersonId','',op);
 }else if(d.action==='hold'||d.action==='retake'){
  ready_(p,mine,checkins);const targets=d.action==='retake'?mine:mine.filter(t=>t.status!=='已完成');targets.forEach(t=>write_('tasks',{status:d.action==='retake'?'待補拍':'暫緩',checkedItems:d.action==='retake'?'[]':t.checkedItems,updatedAt:now_()},t._row));if(live.activePersonId===p.personId)setLive_('activePersonId','',op);if(live.calledPersonId===p.personId)setLive_('calledPersonId','',op);
 }else if(d.action==='requeue'){
  ready_(p,mine,checkins);mine.filter(t=>['暫緩','待補拍'].includes(t.status)).forEach(t=>write_('tasks',{status:'待拍攝',updatedAt:now_()},t._row));
 }else throw Error('不支援的操作');
 logAction_(d.action,p.personId,op,{taskId:text_(d.taskId)});return snapshot_();
 }finally{lock.releaseLock();}
}
function testPhotoConnection(){const b=book_();console.log(JSON.stringify({version:VERSION,title:b.getName(),people:rows_('people').length,tasks:rows_('tasks').length,configured:!!PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD')}));}
