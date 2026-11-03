/** 在新建的 Google 試算表 → 擴充功能 → Apps Script 使用。 */
function setup(){
 const props=PropertiesService.getScriptProperties();const ss=SpreadsheetApp.getActiveSpreadsheet();if(!ss)throw Error('請從 Google 試算表開啟 Apps Script');props.setProperty('SPREADSHEET_ID',ss.getId());if(!props.getProperty('ACCESS_KEY'))props.setProperty('ACCESS_KEY',Utilities.getUuid()+Utilities.getUuid());if(!props.getProperty('ALLOWED_ORIGIN'))props.setProperty('ALLOWED_ORIGIN','https://chimeivera.github.io');
 let sheet=ss.getSheetByName('賓客資料');if(!sheet){sheet=ss.insertSheet('賓客資料');const cols=['id','name','category','department','title','expected','staff','award','awardGroup','awardDeclined','meal','shortcode','phone','note','checked','seat','escorted','explained','notified','contact_status','operator','updated','现场'];sheet.appendRow(cols);sheet.getRange(2,1,INITIAL_ROSTER.length,cols.length).setNumberFormat('@').setValues(INITIAL_ROSTER.map(p=>cols.map(k=>p[k]===undefined?'':String(p[k]))));sheet.setFrozenRows(1)}
 if(!ss.getSheetByName('操作紀錄'))ss.insertSheet('操作紀錄').appendRow(['時間','操作','人員ID','內容']);
}
function doGet(){return HtmlService.createHtmlOutput('餐敘後端已部署。請從 GitHub 簽到網頁使用。')}
function doPost(e){let q={},status=200,data;try{q=JSON.parse(e.parameter.payload||'{}');const props=PropertiesService.getScriptProperties();authorizeRequest(q,props);data=dispatch(q)}catch(err){status=err.status||400;data={error:err.message,code:err.code||''}}const reply=JSON.stringify({requestId:q.requestId,status,data}).replace(/</g,'\\u003c');if(q.transport==='json')return ContentService.createTextOutput(reply).setMimeType(ContentService.MimeType.JSON);const origin=JSON.stringify(PropertiesService.getScriptProperties().getProperty('ALLOWED_ORIGIN')||'https://chimeivera.github.io');return HtmlService.createHtmlOutput('<script>window.top.postMessage('+reply+','+origin+');</script>').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)}
function dispatch(q){
 if(q.method==='POST'&&q.path==='/api/admin/login')return adminLogin(q.body||{});
 if(q.method==='POST'&&q.path==='/api/public/event')return publicEvent();
 if(q.method==='POST'&&q.path==='/api/public/registrations')return publicRegistrations();
 if(q.method==='POST'&&q.path==='/api/admin/event')return saveEvent(q.body||{});
 if(q.method==='POST'&&q.path==='/api/public/lookup')return publicLookup(q.body||{});
 if(q.method==='GET'&&q.path==='/api/people'){const ss=SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'));return readRoster(ss.getSheetByName('賓客資料'))}
 const lock=LockService.getScriptLock(),guest=q.path==='/api/public/checkin';
 if(!lock.tryLock(guest?3000:25000)){const e=Error('目前多人報到中，系統會自動重試，請稍候');e.status=503;e.code='busy';throw e}
 try{return operate(q)}finally{lock.releaseLock()}
}
function fail(message,code){const err=Error(message);err.status=409;err.code=code;throw err}
function readRoster(sheet){const values=sheet.getDataRange().getValues(),cols=values.shift();const rowNumbers=[];const roster=values.flatMap((r,i)=>{if(!r[0])return [];rowNumbers.push(i+2);const p={};cols.forEach((k,i)=>p[k]=String(r[i]??''));['expected','staff','awardDeclined'].forEach(k=>p[k]=p[k]==='true');['checked','escorted','explained','notified'].forEach(k=>p[k]=Number(p[k]||0));return [p]});Object.defineProperties(roster,{columns:{value:cols},rowNumbers:{value:rowNumbers}});return roster}
function writePerson(sheet,p,row,cols,keys){
 // Reuse the header captured by the roster read. Existing rows already use text format.
 const indexes=keys?keys.map(k=>cols.indexOf(k)).filter(i=>i>=0):cols.map((_,i)=>i);
 if(!indexes.length)return;
 const first=Math.min(...indexes),last=Math.max(...indexes),range=sheet.getRange(row,first+1,1,last-first+1);
 if(!keys)range.setNumberFormat('@');
 range.setValues([cols.slice(first,last+1).map(k=>p[k]===undefined?'':String(p[k]))]);
}

// Only these two public operations bypass the staff access code.
function authorizeRequest(q,props){
 if(q.origin!==props.getProperty('ALLOWED_ORIGIN'))throw Error('網站來源不正確');
 const publicAction=q.method==='POST'&&['/api/public/lookup','/api/public/checkin','/api/public/event','/api/public/registrations','/api/admin/login'].includes(q.path);
 if(!publicAction&&!validAdminToken(q.key)){const e=Error('請登入管理後台');e.status=401;e.code='auth';throw e}
}
function adminSignature(value){const props=PropertiesService.getScriptProperties();return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(value,(props.getProperty('ACCESS_KEY')||'')+'|'+(props.getProperty('ADMIN_PASSWORD')||''))).replace(/=+$/,'')}
function validAdminToken(token){if(!PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD'))return false;const p=String(token||'').split('.');return p.length===4&&p[0]==='admin'&&Number(p[1])>Date.now()&&Number(p[1])<=Date.now()+14400000&&p[3]===adminSignature(p.slice(0,3).join('.'))}
function adminLogin(body){
 const password=PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
 if(!password)throw Error('尚未設定後台密碼，請主辦人於 Apps Script 指令碼屬性設定 ADMIN_PASSWORD');
 if(typeof body.password!=='string'||body.password!==password)throw Error('密碼不正確，請重新輸入');
 const expires=Date.now()+14400000,head='admin.'+expires+'.'+Utilities.getUuid();return {token:head+'.'+adminSignature(head),expires};
}
const EVENT_FIELDS=['newDoctors','excellentTeachers','outstandingTeachers','satisfaction'];

function maskRegistrationName(name){const chars=Array.from(String(name||'').replace(/[（(][^）)]*[）)]/g,'').trim());if(chars.length<=1)return '○';if(chars.length===2)return chars[0]+'○';return chars[0]+'○'.repeat(Math.min(chars.length-2,3))+chars[chars.length-1]}
function registrationRemarks(p){const labels=[];if(['優良教師','傑出教師'].includes(p.award))labels.push(p.award);if(/新進(?:主治)?醫師/.test(p.note||''))labels.push('新進醫師');return labels.join('、')}
function publicRegistrations(){
 const cache=typeof CacheService==='undefined'?null:CacheService.getScriptCache();const cached=cache?.get('masked-registrations-v3');if(cached){try{return JSON.parse(cached)}catch{}}
 const ss=SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'));
 const people=readRoster(ss.getSheetByName('賓客資料')).filter(p=>p.expected).map(p=>({name:maskRegistrationName(p.name),department:p.department,title:p.title,category:p.category,remarks:registrationRemarks(p)}));
 const result={people};const json=JSON.stringify(result);if(json.length<85000)cache?.put('masked-registrations-v3',json,10);return result;
}

function publicEvent(){let links={};try{links=JSON.parse(PropertiesService.getScriptProperties().getProperty('EVENT_LINKS')||'{}')}catch{}return {links:Object.fromEntries(EVENT_FIELDS.map(k=>[k,typeof links[k]==='string'?links[k]:'']))}}
function saveEvent(body){const links={};for(const k of EVENT_FIELDS){const url=String(body.links?.[k]||'').trim();if(url&&(!/^https:\/\/[^\s<>"']+$/i.test(url)||url.length>1500))throw Error('請填入完整 HTTPS 連結，或留白');links[k]=url}PropertiesService.getScriptProperties().setProperty('EVENT_LINKS',JSON.stringify(links));return {ok:true,...publicEvent()}}
function publicLookup(body){
 const query=String(body.query||'').trim();if(query.length<2||query.length>80)throw Error('請輸入完整姓名或人事號');
 let roster;const cache=typeof CacheService==='undefined'?null:CacheService.getScriptCache();const cached=cache?.get('guest-lookup-v2');
 if(cached){try{roster=JSON.parse(cached)}catch{}}
 if(!roster){const ss=SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'));roster=readRoster(ss.getSheetByName('賓客資料')).map(p=>({id:p.id,name:p.name,department:p.department,title:p.title,expected:p.expected,checked:p.checked}));const text=JSON.stringify(roster);if(text.length<85000)cache?.put('guest-lookup-v2',text,15)}
 const found=roster.filter(p=>p.id.toUpperCase()===query.toUpperCase()||p.name===query);if(found.length>5)throw Error('請改用人事號查詢，或洽現場工作人員');return {people:found.map(guestIdentity)};
}
function guestSignature(id,expires){return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(id+'|'+expires,PropertiesService.getScriptProperties().getProperty('ACCESS_KEY'))).replace(/=+$/,'')}
function guestToken(id){const expires=Date.now()+600000;return encodeURIComponent(id)+'.'+expires+'.'+guestSignature(id,expires)}
function verifyGuestToken(token){const parts=String(token||'').split('.');if(parts.length!==3)throw Error('請重新查詢本人資料');const id=decodeURIComponent(parts[0]),expires=Number(parts[1]);if(!Number.isFinite(expires)||expires<Date.now()||expires>Date.now()+600000||parts[2]!==guestSignature(id,expires))throw Error('資料確認已逾時，請重新查詢');return id}
function guestIdentity(p){return {name:p.name,department:p.department,title:p.title,expected:p.expected,checked:p.checked,token:guestToken(p.id)}}
function guestReceipt(p){return {name:p.name,department:p.department,award:p.award||'',meal:p.meal||'葷',checked:p.checked,isMainTable:/^主桌(?:$|[\s　0-9０-９一二三四五六七八九十（(])/u.test(String(p.seat||'').trim())}}

function operate(q){const ss=SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID')),sheet=ss.getSheetByName('賓客資料'),roster=readRoster(sheet),b=q.body||{};const remaining=164-roster.filter(p=>p.expected).length-roster.filter(p=>!p.expected&&p.checked).length;const now=Utilities.formatDate(new Date(),'Asia/Taipei','yyyy/MM/dd HH:mm:ss');const log=(op,id,info)=>ss.getSheetByName('操作紀錄').appendRow([now,op,id,info]);
 if(q.method==='GET'&&q.path==='/api/people')return roster;
 if(q.method==='POST'&&q.path==='/api/public/lookup'){
  const query=String(b.query||'').trim();if(query.length<2||query.length>80)throw Error('請輸入完整姓名或人事號');
  const found=roster.filter(p=>p.id.toUpperCase()===query.toUpperCase()||p.name===query);
  if(found.length>5)throw Error('請改用人事號查詢，或洽現場工作人員');
  return {people:found.map(guestIdentity)};
 }
 if(q.method==='POST'&&q.path==='/api/public/checkin'){
  const id=verifyGuestToken(b.token),index=roster.findIndex(p=>p.id===id);if(index<0)throw Error('查無資料，請洽工作人員');
  const p=roster[index];if(p.checked)return {ok:true,alreadyChecked:true,person:guestReceipt(p)};
  if(!p.expected&&remaining<=0)fail('請先找院內工作人員協助確認座位，再辦理報到','capacity');
  const keys=['checked','operator','updated'];
  if(!p.expected){if(!['葷','素食','蛋奶素'].includes(b.meal))throw Error('請先確認餐別');p.meal=b.meal;keys.push('meal')}
  p.checked=1;p.operator='賓客自行報到';p.updated=now;
  writePerson(sheet,p,roster.rowNumbers[index],roster.columns,keys);log('賓客自行報到',id,p.meal||'葷');SpreadsheetApp.flush();
  return {ok:true,alreadyChecked:false,person:guestReceipt(p)};
 }

 if(q.method==='POST'&&q.path==='/api/reset'){if(b.confirmation!=='重設測試紀錄')throw Error('缺少重設確認');const stamp=Utilities.formatDate(new Date(),'Asia/Taipei','yyyyMMdd_HHmmss')+'_'+Utilities.getUuid().slice(0,5);sheet.copyTo(ss).setName('備份_'+stamp);roster.forEach(p=>{const original=INITIAL_ROSTER.find(x=>x.id===p.id);Object.assign(p,{checked:0,seat:'',note:original?.note||'',现场:'',escorted:0,explained:0,notified:0,contact_status:'',operator:'',updated:'',meal:original?.meal||p.meal})});const cols=roster.columns;if(roster.length)sheet.getRange(2,1,roster.length,cols.length).setNumberFormat('@').setValues(roster.map(p=>cols.map(k=>String(p[k]??''))));log('重設','','備份_'+stamp);SpreadsheetApp.flush();return {ok:true,backupId:stamp}}
 if(q.method==='POST'&&q.path==='/api/walkins'){const name=String(b.name||'').trim(),id=String(b.id||'').trim().toUpperCase();if(!name||name.length>80||!['葷','素食','蛋奶素'].includes(b.meal))throw Error('請確認姓名及餐別');if(roster.some(p=>p.name===name||(id&&p.id===id)))fail('名單已有此人，請搜尋後簽到','duplicate');if(remaining<=0&&!b.capacityConfirmed)fail('容量已用盡，請先確認座位','capacity');const p={id:id||'walkin-'+Utilities.getUuid(),name,category:String(b.category||'其他').slice(0,40),department:String(b.department||'').slice(0,100),title:'',expected:false,staff:false,meal:b.meal,checked:1,operator:'未指定工作人員',updated:now};writePerson(sheet,p,sheet.getLastRow()+1,roster.columns);log('臨時新增',p.id,p.meal);SpreadsheetApp.flush();return p}
 if(q.method==='PATCH'&&q.path.startsWith('/api/people/')){const id=decodeURIComponent(q.path.slice('/api/people/'.length)),index=roster.findIndex(p=>p.id===id);if(index<0)throw Error('查無此人');const p=roster[index];if(b.checked===1&&p.checked)return {ok:true,alreadyChecked:true};if(b.checked===1&&!p.expected&&remaining<=0&&!b.capacityConfirmed)fail('容量已用盡','capacity');const allowed=['checked','seat','note','meal','escorted','explained','notified','contact_status'];for(const k of allowed){if(!(k in b))continue;if(['checked','escorted','explained','notified'].includes(k)&&![0,1].includes(b[k]))throw Error('無效狀態');if(['seat','note'].includes(k)&&(typeof b[k]!=='string'||b[k].length>2000))throw Error('文字過長');if(k==='meal'&&!['葷','素食','蛋奶素'].includes(b[k]))throw Error('無效餐別');if(k==='contact_status'&&!['尚未聯繫','聯繫中','未接','稍後抵達','臨時不克'].includes(b[k]))throw Error('無效聯繫狀態');if(k==='note')p.现场=b[k];else p[k]=b[k]}
 p.operator='未指定工作人員';p.updated=now;writePerson(sheet,p,roster.rowNumbers[index],roster.columns,[...allowed.filter(k=>k in b).map(k=>k==='note'?'现场':k),'operator','updated']);log('更新',id,JSON.stringify(b));SpreadsheetApp.flush();return {ok:true,alreadyChecked:false}}
 throw Error('不支援的操作')}




