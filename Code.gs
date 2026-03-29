// ============================================================
// HIRING DASHBOARD — Code.gs
// ============================================================
// Paste this entire file into your Google Apps Script editor.
// File name must be: Code.gs
// ============================================================

var SN_CAND='Candidates',SN_POS='Positions',SN_LOG='Activity Log',SN_EXP='Expenses',SN_SET='Settings',SN_CD='Call Data',SN_WF='Workforce';
var DASHBOARD_PASSWORD='YOUR_PASSWORD_HERE'; // ← Set your own password before deploying
function verifyPassword(pw){return pw===DASHBOARD_PASSWORD}

// ============ INIT ============
function initializeSheet(){
  var ss=SpreadsheetApp.getActiveSpreadsheet(),sh;
  sh=ss.getSheetByName(SN_CAND);
  if(!sh){sh=ss.insertSheet(SN_CAND);sh.appendRow(['ID','Name','Phone','Email','Position','Experience','Key Skills','Current Company','Education','Status','Remarks','Called By','Call Date','CV Link','Upload Date','Uploaded By','AI Summary']);sh.getRange(1,1,1,17).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#fff');sh.setFrozenRows(1);sh.getRange('C:C').setNumberFormat('@')}
  sh=ss.getSheetByName(SN_POS);
  if(!sh){sh=ss.insertSheet(SN_POS);sh.appendRow(['Position Name','Department','Status','Created Date']);sh.getRange(1,1,1,4).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#fff');[['Customer Support Executive','Operations'],['Graphic Designer','Marketing'],['Warehouse Executive','Operations'],['Social Media Manager','Marketing'],['Marketplace Manager','E-commerce']].forEach(function(r){sh.appendRow([r[0],r[1],'Active',new Date()])})}
  sh=ss.getSheetByName(SN_LOG);
  if(!sh){sh=ss.insertSheet(SN_LOG);sh.appendRow(['Timestamp','User','Action','Candidate ID','Details']);sh.getRange(1,1,1,5).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#fff')}
  sh=ss.getSheetByName(SN_EXP);
  if(!sh){sh=ss.insertSheet(SN_EXP);sh.appendRow(['ID','Date','Platform','Amount','Description','Added By','Added On']);sh.getRange(1,1,1,7).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#fff')}
  sh=ss.getSheetByName(SN_SET);
  if(!sh){sh=ss.insertSheet(SN_SET);sh.appendRow(['Type','Value','Extra']);sh.getRange(1,1,1,3).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#fff');
    ['New','Shortlisted','Called','Call Not Received','On Hold','Shortlisted - HR','Interview Scheduled','Rejected','Offered'].forEach(function(s){sh.appendRow(['STATUS',s,''])});
    ['LinkedIn','Indeed','Naukri','Internshala','Referral','Walk-in','Other'].forEach(function(s){sh.appendRow(['PLATFORM',s,''])});
    sh.appendRow(['CALLFIELD','From Delhi?','yesno']);sh.appendRow(['CALLFIELD','Last Salary (CTC)','text']);sh.appendRow(['CALLFIELD','Expected Salary','text']);sh.appendRow(['CALLFIELD','Notice Period','dropdown:Immediate,15 days,30 days,60 days,90 days']);sh.appendRow(['CALLFIELD','Available for Interview?','yesno'])}
  sh=ss.getSheetByName(SN_CD);
  if(!sh){sh=ss.insertSheet(SN_CD);sh.appendRow(['Candidate ID','Field Name','Field Value','Updated By','Updated On']);sh.getRange(1,1,1,5).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#fff')}
  sh=ss.getSheetByName(SN_WF);
  if(!sh){sh=ss.insertSheet(SN_WF);sh.appendRow(['ID','Name','Phone','Email','Position','Stage','Salary','Join Date','End Date','Remarks','Transferred By','Transfer Date','Original Candidate ID','CV Link']);sh.getRange(1,1,1,14).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#fff');sh.setFrozenRows(1);sh.getRange('C:C').setNumberFormat('@')}
  return 'Initialized!'
}

function fixPhoneColumn(){var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SN_CAND);if(!sh)return;sh.getRange('C:C').setNumberFormat('@');if(sh.getLastRow()<2)return;sh.getRange(2,3,sh.getLastRow()-1,1).getValues().forEach(function(r,i){if(r[0]&&typeof r[0]==='number')sh.getRange(i+2,3).setValue(String(r[0]))})}

function doGet(){return HtmlService.createHtmlOutputFromFile('Index').setTitle('Hiring Dashboard').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport','width=device-width, initial-scale=1')}

// ============ ATOMIC ID GENERATION ============
function getNextCandidateId_(){
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SN_CAND);
    var nextId = 1;
    if (sh.getLastRow() >= 2) {
      var ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().flat().filter(Number);
      if (ids.length) nextId = Math.max.apply(null, ids) + 1;
    }
    var wsh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SN_WF);
    if (wsh && wsh.getLastRow() >= 2) {
      var wIds = wsh.getRange(2, 13, wsh.getLastRow() - 1, 1).getValues().flat().filter(Number);
      if (wIds.length) {
        var wMax = Math.max.apply(null, wIds) + 1;
        if (wMax > nextId) nextId = wMax;
      }
    }
    var nr = sh.getLastRow() + 1;
    sh.getRange(nr, 1).setValue(nextId);
    SpreadsheetApp.flush();
    return { id: nextId, row: nr };
  } finally {
    lock.releaseLock();
  }
}

function getNextId(){
  var res = getNextCandidateId_();
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SN_CAND);
  sh.deleteRow(res.row);
  return res.id;
}

// ============ FIX DUPLICATE IDs ============
function fixDuplicateIds(){
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SN_CAND);
  if (!sh || sh.getLastRow() < 2) return 'No data';
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var data = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    var seen = {}, maxId = 0, dupes = [];
    for (var i = 0; i < data.length; i++) {
      var id = Number(data[i][0]);
      if (id > maxId) maxId = id;
      if (seen[id]) { dupes.push(i + 2); } else { seen[id] = true; }
    }
    var newId = maxId + 1;
    for (var j = 0; j < dupes.length; j++) { sh.getRange(dupes[j], 1).setValue(newId); newId++; }
    SpreadsheetApp.flush();
    return 'Fixed ' + dupes.length + ' duplicate IDs. New max ID: ' + (newId - 1);
  } finally { lock.releaseLock(); }
}

// ============ SETTINGS ============
function getSettings(){
  var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SN_SET);
  if(!sh||sh.getLastRow()<2)return{statuses:[],platforms:[],callFields:[]};
  var d=sh.getRange(2,1,sh.getLastRow()-1,3).getValues();
  return{statuses:d.filter(function(r){return r[0]==='STATUS'}).map(function(r){return r[1]}),platforms:d.filter(function(r){return r[0]==='PLATFORM'}).map(function(r){return r[1]}),callFields:d.filter(function(r){return r[0]==='CALLFIELD'}).map(function(r){return{label:r[1],fieldType:r[2]||'text'}})}
}
function addSetting(type,value,extra){SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SN_SET).appendRow([type,value,extra||'']);return{success:true}}
function removeSetting(type,value){var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SN_SET),d=sh.getDataRange().getValues();for(var i=d.length-1;i>=1;i--){if(d[i][0]===type&&d[i][1]===value){sh.deleteRow(i+1);return{success:true}}}return{success:false}}

// ============ POSITIONS ============
function getPositions(){var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SN_POS);if(!sh||sh.getLastRow()<2)return[];return sh.getRange(2,1,sh.getLastRow()-1,3).getValues().filter(function(r){return r[2]==='Active'}).map(function(r){return{name:r[0],department:r[1]}})}
function addPosition(n,d){SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SN_POS).appendRow([n,d||'','Active',new Date()]);return{success:true}}
function removePosition(n){var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SN_POS),d=sh.getDataRange().getValues();for(var i=1;i<d.length;i++){if(d[i][0]===n){sh.getRange(i+1,3).setValue('Inactive');return{success:true}}}return{success:false}}

// ============ CANDIDATES ============
function getCandidates(filters){
  var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SN_CAND);if(!sh||sh.getLastRow()<2)return[];
  var tz=Session.getScriptTimeZone(),d=sh.getRange(2,1,sh.getLastRow()-1,17).getValues();
  var c=d.map(function(r){return{id:r[0],name:r[1],phone:String(r[2]||''),email:r[3],position:r[4],experience:r[5],keySkills:r[6],currentCompany:r[7],education:r[8],status:r[9]||'New',remarks:r[10]||'',calledBy:r[11],callDate:r[12]?Utilities.formatDate(new Date(r[12]),tz,'dd MMM yyyy'):'',cvLink:r[13],uploadDate:r[14]?Utilities.formatDate(new Date(r[14]),tz,'dd MMM yyyy HH:mm'):'',uploadedBy:r[15],aiSummary:r[16]}});
  if(filters){if(filters.position&&filters.position!=='All')c=c.filter(function(x){return x.position===filters.position});if(filters.status&&filters.status!=='All')c=c.filter(function(x){return x.status===filters.status});if(filters.search){var s=filters.search.toLowerCase();c=c.filter(function(x){return(x.name||'').toLowerCase().indexOf(s)!==-1||(x.phone||'').indexOf(s)!==-1||(x.email||'').toLowerCase().indexOf(s)!==-1||(x.keySkills||'').toLowerCase().indexOf(s)!==-1})}}
  return c;
}

function checkDuplicate(phone,email){var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SN_CAND);if(!sh||sh.getLastRow()<2)return{duplicate:false};var d=sh.getRange(2,1,sh.getLastRow()-1,4).getValues(),cp=String(phone||'').replace(/[\s\-\(\)\+]/g,'').slice(-10);for(var i=0;i<d.length;i++){var ep=String(d[i][2]||'').replace(/[\s\-\(\)\+]/g,'').slice(-10);if(cp.length>=10&&ep===cp)return{duplicate:true,field:'phone',existingName:d[i][1],existingId:d[i][0]};var ee=String(d[i][3]||'').toLowerCase().trim();if(email&&email.trim()&&ee===email.toLowerCase().trim()&&ee!=='n/a')return{duplicate:true,field:'email',existingName:d[i][1],existingId:d[i][0]}}return{duplicate:false}}

function updateCandidateStatus(cid,status,remark,calledBy){var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SN_CAND),d=sh.getDataRange().getValues();for(var i=1;i<d.length;i++){if(String(d[i][0])==String(cid)){var row=i+1;sh.getRange(row,10).setValue(status);if(remark&&remark.trim()){var ex=d[i][10]||'',ts=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'dd MMM HH:mm');sh.getRange(row,11).setValue(ex?(ex+'\n['+ts+' | '+calledBy+'] '+remark.trim()):('['+ts+' | '+calledBy+'] '+remark.trim()))}sh.getRange(row,12).setValue(calledBy);if(['Called','Interview Scheduled','Rejected','Offered'].indexOf(status)!==-1)sh.getRange(row,13).setValue(new Date());logActivity(calledBy,'Status → '+status,cid,remark||'');return{success:true}}}return{success:false,error:'Not found'}}

function addRemarkOnly(cid,remark,user){var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SN_CAND),d=sh.getDataRange().getValues();for(var i=1;i<d.length;i++){if(String(d[i][0])==String(cid)){var ex=d[i][10]||'',ts=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'dd MMM HH:mm');sh.getRange(i+1,11).setValue(ex?(ex+'\n['+ts+' | '+user+'] '+remark.trim()):('['+ts+' | '+user+'] '+remark.trim()));logActivity(user,'Remark',cid,remark);return{success:true}}}return{success:false}}

function editRemark(cid,oldRemark,newRemark,user){var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SN_CAND),d=sh.getDataRange().getValues();for(var i=1;i<d.length;i++){if(String(d[i][0])==String(cid)){var full=d[i][10]||'';var updated=full.replace(oldRemark,newRemark+' [edited by '+user+']');sh.getRange(i+1,11).setValue(updated);logActivity(user,'Remark Edited',cid,'');return{success:true}}}return{success:false}}

function deleteRemarkByIndex(cid, remarkIndex, user) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SN_CAND);
  var d = sh.getDataRange().getValues();
  for (var i = 1; i < d.length; i++) {
    if (String(d[i][0]) == String(cid)) {
      var full = d[i][10] || '';
      var lines = full.split('\n').filter(function(x) { return x.trim(); });
      if (remarkIndex >= 0 && remarkIndex < lines.length) {
        lines.splice(remarkIndex, 1);
        sh.getRange(i + 1, 11).setValue(lines.join('\n'));
        logActivity(user, 'Remark Deleted', cid, '');
        return { success: true };
      }
    }
  }
  return { success: false };
}

// ============ CALL DATA ============
function getCallData(cid){var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SN_CD);if(!sh||sh.getLastRow()<2)return{};var d=sh.getRange(2,1,sh.getLastRow()-1,3).getValues(),r={};d.forEach(function(x){if(String(x[0])==String(cid))r[x[1]]=x[2]});return r}
function saveCallField(cid,fn,fv,user){var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SN_CD),d=sh.getDataRange().getValues();for(var i=1;i<d.length;i++){if(String(d[i][0])==String(cid)&&d[i][1]===fn){sh.getRange(i+1,3).setValue(fv);sh.getRange(i+1,4).setValue(user);sh.getRange(i+1,5).setValue(new Date());return{success:true}}}sh.appendRow([cid,fn,fv,user,new Date()]);return{success:true}}

// ============ CV PROCESSING (AI) ============
function processCVWithAI(b64,fileName,position,uploadedBy,mimeType){
  try{
    var apiKey=PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
    if(!apiKey)return{success:false,error:'ANTHROPIC_API_KEY not set in Script Properties'};
    var pdfB64=b64;
    if(mimeType&&mimeType!=='application/pdf'){var blob=Utilities.newBlob(Utilities.base64Decode(b64),mimeType,fileName);var tf=DriveApp.createFile(blob);var df=Drive.Files.copy({title:fileName+'_temp',mimeType:'application/vnd.google-apps.document'},tf.getId());var pb=DriveApp.getFileById(df.id).getAs('application/pdf');pdfB64=Utilities.base64Encode(pb.getBytes());DriveApp.getFileById(df.id).setTrashed(true);tf.setTrashed(true)}
    var resp=UrlFetchApp.fetch('https://api.anthropic.com/v1/messages',{method:'post',contentType:'application/json',headers:{'x-api-key':apiKey,'anthropic-version':'2023-06-01'},payload:JSON.stringify({model:'claude-sonnet-4-20250514',max_tokens:1024,messages:[{role:'user',content:[{type:'document',source:{type:'base64',media_type:'application/pdf',data:pdfB64}},{type:'text',text:'Extract from this Indian CV. PHONE: 10-digit starting 6/7/8/9, add +91. Search everywhere. Return ONLY raw JSON: {"name":"","phone":"+91XXXXXXXXXX","email":"","experience":"","key_skills":"top 5-6","current_company":"","education":"","summary":"2-3 lines"}'}]}]}),muteHttpExceptions:true});
    var result=JSON.parse(resp.getContentText());if(result.error)return{success:false,error:result.error.message};
    var ext=JSON.parse(result.content[0].text.replace(/```json\n?/g,'').replace(/```\n?/g,'').trim());
    var phone=String(ext.phone||'N/A').replace(/[\s\-\(\)]/g,'');if(/^\d{10}$/.test(phone))phone='+91'+phone;if(/^91\d{10}$/.test(phone))phone='+'+phone;
    var dup=checkDuplicate(phone,ext.email);if(dup.duplicate)return{success:false,error:'DUPLICATE: '+ext.name+' matches "'+dup.existingName+'" (#'+dup.existingId+') by '+dup.field};
    var cvBlob=Utilities.newBlob(Utilities.base64Decode(b64),mimeType||'application/pdf',fileName);
    var folder=getCVFolder(position);
    var file=folder.createFile(cvBlob);file.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);
    var reserved = getNextCandidateId_();
    var nid = reserved.id, nr = reserved.row;
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SN_CAND);
    sh.getRange(nr,2).setValue(ext.name||'N/A');
    sh.getRange(nr,3).setNumberFormat('@').setValue(phone);
    sh.getRange(nr,4).setValue(ext.email||'N/A');
    sh.getRange(nr,5).setValue(position);
    sh.getRange(nr,6).setValue(ext.experience||'N/A');
    sh.getRange(nr,7).setValue(ext.key_skills||'N/A');
    sh.getRange(nr,8).setValue(ext.current_company||'N/A');
    sh.getRange(nr,9).setValue(ext.education||'N/A');
    sh.getRange(nr,10).setValue('New');
    sh.getRange(nr,14).setValue(file.getUrl());
    sh.getRange(nr,15).setValue(new Date());
    sh.getRange(nr,16).setValue(uploadedBy);
    sh.getRange(nr,17).setValue(ext.summary||'N/A');
    SpreadsheetApp.flush();
    logActivity(uploadedBy,'CV Upload',nid,ext.name+' → '+position);
    return{success:true,candidate:{id:nid,name:ext.name,phone:phone,email:ext.email,position:position}}
  }catch(e){return{success:false,error:e.toString()}}
}

// ============ RETRY AI SUMMARY ============
function retryCVSummary(candidateId) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SN_CAND);
  var d = sh.getDataRange().getValues();
  var rowIdx = -1;
  for (var i = 1; i < d.length; i++) {
    if (String(d[i][0]) == String(candidateId)) { rowIdx = i; break; }
  }
  if (rowIdx === -1) return { success: false, error: 'Candidate not found' };
  var cvLink = d[rowIdx][13];
  if (!cvLink) return { success: false, error: 'No CV link found for this candidate' };
  var match = cvLink.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (!match) match = cvLink.match(/id=([a-zA-Z0-9_-]+)/);
  if (!match) return { success: false, error: 'Cannot extract file ID from CV link' };
  var fileId = match[1];
  var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) return { success: false, error: 'ANTHROPIC_API_KEY not set' };
  try {
    var file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    var mimeType = blob.getContentType();
    var b64;
    if (mimeType === 'application/pdf') {
      b64 = Utilities.base64Encode(blob.getBytes());
    } else {
      var tf = DriveApp.createFile(blob);
      var df = Drive.Files.copy({title: file.getName() + '_temp', mimeType: 'application/vnd.google-apps.document'}, tf.getId());
      var pb = DriveApp.getFileById(df.id).getAs('application/pdf');
      b64 = Utilities.base64Encode(pb.getBytes());
      DriveApp.getFileById(df.id).setTrashed(true);
      tf.setTrashed(true);
    }
    var resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post', contentType: 'application/json',
      headers: {'x-api-key': apiKey, 'anthropic-version': '2023-06-01'},
      payload: JSON.stringify({
        model: 'claude-sonnet-4-20250514', max_tokens: 1024,
        messages: [{role: 'user', content: [
          {type: 'document', source: {type: 'base64', media_type: 'application/pdf', data: b64}},
          {type: 'text', text: 'Extract from this Indian CV. PHONE: 10-digit starting 6/7/8/9, add +91. Search everywhere. Return ONLY raw JSON: {"name":"","phone":"+91XXXXXXXXXX","email":"","experience":"","key_skills":"top 5-6","current_company":"","education":"","summary":"2-3 lines"}'}
        ]}]
      }),
      muteHttpExceptions: true
    });
    var result = JSON.parse(resp.getContentText());
    if (result.error) return { success: false, error: result.error.message };
    var ext = JSON.parse(result.content[0].text.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim());
    var row = rowIdx + 1;
    var current = d[rowIdx];
    if (!current[1] || current[1] === 'N/A') sh.getRange(row, 2).setValue(ext.name || 'N/A');
    var phone = String(ext.phone || '').replace(/[\s\-\(\)]/g, '');
    if (/^\d{10}$/.test(phone)) phone = '+91' + phone;
    if (/^91\d{10}$/.test(phone)) phone = '+' + phone;
    if (!current[2] || String(current[2]).replace(/\D/g,'').length < 10) sh.getRange(row, 3).setNumberFormat('@').setValue(phone);
    if (!current[3] || current[3] === 'N/A') sh.getRange(row, 4).setValue(ext.email || 'N/A');
    if (!current[5] || current[5] === 'N/A') sh.getRange(row, 6).setValue(ext.experience || 'N/A');
    if (!current[6] || current[6] === 'N/A') sh.getRange(row, 7).setValue(ext.key_skills || 'N/A');
    if (!current[7] || current[7] === 'N/A') sh.getRange(row, 8).setValue(ext.current_company || 'N/A');
    if (!current[8] || current[8] === 'N/A') sh.getRange(row, 9).setValue(ext.education || 'N/A');
    sh.getRange(row, 17).setValue(ext.summary || 'N/A');
    SpreadsheetApp.flush();
    logActivity('System', 'AI Retry', candidateId, 'Re-extracted CV data');
    return { success: true, data: ext };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ============ CV FOLDER MANAGEMENT ============
function getCVFolder(position){
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var props=PropertiesService.getScriptProperties();
    var rootId=props.getProperty('CV_ROOT_FOLDER_ID');
    var root;
    if(rootId){try{root=DriveApp.getFolderById(rootId)}catch(e){root=DriveApp.createFolder('Hiring CVs');props.setProperty('CV_ROOT_FOLDER_ID',root.getId())}}
    else{root=DriveApp.createFolder('Hiring CVs');props.setProperty('CV_ROOT_FOLDER_ID',root.getId())}
    var posFolder;var iter=root.getFoldersByName(position);
    if(iter.hasNext()){posFolder=iter.next()}else{posFolder=root.createFolder(position)}
    return posFolder;
  } finally { lock.releaseLock(); }
}

function getCVFolders(){
  var props=PropertiesService.getScriptProperties();
  var rootId=props.getProperty('CV_ROOT_FOLDER_ID');
  if(!rootId)return{rootUrl:'',folders:[]};
  try{
    var root=DriveApp.getFolderById(rootId);
    var folderMap={},iter=root.getFolders();
    while(iter.hasNext()){
      var f=iter.next();
      var name=f.getName();
      if(folderMap[name]){
        var dupeFiles=f.getFiles();
        while(dupeFiles.hasNext()){var df=dupeFiles.next();folderMap[name].folder.addFile(df);f.removeFile(df)}
        f.setTrashed(true);
      } else {
        var fi=f.getFiles();var count=0;while(fi.hasNext()){fi.next();count++}
        folderMap[name]={folder:f,url:f.getUrl(),fileCount:count};
      }
    }
    var folders=[];
    for(var k in folderMap){
      var fi2=folderMap[k].folder.getFiles();var c2=0;while(fi2.hasNext()){fi2.next();c2++}
      folders.push({name:k,url:folderMap[k].url,fileCount:c2});
    }
    return{rootUrl:root.getUrl(),folders:folders}
  }catch(e){return{rootUrl:'',folders:[]}}
}

// ============ SELF-SUBMIT (Candidate form) ============
function candidateSelfSubmit(b64,fileName,mimeType,name,phone,email,position){
  var dup=checkDuplicate(phone,email);if(dup.duplicate)return{success:false,error:'A candidate with this '+dup.field+' already exists.'};
  var blob=Utilities.newBlob(Utilities.base64Decode(b64),mimeType||'application/pdf',fileName);
  var folder=getCVFolder(position);var file=folder.createFile(blob);file.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);
  var cp=String(phone||'').replace(/[\s\-\(\)]/g,'');if(/^\d{10}$/.test(cp))cp='+91'+cp;
  var reserved = getNextCandidateId_();
  var nid = reserved.id, nr = reserved.row;
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SN_CAND);
  sh.getRange(nr,2).setValue(name);
  sh.getRange(nr,3).setNumberFormat('@').setValue(cp);
  sh.getRange(nr,4).setValue(email||'');
  sh.getRange(nr,5).setValue(position);
  sh.getRange(nr,10).setValue('New');
  sh.getRange(nr,14).setValue(file.getUrl());
  sh.getRange(nr,15).setValue(new Date());
  sh.getRange(nr,16).setValue('Self-Submit');
  SpreadsheetApp.flush();
  logActivity('Self-Submit','CV Received',nid,name+' → '+position);return{success:true}
}

// ============ MANUAL ADD ============
function addCandidateManually(data){
  var dup=checkDuplicate(data.phone,data.email);if(dup.duplicate)return{success:false,error:'DUPLICATE: Matches "'+dup.existingName+'" (#'+dup.existingId+') by '+dup.field};
  var reserved = getNextCandidateId_();
  var nid = reserved.id, nr = reserved.row;
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SN_CAND);
  sh.getRange(nr,2).setValue(data.name);
  sh.getRange(nr,3).setNumberFormat('@').setValue(String(data.phone||''));
  sh.getRange(nr,4).setValue(data.email||'');
  sh.getRange(nr,5).setValue(data.position);
  sh.getRange(nr,6).setValue(data.experience||'');
  sh.getRange(nr,7).setValue(data.keySkills||'');
  sh.getRange(nr,10).setValue('New');
  sh.getRange(nr,14).setValue(data.cvLink||'');
  sh.getRange(nr,15).setValue(new Date());
  sh.getRange(nr,16).setValue(data.uploadedBy||'Manual');
  SpreadsheetApp.flush();
  logActivity(data.uploadedBy||'Manual','Manual Add',nid,data.name+' → '+data.position);return{success:true,id:nid}
}

// ============ WORKFORCE ============
function transferToWorkforce(candidateId,stage,salary,joinDate,remarks,userName){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var csh=ss.getSheetByName(SN_CAND),cd=csh.getDataRange().getValues();
  var cand=null,crowIdx=-1;
  for(var i=1;i<cd.length;i++){if(String(cd[i][0])==String(candidateId)){cand=cd[i];crowIdx=i+1;break}}
  if(!cand)return{success:false,error:'Candidate not found'};
  var wsh=ss.getSheetByName(SN_WF);
  var wid=1;if(wsh.getLastRow()>=2){var ids=wsh.getRange(2,1,wsh.getLastRow()-1,1).getValues().flat().filter(Number);if(ids.length)wid=Math.max.apply(null,ids)+1}
  wsh.appendRow([wid,cand[1],String(cand[2]),cand[3],cand[4],stage,salary||'',joinDate?new Date(joinDate):new Date(),'',remarks||'',userName,new Date(),candidateId,cand[13]||'']);
  csh.deleteRow(crowIdx);
  logActivity(userName,'Transfer → '+stage,candidateId,cand[1]+' → '+stage);
  return{success:true}
}

function getWorkforce(filters){
  var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SN_WF);if(!sh||sh.getLastRow()<2)return[];
  var tz=Session.getScriptTimeZone(),d=sh.getRange(2,1,sh.getLastRow()-1,14).getValues();
  var w=d.map(function(r){return{id:r[0],name:r[1],phone:String(r[2]||''),email:r[3],position:r[4],stage:r[5],salary:r[6],joinDate:r[7]?Utilities.formatDate(new Date(r[7]),tz,'dd MMM yyyy'):'',endDate:r[8]?Utilities.formatDate(new Date(r[8]),tz,'dd MMM yyyy'):'',remarks:r[9]||'',transferredBy:r[10],transferDate:r[11]?Utilities.formatDate(new Date(r[11]),tz,'dd MMM yyyy'):'',origCandId:r[12],cvLink:r[13]||''}});
  if(filters){if(filters.stage&&filters.stage!=='All')w=w.filter(function(x){return x.stage===filters.stage});if(filters.position&&filters.position!=='All')w=w.filter(function(x){return x.position===filters.position})}
  return w;
}

function updateWorkforce(wid,stage,salary,endDate,remarks,user){
  var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SN_WF),d=sh.getDataRange().getValues();
  for(var i=1;i<d.length;i++){if(d[i][0]==wid){var row=i+1;
    if(stage)sh.getRange(row,6).setValue(stage);
    if(salary)sh.getRange(row,7).setValue(salary);
    if(endDate)sh.getRange(row,9).setValue(new Date(endDate));
    if(remarks){var ex=d[i][9]||'';var ts=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'dd MMM HH:mm');sh.getRange(row,10).setValue(ex?(ex+'\n['+ts+' | '+user+'] '+remarks):('['+ts+' | '+user+'] '+remarks))}
    logActivity(user,'Workforce Update',wid,stage||'');return{success:true}}}
  return{success:false}
}

function getWorkforceStats(){
  var wf=getWorkforce();
  return{total:wf.length,training:wf.filter(function(x){return x.stage==='Training'}).length,probation:wf.filter(function(x){return x.stage==='Probation'}).length,intern:wf.filter(function(x){return x.stage==='Intern (Active)'}).length,internCompleted:wf.filter(function(x){return x.stage==='Intern (Completed)'}).length,hired:wf.filter(function(x){return x.stage==='Full-time Hired'}).length,rejected:wf.filter(function(x){return x.stage==='Rejected'}).length}
}

// ============ SHEET URL ============
function getSheetUrl(){return SpreadsheetApp.getActiveSpreadsheet().getUrl()}

// ============ REPORTS ============
function getReportData(startDate,endDate){
  var ss=SpreadsheetApp.getActiveSpreadsheet(),tz=Session.getScriptTimeZone();
  var start=new Date(startDate);start.setHours(0,0,0,0);var end=new Date(endDate);end.setHours(23,59,59,999);
  var csh=ss.getSheetByName(SN_CAND),cpd={},ta=0,byPos={};
  if(csh&&csh.getLastRow()>=2){csh.getRange(2,1,csh.getLastRow()-1,17).getValues().forEach(function(r){if(r[14]){var d=new Date(r[14]);if(d>=start&&d<=end){var k=Utilities.formatDate(d,tz,'yyyy-MM-dd');cpd[k]=(cpd[k]||0)+1;ta++;var pos=r[4]||'Unknown';byPos[pos]=(byPos[pos]||0)+1}}})}
  var sc={},lsh=ss.getSheetByName(SN_LOG);
  if(lsh&&lsh.getLastRow()>=2){lsh.getRange(2,1,lsh.getLastRow()-1,5).getValues().forEach(function(r){if(r[0]){var d=new Date(r[0]);if(d>=start&&d<=end){var a=String(r[2]||'');if(a.indexOf('Status →')===0){var s=a.replace('Status → ','').trim();sc[s]=(sc[s]||0)+1}}}})}
  var te=0,ebp={},esh=ss.getSheetByName(SN_EXP);
  if(esh&&esh.getLastRow()>=2){esh.getRange(2,1,esh.getLastRow()-1,6).getValues().forEach(function(r){if(r[1]){var d=new Date(r[1]);if(d>=start&&d<=end){var a=Number(r[3])||0;te+=a;var p=r[2]||'Other';ebp[p]=(ebp[p]||0)+a}}})}
  return{cvsPerDay:cpd,totalAdded:ta,statusChanges:sc,totalExpense:te,expByPlatform:ebp,cvsByPosition:byPos}
}

// ============ FILTERED ACTIVITY LOG ============
function getActivityFiltered(startDate, endDate, searchTerm, lim) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SN_LOG);
  if (!sh || sh.getLastRow() < 2) return [];
  var tz = Session.getScriptTimeZone();
  var d = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
  var result = [], start = startDate ? new Date(startDate) : null, end = endDate ? new Date(endDate) : null;
  if (start) start.setHours(0, 0, 0, 0);
  if (end) end.setHours(23, 59, 59, 999);
  var search = (searchTerm || '').toLowerCase(), limit = lim || 200;
  for (var i = d.length - 1; i >= 0; i--) {
    if (!d[i][0]) continue;
    var ts = new Date(d[i][0]);
    if (start && ts < start) continue;
    if (end && ts > end) continue;
    if (search) {
      var combined = (String(d[i][1]||'')+' '+String(d[i][2]||'')+' '+String(d[i][4]||'')).toLowerCase();
      if (combined.indexOf(search) === -1) continue;
    }
    result.push({timestamp:Utilities.formatDate(ts,tz,'dd MMM yyyy HH:mm'),user:d[i][1],action:d[i][2],candidateId:d[i][3],details:d[i][4]});
    if (result.length >= limit) break;
  }
  return result;
}

// ============ DASHBOARD SUMMARY ============
function getDashboardSummary() {
  var stats = getDashboardStats(), activity = getRecentActivity(8), wfStats = getWorkforceStats();
  var now = new Date(), monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
  var totalExpMonth = 0, expByPlatform = {};
  var esh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SN_EXP);
  if (esh && esh.getLastRow() >= 2) {
    esh.getRange(2,1,esh.getLastRow()-1,6).getValues().forEach(function(r){
      if(r[1]){var d=new Date(r[1]);if(d>=monthStart&&d<=now){var a=Number(r[3])||0;totalExpMonth+=a;var p=r[2]||'Other';expByPlatform[p]=(expByPlatform[p]||0)+a}}
    });
  }
  var weekStart = new Date(now); weekStart.setDate(now.getDate()-now.getDay()); weekStart.setHours(0,0,0,0);
  var today = new Date(); today.setHours(0,0,0,0);
  var weekCVs = 0, todayCVs = 0;
  var csh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SN_CAND);
  if (csh && csh.getLastRow() >= 2) {
    csh.getRange(2,15,csh.getLastRow()-1,1).getValues().forEach(function(r){
      if(r[0]){var d=new Date(r[0]);if(d>=weekStart)weekCVs++;if(d>=today)todayCVs++}
    });
  }
  return {stats:stats, activity:activity, wfStats:wfStats, totalExpMonth:totalExpMonth, expByPlatform:expByPlatform, weekCVs:weekCVs, todayCVs:todayCVs};
}

// ============ EXPENSES ============
function getExpenses(s,e){var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SN_EXP);if(!sh||sh.getLastRow()<2)return[];var tz=Session.getScriptTimeZone(),st=new Date(s),en=new Date(e);st.setHours(0,0,0,0);en.setHours(23,59,59,999);return sh.getRange(2,1,sh.getLastRow()-1,7).getValues().filter(function(r){var d=new Date(r[1]);return d>=st&&d<=en}).map(function(r){return{id:r[0],date:Utilities.formatDate(new Date(r[1]),tz,'dd MMM yyyy'),platform:r[2],amount:r[3],description:r[4],addedBy:r[5]}}).reverse()}
function addExpense(d){var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SN_EXP),mx=0;if(sh.getLastRow()>=2){var ids=sh.getRange(2,1,sh.getLastRow()-1,1).getValues().flat().filter(Number);if(ids.length)mx=Math.max.apply(null,ids)}sh.appendRow([mx+1,new Date(d.date),d.platform,Number(d.amount),d.description,d.addedBy,new Date()]);logActivity(d.addedBy,'Expense','','₹'+d.amount+' on '+d.platform);return{success:true}}
function deleteExpense(id){var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SN_EXP),d=sh.getDataRange().getValues();for(var i=1;i<d.length;i++){if(d[i][0]==id){sh.deleteRow(i+1);return{success:true}}}return{success:false}}

// ============ DASHBOARD STATS ============
function getDashboardStats(){
  var c=getCandidates(),p=getPositions(),s=getSettings();
  var st={total:c.length,byStatus:{},byPosition:{},positions:p.map(function(x){return x.name})};
  s.statuses.forEach(function(x){st.byStatus[x]=c.filter(function(y){return y.status===x}).length});
  p.forEach(function(x){var pc=c.filter(function(y){return y.position===x.name}),ps={total:pc.length};s.statuses.forEach(function(z){ps[z]=pc.filter(function(y){return y.status===z}).length});st.byPosition[x.name]=ps});
  return st;
}

// ============ ACTIVITY LOG ============
function logActivity(u,a,cid,det){var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SN_LOG);if(sh)sh.appendRow([new Date(),u,a,cid,det])}
function getRecentActivity(lim){var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SN_LOG);if(!sh||sh.getLastRow()<2)return[];var tz=Session.getScriptTimeZone(),n=Math.min(lim||20,sh.getLastRow()-1);return sh.getRange(sh.getLastRow()-n+1,1,n,5).getValues().reverse().map(function(r){return{timestamp:Utilities.formatDate(new Date(r[0]),tz,'dd MMM HH:mm'),user:r[1],action:r[2],candidateId:r[3],details:r[4]}})}

// ============ MISC ============
function deleteCandidate(id){var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SN_CAND),d=sh.getDataRange().getValues();for(var i=1;i<d.length;i++){if(String(d[i][0])==String(id)){sh.deleteRow(i+1);return{success:true}}}return{success:false}}
function bulkUpdateStatus(ids,st,rem,cb){return ids.map(function(id){return updateCandidateStatus(id,st,rem,cb)})}
