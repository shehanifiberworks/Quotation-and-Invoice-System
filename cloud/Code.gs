/** SFG quotation / invoice Google Sheets + Drive backend.
 * Required script properties:
 *   SFG_SHEET_ID: ID of your Google Spreadsheet
 *   SFG_ACCESS_KEY: long random key (at least 16 characters)
 * Optional SFG_FOLDER_ID: dedicated Google Drive folder for saved JSON records
 * Deploy as a web app: Execute as Me, with access limited as appropriate.
 * IMPORTANT: Anyone with your deployed URL AND access key can read/change records.
 * Do not store the access key in source code or public version control.
 */
const SFG_HEADERS=['key','docType','qno','rdno','subject','date','savedAt','paid','paidDate','driveFileId','lastRequestId','deleted'];
function sfgProps_(){const p=PropertiesService.getScriptProperties();const id=p.getProperty('SFG_SHEET_ID'),secret=p.getProperty('SFG_ACCESS_KEY');if(!id||!secret||secret.length<16)throw Error('Configure SFG_SHEET_ID and SFG_ACCESS_KEY script properties');return {id,secret};}
function sfgAccess_(token) {const p=sfgProps_();if(String(token||'')!==p.secret)throw Error('Access denied');return p;}
function sfgSheet_(){const p=sfgProps_(), book=SpreadsheetApp.openById(p.id);let sheet=book.getSheetByName('Documents');if(!sheet)sheet=book.insertSheet('Documents');if(sheet.getLastRow()===0)sheet.appendRow(SFG_HEADERS);return sheet;}
function sfgFolder_(){const p=PropertiesService.getScriptProperties();let id=p.getProperty('SFG_FOLDER_ID');if(id)return DriveApp.getFolderById(id);const folder=DriveApp.createFolder('SFG Quotation Records');p.setProperty('SFG_FOLDER_ID',folder.getId());return folder;}
function sfgRows_(){const sh=sfgSheet_(),values=sh.getDataRange().getValues();return {sh, rows:values.slice(1).map((r,i)=>({row:i+2,key:String(r[0]),docType:r[1],qno:r[2],rdno:r[3],subject:r[4],date:r[5],savedAt:r[6],paid:r[7],paidDate:r[8],fileId:String(r[9]),requestId:String(r[10]),deleted:!!r[11]}))};}
function sfgReply_(obj,callback){const json=JSON.stringify(obj);const cb=String(callback||'');if(cb&&/^sfgCloudCB_[A-Za-z0-9_]{1,90}$/.test(cb))return ContentService.createTextOutput(cb+'('+json+');').setMimeType(ContentService.MimeType.JAVASCRIPT);return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);}
function sfgRead_(r){if(!r.fileId||r.deleted)return null;try{return JSON.parse(DriveApp.getFileById(r.fileId).getBlob().getDataAsString('UTF-8'));}catch(e){return null;}}
function doGet(e){const q=e&&e.parameter||{};try{sfgAccess_(q.token);const action=String(q.action||'ping');if(action==='ping')return sfgReply_({ok:true,service:'SFG Quotation Cloud v1'},q.callback);
 const x=sfgRows_(),active=x.rows.filter(r=>!r.deleted);
 if(action==='verify'){const row=x.rows.find(r=>r.key===String(q.key||'')&&r.requestId===String(q.requestId||''));return sfgReply_({ok:true,done:!!row,rejected:false},q.callback);}
 if(action==='nextNumber'){const largest=active.reduce((max,r)=>Math.max(max,Number(r.qno)||0),999);return sfgReply_({ok:true,nextNumber:String(largest+1)},q.callback);}
 if(action==='list'){const documents=active.map(r=>({key:r.key,data:sfgRead_(r)})).filter(v=>v.data);return sfgReply_({ok:true,documents},q.callback);}
 throw Error('Unknown action');
 }catch(error){return sfgReply_({ok:false,error:String(error.message||error)},q.callback);}}
function sfgNormalize_(data){if(!data||typeof data!=='object'||Array.isArray(data))throw Error('Invalid document');if(!Array.isArray(data.items)||data.items.length>50)throw Error('Invalid item count');return data;}
function sfgID_(v){return String(v==null?'':v).trim().replace(/\s+/g,' ').toUpperCase();}
function sfgCheckDuplicates_(records,key,data){const fields=[['rdno','RD No'],['qno','Quotation / Invoice No'],['pono','PR / PO No']];for(const r of records){if(r.key===key||r.deleted)continue;const d=sfgRead_(r);if(!d)continue;for(const [field,label] of fields){const a=sfgID_(data[field]),b=sfgID_(d[field]);if(a&&a.replace(/[.\-]/g,'')&&a===b)throw Error(label+' already belongs to another document');}
 const used=new Set((d.items||[]).map(i=>sfgID_(i.doNo||i.dono)).filter(Boolean));for(const i of data.items||[]){if(i.doNo&&used.has(sfgID_(i.doNo)))throw Error('D.O Number already belongs to another document');}}
}
function doPost(e){let input={};try {input=JSON.parse(e.postData.contents);sfgAccess_(input.token);if(!/^[a-zA-Z0-9_.-]{8,110}$/.test(String(input.requestId||'')))throw Error('Invalid request');const action=String(input.action||'');if(!['save','delete','backup'].includes(action))throw Error('Unsupported write');const lock=LockService.getScriptLock();lock.waitLock(25000);
 try{const x=sfgRows_(),key=String(input.key||'').trim();if(action==='backup'){const folder=sfgFolder_();const f=folder.createFile('SFG_Backup_'+Date.now()+'.json',String(input.data||''),MimeType.PLAIN_TEXT);x.sh.appendRow([String(input.key||'backup_'+Date.now()),'BACKUP','','','','',new Date().toISOString(),false,'',f.getId(),input.requestId,true]);return sfgReply_({ok:true});}if(!key||key.length>200)throw Error('Invalid document key');const existing=x.rows.find(r=>r.key===key);if(action==='delete'){if(existing){const row=x.sh.getRange(existing.row,1,1,SFG_HEADERS.length).getValues()[0];row[10]=input.requestId;row[11]=true;x.sh.getRange(existing.row,1,1,SFG_HEADERS.length).setValues([row]);}else{x.sh.appendRow([key,'','','','','','','','','',input.requestId,true]);}return sfgReply_({ok:true});}
 const data=sfgNormalize_(input.data);sfgCheckDuplicates_(x.rows,key,data);const folder=sfgFolder_(),json=JSON.stringify(data),filename='SFG_'+key.replace(/[^a-z0-9_-]/gi,'_')+'.json';let file;
 if(existing&&existing.fileId){file=DriveApp.getFileById(existing.fileId);file.setContent(json);}else{file=folder.createFile(filename,json,'application/json');}
 const values=[key,data.docType||'QUOTATION',String(data.qno||''),String(data.rdno||''),String(data.subject||''),String(data.date||''),String(data.savedAt||new Date().toISOString()),!!data.paid,String(data.paidDate||''),file.getId(),input.requestId,false];
 if(existing)x.sh.getRange(existing.row,1,1,values.length).setValues([values]);else x.sh.appendRow(values);
 return sfgReply_({ok:true});
 }finally{lock.releaseLock();}
 }catch(error){console.error(error);return sfgReply_({ok:false,error:String(error.message||error)});}}
