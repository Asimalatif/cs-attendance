const vm=require('vm'),fs=require('fs');
const order=process.argv.slice(2);
const stub=new Proxy(function(){},{get:(t,k)=>k===Symbol.toPrimitive?()=>'' :stub,apply:()=>stub,construct:()=>stub});
const ctx={console};
for(const g of ['SpreadsheetApp','HtmlService','DriveApp','PropertiesService','CacheService','Utilities','LockService','ContentService','Session','ScriptApp','UrlFetchApp','Logger','MailApp','GmailApp','DocumentApp','Sheets','Drive','CalendarApp','XmlService','Charts','Browser','FormApp','SlidesApp']) ctx[g]=stub;
vm.createContext(ctx);
for(const f of order){try{vm.runInContext(fs.readFileSync(f,'utf8'),ctx,{filename:f});}catch(e){console.log('LOAD ERROR in',f,'->',e.message)}}
console.log('done');
