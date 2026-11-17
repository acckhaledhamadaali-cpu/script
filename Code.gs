const SPREADSHEET_ID='1RZkSWsb0uAnjYq1zje6iVIVCeKYxh3lUPSepHRVTxNU';
const S={items:'الأصناف',places:'الأماكن',users:'الصلاحيات',moves:'الحركات',pending:'الحركات_المعلقة',audit:'سجل_التدقيق',settings:'الإعدادات'};
function doGet(){return HtmlService.createHtmlOutputFromFile('Index').setTitle('شاورمتي | إدارة المخزون').addMetaTag('viewport','width=device-width, initial-scale=1');}
function ss(){return SpreadsheetApp.openById(SPREADSHEET_ID);}
function sh(n){const x=ss().getSheetByName(n);if(!x)throw Error('الورقة غير موجودة: '+n);return x;}
function head(r){const o={};r.forEach((v,i)=>o[String(v).trim()]=i);return o;}
function rows(n){const a=sh(n).getDataRange().getValues();return {a,h:head(a[0])};}
function user(){const email=(Session.getActiveUser().getEmail()||'').toLowerCase();if(!email)throw Error('تعذر تحديد بريد المستخدم. تحقق من إعدادات نشر Web App.');const {a,h}=rows(S.users);const r=a.slice(1).find(x=>String(x[h['الإيميل']]||'').trim().toLowerCase()===email);if(!r||!['نعم','yes','true','1'].includes(String(r[h['فعّال']]||'').toLowerCase()))throw Error('المستخدم غير مضاف أو غير مفعل في ورقة الصلاحيات: '+email);return {email,name:r[h['الاسم']]||email,place:r[h['المكان']]||'',role:r[h['الدور']]||'مستخدم مكان'};}
function refs(){const {a,h}=rows(S.items);const items=a.slice(1).filter(r=>r[h['اسم الصنف (فريد)']]).map(r=>({name:String(r[h['اسم الصنف (فريد)']]),unit:String(r[h['الوحدة']]||''),avgCost:Number(r[h['متوسط التكلفة (تلقائي)']]||r[h['سعر مبدئي تقديري (عدّله)']])||0}));const p=rows(S.places);return {items,places:p.a.slice(1).filter(r=>r[p.h['المكان']]&&!String(r[p.h['المكان']]).startsWith('ملاحظة:')).map(r=>({name:String(r[p.h['المكان']]),type:String(r[p.h['النوع']]||'')}))};}
function getInitialData(){const u=user(),d=refs(),p=pending(u),m=recent();const balances={};d.places.forEach(x=>balances[x.name]={qty:0,value:0});const x=rows(S.moves);x.a.slice(1).forEach(r=>{const q=Number(r[x.h['الكمية']])||0,price=Number(r[x.h['سعر الوحدة (للشراء/الافتتاحي فقط)']])||0,from=r[x.h['من (مكان)']],to=r[x.h['إلى (مكان)']];if(to&&balances[to]){balances[to].qty+=q;balances[to].value+=q*price;}if(from&&balances[from]){balances[from].qty-=q;balances[from].value-=q*price;}});const vals=Object.values(balances);return {user:u,items:d.items,places:d.places,summary:{locations:balances,totalQty:vals.reduce((s,v)=>s+v.qty,0),totalValue:vals.reduce((s,v)=>s+v.value,0)},pending:p,recent:m};}
function pending(u){
 const {a,h}=rows(S.pending),groups={};
 a.slice(1).forEach(r=>{
  if(r[h['الحالة']]!=='بانتظار اعتماد'||!(u.role==='مالك'||r[h['إلى']]===u.place))return;
  const id=String(r[h['معرف الحركة']]);if(!groups[id])groups[id]={id,date:fmtDate(r[h['التاريخ']]),type:r[h['نوع الحركة']],from:r[h['من']],to:r[h['إلى']],creator:r[h['المدخل']],notes:r[h['ملاحظات']],items:[]};
  groups[id].items.push({item:r[h['الصنف']],qty:r[h['الكمية']],unit:r[h['الوحدة']]});
 });
 return Object.values(groups);
}
function recent(){const {a,h}=rows(S.moves);return a.slice(1).filter(r=>r[h['الصنف']]).slice(-30).reverse().map(r=>({date:fmtDate(r[h['التاريخ']]),type:r[h['نوع الحركة']],from:r[h['من (مكان)']],to:r[h['إلى (مكان)']],item:r[h['الصنف']],qty:r[h['الكمية']],unit:r[h['الوحدة']]}));}
function fmtDate(d){return d instanceof Date?Utilities.formatDate(d,'Asia/Riyadh','yyyy-MM-dd'):String(d||'');}
function audit(u,act,id,detail){const x=sh(S.audit);const a=[new Date(),u,act,id,detail];while(a.length<x.getLastColumn())a.push('');x.appendRow(a);}
function available(place,item){const {a,h}=rows(S.moves);let n=0;a.slice(1).forEach(r=>{if(r[h['الصنف']]!==item)return;const q=Number(r[h['الكمية']])||0;if(r[h['إلى (مكان)']]===place)n+=q;if(r[h['من (مكان)']]===place)n-=q;});const p=rows(S.pending);p.a.slice(1).forEach(r=>{if(r[p.h['الحالة']]==='بانتظار اعتماد'&&r[p.h['الصنف']]===item&&r[p.h['من']]===place)n-=Number(r[p.h['الكمية']])||0;});return n;}
function submitMovement(f){
 const lock=LockService.getScriptLock();lock.waitLock(30000);
 try{
  const u=user(),d=refs(),type=f.type,from=String(f.from||'').trim(),to=String(f.to||'').trim(),items=Array.isArray(f.items)?f.items:[];
  if(!items.length)throw Error('أضف صنفًا واحدًا على الأقل.');
  if(type==='شراء'&&!from)throw Error('اكتب اسم المورد.');
  if(type==='تحويل'){
   if(!from||!to||from===to)throw Error('حدد المصدر والمستلم بشكل صحيح.');
   if(u.role!=='مالك'&&u.place!==from)throw Error('لا يمكنك التحويل من مكان آخر.');
  }else if(u.role!=='مالك'&&u.place!==(to||from))throw Error('الحركة خارج صلاحية مكانك.');
  if(type==='رصيد افتتاحي'&&u.role!=='مالك')throw Error('الرصيد الافتتاحي للمالك فقط.');
  const clean=items.map(x=>{
   const item=d.items.find(y=>y.name===x.item),q=Number(x.qty),price=Number(x.price)||0;
   if(!item||!(q>0))throw Error('تحقق من الصنف والكمية في جميع السطور.');
   if(['شراء','رصيد افتتاحي'].includes(type)&&!(price>=0))throw Error('تحقق من سعر الوحدة.');
   if(type==='تحويل'&&q>available(from,item.name))throw Error('الكمية المطلوبة من الصنف '+item.name+' أكبر من المتاح في '+from+'.');
   return {item,q,price};
  });
  const id='MV-'+Utilities.formatDate(new Date(),'Asia/Riyadh','yyyyMMdd-HHmmss')+'-'+Math.floor(Math.random()*900+100),ps=sh(S.pending),now=new Date();
  const rowsToAdd=clean.map(x=>{const row=[id,new Date(f.date+'T12:00:00'),type,from,to,x.item.name,x.q,x.price,x.item.unit,f.reference||'',f.notes||'',u.email,now,'بانتظار اعتماد','','','',''];while(row.length<ps.getLastColumn())row.push('');return row;});
  ps.getRange(ps.getLastRow()+1,1,rowsToAdd.length,rowsToAdd[0].length).setValues(rowsToAdd);
  if(type!=='تحويل'&&u.role==='مالك')decideMovement(id,'approve','');
  audit(u.email,'إنشاء حركة',id,type+' | '+clean.length+' أصناف');
  return {ok:true,id,status:'بانتظار اعتماد',message:'تم إرسال الحركة ('+clean.length+' أصناف) للاعتماد.'};
 }finally{lock.releaseLock();}
}
function decideMovement(id,decision,reason){
 const u=user(),x=rows(S.pending),matches=[];
 x.a.forEach((r,i)=>{if(i>0&&String(r[x.h['معرف الحركة']])===String(id))matches.push({r,i});});
 if(!matches.length)throw Error('الحركة غير موجودة.');
 const first=matches[0].r,from=first[x.h['من']],to=first[x.h['إلى']];
 if(matches.some(z=>z.r[x.h['الحالة']]!=='بانتظار اعتماد'))throw Error('تمت معالجة جزء من الطلب مسبقًا؛ راجع سجل الحركات قبل إعادة المحاولة.');
 if(u.role!=='مالك'&&u.place!==to)throw Error('الاعتماد متاح للمكان المستلم فقط.');
 if(u.role!=='مالك'&&String(first[x.h['المدخل']]).toLowerCase()===u.email)throw Error('لا يمكن اعتماد حركة أنشأتها بنفسك.');
 const shx=sh(S.pending),now=new Date();
 if(decision==='approve'){
  const mv=sh(S.moves),toAdd=[];
  matches.forEach(({r,i})=>{
   const item=r[x.h['الصنف']],qty=Number(r[x.h['الكمية']])||0;
   if(r[x.h['نوع الحركة']]==='تحويل'&&qty>available(from,item)+qty)throw Error('تعذر التحقق من الرصيد للصنف: '+item);
   shx.getRange(i+1,x.h['الحالة']+1).setValue('معتمد');shx.getRange(i+1,x.h['المعتمد']+1).setValue(u.email);shx.getRange(i+1,x.h['وقت الاعتماد']+1).setValue(now);
   const v=[r[x.h['التاريخ']],r[x.h['نوع الحركة']],from,to,item,qty,r[x.h['سعر الوحدة']],r[x.h['الوحدة']],qty*Number(r[x.h['سعر الوحدة']]||0),r[x.h['رقم المرجع']],r[x.h['المدخل']],r[x.h['ملاحظات']]];
   while(v.length<mv.getLastColumn())v.push('');toAdd.push(v);
  });
  if(toAdd.length)mv.getRange(mv.getLastRow()+1,1,toAdd.length,toAdd[0].length).setValues(toAdd);
  audit(u.email,'اعتماد حركة',id,matches.length+' أصناف');return {ok:true,message:'تم اعتماد وترحيل الطلب كاملًا ('+matches.length+' أصناف).'};
 }
 if(decision==='reject'){
  if(!String(reason||'').trim())throw Error('اكتب سبب الرفض.');
  matches.forEach(({i})=>{shx.getRange(i+1,x.h['الحالة']+1).setValue('مرفوض');shx.getRange(i+1,x.h['المعتمد']+1).setValue(u.email);shx.getRange(i+1,x.h['وقت الاعتماد']+1).setValue(now);if(x.h['سبب الرفض']!==undefined)shx.getRange(i+1,x.h['سبب الرفض']+1).setValue(reason);});
  audit(u.email,'رفض حركة',id,reason+' | '+matches.length+' أصناف');return {ok:true,message:'تم رفض الطلب كاملًا.'};
 }
 throw Error('قرار غير صحيح.');
}

function debugConnection(){
  const out={spreadsheetId:SPREADSHEET_ID,activeEmail:String(Session.getActiveUser().getEmail()||''),effectiveEmail:String(Session.getEffectiveUser().getEmail()||''),sheets:{},authorization:null};
  const book=SpreadsheetApp.openById(SPREADSHEET_ID);
  ['الأصناف','الأماكن','الصلاحيات','الحركات','الحركات_المعلقة','سجل_التدقيق'].forEach(n=>{
    const s=book.getSheetByName(n);
    if(!s){out.sheets[n]={exists:false};return;}
    const v=s.getDataRange().getDisplayValues();
    out.sheets[n]={exists:true,headers:v[0]||[],nonEmptyRows:Math.max(0,v.length-1)};
  });
  try{out.authorization=user();}catch(e){out.authorizationError=String(e.message||e);}
  const it=out.sheets['الأصناف'],pl=out.sheets['الأماكن'];
  if(it&&it.exists){const s=book.getSheetByName('الأصناف');out.itemCount=s.getRange(3,2,Math.max(1,s.getLastRow()-2),1).getDisplayValues().flat().filter(Boolean).length;}
  if(pl&&pl.exists){const s=book.getSheetByName('الأماكن');out.placeCount=s.getRange(2,1,Math.max(1,s.getLastRow()-1),1).getDisplayValues().flat().filter(x=>x&&!String(x).startsWith('ملاحظة:')).length;}
  return out;
}
