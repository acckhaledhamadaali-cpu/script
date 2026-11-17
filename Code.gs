const SPREADSHEET_ID='1RZkSWsb0uAnjYq1zje6iVIVCeKYxh3lUPSepHRVTxNU';
const S={items:'الأصناف',places:'الأماكن',users:'الصلاحيات',moves:'الحركات',pending:'الحركات_المعلقة',audit:'سجل_التدقيق',settings:'الإعدادات'};
function doGet(){return HtmlService.createHtmlOutputFromFile('Index').setTitle('شاورمتي | إدارة المخزون').addMetaTag('viewport','width=device-width, initial-scale=1');}
function ss(){return SpreadsheetApp.openById(SPREADSHEET_ID);}
function sh(n){const x=ss().getSheetByName(n);if(!x)throw Error('الورقة غير موجودة: '+n);return x;}
function head(r){const o={};r.forEach((v,i)=>o[String(v).trim()]=i);return o;}
function rows(n){const a=sh(n).getDataRange().getValues();return {a,h:head(a[0])};}
function ensureHeaders(sheet,required){
 const lastCol=Math.max(sheet.getLastColumn(),1);
 const header=sheet.getRange(1,1,1,lastCol).getValues()[0].map(v=>String(v||'').trim());
 required.forEach(name=>{if(!header.includes(name)){header.push(name);sheet.getRange(1,header.length).setValue(name);}});
 return head(header);
}
function makeRow(headers,values,width){
 const row=Array(width).fill('');
 Object.keys(values).forEach(k=>{if(headers[k]!==undefined)row[headers[k]]=values[k];});
 return row;
}
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
  if(type==='تحويل'){
   const totals={};clean.forEach(x=>totals[x.item.name]=(totals[x.item.name]||0)+x.q);
   Object.keys(totals).forEach(name=>{if(totals[name]>available(from,name))throw Error('إجمالي الكمية المطلوبة من الصنف '+name+' أكبر من المتاح في '+from+'.');});
  }
  const id='MV-'+Utilities.formatDate(new Date(),'Asia/Riyadh','yyyyMMdd-HHmmss')+'-'+Math.floor(Math.random()*900+100),ps=sh(S.pending),now=new Date();
  const ph=ensureHeaders(ps,['معرف الحركة','التاريخ','نوع الحركة','من','إلى','الصنف','الكمية','سعر الوحدة','الوحدة','رقم المرجع','ملاحظات','المدخل','وقت الإدخال','الحالة','المعتمد','وقت الاعتماد','آخر تعديل','سبب الرفض']);
  const fields=clean.map(x=>makeRow(ph,{'معرف الحركة':id,'التاريخ':new Date(f.date+'T12:00:00'),'نوع الحركة':type,'من':from,'إلى':to,'الصنف':x.item.name,'الكمية':x.q,'سعر الوحدة':x.price,'الوحدة':x.item.unit,'رقم المرجع':f.reference||'','ملاحظات':f.notes||'','المدخل':u.email,'وقت الإدخال':now,'الحالة':'بانتظار اعتماد'},ps.getLastColumn()));
  ps.getRange(ps.getLastRow()+1,1,fields.length,fields[0].length).setValues(fields);
  if(type!=='تحويل'&&u.role==='مالك')decideMovement(id,'approve','');
  audit(u.email,'إنشاء حركة',id,type+' | '+clean.length+' أصناف');
  return {ok:true,id,status:'بانتظار اعتماد',message:'تم إرسال الحركة ('+clean.length+' أصناف) للاعتماد.'};
 }finally{lock.releaseLock();}
}
function decideMovement(id,decision,reason,editedItems){
 const pendingSheet=sh(S.pending),movesSheet=sh(S.moves);
 ensureHeaders(pendingSheet,['معرف الحركة','التاريخ','نوع الحركة','من','إلى','الصنف','الكمية','سعر الوحدة','الوحدة','رقم المرجع','ملاحظات','المدخل','وقت الإدخال','الحالة','المعتمد','وقت الاعتماد','آخر تعديل','سبب الرفض']);
 ensureHeaders(movesSheet,['معرف الحركة','التاريخ','نوع الحركة','من (مكان)','إلى (مكان)','الصنف','الكمية','سعر الوحدة (للشراء/الافتتاحي فقط)','الوحدة','القيمة (تلقائي)','رقم المرجع (فاتورة/إذن)','المسؤول','ملاحظات']);
 const u=user(),x=rows(S.pending),matches=[];
 x.a.forEach((r,i)=>{if(i>0&&String(r[x.h['معرف الحركة']])===String(id))matches.push({r,i});});
 if(!matches.length)throw Error('الحركة غير موجودة.');
 const first=matches[0].r,from=first[x.h['من']],to=first[x.h['إلى']];
 if(matches.some(z=>z.r[x.h['الحالة']]!=='بانتظار اعتماد'))throw Error('تمت معالجة جزء من الطلب مسبقًا؛ راجع سجل الحركات قبل إعادة المحاولة.');
 if(u.role!=='مالك'&&u.place!==to)throw Error('الاعتماد متاح للمكان المستلم فقط.');
 if(u.role!=='مالك'&&String(first[x.h['المدخل']]).toLowerCase()===u.email)throw Error('لا يمكن اعتماد حركة أنشأتها بنفسك.');
 const shx=sh(S.pending),now=new Date();
 if(decision==='approve'){
  if(Array.isArray(editedItems)&&editedItems.length){
   if(editedItems.length!==matches.length)throw Error('عدد الأصناف المعدلة لا يطابق الطلب الأصلي.');
   editedItems.forEach((v,j)=>{const q=Number(v.qty);if(!(q>0))throw Error('الكمية المستلمة يجب أن تكون أكبر من صفر.');shx.getRange(matches[j].i+1,x.h['الكمية']+1).setValue(q);matches[j].r[x.h['الكمية']]=q;});
  }
  const mv=sh(S.moves),mh=head(mv.getRange(1,1,1,mv.getLastColumn()).getValues()[0]);
  const existingIds=mh['معرف الحركة']===undefined?new Set():new Set(mv.getDataRange().getValues().slice(1).map(r=>String(r[mh['معرف الحركة']]||'')));
  const toAdd=[];
  matches.forEach(({r})=>{
   const item=r[x.h['الصنف']],qty=Number(r[x.h['الكمية']])||0;
   if(!r[x.h['نوع الحركة']]||!r[x.h['التاريخ']])throw Error('بيانات الحركة ناقصة: نوع الحركة أو التاريخ. لا يمكن ترحيلها بأمان.');
   if(r[x.h['نوع الحركة']]==='تحويل'&&qty>available(from,item)+qty)throw Error('تعذر التحقق من الرصيد للصنف: '+item);
   if(!existingIds.has(String(id))){
    toAdd.push(makeRow(mh,{'معرف الحركة':id,'التاريخ':r[x.h['التاريخ']],'نوع الحركة':r[x.h['نوع الحركة']],'من (مكان)':from,'إلى (مكان)':to,'الصنف':item,'الكمية':qty,'سعر الوحدة (للشراء/الافتتاحي فقط)':Number(r[x.h['سعر الوحدة']])||0,'الوحدة':r[x.h['الوحدة']],'القيمة (تلقائي)':qty*(Number(r[x.h['سعر الوحدة']])||0),'رقم المرجع (فاتورة/إذن)':r[x.h['رقم المرجع']],'المسؤول':r[x.h['المدخل']],'ملاحظات':r[x.h['ملاحظات']]},mv.getLastColumn()));
   }
  });
  if(toAdd.length)mv.getRange(mv.getLastRow()+1,1,toAdd.length,toAdd[0].length).setValues(toAdd);
  matches.forEach(({i})=>{shx.getRange(i+1,x.h['الحالة']+1).setValue('معتمد');if(x.h['المعتمد']!==undefined)shx.getRange(i+1,x.h['المعتمد']+1).setValue(u.email);if(x.h['وقت الاعتماد']!==undefined)shx.getRange(i+1,x.h['وقت الاعتماد']+1).setValue(now);});
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


// ترحيل آمن لسجل رصيد افتتاحي قديم لا يحتوي على معرف/تاريخ/نوع الحركة.
// شغّل هذه الدالة مرة واحدة من محرر Apps Script مع رقم صف السجل (غالبًا 2).
function repairLegacyOpeningBalance(rowNumber){
 const lock=LockService.getScriptLock();lock.waitLock(30000);
 try{
  const ps=sh(S.pending),mh=ensureHeaders(ps,['معرف الحركة','التاريخ','نوع الحركة','من','إلى','الصنف','الكمية','سعر الوحدة','الوحدة','رقم المرجع','ملاحظات','المدخل','وقت الإدخال','الحالة','المعتمد','وقت الاعتماد','آخر تعديل','سبب الرفض']);
  const row=Number(rowNumber);if(!(row>=2))throw Error('حدد رقم صف صحيحًا.');
  const values=ps.getRange(row,1,1,ps.getLastColumn()).getValues()[0];
  if(String(values[mh['الحالة']]||'')!=='معتمد')throw Error('السجل ليس معتمدًا؛ لم يتم إجراء أي تعديل.');
  if(String(values[mh['نوع الحركة']]||'')&&String(values[mh['معرف الحركة']]||''))throw Error('السجل يحتوي بالفعل على نوع ومعرف حركة.');
  if(!values[mh['إلى']]||!values[mh['الصنف']]||!(Number(values[mh['الكمية']])>0))throw Error('بيانات المكان أو الصنف أو الكمية غير مكتملة.');
  const id='MV-LEGACY-'+Utilities.formatDate(new Date(),'Asia/Riyadh','yyyyMMdd-HHmmss')+'-'+row;
  const rawDate=values[mh['وقت الإدخال']]||new Date();
  const date=rawDate instanceof Date?rawDate:new Date(rawDate);
  if(isNaN(date.getTime()))throw Error('تعذر قراءة تاريخ الإدخال.');
  ps.getRange(row,mh['معرف الحركة']+1).setValue(id);
  ps.getRange(row,mh['التاريخ']+1).setValue(date);
  ps.getRange(row,mh['نوع الحركة']+1).setValue('رصيد افتتاحي');
  const mv=sh(S.moves);
  const headers=ensureHeaders(mv,['معرف الحركة','التاريخ','نوع الحركة','من (مكان)','إلى (مكان)','الصنف','الكمية','سعر الوحدة (للشراء/الافتتاحي فقط)','الوحدة','القيمة (تلقائي)','رقم المرجع (فاتورة/إذن)','المسؤول','ملاحظات']);
  const existing=mv.getDataRange().getValues().slice(1).some(r=>String(r[headers['معرف الحركة']]||'')===id);
  if(!existing){
   const q=Number(values[mh['الكمية']]),price=Number(values[mh['سعر الوحدة']])||0;
   const out=makeRow(headers,{'معرف الحركة':id,'التاريخ':date,'نوع الحركة':'رصيد افتتاحي','من (مكان)':'','إلى (مكان)':values[mh['إلى']],'الصنف':values[mh['الصنف']],'الكمية':q,'سعر الوحدة (للشراء/الافتتاحي فقط)':price,'الوحدة':values[mh['الوحدة']],'القيمة (تلقائي)':q*price,'رقم المرجع (فاتورة/إذن)':values[mh['رقم المرجع']],'المسؤول':values[mh['المدخل']],'ملاحظات':values[mh['ملاحظات']]},mv.getLastColumn());
   mv.getRange(mv.getLastRow()+1,1,1,out.length).setValues([out]);
  }
  audit(String(values[mh['المعتمد']]||values[mh['المدخل']]||'system'),'استعادة رصيد افتتاحي قديم',id,'ترحيل آمن من الصف '+row);
  return {ok:true,id,message:'تم استعادة الرصيد الافتتاحي وترحيله مرة واحدة.'};
 }finally{lock.releaseLock();}
}
