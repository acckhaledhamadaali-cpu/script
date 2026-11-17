const SPREADSHEET_ID='1RZkSWsb0uAnjYq1zje6iVIVCeKYxh3lUPSepHRVTxNU';
const S={items:'الأصناف',places:'الأماكن',users:'الصلاحيات',moves:'الحركات',pending:'الحركات_المعلقة',audit:'سجل_التدقيق',settings:'الإعدادات'};
function doGet(){return HtmlService.createHtmlOutputFromFile('Index').setTitle('شاورمتي | إدارة المخزون').addMetaTag('viewport','width=device-width, initial-scale=1');}
function ss_(){return SpreadsheetApp.openById(SPREADSHEET_ID);}
function sh_(n){const x=ss_().getSheetByName(n);if(!x)throw Error('الورقة غير موجودة: '+n);return x;}
function head_(r){const o={};r.forEach((v,i)=>o[String(v).trim()]=i);return o;}
function rows_(n){const a=sh_(n).getDataRange().getValues();return {a,h:head_(a[0])};}
function ensureHeaders_(sheet,required){
 const lastCol=Math.max(sheet.getLastColumn(),1);
 const header=sheet.getRange(1,1,1,lastCol).getValues()[0].map(v=>String(v||'').trim());
 required.forEach(name=>{if(!header.includes(name)){header.push(name);sheet.getRange(1,header.length).setValue(name);}});
 return head_(header);
}

/**
 * يحدد أول مساحة فارغة اعتمادًا على أعمدة البيانات الفعلية، مع تجاهل أعمدة الصيغ.
 * segments تحدد الأعمدة التي يسمح بالكتابة فيها فقط للحفاظ على الصيغ.
 */
function appendAtFirstDataGap_(sheet, dataRows, keyColumns, segments) {
  if (!dataRows || !dataRows.length) return;
  const needed = dataRows.length;
  const maxRows = sheet.getMaxRows();
  const scanRows = Math.max(0, maxRows - 1);
  const keyValues = scanRows
    ? sheet.getRange(2, 1, scanRows, sheet.getLastColumn()).getValues()
    : [];
  let startRow = 2;
  let found = false;
  for (let i = 0; i <= keyValues.length - needed; i++) {
    let blockIsEmpty = true;
    for (let j = 0; j < needed && blockIsEmpty; j++) {
      if (keyColumns.some(col => keyValues[i + j][col - 1] !== '' &&
          keyValues[i + j][col - 1] !== null)) blockIsEmpty = false;
    }
    if (blockIsEmpty) {
      startRow = i + 2;
      found = true;
      break;
    }
  }
  if (!found) {
    startRow = maxRows + 1;
    sheet.insertRowsAfter(maxRows, needed);
  }
  segments.forEach(([startCol, width]) => {
    const block = dataRows.map(row => row.slice(startCol - 1, startCol - 1 + width));
    sheet.getRange(startRow, startCol, needed, width).setValues(block);
  });
  return startRow;
}
function appendMovementRows_(sheet, dataRows) {
  // تجاهل أعمدة الصيغ H وI وM وN، والكتابة في بقية أعمدة الحركة فقط.
  return appendAtFirstDataGap_(sheet, dataRows,
    [1,2,3,4,5,6,7,10,11,12,15],
    [[1,7],[10,3],[15,1]]);
}

function makeRow_(headers,values,width){
 const row=Array(width).fill('');
 Object.keys(values).forEach(k=>{if(headers[k]!==undefined)row[headers[k]]=values[k];});
 return row;
}
function user_(){
 const email=(Session.getActiveUser().getEmail()||'').trim().toLowerCase();
 if(!email)throw Error('تعذر تحديد بريد المستخدم. تحقق من إعدادات نشر Web App.');
 const {a,h}=rows_(S.users),active=a.slice(1).filter(r=>['نعم','yes','true','1'].includes(String(r[h['فعّال']]||'').trim().toLowerCase()));
 const mine=active.filter(r=>String(r[h['الإيميل']]||'').trim().toLowerCase()===email);
 if(mine.length!==1)throw Error(mine.length?'يوجد تكرار لحسابك في ورقة الصلاحيات.':'المستخدم غير مضاف أو غير مفعل في ورقة الصلاحيات: '+email);
 const r=mine[0],role=String(r[h['الدور']]||'مسؤول مكان').trim(),place=String(r[h['المكان']]||'').trim();
 if(role==='مالك'){
  if(place)throw Error('حساب المدير العام يجب ألا يرتبط بمكان محدد.');
 }else{
  if(!place)throw Error('يجب تحديد مكان واحد للمستخدم.');
  const same=active.filter(x=>String(x[h['المكان']]||'').trim()===place);
  if(same.length!==1)throw Error('يوجد أكثر من مستخدم فعّال للمكان '+place+'. يجب أن يكون لكل مكان مسؤول واحد فقط.');
  if(!['مسؤول مكان','مستخدم مكان'].includes(role))throw Error('الدور غير معتمد. استخدم مالك أو مسؤول مكان.');
 }
 return {email,name:r[h['الاسم']]||email,place,role};
}
function refs_(){const {a,h}=rows_(S.items);const items=a.slice(1).filter(r=>r[h['اسم الصنف (فريد)']]).map(r=>({name:String(r[h['اسم الصنف (فريد)']]),unit:String(r[h['الوحدة']]||''),avgCost:Number(r[h['متوسط التكلفة (تلقائي)']]||r[h['سعر مبدئي تقديري (عدّله)']])||0}));const p=rows_(S.places);return {items,places:p.a.slice(1).filter(r=>r[p.h['المكان']]&&!String(r[p.h['المكان']]).startsWith('ملاحظة:')).map(r=>({name:String(r[p.h['المكان']]),type:String(r[p.h['النوع']]||'')}))};}
function verifyLogin(enteredEmail){
 const u=user_();
 const entered=String(enteredEmail||'').trim().toLowerCase();
 if(!entered)throw Error('أدخل بريدك الإلكتروني.');
 if(entered!==u.email)throw Error('البريد المدخل لا يطابق حساب Google المسجل حاليًا. سجّل الدخول بحسابك الصحيح.');
 return {ok:true,user:{name:u.name,email:u.email,role:u.role,place:u.place}};
}
function getInitialData(){
 const u=user_();
 repairKnownLegacyOpeningBalance_();
 const d=refs_(),p=pending_(u),m=recent_(u),allBalances={};
 d.places.forEach(x=>allBalances[x.name]={qty:0,value:0});
 const x=rows_(S.moves);
 x.a.slice(1).forEach(r=>{
  const q=Number(r[x.h['الكمية']])||0,price=Number(r[x.h['سعر الوحدة (للشراء/الافتتاحي فقط)']])||0;
  const from=r[x.h['من (مكان)']],to=r[x.h['إلى (مكان)']];
  if(to&&allBalances[to]){allBalances[to].qty+=q;allBalances[to].value+=q*price;}
  if(from&&allBalances[from]){allBalances[from].qty-=q;allBalances[from].value-=q*price;}
 });
 const balances=u.role==='مالك'?allBalances:{[u.place]:allBalances[u.place]||{qty:0,value:0}};
 const vals=Object.values(balances);
 return {user:u,items:d.items,places:d.places,summary:{locations:balances,totalQty:vals.reduce((s,v)=>s+v.qty,0),totalValue:vals.reduce((s,v)=>s+v.value,0)},pending:p,recent:m};
}
function pending_(u){
 const {a,h}=rows_(S.pending),groups={};
 a.slice(1).forEach(r=>{
  if(r[h['الحالة']]!=='بانتظار اعتماد'||!(u.role==='مالك'||r[h['إلى']]===u.place))return;
  const id=String(r[h['معرف الحركة']]);if(!groups[id])groups[id]={id,date:fmtDate_(r[h['التاريخ']]),type:r[h['نوع الحركة']],from:r[h['من']],to:r[h['إلى']],creator:r[h['المدخل']],notes:r[h['ملاحظات']],items:[]};
  groups[id].items.push({item:r[h['الصنف']],qty:r[h['الكمية']],unit:r[h['الوحدة']]});
 });
 return Object.values(groups);
}
function recent_(u){
 const {a,h}=rows_(S.moves);
 return a.slice(1).filter(r=>r[h['الصنف']]&&(u.role==='مالك'||r[h['من (مكان)']]===u.place||r[h['إلى (مكان)']]===u.place)).slice(-30).reverse().map(r=>({date:fmtDate_(r[h['التاريخ']]),type:r[h['نوع الحركة']],from:r[h['من (مكان)']],to:r[h['إلى (مكان)']],item:r[h['الصنف']],qty:r[h['الكمية']],unit:r[h['الوحدة']]}));
}
function fmtDate_(d){return d instanceof Date?Utilities.formatDate(d,'Asia/Riyadh','yyyy-MM-dd'):String(d||'');}
function audit_(u,act,id,detail){const x=sh_(S.audit);const a=[new Date(),u,act,id,detail];while(a.length<x.getLastColumn())a.push('');x.appendRow(a);}
function available_(place,item){const {a,h}=rows_(S.moves);let n=0;a.slice(1).forEach(r=>{if(r[h['الصنف']]!==item)return;const q=Number(r[h['الكمية']])||0;if(r[h['إلى (مكان)']]===place)n+=q;if(r[h['من (مكان)']]===place)n-=q;});const p=rows_(S.pending);p.a.slice(1).forEach(r=>{if(r[p.h['الحالة']]==='بانتظار اعتماد'&&r[p.h['الصنف']]===item&&r[p.h['من']]===place)n-=Number(r[p.h['الكمية']])||0;});return n;}
function submitMovement(f){
 const lock=LockService.getScriptLock();lock.waitLock(30000);
 try{
  const u=user_(),d=refs_(),type=String(f.type||''),validPlaces=d.places.map(x=>x.name);
  let from=String(f.from||'').trim(),to=String(f.to||'').trim();
  const items=Array.isArray(f.items)?f.items:[];
  const transfer=type==='تحويل',purchase=type==='شراء';
  const inbound=['شراء','رصيد افتتاحي','تسوية زيادة'].includes(type);
  const outbound=['صرف للتشغيل','هالك','مرتجع للمورد','تسوية عجز'].includes(type);
  if(!items.length)throw Error('أضف صنفًا واحدًا على الأقل.');
  if(!transfer&&!purchase&&!inbound&&!outbound)throw Error('نوع الحركة غير معتمد.');
  if(u.role!=='مالك'){
   if(transfer){if(from!==u.place)throw Error('لا يمكنك التحويل إلا من مخزون مكانك.');}
   else if(inbound){to=u.place;if(purchase){}}
   else if(outbound){from=u.place;to='';}
   else throw Error('هذه الحركة غير متاحة لصلاحيتك.');
  }
  if(transfer){
   if(!validPlaces.includes(from)||!validPlaces.includes(to)||from===to)throw Error('حدد مصدرًا ومستلمًا صحيحين.');
   if(u.role!=='مالك'&&from!==u.place)throw Error('لا يمكنك التحويل من مكان آخر.');
  }else if(purchase){
   if(!from)throw Error('اكتب اسم المورد.');
   if(!validPlaces.includes(to))throw Error('حدد المكان الذي استلم المشتريات.');
  }else if(inbound){
   if(type==='رصيد افتتاحي'&&u.role!=='مالك')throw Error('الرصيد الافتتاحي للمدير فقط.');
   if(!validPlaces.includes(to))throw Error('حدد المكان المستلم للحركة.');
   from='';
  }else if(outbound){
   if(!validPlaces.includes(from))throw Error('حدد المكان الذي ستُصرف منه الكمية.');
   to='';
  }
  const clean=items.map(x=>{
   const item=d.items.find(y=>y.name===x.item),q=Number(x.qty),price=Number(x.price)||0;
   if(!item||!(q>0))throw Error('تحقق من الصنف والكمية في جميع السطور.');
   if(purchase||type==='رصيد افتتاحي'){if(!(Number(x.price)>=0))throw Error('أدخل سعر الوحدة بصورة صحيحة.');}
   return {item,q,price};
  });
  if(transfer||outbound){
   const totals={};clean.forEach(x=>totals[x.item.name]=(totals[x.item.name]||0)+x.q);
   Object.keys(totals).forEach(name=>{if(totals[name]>available_(from,name))throw Error('الكمية المطلوبة من الصنف '+name+' أكبر من المتاح في '+from+'.');});
  }
  const id='MV-'+Utilities.formatDate(new Date(),'Asia/Riyadh','yyyyMMdd-HHmmss')+'-'+Math.floor(Math.random()*900+100);
  const now=new Date(),date=new Date(String(f.date||'')+'T12:00:00');
  if(isNaN(date.getTime()))throw Error('أدخل تاريخًا صحيحًا.');
  const mv=sh_(S.moves);
  const mh=ensureHeaders_(mv,['معرف الحركة','التاريخ','نوع الحركة','من (مكان)','إلى (مكان)','الصنف','الكمية','سعر الوحدة (للشراء/الافتتاحي فقط)','الوحدة','القيمة (تلقائي)','رقم المرجع (فاتورة/إذن)','المسؤول','ملاحظات']);
  const values=clean.map(x=>makeRow_(mh,{'معرف الحركة':id,'التاريخ':date,'نوع الحركة':type,'من (مكان)':from,'إلى (مكان)':to,'الصنف':x.item.name,'الكمية':x.q,'سعر الوحدة (للشراء/الافتتاحي فقط)':x.price,'الوحدة':x.item.unit,'القيمة (تلقائي)':x.q*x.price,'رقم المرجع (فاتورة/إذن)':f.reference||'','المسؤول':u.email,'ملاحظات':(purchase?'المورد: '+from+' | ':'')+(f.notes||'')},mv.getLastColumn()));
  if(transfer){
   const ps=sh_(S.pending),ph=ensureHeaders_(ps,['معرف الحركة','التاريخ','نوع الحركة','من','إلى','الصنف','الكمية','سعر الوحدة','الوحدة','رقم المرجع','ملاحظات','المدخل','وقت الإدخال','الحالة','المعتمد','وقت الاعتماد','آخر تعديل','سبب الرفض']);
   const pendingRows=clean.map(x=>makeRow_(ph,{'معرف الحركة':id,'التاريخ':date,'نوع الحركة':type,'من':from,'إلى':to,'الصنف':x.item.name,'الكمية':x.q,'سعر الوحدة':x.price,'الوحدة':x.item.unit,'رقم المرجع':f.reference||'','ملاحظات':f.notes||'','المدخل':u.email,'وقت الإدخال':now,'الحالة':'بانتظار اعتماد'},ps.getLastColumn()));
   appendAtFirstDataGap_(ps,pendingRows,Array.from({length:pendingRows[0].length},(_,i)=>i+1),[[1,pendingRows[0].length]]);
   audit_(u.email,'إنشاء تحويل',id,from+' إلى '+to+' | '+clean.length+' أصناف');
   return {ok:true,id,status:'بانتظار اعتماد',message:'تم إرسال التحويل إلى '+to+' وينتظر تأكيد الاستلام.'};
  }
  appendMovementRows_(mv,values);
  audit_(u.email,'تسجيل حركة مباشرة',id,type+' | المكان: '+(to||from)+' | '+clean.length+' أصناف');
  return {ok:true,id,status:'معتمد',message:'تم تسجيل الحركة وتحديث المخزون.'};
 }finally{lock.releaseLock();}
}
function decideMovement(id,decision,reason,editedItems){
 const lock=LockService.getScriptLock();lock.waitLock(30000);
 try{
  const u=user_(),pendingSheet=sh_(S.pending),movesSheet=sh_(S.moves);
  ensureHeaders_(pendingSheet,['معرف الحركة','التاريخ','نوع الحركة','من','إلى','الصنف','الكمية','سعر الوحدة','الوحدة','رقم المرجع','ملاحظات','المدخل','وقت الإدخال','الحالة','المعتمد','وقت الاعتماد','آخر تعديل','سبب الرفض','الكمية المستلمة','فرق الاستلام']);
  const x=rows_(S.pending),matches=[];
  x.a.forEach((r,i)=>{if(i>0&&String(r[x.h['معرف الحركة']])===String(id))matches.push({r,i});});
  if(!matches.length)throw Error('التحويل غير موجود.');
  const first=matches[0].r,from=String(first[x.h['من']]||''),to=String(first[x.h['إلى']]||'');
  if(first[x.h['نوع الحركة']]!=='تحويل')throw Error('اعتماد الطلبات متاح للتحويلات فقط.');
  if(matches.some(z=>z.r[x.h['الحالة']]!=='بانتظار اعتماد'))throw Error('تمت معالجة هذا التحويل مسبقًا.');
  if(u.role!=='مالك'&&u.place!==to)throw Error('تأكيد الاستلام متاح لمسؤول المكان المستلم فقط.');
  if(u.role!=='مالك'&&String(first[x.h['المدخل']]||'').toLowerCase()===u.email)throw Error('لا يمكنك تأكيد تحويل أنشأته بنفسك.');
  if(decision==='approve'){
   if(!Array.isArray(editedItems)||editedItems.length!==matches.length)throw Error('راجع كميات جميع الأصناف قبل التأكيد.');
   const sentTotals={};
   editedItems.forEach((v,j)=>{
    const received=Number(v.qty),sent=Number(matches[j].r[x.h['الكمية']])||0;
    if(!Number.isFinite(received)||received<0||received>sent)throw Error('الكمية المستلمة يجب أن تكون بين صفر والكمية المرسلة.');
    const item=String(matches[j].r[x.h['الصنف']]);sentTotals[item]=(sentTotals[item]||0)+sent;
   });
   Object.keys(sentTotals).forEach(item=>{if(sentTotals[item]>available_(from,item)+sentTotals[item])throw Error('الرصيد المتاح في المصدر غير كافٍ للصنف '+item+'.');});
   const mh=ensureHeaders_(movesSheet,['معرف الحركة','التاريخ','نوع الحركة','من (مكان)','إلى (مكان)','الصنف','الكمية','سعر الوحدة (للشراء/الافتتاحي فقط)','الوحدة','القيمة (تلقائي)','رقم المرجع (فاتورة/إذن)','المسؤول','ملاحظات']);
   const toAdd=[];
   matches.forEach(({r},j)=>{
    const sent=Number(r[x.h['الكمية']])||0,received=Number(editedItems[j].qty),price=Number(r[x.h['سعر الوحدة']])||0;
    const base={'معرف الحركة':id,'التاريخ':r[x.h['التاريخ']],'سعر الوحدة (للشراء/الافتتاحي فقط)':price,'الوحدة':r[x.h['الوحدة']],'رقم المرجع (فاتورة/إذن)':r[x.h['رقم المرجع']],'المسؤول':r[x.h['المدخل']]};
    toAdd.push(makeRow_(mh,{...base,'نوع الحركة':'تحويل صادر','من (مكان)':from,'إلى (مكان)':'','الصنف':r[x.h['الصنف']],'الكمية':sent,'القيمة (تلقائي)':sent*price,'ملاحظات':'صادر إلى '+to+' | '+String(r[x.h['ملاحظات']]||'')},movesSheet.getLastColumn()));
    if(received>0)toAdd.push(makeRow_(mh,{...base,'نوع الحركة':'تحويل وارد','من (مكان)':'','إلى (مكان)':to,'الصنف':r[x.h['الصنف']],'الكمية':received,'القيمة (تلقائي)':received*price,'ملاحظات':'وارد من '+from+(received<sent?' | عجز استلام: '+(sent-received):'')+' | '+String(r[x.h['ملاحظات']]||'')},movesSheet.getLastColumn()));
   });
   appendMovementRows_(movesSheet,toAdd);
   matches.forEach(({i,r},j)=>{
    const received=Number(editedItems[j].qty),sent=Number(r[x.h['الكمية']])||0,row=i+1;
    pendingSheet.getRange(row,x.h['الكمية المستلمة']+1).setValue(received);
    pendingSheet.getRange(row,x.h['فرق الاستلام']+1).setValue(sent-received);
    pendingSheet.getRange(row,x.h['الحالة']+1).setValue('معتمد');
    if(x.h['المعتمد']!==undefined)pendingSheet.getRange(row,x.h['المعتمد']+1).setValue(u.email);
    if(x.h['وقت الاعتماد']!==undefined)pendingSheet.getRange(row,x.h['وقت الاعتماد']+1).setValue(new Date());
   });
   audit_(u.email,'تأكيد استلام تحويل',id,from+' إلى '+to+' | '+matches.length+' أصناف');
   return {ok:true,message:'تم تأكيد الاستلام. خُصمت الكمية المرسلة وأضيفت الكمية المستلمة الفعلية.'};
  }
  if(decision==='reject'){
   if(!String(reason||'').trim())throw Error('اكتب سبب الرفض.');
   matches.forEach(({i})=>{const row=i+1;pendingSheet.getRange(row,x.h['الحالة']+1).setValue('مرفوض');if(x.h['المعتمد']!==undefined)pendingSheet.getRange(row,x.h['المعتمد']+1).setValue(u.email);if(x.h['وقت الاعتماد']!==undefined)pendingSheet.getRange(row,x.h['وقت الاعتماد']+1).setValue(new Date());if(x.h['سبب الرفض']!==undefined)pendingSheet.getRange(row,x.h['سبب الرفض']+1).setValue(reason);});
   audit_(u.email,'رفض تحويل',id,reason);
   return {ok:true,message:'تم رفض التحويل مع تسجيل السبب.'};
  }
  throw Error('قرار غير صحيح.');
 }finally{lock.releaseLock();}
}

function debugConnection(){
 const auth=user_(),out={activeEmail:String(Session.getActiveUser().getEmail()||''),authorization:auth,sheets:{}},book=ss_();
 ['الأصناف','الأماكن','الصلاحيات','الحركات','الحركات_المعلقة','سجل_التدقيق'].forEach(n=>{
  const s=book.getSheetByName(n);if(!s){out.sheets[n]={exists:false};return;}
  const v=s.getDataRange().getDisplayValues();out.sheets[n]={exists:true,headers:v[0]||[],nonEmptyRows:Math.max(0,v.length-1)};
 });
 const it=book.getSheetByName('الأصناف'),pl=book.getSheetByName('الأماكن');
 if(it)out.itemCount=it.getRange(3,2,Math.max(1,it.getLastRow()-2),1).getDisplayValues().flat().filter(Boolean).length;
 if(pl)out.placeCount=pl.getRange(2,1,Math.max(1,pl.getLastRow()-1),1).getDisplayValues().flat().filter(x=>x&&!String(x).startsWith('ملاحظة:')).length;
 return out;
}


// ترحيل آمن لسجل رصيد افتتاحي قديم لا يحتوي على معرف/تاريخ/نوع الحركة.
// شغّل هذه الدالة مرة واحدة من محرر Apps Script مع رقم صف السجل (غالبًا 2).
function repairLegacyOpeningBalance_(rowNumber){
 const lock=LockService.getScriptLock();lock.waitLock(30000);
 try{
  const auth=user_();if(auth.role!=='مالك')throw Error('استعادة الأرصدة القديمة متاحة للمدير فقط.');
  const ps=sh_(S.pending),mh=ensureHeaders_(ps,['معرف الحركة','التاريخ','نوع الحركة','من','إلى','الصنف','الكمية','سعر الوحدة','الوحدة','رقم المرجع','ملاحظات','المدخل','وقت الإدخال','الحالة','المعتمد','وقت الاعتماد','آخر تعديل','سبب الرفض']);
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
  const mv=sh_(S.moves);
  const headers=ensureHeaders_(mv,['معرف الحركة','التاريخ','نوع الحركة','من (مكان)','إلى (مكان)','الصنف','الكمية','سعر الوحدة (للشراء/الافتتاحي فقط)','الوحدة','القيمة (تلقائي)','رقم المرجع (فاتورة/إذن)','المسؤول','ملاحظات']);
  const existing=mv.getDataRange().getValues().slice(1).some(r=>String(r[headers['معرف الحركة']]||'')===id);
  if(!existing){
   const q=Number(values[mh['الكمية']]),price=Number(values[mh['سعر الوحدة']])||0;
   const out=makeRow_(headers,{'معرف الحركة':id,'التاريخ':date,'نوع الحركة':'رصيد افتتاحي','من (مكان)':'','إلى (مكان)':values[mh['إلى']],'الصنف':values[mh['الصنف']],'الكمية':q,'سعر الوحدة (للشراء/الافتتاحي فقط)':price,'الوحدة':values[mh['الوحدة']],'القيمة (تلقائي)':q*price,'رقم المرجع (فاتورة/إذن)':values[mh['رقم المرجع']],'المسؤول':values[mh['المدخل']],'ملاحظات':values[mh['ملاحظات']]},mv.getLastColumn());
   appendMovementRows_(mv,[out]);
  }
  audit_(String(values[mh['المعتمد']]||values[mh['المدخل']]||'system'),'استعادة رصيد افتتاحي قديم',id,'ترحيل آمن من الصف '+row);
  return {ok:true,id,message:'تم استعادة الرصيد الافتتاحي وترحيله مرة واحدة.'};
 }finally{lock.releaseLock();}
}


// استعادة تلقائية وآمنة للسجل القديم المحدد الذي أرسله المستخدم.
// لا تمس أي سجل آخر، وتعمل مرة واحدة فقط بفضل معرف الحركة.
function parseSheetNumber_(v){if(typeof v==='number')return v;const s=String(v??'').trim().replace(/[٠-٩]/g,d=>'٠١٢٣٤٥٦٧٨٩'.indexOf(d)).replace(/[۰-۹]/g,d=>'۰۱۲۳۴۵۶۷۸۹'.indexOf(d)).replace(/[٫٬]/g,(m)=>m==='٫'?'.':'').replace(/,/g,'');return Number(s);}
function repairKnownLegacyOpeningBalance_(){
 const ps=sh_(S.pending);
 const h=ensureHeaders_(ps,['معرف الحركة','التاريخ','نوع الحركة','من','إلى','الصنف','الكمية','سعر الوحدة','الوحدة','رقم المرجع','ملاحظات','المدخل','وقت الإدخال','الحالة','المعتمد','وقت الاعتماد','آخر تعديل','سبب الرفض']);
 const last=ps.getLastRow();if(last<2)return;
 const data=ps.getRange(2,1,last-1,ps.getLastColumn()).getValues();
 for(let j=0;j<data.length;j++){
  const r=data[j],status=String(r[h['الحالة']]||'').trim();
  const isTarget=status==='معتمد'&&!r[h['معرف الحركة']]&&!r[h['نوع الحركة']]&&
   String(r[h['إلى']]||'').trim()==='المطبخ المركزي'&&
   String(r[h['الصنف']]||'').trim()==='كاتشب'&&
   Math.abs(parseSheetNumber_(r[h['الكمية']])-0.5)<0.000001&&
   Math.abs(parseSheetNumber_(r[h['سعر الوحدة']])-5)<0.000001&&
   String(r[h['الوحدة']]||'').trim()==='كجم';
  if(!isTarget)continue;
  const rowNumber=j+2,id='MV-LEGACY-OPENING-'+rowNumber;
  const rawDate=r[h['وقت الإدخال']]||new Date('2026-11-17T12:00:00');
  const date=rawDate instanceof Date?rawDate:new Date(rawDate);
  if(isNaN(date.getTime()))throw Error('تعذر قراءة تاريخ سجل الرصيد الافتتاحي القديم.');
  const mv=sh_(S.moves);
  const mh=ensureHeaders_(mv,['معرف الحركة','التاريخ','نوع الحركة','من (مكان)','إلى (مكان)','الصنف','الكمية','سعر الوحدة (للشراء/الافتتاحي فقط)','الوحدة','القيمة (تلقائي)','رقم المرجع (فاتورة/إذن)','المسؤول','ملاحظات']);
  const movementValues=mv.getDataRange().getValues();
  const alreadyById=movementValues.some((x,i)=>i>0&&String(x[mh['معرف الحركة']]||'')===id);
  const q=parseSheetNumber_(r[h['الكمية']]),price=parseSheetNumber_(r[h['سعر الوحدة']])||0;
  // إذا كان الرصيد قد رُحّل سابقًا دون معرف، اربطه بالسجل بدل إنشاء حركة مكررة.
  const legacyRow=movementValues.findIndex((x,i)=>i>0&&
   !String(x[mh['معرف الحركة']]||'')&&
   String(x[mh['نوع الحركة']]||'').trim()==='رصيد افتتاحي'&&
   String(x[mh['إلى (مكان)']]||'').trim()===String(r[h['إلى']]||'').trim()&&
   String(x[mh['الصنف']]||'').trim()===String(r[h['الصنف']]||'').trim()&&
   Math.abs(parseSheetNumber_(x[mh['الكمية']])-q)<0.000001&&
   Math.abs(parseSheetNumber_(x[mh['سعر الوحدة (للشراء/الافتتاحي فقط)']])-price)<0.000001);
  if(!alreadyById&&legacyRow>=1){
   mv.getRange(legacyRow+1,mh['معرف الحركة']+1).setValue(id);
  }else if(!alreadyById){
   const out=makeRow_(mh,{'معرف الحركة':id,'التاريخ':date,'نوع الحركة':'رصيد افتتاحي','من (مكان)':'','إلى (مكان)':r[h['إلى']],'الصنف':r[h['الصنف']],'الكمية':q,'سعر الوحدة (للشراء/الافتتاحي فقط)':price,'الوحدة':r[h['الوحدة']],'القيمة (تلقائي)':q*price,'رقم المرجع (فاتورة/إذن)':r[h['رقم المرجع']],'المسؤول':r[h['المدخل']],'ملاحظات':r[h['ملاحظات']]},mv.getLastColumn());
   appendMovementRows_(mv,[out]);
  }
  ps.getRange(rowNumber,h['معرف الحركة']+1).setValue(id);
  ps.getRange(rowNumber,h['التاريخ']+1).setValue(date);
  ps.getRange(rowNumber,h['نوع الحركة']+1).setValue('رصيد افتتاحي');
  break;
 }
}

// فحص قراءة شيت الحركات والحركات المعلقة دون تعديل أي بيانات.

function testMovementSheet_() {
  const s = SpreadsheetApp.openById(SPREADSHEET_ID);
  const p = s.getSheetByName('الحركات_المعلقة');
  const m = s.getSheetByName('الحركات');
  const pendingData = p.getDataRange().getDisplayValues();
  const movementData = m.getDataRange().getDisplayValues();
  const pHead = pendingData[0] || [];
  const mHead = movementData[0] || [];
  const pId = pHead.indexOf('معرف الحركة');
  const mId = mHead.indexOf('معرف الحركة');
  const pendingIds = pendingData.slice(1).filter(r => r[pId]).map(r => r[pId]);
  const movementMatches = movementData.slice(1)
    .map((r, i) => ({ row: i + 2, id: mId >= 0 ? r[mId] : '', values: r }))
    .filter(x => pendingIds.includes(x.id));
  const nonEmptyMovementRows = movementData.slice(1)
    .map((r, i) => ({ row: i + 2, values: r }))
    .filter(x => x.values.some(v => String(v).trim() !== ''));
  const result = {
    pendingIds: pendingIds,
    pendingRows: pendingData.slice(1).filter(r => r.some(v => String(v).trim() !== '')),
    movementNonEmptyRowCount: nonEmptyMovementRows.length,
    movementLastNonEmptyRows: nonEmptyMovementRows.slice(-8),
    movementMatchesForPending: movementMatches,
    movementHeaders: mHead
  };
  Logger.log(JSON.stringify(result, null, 2));
  console.log(JSON.stringify(result, null, 2));
  return result;
}




function inspectMovementTopRows_() {
  const s = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = s.getSheetByName('الحركات');
  const lastCol = sh.getLastColumn();
  const rows = Math.min(8, sh.getMaxRows());
  const range = sh.getRange(1, 1, rows, lastCol);
  const values = range.getDisplayValues();
  const formulas = range.getFormulas();
  const result = [];
  for (let r = 0; r < rows; r++) {
    const cells = [];
    for (let col = 0; col < lastCol; col++) {
      if (values[r][col] !== '' || formulas[r][col] !== '') {
        cells.push({
          column: col + 1,
          header: values[0][col] || '',
          value: values[r][col],
          formula: formulas[r][col]
        });
      }
    }
    result.push({ row: r + 1, cells: cells });
  }
  Logger.log(JSON.stringify(result, null, 2));
  console.log(JSON.stringify(result, null, 2));
  return result;
}
