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
 if(!mine.length)throw Error('المستخدم غير مضاف أو غير مفعل في ورقة الصلاحيات: '+email);
 const roles=[...new Set(mine.map(r=>String(r[h['الدور']]||'مسؤول مكان').trim()))];
 if(roles.length!==1)throw Error('يجب توحيد الدور لكل سجلات البريد الإلكتروني نفسه.');
 const role=roles[0],places=mine.map(r=>String(r[h['المكان']]||'').trim()).filter(Boolean);
 if(role==='مالك'){
  if(mine.some(r=>String(r[h['المكان']]||'').trim()))throw Error('حساب المالك يجب ألا يرتبط بمكان محدد.');
  if(mine.length!==1)throw Error('يجب أن يظهر حساب المالك مرة واحدة فقط في ورقة الصلاحيات.');
  return {email,name:mine[0][h['الاسم']]||email,place:'',places:[],role};
 }
 if(!['مسؤول مكان','مستخدم مكان'].includes(role))throw Error('الدور غير معتمد. استخدم مالك أو مسؤول مكان.');
 if(places.length!==mine.length)throw Error('يجب تحديد مكان لكل سجل مستخدم.');
 if(new Set(places).size!==places.length)throw Error('يوجد تكرار لنفس المكان في سجلات هذا البريد الإلكتروني.');
 places.forEach(place=>{
  const assigned=active.filter(r=>String(r[h['المكان']]||'').trim()===place);
  const assignedEmails=[...new Set(assigned.map(r=>String(r[h['الإيميل']]||'').trim().toLowerCase()))];
  if(assignedEmails.length!==1||assignedEmails[0]!==email)throw Error('المكان '+place+' مرتبط بأكثر من حساب فعّال. يجب أن يكون له حساب مسؤول واحد فقط.');
 });
 return {email,name:mine.map(r=>r[h['الاسم']]).find(Boolean)||email,place:places.length===1?places[0]:'',places,role};
}
function refs_(){const {a,h}=rows_(S.items);const items=a.slice(1).filter(r=>r[h['اسم الصنف (فريد)']]).map(r=>({name:String(r[h['اسم الصنف (فريد)']]),unit:String(r[h['الوحدة']]||''),avgCost:Number(r[h['متوسط التكلفة (تلقائي)']]||r[h['سعر مبدئي تقديري (عدّله)']])||0}));const p=rows_(S.places);return {items,places:p.a.slice(1).filter(r=>r[p.h['المكان']]&&!String(r[p.h['المكان']]).startsWith('ملاحظة:')).map(r=>({name:String(r[p.h['المكان']]),type:String(r[p.h['النوع']]||'')}))};}
function verifyLogin(enteredEmail){
 const u=user_();
 const entered=String(enteredEmail||'').trim().toLowerCase();
 if(!entered)throw Error('أدخل بريدك الإلكتروني.');
 if(entered!==u.email)throw Error('البريد المدخل لا يطابق حساب Google المسجل حاليًا. سجّل الدخول بحسابك الصحيح.');
 return {ok:true,user:{name:u.name,email:u.email,role:u.role,place:u.place,places:u.places}};
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
 const balances=u.role==='مالك'?allBalances:Object.fromEntries(u.places.map(place=>[place,allBalances[place]||{qty:0,value:0}]));
 const vals=Object.values(balances);
 return {user:u,items:d.items,places:d.places,summary:{locations:balances,totalQty:vals.reduce((s,v)=>s+v.qty,0),totalValue:vals.reduce((s,v)=>s+v.value,0)},pending:p,recent:m,invoices:invoiceArchive_(u)};
}
function pending_(u){
 const {a,h}=rows_(S.pending),groups={};
 a.slice(1).forEach(r=>{
  // الطلبات تظهر للمالك للمتابعة، ولمسؤول المكان المستلم فقط للاعتماد.
  if(r[h['الحالة']]!=='بانتظار اعتماد'||!(u.role==='مالك'||(u.role==='مسؤول مكان'&&u.places.includes(String(r[h['إلى']]||'')))))return;
  const id=String(r[h['معرف الحركة']]);if(!groups[id])groups[id]={id,date:fmtDate_(r[h['التاريخ']]),type:r[h['نوع الحركة']],from:r[h['من']],to:r[h['إلى']],creator:r[h['المدخل']],notes:r[h['ملاحظات']],items:[]};
  groups[id].items.push({item:r[h['الصنف']],qty:r[h['الكمية']],unit:r[h['الوحدة']]});
 });
 return Object.values(groups);
}
function recent_(u){
 const {a,h}=rows_(S.moves);
 return a.slice(1).filter(r=>r[h['الصنف']]&&(u.role==='مالك'||u.places.includes(String(r[h['من (مكان)']]||''))||u.places.includes(String(r[h['إلى (مكان)']]||'')))).slice(-30).reverse().map(r=>({date:fmtDate_(r[h['التاريخ']]),type:r[h['نوع الحركة']],from:r[h['من (مكان)']],to:r[h['إلى (مكان)']],item:r[h['الصنف']],qty:r[h['الكمية']],unit:r[h['الوحدة']]}));
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
   if(transfer){if(!u.places.includes(from))throw Error('لا يمكنك التحويل إلا من أحد الأماكن المصرح لك بها.');}
   else if(inbound){if(!u.places.includes(to))throw Error('اختر مكانًا مصرحًا لك باستلام المشتريات أو الزيادة فيه.');}
   else if(outbound){if(!u.places.includes(from))throw Error('اختر مكانًا مصرحًا لك بالصرف منه.');to='';}
   else throw Error('هذه الحركة غير متاحة لصلاحيتك.');
  }
  if(transfer){
   if(!validPlaces.includes(from)||!validPlaces.includes(to)||from===to)throw Error('حدد مصدرًا ومستلمًا صحيحين.');
   if(u.role!=='مالك'&&!u.places.includes(from))throw Error('لا يمكنك التحويل من مكان غير مصرح لك به.');
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
  const mh=ensureHeaders_(mv,['معرف الحركة','التاريخ','نوع الحركة','من (مكان)','إلى (مكان)','الصنف','الكمية','سعر الوحدة (للشراء/الافتتاحي فقط)','الوحدة','القيمة (تلقائي)','رقم المرجع (فاتورة/إذن)','المسؤول','ملاحظات','رابط الفاتورة']);
  const values=clean.map(x=>makeRow_(mh,{'معرف الحركة':id,'التاريخ':date,'نوع الحركة':type,'من (مكان)':from,'إلى (مكان)':to,'الصنف':x.item.name,'الكمية':x.q,'سعر الوحدة (للشراء/الافتتاحي فقط)':x.price,'الوحدة':x.item.unit,'القيمة (تلقائي)':x.q*x.price,'رقم المرجع (فاتورة/إذن)':f.reference||'','المسؤول':u.email,'ملاحظات':(purchase?'المورد: '+from+' | ':'')+(f.notes||''),'رابط الفاتورة':invoiceUrlForMovement_(f.invoiceFileId,u,to||from,{reference:f.reference,date:date,supplier:purchase?from:'',movementId:id})},mv.getLastColumn()));
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
  if(matches.some(z=>String(z.r[x.h['من']]||'')!==from||String(z.r[x.h['إلى']]||'')!==to||String(z.r[x.h['المدخل']]||'')!==String(first[x.h['المدخل']]||'')))throw Error('بيانات التحويل غير متطابقة. تم إيقاف الاعتماد لحماية المخزون.');
  if(u.role!=='مالك'&&(u.role!=='مسؤول مكان'||!u.places.includes(to)))throw Error('هذه الحركة مخصصة لمسؤول المكان المستلم فقط، ولا تملك صلاحية اعتمادها.');
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


const INVOICE_FOLDER_PROP='SHAWARMATI_INVOICE_FOLDER_ID';
function invoiceFolder_(u){
 const props=PropertiesService.getScriptProperties();
 let id=props.getProperty(INVOICE_FOLDER_PROP),folder=null;
 if(id){try{folder=DriveApp.getFolderById(id);folder.getName();}catch(e){folder=null;}}
 if(!folder){
  folder=DriveApp.createFolder('شاورمتي - فواتير المشتريات');
  props.setProperty(INVOICE_FOLDER_PROP,folder.getId());
 }
 // Grant the authorized app users access to the shared invoice archive.
 const {a,h}=rows_(S.users);
 const emails=[u.email].concat(a.slice(1).filter(r=>['نعم','yes','true','1'].includes(String(r[h['فعّال']]||'').trim().toLowerCase())).map(r=>String(r[h['الإيميل']]||'').trim().toLowerCase()).filter(Boolean));
 [...new Set(emails)].forEach(email=>{try{folder.addEditor(email);}catch(e){}});
 return folder;
}
function invoiceSheet_(){
 const book=ss_();let sheet=book.getSheetByName('أرشيف الفواتير');
 if(!sheet)sheet=book.insertSheet('أرشيف الفواتير');
 ensureHeaders_(sheet,['معرف المستند','تاريخ الرفع','رقم الفاتورة','تاريخ الفاتورة','المورد','المكان','اسم الملف','نوع الملف','رابط المستند','معرف ملف Drive','رافع المستند','معرف حركة المخزون','الحالة']);
 return sheet;
}
function uploadInvoice(payload){
 const u=user_();
 if(!payload||!payload.data||!payload.name||!payload.mimeType)throw Error('اختر صورة أو ملف PDF أولاً.');
 const mime=String(payload.mimeType).toLowerCase();
 if(!['application/pdf','image/jpeg','image/png','image/webp'].includes(mime))throw Error('الملفات المسموحة: PDF وJPG وPNG وWEBP فقط.');
 const base64=String(payload.data).replace(/^data:[^,]*,/, '');
 if(base64.length>11*1024*1024)throw Error('حجم الملف كبير. الحد الأقصى 8 ميجابايت.');
 const bytes=Utilities.base64Decode(base64);
 const blob=Utilities.newBlob(bytes,mime,String(payload.name).replace(/[\\/:*?"<>|]/g,'_'));
 const folder=invoiceFolder_(u);
 const file=folder.createFile(blob);
 const now=new Date(),id='INV-'+Utilities.formatDate(now,'Asia/Riyadh','yyyyMMdd-HHmmss')+'-'+Math.floor(Math.random()*900+100);
 const sheet=invoiceSheet_(),headers=head_(sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0]);
 const place=String(payload.place||'').trim();
 if(u.role==='مالك'){
  if(place&&!refs_().places.some(p=>p.name===place))throw Error('المكان المحدد غير صحيح.');
 }else if(!u.places.includes(place))throw Error('اختر مكانًا مصرحًا لك برفع الفاتورة عليه.');
 const row=makeRow_(headers,{'معرف المستند':id,'تاريخ الرفع':now,'اسم الملف':file.getName(),'نوع الملف':mime,'رابط المستند':file.getUrl(),'معرف ملف Drive':file.getId(),'رافع المستند':u.email,'المكان':place,'الحالة':'محفوظ - بانتظار ربطه بحركة'},sheet.getLastColumn());
 appendAtFirstDataGap_(sheet,[row],Array.from({length:row.length},(_,i)=>i+1),[[1,row.length]]);
 audit_(u.email,'رفع فاتورة',id,file.getName());
 return {ok:true,id,fileId:file.getId(),url:file.getUrl(),name:file.getName(),message:'تم حفظ المستند في Google Drive وإضافته إلى أرشيف الفواتير.'};
}
function invoiceUrlForMovement_(fileId,u,place,meta){
 if(!fileId)return '';
 const sheet=invoiceSheet_(),data=rows_('أرشيف الفواتير'),index=data.a.slice(1).findIndex(r=>String(r[data.h['معرف ملف Drive']])===String(fileId));
 if(index<0)throw Error('ملف الفاتورة غير موجود في الأرشيف.');
 const row=data.a[index+1];
 if(u.role!=='مالك'&&(!u.places.includes(place)||String(row[data.h['المكان']]||'')!==place))throw Error('الفاتورة مرتبطة بمكان مختلف عن الحركة أو لا تملك صلاحية هذا المكان.');
 const url=String(row[data.h['رابط المستند']]||'');
 if(!url)throw Error('رابط المستند غير متاح.');
 if(meta){const r=index+2;const set=(name,value)=>{if(data.h[name]!==undefined&&value!==undefined&&value!==null)sheet.getRange(r,data.h[name]+1).setValue(value);};set('رقم الفاتورة',meta.reference||'');set('تاريخ الفاتورة',meta.date||'');set('المورد',meta.supplier||'');set('المكان',place||'');set('معرف حركة المخزون',meta.movementId||'');set('الحالة','مرتبط بحركة مخزون');}
 return url;
}
function invoiceArchive_(u){
 const sheet=invoiceSheet_(),{a,h}=rows_('أرشيف الفواتير');
 return a.slice(1).filter(r=>r[h['معرف المستند']]&&(u.role==='مالك'||u.places.includes(String(r[h['المكان']]||''))||String(r[h['رافع المستند']]||'').toLowerCase()===u.email)).slice(-100).reverse().map(r=>({
  id:String(r[h['معرف المستند']]||''),uploaded:fmtDate_(r[h['تاريخ الرفع']]),invoiceNo:String(r[h['رقم الفاتورة']]||''),supplier:String(r[h['المورد']]||''),place:String(r[h['المكان']]||''),name:String(r[h['اسم الملف']]||''),url:String(r[h['رابط المستند']]||''),status:String(r[h['الحالة']]||'')
 }));
}


// إعدادات Gemini محفوظة على الخادم فقط؛ لا تُرسل إلى المتصفح.
const GEMINI_MODEL_='gemini-3.8-flash';
const GEMINI_KEY_PROPERTY_='GEMINI_API_KEY';

function invoiceArchiveRow_(fileId,u){
 const data=rows_('أرشيف الفواتير');
 const index=data.a.slice(1).findIndex(r=>String(r[data.h['معرف ملف Drive']]||'')===String(fileId));
 if(index<0)throw Error('الفاتورة غير موجودة في الأرشيف.');
 const row=data.a[index+1];
 if(u.role!=='مالك'&&String(row[data.h['رافع المستند']]||'').toLowerCase()!==u.email&&!u.places.includes(String(row[data.h['المكان']]||'')))throw Error('لا تملك صلاحية الوصول إلى هذه الفاتورة.');
 return {row,index,headers:data.h,sheet:sh_('أرشيف الفواتير')};
}

function invoiceLearningExamples_(){
 const book=ss_();let sheet=book.getSheetByName('تعلم_الفواتير');
 if(!sheet)return [];
 const values=sheet.getDataRange().getValues();
 if(values.length<2)return [];
 const h=head_(values[0]);
 return values.slice(1).filter(r=>r[h['وصف الفاتورة']]&&r[h['الصنف الصحيح']])
  .slice(-30).map(r=>({description:String(r[h['وصف الفاتورة']]),item:String(r[h['الصنف الصحيح']]),unit:String(r[h['الوحدة الصحيحة']]||'')})).slice(-20);
}

function extractInvoiceData(fileId){
 const u=user_();
 if(!fileId)throw Error('لم يتم تحديد ملف الفاتورة.');
 const archived=invoiceArchiveRow_(fileId,u);
 const apiKey=PropertiesService.getScriptProperties().getProperty(GEMINI_KEY_PROPERTY_);
 if(!apiKey)throw Error('لم يتم إعداد مفتاح Gemini على الخادم. أضف GEMINI_API_KEY في Script Properties داخل Apps Script ثم أعد المحاولة.');
 const file=DriveApp.getFileById(String(fileId));
 const blob=file.getBlob(),bytes=blob.getBytes(),mime=String(blob.getContentType()||'').toLowerCase();
 if(!['application/pdf','image/jpeg','image/png','image/webp'].includes(mime))throw Error('صيغة الفاتورة غير مدعومة للاستخراج.');
 if(bytes.length>8*1024*1024)throw Error('حجم الفاتورة يتجاوز حد الاستخراج البالغ 8 ميجابايت.');
 const examples=invoiceLearningExamples_();
 const catalog=refs_().items.map(x=>({name:x.name,unit:x.unit}));
 const prompt=[
  'أنت محاسب ومختص OCR لفواتير المشتريات العربية والسعودية. اقرأ المستند المرفق بصريًا واستخرج البيانات الظاهرة فقط.',
  'أعد JSON صالحًا فقط بالمفاتيح التالية: supplier, invoice_number, invoice_date, currency, subtotal, vat_amount, total, items.',
  'كل عنصر في items يجب أن يحتوي: description, suggested_item, quantity, unit, unit_price, line_total, confidence, notes.',
  'قواعد صارمة: لا تخمّن ولا تكمل رقمًا غير واضح. إذا لم تستطع قراءة قيمة أعد null. الأرقام كقيم رقمية لا كنصوص. التاريخ بصيغة YYYY-MM-DD إذا أمكن وإلا null.',
  'انتبه لاتجاه العربية، الفواصل العشرية، الكميات العشرية، الوحدات، الخصومات، وضريبة القيمة المضافة. لا تعتبر رقم الباركود كمية أو سعرًا.',
  'لا تخلط بين إجمالي السطر وسعر الوحدة. إذا كان الإجمالي ظاهرًا فقط اترك unit_price null. لا تضف الأصناف غير الموجودة في الفاتورة.',
  'اقرأ كل صفحات PDF وكل أجزاء الفاتورة. إذا كانت الصورة غير واضحة أشر لذلك في notes ولا تخترع بيانات.',
  'كتالوج الأصناف المعتمد بالمخزون (استخدم suggested_item فقط عند وجود تطابق واضح، وإلا null): '+JSON.stringify(catalog),
  'أمثلة تصحيحات بشرية سابقة (استفد منها لمطابقة أوصاف الموردين بالأصناف، ولا تنسخ كميات أو أسعار): '+JSON.stringify(examples),
  'أعد البنية التالية فقط: {"supplier":null,"invoice_number":null,"invoice_date":null,"currency":"SAR","subtotal":null,"vat_amount":null,"total":null,"items":[{"description":"","suggested_item":null,"quantity":null,"unit":null,"unit_price":null,"line_total":null,"confidence":"low|medium|high","notes":""}]}'
 ].join('\n');
 const body={
  contents:[{role:'user',parts:[{text:prompt},{inline_data:{mime_type:mime,data:Utilities.base64Encode(bytes)}}]}],
  generationConfig:{temperature:0.1,maxOutputTokens:8192,responseMimeType:'application/json'}
 };
 const endpoint='https://generativelanguage.googleapis.com/v1beta/models/'+GEMINI_MODEL_+':generateContent';
 const response=UrlFetchApp.fetch(endpoint,{method:'post',contentType:'application/json',headers:{'x-goog-api-key':apiKey},payload:JSON.stringify(body),muteHttpExceptions:true});
 const status=response.getResponseCode(),responseText=response.getContentText();
 if(status<200||status>=300){
  let detail=responseText;try{const parsed=JSON.parse(responseText);detail=parsed.error?.message||responseText;}catch(e){}
  throw Error('تعذر استخراج الفاتورة من Gemini ('+status+'): '+String(detail).slice(0,500));
 }
 let envelope;try{envelope=JSON.parse(responseText);}catch(e){throw Error('استجابة Gemini غير مفهومة.');}
 const raw=(envelope.candidates?.[0]?.content?.parts||[]).map(p=>p.text||'').join('').trim();
 if(!raw)throw Error('لم يُرجع Gemini نصًا قابلًا للاستخراج. قد تكون الصورة غير واضحة أو محجوبة.');
 let result;try{result=JSON.parse(raw.replace(/^\x60\x60\x60(?:json)?\s*/i,'').replace(/\s*\x60\x60\x60$/,''));}catch(e){throw Error('تعذر تحليل نتيجة استخراج الفاتورة. أعد تصويرها بصورة أوضح.');}
 if(!Array.isArray(result.items))throw Error('لم يتمكن Gemini من تحديد بنود الفاتورة.');
 const items=result.items.map(x=>({
  description:String(x.description||'').trim(),
  suggested_item:x.suggested_item?String(x.suggested_item).trim():'',
  quantity:x.quantity===null||x.quantity===undefined?'':Number(x.quantity),
  unit:String(x.unit||''),
  unit_price:x.unit_price===null||x.unit_price===undefined?'':Number(x.unit_price),
  line_total:x.line_total===null||x.line_total===undefined?'':Number(x.line_total),
  confidence:['high','medium','low'].includes(String(x.confidence||'').toLowerCase())?String(x.confidence).toLowerCase():'low',
  notes:String(x.notes||'')
 })).filter(x=>x.description||x.quantity!==''||x.line_total!=='');
 if(!items.length)throw Error('لم يتم العثور على بنود مقروءة. جرّب تصوير الفاتورة بشكل مستقيم وبإضاءة جيدة.');
 return {
  ok:true,model:GEMINI_MODEL_,
  invoice:{supplier:String(result.supplier||''),invoice_number:String(result.invoice_number||''),invoice_date:String(result.invoice_date||''),currency:String(result.currency||'SAR'),subtotal:result.subtotal??'',vat_amount:result.vat_amount??'',total:result.total??''},
  items
 };
}


function invoicePdfEscape_(value){
 return String(value===null||value===undefined?'':value)
  .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
  .replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}
function generateInvoicePdf(payload){
 const u=user_();
 if(!payload||!payload.fileId||!Array.isArray(payload.items)||!payload.items.length)throw Error('لا توجد بيانات فاتورة صالحة لإنشاء PDF.');
 invoiceArchiveRow_(payload.fileId,u);
 if(payload.items.length>300)throw Error('عدد البنود يتجاوز الحد المسموح به للتقرير.');
 const inv=payload.invoice||{};
 const safe=(v)=>invoicePdfEscape_(v);
 const n=(v)=>{if(v===null||v===undefined||v==='')return '—';const x=Number(v);return Number.isFinite(x)?x.toLocaleString('en-US',{maximumFractionDigits:4}):'—';};
 const rows=payload.items.map((x,i)=>'<tr><td>'+(i+1)+'</td><td>'+safe(x.description)+'</td><td>'+safe(x.item)+'</td><td>'+n(x.quantity)+'</td><td>'+safe(x.unit)+'</td><td>'+n(x.unitPrice)+'</td><td>'+n(x.lineTotal)+'</td></tr>').join('');
 const now=Utilities.formatDate(new Date(),'Asia/Riyadh','yyyy-MM-dd HH:mm');
 const doc='<!doctype html><html lang="ar" dir="rtl"><head><meta charset="utf-8"><style>'+
 'body{font-family:Arial,Tahoma,sans-serif;color:#172b4d;padding:26px;font-size:12px}h1{font-size:21px;margin:0 0 5px}.sub{color:#667085;margin-bottom:22px}.meta{display:grid;grid-template-columns:1fr 1fr;gap:9px;margin-bottom:18px}.box{border:1px solid #dce3ec;border-radius:7px;padding:10px}.label{color:#667085;font-size:10px;margin-bottom:4px}.value{font-weight:bold}table{width:100%;border-collapse:collapse;margin-top:12px}th,td{border:1px solid #dce3ec;padding:8px;text-align:right;font-size:10px}th{background:#eef2f7}tfoot td{font-weight:bold;background:#f7f9fc}.totals{margin-top:18px;width:48%;margin-right:auto}.foot{margin-top:24px;color:#667085;font-size:10px;border-top:1px solid #dce3ec;padding-top:10px}</style></head><body>'+
 '<h1>شاورمتي | تقرير بيانات الفاتورة</h1><div class="sub">تقرير بيانات مستخرجة ومراجعة من المستخدم — لا يُعد مستندًا ضريبيًا بديلًا عن الفاتورة الأصلية</div>'+
 '<div class="meta"><div class="box"><div class="label">المورد</div><div class="value">'+safe(inv.supplier||'غير محدد')+'</div></div>'+
 '<div class="box"><div class="label">رقم الفاتورة</div><div class="value">'+safe(inv.invoiceNumber||'غير محدد')+'</div></div>'+
 '<div class="box"><div class="label">تاريخ الفاتورة</div><div class="value">'+safe(inv.invoiceDate||'غير محدد')+'</div></div>'+
 '<div class="box"><div class="label">العملة</div><div class="value">'+safe(inv.currency||'SAR')+'</div></div></div>'+
 '<table><thead><tr><th>#</th><th>الوصف بالفاتورة</th><th>الصنف بالمخزون</th><th>الكمية</th><th>الوحدة</th><th>سعر الوحدة</th><th>إجمالي السطر</th></tr></thead><tbody>'+rows+'</tbody></table>'+
 '<table class="totals"><tbody><tr><td>الإجمالي قبل الضريبة</td><td>'+n(inv.subtotal)+'</td></tr><tr><td>ضريبة القيمة المضافة</td><td>'+n(inv.vatAmount)+'</td></tr><tr><td>الإجمالي النهائي</td><td><b>'+n(inv.total)+' '+safe(inv.currency||'SAR')+'</b></td></tr></tbody></table>'+
 '<div class="foot">تاريخ إنشاء التقرير: '+safe(now)+'<br>تم إنشاء هذا التقرير من البيانات التي راجعها المستخدم. يرجى الرجوع إلى صورة الفاتورة الأصلية للتحقق النهائي.</div></body></html>';
 const blob=HtmlService.createHtmlOutput(doc).getBlob().getAs(MimeType.PDF).setName('Invoice-Report-'+Utilities.formatDate(new Date(),'Asia/Riyadh','yyyyMMdd-HHmmss')+'.pdf');
 const file=invoiceFolder_(u).createFile(blob);
 audit_(u.email,'إنشاء تقرير PDF للفاتورة',String(payload.fileId),file.getName());
 return {ok:true,url:file.getUrl(),name:file.getName(),message:'تم إنشاء تقرير PDF وحفظه في مجلد فواتير المشتريات.'};
}

function saveInvoiceLearning(payload){
 const u=user_();
 if(!payload||!payload.fileId||!Array.isArray(payload.originalItems)||!Array.isArray(payload.correctedItems))throw Error('بيانات التصحيح غير مكتملة.');
 invoiceArchiveRow_(payload.fileId,u);
 if(payload.originalItems.length!==payload.correctedItems.length)throw Error('عدد البنود المصححة لا يطابق عدد البنود المستخرجة.');
 const book=ss_();let sheet=book.getSheetByName('تعلم_الفواتير');
 if(!sheet){sheet=book.insertSheet('تعلم_الفواتير');sheet.appendRow(['تاريخ التصحيح','المستخدم','معرف الفاتورة','وصف الفاتورة','الصنف المقترح من AI','الصنف الصحيح','الوحدة المستخرجة','الوحدة الصحيحة']);}
 const now=new Date(),rows=[];
 payload.correctedItems.forEach((x,i)=>{
  const original=payload.originalItems[i]||{};
  const description=String(original.description||'').trim(),item=String(x.item||'').trim();
  if(description&&item)rows.push([now,u.email,String(payload.fileId),description,String(original.suggested_item||''),item,String(original.unit||''),String(x.unit||'')]);
 });
 if(!rows.length)throw Error('لا توجد تصحيحات صنف قابلة للحفظ.');
 sheet.getRange(sheet.getLastRow()+1,1,rows.length,rows[0].length).setValues(rows);
 audit_(u.email,'تعليم مستمر لاستخراج الفواتير',String(payload.fileId),'حفظ '+rows.length+' تصحيحًا بشريًا للاستفادة منها في المطابقات القادمة.');
 return {ok:true,message:'تم حفظ تصحيحاتك لتُستخدم كأمثلة في عمليات استخراج الفواتير القادمة.'};
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
