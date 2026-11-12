/**
 * نواة النظام (Core): تعمل كتطبيق ويب "باسم المالك" وتستقبل الطلبات من البوابة فقط.
 * كل القواعد (الصلاحيات، قفل التعديل، الاعتماد، الأرصدة) تُنفَّذ هنا في الكود.
 * تعتمد على الثوابت TAB وlockToOwner_ الموجودة في 02_Setup_Phase1.gs
 */

const LEDGER = 'الحركات';
const ITEMS_TAB = 'الأصناف';
const Q = { ID: 0, DATE: 1, TYPE: 2, FROM: 3, TO: 4, ITEM: 5, QTY: 6, PRICE: 7, UNIT: 8, REF: 9,
  NOTES: 10, BY: 11, AT: 12, STATUS: 13, APPR: 14, APPRAT: 15, EDITAT: 16, REASON: 17, LROW: 18 };
const QCOLS = 19;
const ST = { PEND: 'قيد الانتظار', OK: 'معتمد', REJ: 'مرفوض', CAN: 'ملغي' };
const SHAPE = {
  'رصيد افتتاحي': { from: false, to: true, price: true },
  'شراء': { from: false, to: true, price: true },
  'تسوية زيادة': { from: false, to: true, price: false },
  'تحويل': { from: true, to: true, price: false },
  'صرف للتشغيل': { from: true, to: false, price: false },
  'هالك': { from: true, to: false, price: false },
  'مرتجع للمورد': { from: true, to: false, price: false },
  'تسوية عجز': { from: true, to: false, price: false }
};

/* ============ تهيئة (تُشغَّل مرة واحدة من المحرر) ============ */
function initCore() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('SS_ID', SpreadsheetApp.getActiveSpreadsheet().getId());
  if (!props.getProperty('GATEWAY_SECRET')) {
    props.setProperty('GATEWAY_SECRET', Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, ''));
  }
  Logger.log('انسخ هذا المفتاح السري إلى البوابة (SECRET):\n' + props.getProperty('GATEWAY_SECRET'));
}

// اختبار بدون بوابة: يعمل بحساب المالك ويطبع ملخص البيانات التي ستراها الواجهة
function coreSelfTest() {
  const email = Session.getEffectiveUser().getEmail().toLowerCase();
  const m = handle_(email, 'meta', {});
  Logger.log(JSON.stringify({
    user: m.user,
    places: m.places.map(function (p) { return p.name + ' (' + p.type + ')'; }),
    items: m.items.length,
    types: m.types.map(function (t) { return t.type; }),
    windowMin: m.windowMin
  }));
}

// يُشغَّل بعد نجاح كل الاختبارات: قفل تبويب الحركات للمالك فقط
function lockLedgerToOwner() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  lockToOwner_(ss.getSheetByName(LEDGER), 'دفتر الحركات المعتمدة - للمالك فقط');
  ss.toast('تم قفل تبويب الحركات للمالك', 'الصلاحيات', 6);
}

/* ============ نقطة الدخول ============ */
function doPost(e) {
  let res;
  try {
    const req = JSON.parse(e.postData.contents);
    const secret = PropertiesService.getScriptProperties().getProperty('GATEWAY_SECRET');
    if (!secret || req.secret !== secret) throw new Error('غير مصرح');
    res = { ok: true, data: handle_(String(req.email || '').toLowerCase().trim(), req.action, req.payload || {}) };
  } catch (err) {
    res = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
}

function handle_(email, action, p) {
  const ctx = buildCtx_(email);
  switch (action) {
    case 'meta': return meta_(ctx);
    case 'list': return list_(ctx, p);
    case 'available': return available_(ctx, p);
    case 'submit': return submit_(ctx, p);
    case 'edit': return edit_(ctx, p);
    case 'cancel': return cancel_(ctx, p);
    case 'approve': return approve_(ctx, p);
    case 'reject': return reject_(ctx, p);
    case 'balances': return balances_(ctx, p);
    case 'card': return card_(ctx, p);
    case 'summary': return summary_(ctx);
    default: throw new Error('إجراء غير معروف');
  }
}

/* ============ تحميل السياق ============ */
function ss_() {
  const id = PropertiesService.getScriptProperties().getProperty('SS_ID');
  if (!id) throw new Error('النواة غير مهيأة (شغّل initCore)');
  return SpreadsheetApp.openById(id);
}

function buildCtx_(email) {
  const ss = ss_();
  const ctx = { ss: ss, tz: ss.getSpreadsheetTimeZone() };
  const pl = loadPlaces_(ss);
  ctx.places = pl.list;
  ctx.placeType = pl.types;
  ctx.user = loadUser_(ctx, email);
  ctx.settings = loadSettings_(ss);
  const rawWin = ctx.settings['مدة_سماح_التعديل_بالدقائق'];
  ctx.windowMin = (rawWin === '' || rawWin === undefined || !isFinite(Number(rawWin)) || Number(rawWin) < 0) ? 30 : Number(rawWin);
  ctx.allowNeg = ctx.settings['السماح_بالرصيد_السالب'] === 'نعم';
  const rl = loadRules_(ss);
  ctx.rules = rl.rules;
  ctx.transfers = rl.transfers;
  return ctx;
}

function loadPlaces_(ss) {
  const v = ss.getSheetByName(TAB.PLACES).getRange('A2:B9').getValues();
  const list = [];
  const types = {};
  v.forEach(function (r) {
    const n = String(r[0]).trim();
    if (n) { list.push(n); types[n] = String(r[1]).trim(); }
  });
  return { list: list, types: types };
}

function loadUser_(ctx, email) {
  const sh = ctx.ss.getSheetByName(TAB.PERMS);
  const rows = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 5).getValues();
  ctx.permRows = rows;
  let u = null;
  rows.forEach(function (r) {
    if (!u && email && String(r[0]).toLowerCase().trim() === email) {
      u = { email: email, name: String(r[1] || '').trim(), place: String(r[2] || '').trim(),
        role: String(r[3]).trim(), active: String(r[4]).trim() };
    }
  });
  if (!u) throw new Error('هذا الحساب (' + email + ') غير مسجل في نظام الصلاحيات. تواصل مع المالك');
  if (u.active !== 'نعم') throw new Error('هذا الحساب غير مفعّل');
  u.isOwner = u.role === 'مالك';
  u.canWrite = u.isOwner || u.role === 'مستخدم مكان';
  if (u.role === 'مستخدم مكان' && ctx.places.indexOf(u.place) < 0) {
    throw new Error('المكان المسجل لحسابك غير موجود في تبويب الأماكن');
  }
  return u;
}

function nameOf_(ctx, email) {
  let n = email;
  (ctx.permRows || []).forEach(function (r) {
    if (String(r[0]).toLowerCase().trim() === email && String(r[1]).trim()) n = String(r[1]).trim();
  });
  return n;
}

function loadSettings_(ss) {
  const v = ss.getSheetByName(TAB.SETTINGS).getRange('A2:B20').getValues();
  const m = {};
  v.forEach(function (r) { if (r[0] !== '') m[String(r[0]).trim()] = r[1] === '' ? '' : (r[1] instanceof Date || typeof r[1] === 'number' ? r[1] : String(r[1]).trim()); });
  return m;
}

function loadRules_(ss) {
  const v = ss.getSheetByName(TAB.RULES).getRange(1, 1, 40, 5).getValues();
  const rules = {};
  const transfers = {};
  let i = 1;
  for (; i < v.length && String(v[i][0]).trim() !== ''; i++) {
    rules[String(v[i][0]).trim()] = {
      enter: String(v[i][1]).trim(), approve: String(v[i][2]).trim(), self: String(v[i][3]).trim()
    };
  }
  for (; i < v.length; i++) {
    if (String(v[i][0]).trim() === 'من (نوع المكان)') {
      for (let j = i + 1; j < v.length && String(v[j][0]).trim() !== ''; j++) {
        transfers[String(v[j][0]).trim() + '>' + String(v[j][1]).trim()] = String(v[j][2]).trim() === 'نعم';
      }
      break;
    }
  }
  return { rules: rules, transfers: transfers };
}

function loadItems_(ctx) {
  if (ctx.items) return ctx.items;
  const sh = ctx.ss.getSheetByName(ITEMS_TAB);
  const names = sh.getRange('B3:B102').getValues();
  const units = sh.getRange('D3:D102').getValues();
  const m = {};
  const list = [];
  names.forEach(function (r, i) {
    const n = String(r[0]).trim();
    if (n) { m[n] = String(units[i][0]).trim(); list.push({ name: n, unit: m[n] }); }
  });
  ctx.items = m;
  ctx.itemsList = list;
  return m;
}

/* ============ الدفتر والمعلق ============ */
function readQueue_(ss) {
  const sh = ss.getSheetByName(TAB.QUEUE);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, QCOLS).getValues()
    .map(function (r, i) { return toQ_(r, i + 2); })
    .filter(function (q) { return q.id; });
}

function toQ_(r, rowNum) {
  return {
    row: rowNum, id: String(r[Q.ID]), date: r[Q.DATE], type: r[Q.TYPE], from: r[Q.FROM], to: r[Q.TO],
    item: r[Q.ITEM], qty: Number(r[Q.QTY]), price: r[Q.PRICE], unit: r[Q.UNIT], ref: r[Q.REF],
    notes: r[Q.NOTES], by: String(r[Q.BY]).toLowerCase(), at: r[Q.AT], status: r[Q.STATUS],
    appr: r[Q.APPR], apprAt: r[Q.APPRAT], editAt: r[Q.EDITAT], reason: r[Q.REASON], lrow: r[Q.LROW]
  };
}

function ledgerBalances_(ctx) {
  const v = ctx.ss.getSheetByName(LEDGER).getRange(2, 3, 3000, 4).getValues(); // C من، D إلى، E صنف، F كمية
  const m = {};
  v.forEach(function (r) {
    const from = String(r[0]).trim(), to = String(r[1]).trim(), item = String(r[2]).trim(), q = Number(r[3]);
    if (!item || !q) return;
    if (to) m[to + '|' + item] = (m[to + '|' + item] || 0) + q;
    if (from) m[from + '|' + item] = (m[from + '|' + item] || 0) - q;
  });
  return m;
}

// المتاح = الرصيد المعتمد - التحويلات/الصرف الصادر المعلق
function availableFor_(ctx, place, item, excludeId) {
  let a = ledgerBalances_(ctx)[place + '|' + item] || 0;
  readQueue_(ctx.ss).forEach(function (q) {
    if (q.status === ST.PEND && q.from === place && q.item === item && q.id !== excludeId) a -= q.qty;
  });
  return a;
}

function assertAvailable_(ctx, e, excludeId) {
  if (!e.from || ctx.allowNeg) return;
  const a = availableFor_(ctx, e.from, e.item, excludeId);
  if (e.qty > a + 1e-9) {
    throw new Error('الكمية (' + e.qty + ') أكبر من المتاح في ' + e.from + ' (' + a + ' ' + e.unit + ')');
  }
}

function postToLedger_(ctx, q) {
  const sh = ctx.ss.getSheetByName(LEDGER);
  const v = sh.getRange(2, 1, 3000, 5).getValues(); // A..E
  let row = -1;
  for (let i = 0; i < v.length; i++) {
    if (v[i][0] === '' && v[i][1] === '' && v[i][4] === '') { row = i + 2; break; }
  }
  if (row < 0) throw new Error('جدول الحركات ممتلئ (3000 صف). يلزم أرشفة');
  // نكتب فقط في الأعمدة غير المعادلات: A-G وJ-L (H وI وM وN معادلات)
  sh.getRange(row, 1, 1, 7).setValues([[q.date, q.type, q.from, q.to, q.item, q.qty, q.price]]);
  sh.getRange(row, 10, 1, 3).setValues([[q.ref, nameOf_(ctx, q.by), (q.notes ? q.notes + ' ' : '') + '[' + q.id + ']']]);
  return row;
}

/* ============ الصلاحيات والقواعد ============ */
function sidePlace_(role, e) {
  if (role === 'المكان المرسل') return e.from;
  if (role === 'المكان المستلم') return e.to;
  if (role === 'المكان المعني') return e.from || e.to;
  return '';
}

function mayAct_(user, role, e) {
  if (user.isOwner) return true;
  if (role === 'المالك' || role === 'تلقائي' || !role) return false;
  return user.role === 'مستخدم مكان' && !!user.place && user.place === sidePlace_(role, e);
}

function withinWindow_(ctx, q) {
  if (ctx.windowMin === 0) return true;
  return (Date.now() - new Date(q.at).getTime()) / 60000 <= ctx.windowMin;
}

function flags_(ctx, q) {
  const rule = ctx.rules[q.type];
  let canEdit = false, canApprove = false;
  if (q.status === ST.PEND && rule) {
    if (ctx.user.isOwner) {
      canEdit = true; canApprove = true;
    } else {
      const mine = q.by === ctx.user.email;
      canEdit = mine && withinWindow_(ctx, q) && mayAct_(ctx.user, rule.enter, q);
      canApprove = mayAct_(ctx.user, rule.approve, q) && !(mine && rule.self !== 'نعم');
    }
  }
  return { canEdit: canEdit, canApprove: canApprove };
}

function assertMayEnter_(ctx, e) {
  const rule = ctx.rules[e.type];
  if (!ctx.user.canWrite) throw new Error('صلاحيتك للعرض فقط');
  if (!mayAct_(ctx.user, rule.enter, e)) throw new Error('غير مسموح لك بإدخال حركة (' + e.type + ') لهذا المكان');
  if (e.type === 'تحويل' && !ctx.user.isOwner) {
    const key = (ctx.placeType[e.from] || '') + '>' + (ctx.placeType[e.to] || '');
    if (ctx.transfers[key] !== true) throw new Error('هذا النوع من التحويل مسموح للمالك فقط');
  }
}

function normalize_(ctx, p) {
  const e = {
    type: String(p.type || '').trim(),
    from: String(p.from || '').trim(),
    to: String(p.to || '').trim(),
    item: String(p.item || '').trim(),
    qty: Number(p.qty),
    price: (p.price === '' || p.price === null || p.price === undefined) ? '' : Number(p.price),
    ref: String(p.ref || '').trim(),
    notes: String(p.notes || '').trim()
  };
  const shape = SHAPE[e.type];
  if (!shape || !ctx.rules[e.type]) throw new Error('نوع الحركة غير معروف');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(p.date || ''))) throw new Error('التاريخ غير صحيح');
  e.date = Utilities.parseDate(String(p.date), ctx.tz, 'yyyy-MM-dd');
  assertOpenPeriod_(ctx, e.date);
  if (!isFinite(e.qty) || e.qty <= 0) throw new Error('الكمية يجب أن تكون رقمًا أكبر من صفر');
  const items = loadItems_(ctx);
  if (!Object.prototype.hasOwnProperty.call(items, e.item)) throw new Error('الصنف غير موجود في تبويب الأصناف');
  e.unit = items[e.item];
  if (shape.from) { if (!e.from) throw new Error('حدد المكان (من)'); } else if (e.from) { throw new Error('اترك (من) فارغًا لهذا النوع'); }
  if (shape.to) { if (!e.to) throw new Error('حدد المكان (إلى)'); } else if (e.to) { throw new Error('اترك (إلى) فارغًا لهذا النوع'); }
  if (e.from && ctx.places.indexOf(e.from) < 0) throw new Error('المكان (من) غير معروف');
  if (e.to && ctx.places.indexOf(e.to) < 0) throw new Error('المكان (إلى) غير معروف');
  if (e.from && e.to && e.from === e.to) throw new Error('(من) و(إلى) نفس المكان');
  if (shape.price) { if (!(e.price > 0)) throw new Error('حدد سعر الوحدة'); } else { e.price = ''; }
  return e;
}

/* ============ الإجراءات ============ */
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function nextId_() {
  const props = PropertiesService.getScriptProperties();
  const n = Number(props.getProperty('SEQ') || '0') + 1;
  props.setProperty('SEQ', String(n));
  return 'T' + ('00000' + n).slice(-5);
}

function audit_(ctx, action, id, details) {
  const sh = ctx.ss.getSheetByName(TAB.AUDIT);
  sh.getRange(sh.getLastRow() + 1, 1, 1, 5).setValues([[new Date(), ctx.user.email, action, id, details]]);
}

function findQ_(ctx, id) {
  const q = readQueue_(ctx.ss).filter(function (x) { return x.id === id; })[0];
  if (!q) throw new Error('الحركة غير موجودة');
  return q;
}

function approveRow_(ctx, rowNum, label) {
  const qs = ctx.ss.getSheetByName(TAB.QUEUE);
  const q = toQ_(qs.getRange(rowNum, 1, 1, QCOLS).getValues()[0], rowNum);
  if (q.status !== ST.PEND) throw new Error('الحركة ليست قيد الانتظار');
  assertOpenPeriod_(ctx, q.date);
  if (q.from && !ctx.allowNeg) {
    const bal = ledgerBalances_(ctx)[q.from + '|' + q.item] || 0;
    if (q.qty > bal + 1e-9) throw new Error('لا يمكن الاعتماد: رصيد ' + q.from + ' الحالي (' + bal + ') أقل من الكمية');
  }
  const lrow = postToLedger_(ctx, q);
  qs.getRange(rowNum, Q.STATUS + 1, 1, 2).setValues([[ST.OK, label]]);
  qs.getRange(rowNum, Q.APPRAT + 1).setValue(new Date());
  qs.getRange(rowNum, Q.LROW + 1).setValue(lrow);
  audit_(ctx, 'اعتماد', q.id, label + ' | صف الدفتر ' + lrow);
  return lrow;
}

function submit_(ctx, p) {
  return withLock_(function () {
    const e = normalize_(ctx, p);
    assertMayEnter_(ctx, e);
    assertAvailable_(ctx, e, null);
    const rule = ctx.rules[e.type];
    const qs = ctx.ss.getSheetByName(TAB.QUEUE);
    const id = nextId_();
    const now = new Date();
    const rowNum = qs.getLastRow() + 1;
    qs.getRange(rowNum, 1, 1, QCOLS).setValues([[id, e.date, e.type, e.from, e.to, e.item, e.qty, e.price,
      e.unit, e.ref, e.notes, ctx.user.email, now, ST.PEND, '', '', now, '', '']]);
    audit_(ctx, 'إدخال', id, e.type + ' | ' + e.item + ' | ' + e.qty + ' | ' + (e.from || '-') + ' > ' + (e.to || '-'));
    if (rule.approve !== 'تلقائي') notifyApprovers_(ctx, id, e);
    let status = ST.PEND;
    if (rule.approve === 'تلقائي') { approveRow_(ctx, rowNum, 'تلقائي'); status = ST.OK; }
    return { id: id, status: status };
  });
}

function edit_(ctx, p) {
  return withLock_(function () {
    const q = findQ_(ctx, String(p.id));
    if (!flags_(ctx, q).canEdit) {
      throw new Error('لا يمكن تعديل هذه الحركة (اعتُمدت، أو انتهت مهلة التعديل، أو لا تملك الصلاحية)');
    }
    const e = normalize_(ctx, p);
    assertMayEnter_(ctx, e);
    assertAvailable_(ctx, e, q.id);
    const qs = ctx.ss.getSheetByName(TAB.QUEUE);
    qs.getRange(q.row, Q.DATE + 1, 1, 10).setValues([[e.date, e.type, e.from, e.to, e.item, e.qty, e.price, e.unit, e.ref, e.notes]]);
    qs.getRange(q.row, Q.EDITAT + 1).setValue(new Date());
    audit_(ctx, 'تعديل', q.id, e.type + ' | ' + e.item + ' | ' + e.qty + ' | ' + (e.from || '-') + ' > ' + (e.to || '-'));
    if (ctx.rules[e.type].approve === 'تلقائي') approveRow_(ctx, q.row, 'تلقائي');
    return { id: q.id };
  });
}

function cancel_(ctx, p) {
  return withLock_(function () {
    const q = findQ_(ctx, String(p.id));
    if (!flags_(ctx, q).canEdit) throw new Error('لا يمكن إلغاء هذه الحركة');
    ctx.ss.getSheetByName(TAB.QUEUE).getRange(q.row, Q.STATUS + 1).setValue(ST.CAN);
    audit_(ctx, 'إلغاء', q.id, '');
    return { id: q.id };
  });
}

function approve_(ctx, p) {
  return withLock_(function () {
    const q = findQ_(ctx, String(p.id));
    if (!flags_(ctx, q).canApprove) throw new Error('لا تملك صلاحية اعتماد هذه الحركة');
    approveRow_(ctx, q.row, ctx.user.email);
    return { id: q.id };
  });
}

function reject_(ctx, p) {
  return withLock_(function () {
    const q = findQ_(ctx, String(p.id));
    if (!flags_(ctx, q).canApprove) throw new Error('لا تملك صلاحية رفض هذه الحركة');
    const reason = String(p.reason || '').trim();
    if (!reason) throw new Error('اكتب سبب الرفض');
    const qs = ctx.ss.getSheetByName(TAB.QUEUE);
    qs.getRange(q.row, Q.STATUS + 1, 1, 2).setValues([[ST.REJ, ctx.user.email]]);
    qs.getRange(q.row, Q.APPRAT + 1).setValue(new Date());
    qs.getRange(q.row, Q.REASON + 1).setValue(reason);
    audit_(ctx, 'رفض', q.id, reason);
    return { id: q.id };
  });
}

/* ============ القراءة ============ */
function fmt_(ctx, d, pattern) {
  return (d instanceof Date) ? Utilities.formatDate(d, ctx.tz, pattern) : String(d || '');
}

function out_(ctx, q) {
  const f = flags_(ctx, q);
  let until = '';
  if (q.status === ST.PEND && ctx.windowMin > 0 && q.at instanceof Date) {
    until = Utilities.formatDate(new Date(q.at.getTime() + ctx.windowMin * 60000), ctx.tz, 'HH:mm');
  }
  return {
    id: q.id, date: fmt_(ctx, q.date, 'yyyy-MM-dd'), type: q.type, from: q.from, to: q.to, item: q.item,
    qty: q.qty, price: q.price, unit: q.unit, ref: q.ref, notes: q.notes, by: nameOf_(ctx, q.by),
    at: fmt_(ctx, q.at, 'yyyy-MM-dd HH:mm'), status: q.status, appr: q.appr === 'تلقائي' ? 'تلقائي' : nameOf_(ctx, String(q.appr).toLowerCase()),
    apprAt: fmt_(ctx, q.apprAt, 'yyyy-MM-dd HH:mm'), reason: q.reason,
    canEdit: f.canEdit, canApprove: f.canApprove, editUntil: until
  };
}

function list_(ctx, p) {
  const u = ctx.user;
  let rows = readQueue_(ctx.ss).filter(function (q) {
    if (u.isOwner) return true;
    return q.by === u.email || (u.place && (q.from === u.place || q.to === u.place));
  });
  if (p.scope === 'pending') rows = rows.filter(function (q) { return q.status === ST.PEND; });
  rows.sort(function (a, b) { return new Date(b.at) - new Date(a.at); });
  return rows.slice(0, 80).map(function (q) { return out_(ctx, q); });
}

function available_(ctx, p) {
  const place = String(p.place || '').trim();
  if (!ctx.user.isOwner && place !== ctx.user.place) throw new Error('غير مسموح');
  return { available: availableFor_(ctx, place, String(p.item || '').trim(), null) };
}

function meta_(ctx) {
  loadItems_(ctx);
  const types = Object.keys(ctx.rules).filter(function (t) { return SHAPE[t]; }).map(function (t) {
    const r = ctx.rules[t], s = SHAPE[t];
    let canEnter = false, lock = '';
    if (ctx.user.isOwner) {
      canEnter = true;
    } else if (ctx.user.canWrite && ['المكان المرسل', 'المكان المستلم', 'المكان المعني'].indexOf(r.enter) >= 0) {
      canEnter = true;
      lock = r.enter === 'المكان المرسل' ? 'from' : (r.enter === 'المكان المستلم' ? 'to' : (s.from ? 'from' : 'to'));
    }
    return { type: t, from: s.from, to: s.to, price: s.price, lock: lock, canEnter: canEnter, approve: r.approve };
  });
  return {
    user: { name: ctx.user.name || ctx.user.email, email: ctx.user.email, role: ctx.user.role,
      place: ctx.user.place, isOwner: ctx.user.isOwner, canWrite: ctx.user.canWrite },
    places: ctx.places.map(function (n) { return { name: n, type: ctx.placeType[n] }; }),
    items: ctx.itemsList,
    types: types,
    windowMin: ctx.windowMin
  };
}
