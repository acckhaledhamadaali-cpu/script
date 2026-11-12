/**
 * المرحلة 3 (Core): أرصدة المكان، كارتة الصنف، ملخص التشغيل، إقفال الفترات، تنبيه البريد.
 * تعتمد على: LEDGER, ST, Q, readQueue_, flags_, sidePlace_, loadItems_, fmt_ من 05 core.gs
 * التشغيل مرة واحدة من المحرر: addPhase3Settings ثم authorizeMail (إن أردت تنبيه البريد)
 */

/* ============ إعدادات المرحلة 3 ============ */
function addPhase3Settings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(TAB.SETTINGS);
  const last = Math.max(sh.getLastRow(), 1);
  const keys = sh.getRange(1, 1, last, 1).getValues().map(function (r) { return String(r[0]).trim(); });
  const wanted = [
    ['تاريخ_الإقفال', '', 'لا يقبل النظام من غير المالك حركة بتاريخ يساوي أو يسبق هذا التاريخ. الصيغة 2026-10-31. اتركه فارغًا لعدم الإقفال'],
    ['تنبيه_البريد', 'لا', 'نعم = يصل المعتمد بريد عند إدخال حركة بانتظاره (يتطلب تشغيل authorizeMail مرة واحدة)'],
    ['رابط_التطبيق', '', 'اختياري: رابط الويب آب ليظهر في رسالة التنبيه']
  ];
  const added = [];
  wanted.forEach(function (w) {
    if (keys.indexOf(w[0]) >= 0) return;
    const row = sh.getLastRow() + 1;
    sh.getRange(row, 1, 1, 3).setValues([w]);
    if (w[0] === 'تاريخ_الإقفال') sh.getRange(row, 2).setNumberFormat('@');
    added.push(w[0]);
  });
  const msg = added.length ? 'أُضيفت: ' + added.join('، ') : 'الإعدادات موجودة مسبقًا';
  Logger.log(msg);
  return msg;
}

function authorizeMail() {
  const q = MailApp.getRemainingDailyQuota();
  Logger.log('تمت الموافقة على البريد. الحصة اليومية المتبقية: ' + q);
  return q;
}

/* ============ إقفال الفترات ============ */
// يرجع نص yyyy-MM-dd أو '' إن لم يُضبط
function closeDateStr_(ctx) {
  const raw = ctx.settings['تاريخ_الإقفال'];
  if (raw === '' || raw === undefined || raw === null) return '';
  if (raw instanceof Date) return Utilities.formatDate(raw, ctx.tz, 'yyyy-MM-dd');
  const s = String(raw).trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    throw new Error('صيغة تاريخ_الإقفال في الإعدادات غير صحيحة (المطلوب مثل 2026-10-31)');
  }
  return s;
}

function assertOpenPeriod_(ctx, date) {
  if (ctx.user.isOwner) return;
  const close = closeDateStr_(ctx);
  if (!close) return;
  const d = fmt_(ctx, date, 'yyyy-MM-dd');
  if (d <= close) {
    throw new Error('الفترة مقفلة حتى ' + close + '، لا يمكن إدخال أو اعتماد حركة بتاريخ ' + d);
  }
}

/* ============ دوال حسابية بحتة (قابلة للاختبار) ============ */
function round4_(n) { return Math.round(Number(n) * 10000) / 10000; }

// rows: [من، إلى، صنف، كمية]
function stockByItem_(rows, place) {
  const m = {};
  rows.forEach(function (r) {
    const from = String(r[0]).trim(), to = String(r[1]).trim(), item = String(r[2]).trim(), q = Number(r[3]);
    if (!item || !q) return;
    if (to === place) m[item] = (m[item] || 0) + q;
    if (from === place) m[item] = (m[item] || 0) - q;
  });
  return m;
}

// rows: صفوف الدفتر A..L: [تاريخ، نوع، من، إلى، صنف، كمية، سعر، وحدة، قيمة، مرجع، اسم، ملاحظات]
function cardRows_(rows, place, item) {
  const out = [];
  let bal = 0;
  rows.forEach(function (r) {
    if (String(r[4]).trim() !== item) return;
    const from = String(r[2]).trim(), to = String(r[3]).trim(), q = Number(r[5]);
    if (!q) return;
    const inQ = (to === place) ? q : 0;
    const outQ = (from === place) ? q : 0;
    if (!inQ && !outQ) return;
    bal += inQ - outQ;
    out.push({
      date: r[0], type: r[1], from: from, to: to, inQty: inQ, outQty: outQ, balance: round4_(bal),
      ref: r[9], by: r[10], notes: String(r[11] || '').replace(/\s*\[T\d+\]\s*$/, '')
    });
  });
  return out;
}

/* ============ الإجراءات الجديدة ============ */
function viewPlace_(ctx, p) {
  let place = String(p.place || '').trim();
  if (!ctx.user.isOwner) {
    if (!ctx.user.place) throw new Error('حسابك غير مرتبط بمكان');
    place = ctx.user.place;
  }
  if (!place) place = ctx.places[0];
  if (ctx.places.indexOf(place) < 0) throw new Error('المكان غير معروف');
  return place;
}

function balances_(ctx, p) {
  const place = viewPlace_(ctx, p);
  loadItems_(ctx);
  const led = ctx.ss.getSheetByName(LEDGER).getRange(2, 3, 3000, 4).getValues(); // C..F
  const stock = stockByItem_(led, place);
  const outP = {}, inP = {};
  readQueue_(ctx.ss).forEach(function (q) {
    if (q.status !== ST.PEND) return;
    if (q.from === place) outP[q.item] = (outP[q.item] || 0) + q.qty;
    if (q.to === place) inP[q.item] = (inP[q.item] || 0) + q.qty;
  });
  const names = ctx.itemsList.map(function (i) { return i.name; });
  Object.keys(stock).forEach(function (n) { if (names.indexOf(n) < 0) names.push(n); });
  const rows = [];
  names.forEach(function (n) {
    const qty = round4_(stock[n] || 0), o = round4_(outP[n] || 0), i = round4_(inP[n] || 0);
    if (!qty && !o && !i) return;
    rows.push({ item: n, unit: ctx.items[n] || '', qty: qty, outPending: o, inPending: i, available: round4_(qty - o) });
  });
  return { place: place, rows: rows };
}

function card_(ctx, p) {
  const place = viewPlace_(ctx, p);
  const item = String(p.item || '').trim();
  const items = loadItems_(ctx);
  if (!Object.prototype.hasOwnProperty.call(items, item)) throw new Error('الصنف غير موجود في تبويب الأصناف');
  const led = ctx.ss.getSheetByName(LEDGER).getRange(2, 1, 3000, 12).getValues();
  const all = cardRows_(led, place, item);
  const lastBal = all.length ? all[all.length - 1].balance : 0;
  const rows = all.slice(-150).reverse().map(function (r) {
    return { date: fmt_(ctx, r.date, 'yyyy-MM-dd'), type: r.type, from: r.from, to: r.to,
      inQty: r.inQty, outQty: r.outQty, balance: r.balance, ref: r.ref, by: r.by, notes: r.notes };
  });
  return { place: place, item: item, unit: items[item], balance: lastBal, count: all.length, rows: rows };
}

function summary_(ctx) {
  const queue = readQueue_(ctx.ss);
  const now = Date.now();
  const today = fmt_(ctx, new Date(), 'yyyy-MM-dd');
  let toApprove = 0, mineOpen = 0, approvedToday = 0, rejectedToday = 0, stale = 0, pendingTotal = 0;
  const byWho = {};
  queue.forEach(function (x) {
    if (x.status === ST.PEND) {
      pendingTotal++;
      if (flags_(ctx, x).canApprove) toApprove++;
      if (x.by === ctx.user.email) mineOpen++;
      const rule = ctx.rules[x.type];
      const who = rule ? (rule.approve === 'المالك' ? 'المالك' : (sidePlace_(rule.approve, x) || '—')) : '—';
      const ageH = (now - new Date(x.at).getTime()) / 3600000;
      if (ageH > 24) stale++;
      const g = byWho[who] || (byWho[who] = { who: who, count: 0, oldestHours: 0 });
      g.count++;
      if (ageH > g.oldestHours) g.oldestHours = ageH;
    } else if (x.apprAt instanceof Date && fmt_(ctx, x.apprAt, 'yyyy-MM-dd') === today) {
      if (x.status === ST.OK) approvedToday++;
      else if (x.status === ST.REJ) rejectedToday++;
    }
  });
  const out = { toApprove: toApprove, mineOpen: mineOpen };
  if (ctx.user.isOwner) {
    const used = ctx.ss.getSheetByName(LEDGER).getRange(2, 5, 3000, 1).getValues()
      .filter(function (r) { return String(r[0]).trim() !== ''; }).length;
    let close = '';
    try { close = closeDateStr_(ctx); } catch (e) { close = 'غير صحيح'; }
    out.pendingTotal = pendingTotal;
    out.stale = stale;
    out.approvedToday = approvedToday;
    out.rejectedToday = rejectedToday;
    out.ledgerUsed = used;
    out.ledgerCap = 3000;
    out.closeDate = close;
    out.byWho = Object.keys(byWho).map(function (k) {
      return { who: k, count: byWho[k].count, oldestHours: Math.round(byWho[k].oldestHours * 10) / 10 };
    }).sort(function (a, b) { return b.oldestHours - a.oldestHours; });
  }
  return out;
}

/* ============ تنبيه البريد (اختياري) ============ */
function notifyApprovers_(ctx, id, e) {
  try {
    if (String(ctx.settings['تنبيه_البريد'] || '').trim() !== 'نعم') return;
    const rule = ctx.rules[e.type];
    if (!rule || rule.approve === 'تلقائي') return;
    const place = rule.approve === 'المالك' ? '' : sidePlace_(rule.approve, e);
    const to = [];
    (ctx.permRows || []).forEach(function (r) {
      const mail = String(r[0]).toLowerCase().trim(), role = String(r[3]).trim(), active = String(r[4]).trim();
      if (!mail || active !== 'نعم' || mail === ctx.user.email) return;
      const ok = rule.approve === 'المالك' ? role === 'مالك' : (role === 'مستخدم مكان' && String(r[2]).trim() === place);
      if (ok && to.indexOf(mail) < 0) to.push(mail);
    });
    if (!to.length) return;
    const link = String(ctx.settings['رابط_التطبيق'] || '').trim();
    const body = 'حركة جديدة بانتظار اعتمادك\n\n' +
      'النوع: ' + e.type + '\nالصنف: ' + e.item + '\nالكمية: ' + e.qty + ' ' + e.unit + '\n' +
      'من: ' + (e.from || '-') + '  إلى: ' + (e.to || '-') + '\nالمرجع: ' + (e.ref || '-') + '\n' +
      'أدخلها: ' + nameOf_(ctx, ctx.user.email) + '\nرقم الحركة: ' + id + (link ? '\n\nافتح التطبيق: ' + link : '');
    MailApp.sendEmail({ to: to.join(','), subject: 'حركة بانتظار اعتمادك: ' + e.type + ' · ' + e.item, body: body, name: 'نظام المخازن' });
  } catch (err) { /* التنبيه لا يعطّل الحفظ */ }
}
