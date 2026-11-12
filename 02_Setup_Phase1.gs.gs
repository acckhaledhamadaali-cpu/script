/**
 * المرحلة 1: إنشاء تبويبات النظام وقفلها للمالك.
 * آمن: لا يلمس تبويب "الحركات" ولا أي تبويب موجود، ولا يعيد إنشاء تبويب موجود.
 * التشغيل: شغّل الدالة setupPhase1 مرة واحدة من محرر Apps Script.
 */

const TAB = {
  PERMS: 'الصلاحيات',
  RULES: 'قواعد_الحركات',
  SETTINGS: 'الإعدادات',
  QUEUE: 'الحركات_المعلقة',
  AUDIT: 'سجل_التدقيق',
  PLACES: 'الأماكن'
};

function setupPhase1() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const report = [];

  report.push(buildPerms_(ss));
  report.push(buildRules_(ss));
  report.push(buildSettings_(ss));
  report.push(buildQueue_(ss));
  report.push(buildAudit_(ss));

  [TAB.PERMS, TAB.RULES, TAB.SETTINGS, TAB.QUEUE, TAB.AUDIT].forEach(function (name) {
    lockToOwner_(ss.getSheetByName(name), 'للمالك فقط - نظام صلاحيات المخازن');
  });
  report.push('تم قفل تبويبات النظام للمالك فقط');

  Logger.log(report.join('\n'));
  ss.toast('اكتمل إعداد المرحلة 1، راجع سجل التنفيذ', 'نظام الصلاحيات', 8);
}

/* ---------- الصلاحيات ---------- */
function buildPerms_(ss) {
  const r = ensureSheet_(ss, TAB.PERMS,
    ['الإيميل', 'الاسم', 'المكان', 'الدور', 'فعّال', 'ملاحظات'], [260, 160, 180, 130, 90, 260]);
  if (!r.isNew) return 'موجود مسبقًا: ' + TAB.PERMS;
  const sh = r.sheet;

  const owner = Session.getEffectiveUser().getEmail();
  sh.getRange(2, 1, 1, 6).setValues([[owner, 'المالك', '', 'مالك', 'نعم', 'صلاحية كاملة']]);

  const places = ss.getSheetByName(TAB.PLACES);
  if (places) {
    sh.getRange('C2:C200').setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInRange(places.getRange('A2:A30'), true).setAllowInvalid(false).build());
  }
  sh.getRange('D2:D200').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['مالك', 'مستخدم مكان', 'عرض فقط'], true).setAllowInvalid(false).build());
  sh.getRange('E2:E200').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['نعم', 'لا'], true).setAllowInvalid(false).build());
  sh.getRange('G1').setValue('ملاحظة: اكتب الإيميل بحروف صغيرة. المالك = كل الصلاحيات. مستخدم مكان = يدخل ويعتمد حسب قواعد الحركات للمكان المحدد.');
  return 'تم إنشاء: ' + TAB.PERMS;
}

/* ---------- قواعد الحركات ---------- */
function buildRules_(ss) {
  const r = ensureSheet_(ss, TAB.RULES,
    ['نوع الحركة', 'من يدخل', 'من يعتمد', 'اعتماد ذاتي', 'ملاحظات'], [160, 160, 160, 110, 380]);
  if (!r.isNew) return 'موجود مسبقًا: ' + TAB.RULES;
  r.sheet.getRange(2, 1, 3, 5).setValues([
    ['رصيد افتتاحي', 'المالك', 'تلقائي', 'لا ينطبق', 'يدخله المالك فقط ويُعتمد فورًا'],
    ['شراء', 'المكان المستلم', 'المكان المستلم', 'نعم', 'قرار الإدارة: المدخل والمعتمد نفس المكان. غيّر الاعتماد الذاتي إلى لا لتفعيل فصل المهام'],
    ['تحويل', 'المكان المرسل', 'المكان المستلم', 'لا', 'يشمل: مستودع/مطبخ إلى فرع، وفرع إلى فرع']
  ]);
  r.sheet.getRange('A6').setValue(
    'غير محدد حاليًا ويُسمح به للمالك فقط: تحويل فرع إلى مستودع/مطبخ (مرتجع)، وتحويل مستودع إلى مطبخ والعكس.');
  return 'تم إنشاء: ' + TAB.RULES;
}

/* ---------- الإعدادات ---------- */
function buildSettings_(ss) {
  const r = ensureSheet_(ss, TAB.SETTINGS, ['المفتاح', 'القيمة', 'الشرح'], [260, 120, 420]);
  if (!r.isNew) return 'موجود مسبقًا: ' + TAB.SETTINGS;
  r.sheet.getRange(2, 1, 2, 3).setValues([
    ['مدة_سماح_التعديل_بالدقائق', 30, 'يعدّل المدخل خلال هذه المدة من وقت الإدخال. اكتب 0 للسماح بالتعديل حتى الاعتماد بلا حد زمني'],
    ['السماح_بالرصيد_السالب', 'لا', 'لا = يُرفض أي تحويل أكبر من الرصيد المتاح (الرصيد المعتمد ناقص التحويلات الصادرة المعلقة)']
  ]);
  r.sheet.getRange('B3').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['نعم', 'لا'], true).build());
  return 'تم إنشاء: ' + TAB.SETTINGS;
}

/* ---------- الحركات المعلقة ---------- */
function buildQueue_(ss) {
  const r = ensureSheet_(ss, TAB.QUEUE, [
    'معرف الحركة', 'التاريخ', 'نوع الحركة', 'من', 'إلى', 'الصنف', 'الكمية', 'سعر الوحدة',
    'الوحدة', 'رقم المرجع', 'ملاحظات', 'المدخل', 'وقت الإدخال', 'الحالة', 'المعتمد',
    'وقت الاعتماد', 'آخر تعديل', 'سبب الرفض'
  ], [120, 100, 110, 150, 150, 200, 80, 90, 80, 120, 200, 220, 150, 100, 220, 150, 150, 200]);
  if (!r.isNew) return 'موجود مسبقًا: ' + TAB.QUEUE;
  r.sheet.getRange('N2:N3000').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['قيد الانتظار', 'معتمد', 'مرفوض', 'ملغي'], true).build());
  return 'تم إنشاء: ' + TAB.QUEUE;
}

/* ---------- سجل التدقيق ---------- */
function buildAudit_(ss) {
  const r = ensureSheet_(ss, TAB.AUDIT,
    ['الوقت', 'الإيميل', 'الإجراء', 'معرف الحركة', 'التفاصيل'], [160, 240, 140, 120, 500]);
  return r.isNew ? 'تم إنشاء: ' + TAB.AUDIT : 'موجود مسبقًا: ' + TAB.AUDIT;
}

/* ---------- أدوات مساعدة ---------- */
function ensureSheet_(ss, name, headers, widths) {
  let sh = ss.getSheetByName(name);
  if (sh) return { sheet: sh, isNew: false };
  sh = ss.insertSheet(name);
  sh.setRightToLeft(true);
  const h = sh.getRange(1, 1, 1, headers.length);
  h.setValues([headers]).setFontWeight('bold').setFontColor('#ffffff')
    .setBackground('#1F3A5F').setHorizontalAlignment('center');
  sh.setFrozenRows(1);
  (widths || []).forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  return { sheet: sh, isNew: true };
}

function lockToOwner_(sheet, description) {
  const me = Session.getEffectiveUser();
  const existing = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  const p = existing.length ? existing[0] : sheet.protect();
  p.setDescription(description);
  p.addEditor(me);
  const others = p.getEditors().filter(function (u) { return u.getEmail() !== me.getEmail(); });
  if (others.length) p.removeEditors(others);
  if (p.canDomainEdit()) p.setDomainEdit(false);
}
