/**
 * المرحلة 2: إضافة المطبخ المركزي كمكان ثامن + تحديث قواعد الحركات الثمانية.
 *
 * مبني على المعادلات الفعلية التي استخرجها الفحص. الترتيب:
 *   1) addCentralKitchen   (يُنشئ نسخة احتياطية تلقائيًا قبل أي تغيير)
 *   2) updateRulesPhase2
 * كل دالة تتوقف بخطأ واضح إذا وجدت الملف بحالة غير متوقعة، دون تغيير جزئي قبل النسخة الاحتياطية.
 */

const KITCHEN_NAME = 'المطبخ المركزي';

function addCentralKitchen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const places = ss.getSheetByName('الأماكن');
  const bal = ss.getSheetByName('الأرصدة');
  const val = ss.getSheetByName('القيمة');
  const dash = ss.getSheetByName('لوحة_المتابعة');
  const card = ss.getSheetByName('كارتة_الصنف');
  const inv = ss.getSheetByName('الجرد');
  const report = [];

  // ---- فحوصات أمان قبل أي تغيير ----
  if (String(places.getRange('A9').getValue()).trim() !== '') {
    throw new Error('الخلية A9 في الأماكن ليست فارغة (' + places.getRange('A9').getValue() + '). تم الإيقاف دون أي تغيير.');
  }
  if (String(bal.getRange('K1').getValue()).trim() !== 'الإجمالي') {
    throw new Error('العمود K في الأرصدة ليس "الإجمالي". يبدو أن الهيكل تغيّر. تم الإيقاف.');
  }
  if (String(val.getRange('K1').getValue()).trim() !== 'إجمالي القيمة') {
    throw new Error('العمود K في القيمة ليس "إجمالي القيمة". يبدو أن الهيكل تغيّر. تم الإيقاف.');
  }

  // ---- نسخة احتياطية ----
  const stamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm');
  const backup = DriveApp.getFileById(ss.getId()).makeCopy('نسخة احتياطية قبل إضافة المطبخ - ' + stamp);
  report.push('النسخة الاحتياطية: ' + backup.getUrl());

  // ---- الأماكن ----
  places.getRange('A8:C8').copyTo(places.getRange('A9:C9'), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  places.getRange('A9:B9').setValues([[KITCHEN_NAME, 'مطبخ']]);
  const note = String(places.getRange('A10').getValue());
  if (note.indexOf('السبعة') >= 0) places.getRange('A10').setValue(note.replace('السبعة', 'الثمانية'));
  report.push('أُضيف المطبخ المركزي في الأماكن (A9)');

  // ---- الأرصدة: عمود جديد K للمطبخ، والإجمالي يصبح L وحد الطلب M والحالة N ----
  bal.insertColumnAfter(10);
  bal.setColumnWidth(11, bal.getColumnWidth(10));
  bal.getRange('J1').copyTo(bal.getRange('K1'), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  bal.getRange('K1').setFormula("='الأماكن'!A9");
  bal.getRange('J3:J102').copyTo(bal.getRange('K3:K102'), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  rewriteFormulas_(bal.getRange('L3:L102'), widen_);
  rewriteFormulas_(bal.getRange('N3:N102'), widen_);
  SpreadsheetApp.flush();
  report.push('الأرصدة: أُضيف عمود المطبخ وتم توسيع الإجمالي والحالة');

  // ---- القيمة: نفس الفكرة ----
  val.insertColumnAfter(10);
  val.setColumnWidth(11, val.getColumnWidth(10));
  val.getRange('J1').copyTo(val.getRange('K1'), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  val.getRange('K1').setFormula("='الأماكن'!A9");
  val.getRange('J2:J102').copyTo(val.getRange('K2:K102'), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  rewriteFormulas_(val.getRange('L3:L102'), widen_);
  SpreadsheetApp.flush();
  report.push('القيمة: أُضيف عمود المطبخ وتم توسيع الإجمالي');

  // ---- كارتة_الصنف والجرد: توسيع نطاقات البحث من J إلى K ----
  rewriteFormulas_(card.getRange('I4'), widenLookup_);
  rewriteFormulas_(inv.getRange('C7:C106'), widenLookup_);
  report.push('الكارتة والجرد: توسيع نطاقات الأرصدة');

  // ---- قوائم اختيار الأماكن ----
  extendPlaceValidation_(ss);
  report.push('قوائم اختيار الأماكن تشمل المطبخ');

  // ---- لوحة المتابعة ----
  addDashboardKitchen_(dash, report);

  // ---- فحص ذاتي وإصلاح المراجع ----
  verifyKitchen_(ss, report);

  Logger.log(report.join('\n'));
  ss.toast('اكتملت إضافة المطبخ، راجع سجل التنفيذ', 'المرحلة 2', 8);
}

/* ---------- أدوات ---------- */
function widen_(f) {
  return f.replace(/D(\d+):J(\d+)/g, 'D$1:K$2');
}

function widenLookup_(f) {
  return f
    .replace(/\$D\$3:\$J\$102/g, function () { return '$D$3:$K$102'; })
    .replace(/\$D\$1:\$J\$1/g, function () { return '$D$1:$K$1'; });
}

function rewriteFormulas_(range, fn) {
  const f = range.getFormulas();
  const allFormulas = f.every(function (row) { return row.every(function (x) { return x !== ''; }); });
  if (allFormulas) {
    range.setFormulas(f.map(function (row) { return row.map(fn); }));
    return;
  }
  const sh = range.getSheet();
  const r0 = range.getRow();
  const c0 = range.getColumn();
  for (let i = 0; i < f.length; i++) {
    for (let j = 0; j < f[i].length; j++) {
      if (!f[i][j]) continue;
      const nf = fn(f[i][j]);
      if (nf !== f[i][j]) sh.getRange(r0 + i, c0 + j).setFormula(nf);
    }
  }
}

function extendPlaceValidation_(ss) {
  const placesRange = ss.getSheetByName('الأماكن').getRange('A2:A9');
  [['الحركات', 'C2:D3001'], ['كارتة_الصنف', 'C4'], ['الجرد', 'B3']].forEach(function (p) {
    const rng = ss.getSheetByName(p[0]).getRange(p[1]);
    const rules = rng.getDataValidations().map(function (row) {
      return row.map(function (r) {
        if (!r) return r;
        if (r.getCriteriaType() !== SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) return r;
        return r.copy().requireValueInRange(placesRange, true).build();
      });
    });
    rng.setDataValidations(rules);
  });
}

function addDashboardKitchen_(dash, report) {
  if (dash.getRange('F3:F4').isBlank()) {
    dash.getRange('E3:E4').copyTo(dash.getRange('F3:F4'), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    dash.getRange('F3').setValue('قيمة مخزون المطبخ');
    dash.getRange('F4').setFormula("='القيمة'!K2");
    report.push('اللوحة: أُضيفت بطاقة قيمة مخزون المطبخ (F3:F4)');
  } else {
    report.push('اللوحة: F3:F4 غير فارغة، لم تُضَف بطاقة المطبخ (راجعها يدويًا)');
  }
  if (dash.getRange('A15:E15').isBlank()) {
    dash.getRange('A14:E14').copyTo(dash.getRange('A15:E15'), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    dash.getRange('A15').setFormula("='الأماكن'!A9");
    dash.getRange('B15').setFormula("='القيمة'!K2");
    dash.getRange('C15').setFormula("=SUMIFS('الحركات'!$I$2:$I$3001,'الحركات'!$B$2:$B$3001,\"شراء\",'الحركات'!$D$2:$D$3001,$A15)");
    dash.getRange('D15').setFormula("=SUMIFS('الحركات'!$I$2:$I$3001,'الحركات'!$B$2:$B$3001,\"صرف للتشغيل\",'الحركات'!$C$2:$C$3001,$A15)");
    dash.getRange('E15').setFormula("=SUMIFS('الحركات'!$I$2:$I$3001,'الحركات'!$B$2:$B$3001,\"هالك\",'الحركات'!$C$2:$C$3001,$A15)");
    report.push('اللوحة: أُضيف صف المطبخ في جدول الأماكن (الصف 15)');
  } else {
    report.push('اللوحة: الصف 15 غير فارغ، لم يُضَف صف المطبخ في جدول الأماكن (راجعه يدويًا)');
  }
}

function verifyKitchen_(ss, report) {
  const bal = ss.getSheetByName('الأرصدة');
  const val = ss.getSheetByName('القيمة');
  const items = ss.getSheetByName('الأصناف');
  const dash = ss.getSheetByName('لوحة_المتابعة');

  // إصلاح مراجع قد لا يعدّلها جوجل تلقائيًا
  const colOf = function (f) {
    const m = /الأرصدة'!\$?([A-Z]+)\$?\d+/.exec(f);
    return m ? m[1] : '';
  };
  if (colOf(items.getRange('I3').getFormula()) !== 'L') {
    const f = [];
    for (let r = 3; r <= 102; r++) f.push(["='الأرصدة'!L" + r]);
    items.getRange('I3:I102').setFormulas(f);
    report.push('الأصناف: صُحّح مرجع إجمالي الكمية إلى عمود L');
  }
  if (colOf(items.getRange('K3').getFormula()) !== 'N') {
    const f = [];
    for (let r = 3; r <= 102; r++) f.push(["='الأرصدة'!N" + r]);
    items.getRange('K3:K102').setFormulas(f);
    report.push('الأصناف: صُحّح مرجع الحالة إلى عمود N');
  }
  if (/القيمة'!K2/.test(dash.getRange('A4').getFormula())) {
    dash.getRange('A4').setFormula("='القيمة'!L2");
    report.push('اللوحة: صُحّح مرجع إجمالي المخزون');
  }
  if (/M3:M102/.test(dash.getRange('D4').getFormula())) {
    dash.getRange('D4').setFormula("=COUNTIF('الأرصدة'!N3:N102,\"اطلب\")");
  }
  if (/M3:M102/.test(dash.getRange('E4').getFormula())) {
    dash.getRange('E4').setFormula("=COUNTIF('الأرصدة'!N3:N102,\"⚠*\")");
  }

  const checks = [
    ['الأرصدة!K1', bal.getRange('K1').getFormula(), /الأماكن'!A9/],
    ['الأرصدة!K3', bal.getRange('K3').getFormula(), /K\$1/],
    ['الأرصدة!L3 (الإجمالي)', bal.getRange('L3').getFormula(), /D3:K3/],
    ['الأرصدة!N3 (الحالة)', bal.getRange('N3').getFormula(), /D3:K3/],
    ['القيمة!K3', val.getRange('K3').getFormula(), /الأرصدة'!K3/],
    ['القيمة!L3 (الإجمالي)', val.getRange('L3').getFormula(), /D3:K3/],
    ['كارتة_الصنف!I4', ss.getSheetByName('كارتة_الصنف').getRange('I4').getFormula(), /\$D\$3:\$K\$102/],
    ['الجرد!C7', ss.getSheetByName('الجرد').getRange('C7').getFormula(), /\$K\$102/]
  ];
  checks.forEach(function (c) {
    report.push((c[2].test(c[1]) ? 'سليم: ' : 'يحتاج مراجعة: ') + c[0]);
  });
}

/* ======================================================================
 * قواعد الحركات الثمانية + جدول التحويلات المسموحة + عمود "صف الحركات" في المعلق
 * ====================================================================== */
function updateRulesPhase2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(TAB.RULES);

  sh.getRange('A2:E40').clearContent();
  sh.getRange(2, 1, 8, 5).setValues([
    ['رصيد افتتاحي', 'المالك', 'تلقائي', 'لا ينطبق', 'يدخله المالك فقط ويُعتمد فورًا'],
    ['شراء', 'المكان المستلم', 'المكان المستلم', 'نعم', 'قرار الإدارة: المدخل والمعتمد من نفس المكان. غيّر الاعتماد الذاتي إلى لا لفصل المهام'],
    ['تحويل', 'المكان المرسل', 'المكان المستلم', 'لا', 'انظر جدول التحويلات المسموحة أدناه'],
    ['صرف للتشغيل', 'المكان المعني', 'تلقائي', 'لا ينطبق', 'استهلاك يومي، ويُراجع عبر الجرد'],
    ['هالك', 'المكان المعني', 'المالك', 'لا', 'يعتمده المالك'],
    ['مرتجع للمورد', 'المكان المعني', 'المالك', 'لا', 'يعتمده المالك'],
    ['تسوية زيادة', 'المالك', 'تلقائي', 'لا ينطبق', 'تسوية من الجرد، للمالك فقط'],
    ['تسوية عجز', 'المالك', 'تلقائي', 'لا ينطبق', 'تسوية من الجرد، للمالك فقط']
  ]);

  sh.getRange('A11').setValue('التحويلات المسموحة لغير المالك (حسب نوع المكان في تبويب الأماكن). غير المذكور للمالك فقط.')
    .setFontWeight('bold');
  sh.getRange(12, 1, 1, 3).setValues([['من (نوع المكان)', 'إلى (نوع المكان)', 'مسموح']])
    .setFontWeight('bold').setBackground('#1F3A5F').setFontColor('#ffffff');
  sh.getRange(13, 1, 7, 3).setValues([
    ['مستودع', 'فرع', 'نعم'],
    ['مطبخ', 'فرع', 'نعم'],
    ['فرع', 'فرع', 'نعم'],
    ['مستودع', 'مطبخ', 'لا'],
    ['مطبخ', 'مستودع', 'لا'],
    ['فرع', 'مستودع', 'لا'],
    ['فرع', 'مطبخ', 'لا']
  ]);
  sh.getRange('C13:C19').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['نعم', 'لا'], true).build());

  sh.getRange('A21').setValue(
    'المصطلحات: المكان المرسل = (من)، المكان المستلم = (إلى)، المكان المعني = المكان الوحيد في الحركة. ' +
    'تلقائي = يُعتمد فور الإدخال. اعتماد ذاتي = هل يجوز لمن أدخل أن يعتمد بنفسه.');

  // عمود إضافي في المعلق لربط الحركة بصفها في الدفتر
  const q = ss.getSheetByName(TAB.QUEUE);
  if (String(q.getRange('S1').getValue()).trim() === '') {
    q.getRange('R1').copyTo(q.getRange('S1'), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    q.getRange('S1').setValue('صف الحركات');
    q.setColumnWidth(19, 100);
  }
  ss.toast('تم تحديث القواعد', 'المرحلة 2', 6);
}
