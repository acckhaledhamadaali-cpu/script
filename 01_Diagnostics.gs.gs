/**
 * المرحلة 0: فحص تشخيصي (قراءة فقط، لا يغيّر شيئًا في الملف)
 *
 * الخطوات:
 * 1) في الشيت: الإضافات > Apps Script
 * 2) الصق هذا الملف، ثم شغّل الدالة dumpDiagnostics ووافق على الأذونات
 * 3) افتح "سجل التنفيذ" وانسخ الناتج كاملًا
 */

function dumpDiagnostics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const out = [];

  out.push('== التبويبات ==');
  ss.getSheets().forEach(function (s) {
    out.push(s.getName() + ' | صفوف ' + s.getMaxRows() + ' | أعمدة ' + s.getMaxColumns() +
      ' | آخر صف به بيانات ' + s.getLastRow());
  });

  out.push('== الأماكن (قيم) ==');
  const pl = ss.getSheetByName('الأماكن');
  if (pl) {
    pl.getRange('A1:G10').getValues().forEach(function (r, i) {
      out.push((i + 1) + ': ' + r.join(' ; '));
    });
  }

  out.push('== المعادلات ==');
  [
    ['الحركات', 'A1:N3'],
    ['الأرصدة', 'A1:M4'],
    ['القيمة', 'A1:L4'],
    ['الأصناف', 'A1:L4'],
    ['كارتة_الصنف', 'A1:K12'],
    ['لوحة_المتابعة', 'A1:G8'],
    ['الجرد', 'A1:G12']
  ].forEach(function (p) { dumpFormulasIn_(ss, p[0], p[1], out); });

  const mv = ss.getSheetByName('الحركات');
  if (mv) {
    out.push('== قوائم الاختيار في الحركات (الصف 2) ==');
    mv.getRange('A2:N2').getDataValidations()[0].forEach(function (rule, i) {
      out.push('عمود ' + (i + 1) + ': ' + describeRule_(rule));
    });
    const filled = mv.getRange(2, 2, mv.getMaxRows() - 1, 1).getValues()
      .filter(function (r) { return r[0] !== ''; }).length;
    out.push('عدد صفوف الحركات المعبأة حاليًا: ' + filled);
  }

  out.push('== نطاقات مسماة ==');
  ss.getNamedRanges().forEach(function (n) {
    out.push(n.getName() + ' -> ' + n.getRange().getSheet().getName() + '!' + n.getRange().getA1Notation());
  });

  out.push('== حماية موجودة ==');
  ss.getSheets().forEach(function (s) {
    const sp = s.getProtections(SpreadsheetApp.ProtectionType.SHEET).length;
    const rp = s.getProtections(SpreadsheetApp.ProtectionType.RANGE).length;
    if (sp || rp) out.push(s.getName() + ' | حماية تبويب: ' + sp + ' | حماية نطاقات: ' + rp);
  });

  out.push('== المالك ==');
  out.push('مشغّل السكربت: ' + Session.getEffectiveUser().getEmail());
  out.push('المنطقة الزمنية للملف: ' + ss.getSpreadsheetTimeZone());

  Logger.log(out.join('\n'));
}

function dumpFormulasIn_(ss, sheetName, a1, out) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) { out.push('تبويب غير موجود: ' + sheetName); return; }
  const rng = sh.getRange(a1);
  const f = rng.getFormulas();
  const r0 = rng.getRow();
  const c0 = rng.getColumn();
  for (let i = 0; i < f.length; i++) {
    for (let j = 0; j < f[i].length; j++) {
      if (f[i][j]) {
        out.push(sheetName + '!' + sh.getRange(r0 + i, c0 + j).getA1Notation() + '  ' + f[i][j].slice(0, 300));
      }
    }
  }
}

function describeRule_(rule) {
  if (!rule) return 'بدون';
  const vals = rule.getCriteriaValues().map(function (v) {
    return (v && v.getA1Notation) ? v.getSheet().getName() + '!' + v.getA1Notation() : String(v);
  });
  return rule.getCriteriaType() + ' ' + JSON.stringify(vals).slice(0, 150);
}

/**
 * اختبار الهوية (مؤقت، سيُحذف لاحقًا):
 * نشر > نشر جديد > تطبيق ويب
 *   التنفيذ باسم: أنا
 *   من لديه حق الوصول: أي شخص لديه حساب Google
 * افتح الرابط من حساب Gmail مختلف عن حساب المالك وشاهد النتيجة.
 */
function doGet() {
  const active = Session.getActiveUser().getEmail();
  const effective = Session.getEffectiveUser().getEmail();
  return HtmlService.createHtmlOutput(
    '<div dir="rtl" style="font:18px sans-serif;padding:24px">' +
    '<p>المستخدم الذي فتح الرابط (Active): <b>' + (active || '(فارغ)') + '</b></p>' +
    '<p>المنفّذ (Effective): <b>' + effective + '</b></p></div>'
  );
}
