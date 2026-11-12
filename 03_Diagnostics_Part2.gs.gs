/**
 * المرحلة 0 - الجزء الثاني (قراءة فقط، لا يغيّر شيئًا).
 * يكمل ما اقتُطع من الجزء الأول، ويطبع المخرجات على دفعات صغيرة.
 * الاستخدام: الصق الملف في نفس المشروع (بجانب الملفين الآخرين)،
 * شغّل dumpDiagnosticsPart2، ثم انسخ كل ما في "سجل التنفيذ".
 * يعتمد على الدالة describeRule_ الموجودة في 01_Diagnostics.gs
 */

function dumpDiagnosticsPart2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const buf = [];

  buf.push('== الحركات!M2 كاملة ==');
  addFormulas_(ss, 'الحركات', 'M2', 2000, buf);

  buf.push('== الأصناف (الصفوف 1-4) ==');
  addFormulas_(ss, 'الأصناف', 'A1:L4', 1500, buf);

  buf.push('== كارتة_الصنف (الصفوف 1-14) ==');
  addFormulas_(ss, 'كارتة_الصنف', 'A1:K14', 500, buf);

  buf.push('== لوحة_المتابعة (الصفوف 1-12) ==');
  addFormulas_(ss, 'لوحة_المتابعة', 'A1:G12', 500, buf);

  buf.push('== الجرد (الصفوف 1-8) ==');
  addFormulas_(ss, 'الجرد', 'A1:G8', 500, buf);

  buf.push('== قوائم الاختيار ==');
  addRules_(ss, 'الحركات', 'A2:N2', buf);
  addRules_(ss, 'كارتة_الصنف', 'C3:C4', buf);
  addRules_(ss, 'الجرد', 'B3', buf);

  buf.push('== حماية موجودة ==');
  ss.getSheets().forEach(function (s) {
    const sp = s.getProtections(SpreadsheetApp.ProtectionType.SHEET).length;
    const rp = s.getProtections(SpreadsheetApp.ProtectionType.RANGE).length;
    if (sp || rp) buf.push(s.getName() + ' | تبويب: ' + sp + ' | نطاقات: ' + rp);
  });

  flush_(buf);
}

function addFormulas_(ss, sheetName, a1, limit, buf) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) { buf.push('تبويب غير موجود: ' + sheetName); return; }
  const rng = sh.getRange(a1);
  const f = rng.getFormulas();
  const r0 = rng.getRow();
  const c0 = rng.getColumn();
  for (let i = 0; i < f.length; i++) {
    for (let j = 0; j < f[i].length; j++) {
      if (f[i][j]) {
        buf.push(sheetName + '!' + sh.getRange(r0 + i, c0 + j).getA1Notation() + '  ' + f[i][j].slice(0, limit));
      }
    }
  }
}

function addRules_(ss, sheetName, a1, buf) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) { buf.push('تبويب غير موجود: ' + sheetName); return; }
  const rng = sh.getRange(a1);
  const v = rng.getDataValidations();
  const r0 = rng.getRow();
  const c0 = rng.getColumn();
  for (let i = 0; i < v.length; i++) {
    for (let j = 0; j < v[i].length; j++) {
      buf.push(sheetName + '!' + sh.getRange(r0 + i, c0 + j).getA1Notation() + '  ' + describeRule_(v[i][j]));
    }
  }
}

// يطبع على دفعات أقل من 6000 حرف لتفادي اقتطاع السجل
function flush_(buf) {
  let chunk = '';
  buf.forEach(function (line) {
    if (chunk.length + line.length > 6000) { Logger.log(chunk); chunk = ''; }
    chunk += line + '\n';
  });
  if (chunk) Logger.log(chunk);
}
