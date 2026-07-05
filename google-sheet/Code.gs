/*******************************************************
 *  نظام تتبع العملاء — Rachid DevWorks Pro CRM  (v2)
 *  ----------------------------------------------------
 *  هذه نسخة مطابقة للكود المنشور في Apps Script
 *  المرتبط بالشيت — للمرجع والنسخ الاحتياطي فقط.
 *******************************************************/

const SHEET_LEADS = '📋 الطلبات';
const SHEET_DASH  = '📊 لوحة التحكم';
const SHEET_CONF  = '⚙️ الإعدادات';

const STATUSES = ['🆕 جديد', '📞 تم التواصل', '🤝 اتفاق', '⚙️ قيد التنفيذ', '✅ تم التسليم', '❌ ملغي'];
const SOURCES  = ['🌐 فورم الموقع', '💬 واتساب مباشر', '📱 وسائل التواصل', '👥 توصية'];
const SERVICES = ['موقع إلكتروني', 'صفحة هبوط', 'تصميم صور إعلانية', 'فيديو إعلاني', 'حملة إعلانية', 'خدمة أخرى'];

const HEADERS = ['#', 'التاريخ', 'الاسم', 'الواتساب', 'الخدمة', 'الميزانية',
                 'وصف المشروع', 'المصدر', 'الحالة', 'السعر المتفق (درهم)',
                 'المدفوع (درهم)', 'المتبقي (درهم)', 'تاريخ التسليم', 'ملاحظات'];

/* صيغة "المتبقي" لصف واحد */
function setRemainFormula(sh, r) {
  sh.getRange(r, 12).setFormula('=IF(J' + r + '="","",J' + r + '-IF(K' + r + '="",0,K' + r + '))');
}

/* ─────────── إصلاح: نقل الطلبات العالقة تحت الصف 1001 إلى الأعلى ─────────── */
function fixSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_LEADS);

  // 1) مسح صيغ "المتبقي" المعبأة مسبقاً (كانت سبب كتابة الطلبات تحت الصف 1001)
  sh.getRange(2, 12, 1000).clearContent();

  // 2) نقل أي صفوف عالقة بعد الصف 1001
  const last = sh.getLastRow();
  if (last > 1001) {
    const n = last - 1001;
    const data = sh.getRange(1002, 1, n, 14).getValues();
    let target = 2;
    const names = sh.getRange(2, 3, 1000).getValues();
    while (target - 2 < 1000 && names[target - 2][0] !== '') target++;
    data.forEach((row, i) => {
      row[0] = target - 1 + i;
      sh.getRange(target + i, 1, 1, 14).setValues([row]);
      setRemainFormula(sh, target + i);
    });
    sh.deleteRows(1002, n);
  }
}

/* ─────────── الإعداد الأولي (بناء كامل — يمسح البيانات!) ─────────── */
function setupCRM() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetLocale('en_US');  // باش المعادلات تخدم بالفواصل العادية

  let leads = ss.getSheetByName(SHEET_LEADS) || ss.insertSheet(SHEET_LEADS, 0);
  leads.clear();
  leads.setRightToLeft(true);
  leads.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
       .setBackground('#0e1b30').setFontColor('#ffffff')
       .setFontWeight('bold').setFontSize(11)
       .setHorizontalAlignment('center').setVerticalAlignment('middle');
  leads.setFrozenRows(1);
  leads.setRowHeight(1, 42);

  const widths = [45, 130, 150, 130, 150, 140, 320, 130, 130, 130, 110, 110, 110, 220];
  widths.forEach((w, i) => leads.setColumnWidth(i + 1, w));

  const maxRows = 1000;
  setDropdown(leads, 5,  SERVICES, maxRows);
  setDropdown(leads, 8,  SOURCES,  maxRows);
  setDropdown(leads, 9,  STATUSES, maxRows);
  // ملاحظة: بلا تعبئة مسبقة لصيغ العمود L — كتضاف أوتوماتيكياً لكل صف جديد

  const colors = {
    '🆕 جديد':        '#fff3cd',
    '📞 تم التواصل':  '#cfe2ff',
    '🤝 اتفاق':       '#e2d9f3',
    '⚙️ قيد التنفيذ': '#ffe5d0',
    '✅ تم التسليم':  '#d1e7dd',
    '❌ ملغي':        '#f8d7da'
  };
  const rules = [];
  const rowRange = leads.getRange(2, 1, maxRows, HEADERS.length);
  Object.keys(colors).forEach(st => {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$I2="' + st + '"')
      .setBackground(colors[st])
      .setRanges([rowRange]).build());
  });
  leads.setConditionalFormatRules(rules);

  leads.getRange(2, 2, maxRows).setNumberFormat('yyyy/mm/dd hh:mm');
  leads.getRange(2, 10, maxRows, 3).setNumberFormat('#,##0 "درهم"');
  leads.getRange(2, 13, maxRows).setNumberFormat('yyyy/mm/dd');
  leads.getRange(2, 7, maxRows).setWrap(true);
  leads.getRange(2, 1, maxRows, HEADERS.length).setVerticalAlignment('middle');

  let conf = ss.getSheetByName(SHEET_CONF) || ss.insertSheet(SHEET_CONF);
  conf.clear();
  conf.setRightToLeft(true);
  conf.getRange('A1:C1').setValues([['الحالات', 'المصادر', 'الخدمات']])
      .setFontWeight('bold').setBackground('#0e1b30').setFontColor('#ffffff');
  conf.getRange(2, 1, STATUSES.length).setValues(STATUSES.map(s => [s]));
  conf.getRange(2, 2, SOURCES.length ).setValues(SOURCES.map(s => [s]));
  conf.getRange(2, 3, SERVICES.length).setValues(SERVICES.map(s => [s]));
  conf.setColumnWidths(1, 3, 170);

  let dash = ss.getSheetByName(SHEET_DASH) || ss.insertSheet(SHEET_DASH, 1);
  dash.clear();
  dash.setRightToLeft(true);
  dash.getRange('B2').setValue('📊 لوحة تحكم المشاريع — Rachid DevWorks Pro')
      .setFontSize(16).setFontWeight('bold').setFontColor('#0e1b30');

  const L = "'" + SHEET_LEADS + "'";
  const labels = [
    'إجمالي الطلبات', 'طلبات هذا الشهر', '🆕 جديد', '📞 تم التواصل',
    '🤝 اتفاق', '⚙️ قيد التنفيذ', '✅ تم التسليم', '❌ ملغي',
    '💰 إجمالي المداخيل المتفق عليها', '✅ إجمالي المدفوع', '⏳ إجمالي المتبقي'
  ];
  const formulas = [
    '=COUNTA(' + L + '!C2:C)',
    '=COUNTIFS(' + L + '!B2:B,">="&EOMONTH(TODAY(),-1)+1)',
    '=COUNTIF(' + L + '!I2:I,"🆕 جديد")',
    '=COUNTIF(' + L + '!I2:I,"📞 تم التواصل")',
    '=COUNTIF(' + L + '!I2:I,"🤝 اتفاق")',
    '=COUNTIF(' + L + '!I2:I,"⚙️ قيد التنفيذ")',
    '=COUNTIF(' + L + '!I2:I,"✅ تم التسليم")',
    '=COUNTIF(' + L + '!I2:I,"❌ ملغي")',
    '=SUM(' + L + '!J2:J)',
    '=SUM(' + L + '!K2:K)',
    '=SUM(' + L + '!L2:L)'
  ];
  dash.getRange(4, 2, labels.length, 1).setValues(labels.map(l => [l]));
  dash.getRange(4, 3, formulas.length, 1).setFormulas(formulas.map(f => [f]));
  dash.getRange(4, 2, labels.length, 1).setFontWeight('bold').setFontSize(11).setBackground('#f1f5fb');
  dash.getRange(4, 3, formulas.length, 1).setFontSize(12).setFontWeight('bold')
      .setHorizontalAlignment('center').setBackground('#ffffff');
  dash.getRange(12, 3, 3, 1).setNumberFormat('#,##0 "درهم"');
  dash.setColumnWidth(2, 260);
  dash.setColumnWidth(3, 160);
  dash.getRange(4, 2, labels.length, 2).setBorder(true, true, true, true, true, true, '#d0d7e2', SpreadsheetApp.BorderStyle.SOLID);

  const def = ss.getSheetByName('Sheet1') || ss.getSheetByName('ورقة1') || ss.getSheetByName('Feuille 1');
  if (def && def.getLastRow() === 0) ss.deleteSheet(def);

  try { SpreadsheetApp.getUi().alert('✅ تم بناء نظام تتبع العملاء بنجاح!'); } catch (e) {}
}

function setDropdown(sheet, col, list, rows) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(list, true).setAllowInvalid(false).build();
  sheet.getRange(2, col, rows).setDataValidation(rule);
}

/* عند تعديل السعر أو المدفوع يدوياً ← إضافة صيغة المتبقي أوتوماتيكياً */
function onEdit(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== SHEET_LEADS) return;
  const r = e.range.getRow(), c = e.range.getColumn();
  if (r >= 2 && (c === 10 || c === 11)) setRemainFormula(sh, r);
}

/* ─────────── استقبال طلبات فورم الموقع ─────────── */
function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_LEADS);
    const p = e.parameter;
    const r = sh.getLastRow() + 1;

    sh.getRange(r, 1, 1, 14).setValues([[
      r - 1,
      new Date(),
      p.name    || '',
      p.phone   || '',
      p.service || '',
      p.budget  || 'غير محددة',
      p.desc    || '',
      '🌐 فورم الموقع',
      '🆕 جديد',
      '', '', '', '', ''
    ]]);
    setRemainFormula(sh, r);

    return ContentService.createTextOutput(JSON.stringify({ ok: true, row: r }))
                         .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}

/* فحص سريع: يرجع عدد الطلبات وآخر طلب */
function doGet() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LEADS);
  const last = sh.getLastRow();
  const info = { ok: true, leads: Math.max(0, last - 1) };
  if (last >= 2) {
    const v = sh.getRange(last, 1, 1, 9).getValues()[0];
    info.lastLead = { num: v[0], name: v[2], service: v[4], status: v[8] };
  }
  return ContentService.createTextOutput(JSON.stringify(info))
                       .setMimeType(ContentService.MimeType.JSON);
}
