/*******************************************************
 *  Système de suivi clients — Rachid DevWorks Pro CRM (v3 FR)
 *  ----------------------------------------------------
 *  Copie de référence du code déployé dans Apps Script.
 *******************************************************/

const SHEET_LEADS = '📋 Commandes';
const SHEET_DASH  = '📊 Tableau de bord';
const SHEET_CONF  = '⚙️ Paramètres';

const STATUSES = ['🆕 Nouveau', '📞 Contacté', '🤝 Accord', '⚙️ En cours', '✅ Livré', '❌ Annulé'];
const SOURCES  = ['🌐 Formulaire du site', '💬 WhatsApp direct', '📱 Réseaux sociaux', '👥 Recommandation'];
const SERVICES = ['Site web', 'Landing page', 'Visuels publicitaires', 'Vidéo publicitaire', 'Campagne publicitaire', 'Autre service'];

// الفورم كيصيفط القيم بالعربية إلا كان الزائر مختار العربية — كنحولوها للفرنسية
const SERVICE_MAP = {
  'موقع إلكتروني':      'Site web',
  'صفحة هبوط':          'Landing page',
  'تصميم صور إعلانية':  'Visuels publicitaires',
  'فيديو إعلاني':       'Vidéo publicitaire',
  'حملة إعلانية':       'Campagne publicitaire',
  'خدمة أخرى':          'Autre service'
};
const BUDGET_MAP = {
  'أقل من 1000 درهم':   'Moins de 1000 DH',
  '1000 – 3000 درهم':   '1000 – 3000 DH',
  '3000 – 7000 درهم':   '3000 – 7000 DH',
  'أكثر من 7000 درهم':  'Plus de 7000 DH',
  'غير محددة':          'Non défini'
};

const HEADERS = ['#', 'Date', 'Nom', 'WhatsApp', 'Service', 'Budget',
                 'Description du projet', 'Source', 'Statut', 'Prix convenu (DH)',
                 'Payé (DH)', 'Reste (DH)', 'Livraison', 'Notes'];

/* Formule "Reste" pour une seule ligne */
function setRemainFormula(sh, r) {
  sh.getRange(r, 12).setFormula('=IF(J' + r + '="","",J' + r + '-IF(K' + r + '="",0,K' + r + '))');
}

/* ─────────── Construction complète (efface les données !) ─────────── */
function setupCRM() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetLocale('en_US');  // séparateurs de formules stables

  // Renommer les anciennes feuilles arabes si présentes (sinon en créer)
  const oldLeads = ss.getSheetByName('📋 الطلبات');
  if (oldLeads) oldLeads.setName(SHEET_LEADS);
  const oldDash = ss.getSheetByName('📊 لوحة التحكم');
  if (oldDash) oldDash.setName(SHEET_DASH);
  const oldConf = ss.getSheetByName('⚙️ الإعدادات');
  if (oldConf) oldConf.setName(SHEET_CONF);

  let leads = ss.getSheetByName(SHEET_LEADS) || ss.insertSheet(SHEET_LEADS, 0);
  leads.clear();
  leads.setRightToLeft(false);
  leads.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
       .setBackground('#0e1b30').setFontColor('#ffffff')
       .setFontWeight('bold').setFontSize(11)
       .setHorizontalAlignment('center').setVerticalAlignment('middle');
  leads.setFrozenRows(1);
  leads.setRowHeight(1, 42);

  const widths = [45, 130, 150, 130, 160, 140, 320, 160, 130, 130, 110, 110, 110, 220];
  widths.forEach((w, i) => leads.setColumnWidth(i + 1, w));

  const maxRows = 1000;
  setDropdown(leads, 5,  SERVICES, maxRows);
  setDropdown(leads, 8,  SOURCES,  maxRows);
  setDropdown(leads, 9,  STATUSES, maxRows);
  // NB : pas de préremplissage de formules en colonne L (bug appendRow)

  const colors = {
    '🆕 Nouveau':  '#fff3cd',
    '📞 Contacté': '#cfe2ff',
    '🤝 Accord':   '#e2d9f3',
    '⚙️ En cours': '#ffe5d0',
    '✅ Livré':    '#d1e7dd',
    '❌ Annulé':   '#f8d7da'
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
  leads.getRange(2, 10, maxRows, 3).setNumberFormat('#,##0 "DH"');
  leads.getRange(2, 13, maxRows).setNumberFormat('yyyy/mm/dd');
  leads.getRange(2, 7, maxRows).setWrap(true);
  leads.getRange(2, 1, maxRows, HEADERS.length).setVerticalAlignment('middle');

  let conf = ss.getSheetByName(SHEET_CONF) || ss.insertSheet(SHEET_CONF);
  conf.clear();
  conf.setRightToLeft(false);
  conf.getRange('A1:C1').setValues([['Statuts', 'Sources', 'Services']])
      .setFontWeight('bold').setBackground('#0e1b30').setFontColor('#ffffff');
  conf.getRange(2, 1, STATUSES.length).setValues(STATUSES.map(s => [s]));
  conf.getRange(2, 2, SOURCES.length ).setValues(SOURCES.map(s => [s]));
  conf.getRange(2, 3, SERVICES.length).setValues(SERVICES.map(s => [s]));
  conf.setColumnWidths(1, 3, 190);

  let dash = ss.getSheetByName(SHEET_DASH) || ss.insertSheet(SHEET_DASH, 1);
  dash.clear();
  dash.setRightToLeft(false);
  dash.getRange('B2').setValue('📊 Tableau de bord — Rachid DevWorks Pro')
      .setFontSize(16).setFontWeight('bold').setFontColor('#0e1b30');

  const L = "'" + SHEET_LEADS + "'";
  const labels = [
    'Total des commandes', 'Commandes ce mois-ci', '🆕 Nouveau', '📞 Contacté',
    '🤝 Accord', '⚙️ En cours', '✅ Livré', '❌ Annulé',
    '💰 Total des revenus convenus', '✅ Total payé', '⏳ Total restant'
  ];
  const formulas = [
    '=COUNTA(' + L + '!C2:C)',
    '=COUNTIFS(' + L + '!B2:B,">="&EOMONTH(TODAY(),-1)+1)',
    '=COUNTIF(' + L + '!I2:I,"🆕 Nouveau")',
    '=COUNTIF(' + L + '!I2:I,"📞 Contacté")',
    '=COUNTIF(' + L + '!I2:I,"🤝 Accord")',
    '=COUNTIF(' + L + '!I2:I,"⚙️ En cours")',
    '=COUNTIF(' + L + '!I2:I,"✅ Livré")',
    '=COUNTIF(' + L + '!I2:I,"❌ Annulé")',
    '=SUM(' + L + '!J2:J)',
    '=SUM(' + L + '!K2:K)',
    '=SUM(' + L + '!L2:L)'
  ];
  dash.getRange(4, 2, labels.length, 1).setValues(labels.map(l => [l]));
  dash.getRange(4, 3, formulas.length, 1).setFormulas(formulas.map(f => [f]));
  dash.getRange(4, 2, labels.length, 1).setFontWeight('bold').setFontSize(11).setBackground('#f1f5fb');
  dash.getRange(4, 3, formulas.length, 1).setFontSize(12).setFontWeight('bold')
      .setHorizontalAlignment('center').setBackground('#ffffff');
  dash.getRange(12, 3, 3, 1).setNumberFormat('#,##0 "DH"');
  dash.setColumnWidth(2, 280);
  dash.setColumnWidth(3, 160);
  dash.getRange(4, 2, labels.length, 2).setBorder(true, true, true, true, true, true, '#d0d7e2', SpreadsheetApp.BorderStyle.SOLID);

  const def = ss.getSheetByName('Sheet1') || ss.getSheetByName('ورقة1') || ss.getSheetByName('Feuille 1');
  if (def && def.getLastRow() === 0) ss.deleteSheet(def);

  try { SpreadsheetApp.getUi().alert('✅ Système CRM (FR) construit avec succès !'); } catch (e) {}
}

function setDropdown(sheet, col, list, rows) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(list, true).setAllowInvalid(false).build();
  sheet.getRange(2, col, rows).setDataValidation(rule);
}

/* Formule "Reste" auto lors d'une édition manuelle de Prix/Payé */
function onEdit(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== SHEET_LEADS) return;
  const r = e.range.getRow(), c = e.range.getColumn();
  if (r >= 2 && (c === 10 || c === 11)) setRemainFormula(sh, r);
}

/* ─────────── Réception du formulaire du site ─────────── */
function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_LEADS);
    const p = e.parameter;
    const r = sh.getLastRow() + 1;

    const service = SERVICE_MAP[p.service] || p.service || '';
    const budget  = BUDGET_MAP[p.budget]  || p.budget  || 'Non défini';

    sh.getRange(r, 1, 1, 14).setValues([[
      r - 1,
      new Date(),
      p.name  || '',
      p.phone || '',
      service,
      budget,
      p.desc  || '',
      '🌐 Formulaire du site',
      '🆕 Nouveau',
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

/* Vérification rapide : nombre de commandes + dernière commande */
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
