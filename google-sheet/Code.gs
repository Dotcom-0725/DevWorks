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

/* ── Module Finances (v1) — additif, indépendant des Commandes ── */
const SHEET_FIN = '💸 Finances';
const FIN_TYPES = ['💰 Revenu', '💸 Dépense'];
const FIN_CAT_REVENU  = ['Projet client', 'Autre revenu'];
const FIN_CAT_DEPENSE = ['Hébergement & Domaine', 'Logiciels & Abonnements', 'Publicité (Ads)', 'Matériel', 'Formation', 'Transport', 'Autre dépense'];
const FIN_PAYMENT_MODES = ['Espèces', 'Virement bancaire', 'Carte bancaire', 'PayPal', 'Autre'];
const FIN_HEADERS = ['#', 'Date', 'Type', 'Catégorie', 'Description', 'Montant (DH)', 'Mode de paiement', 'Notes'];

/* ── Jeton d'accès au tableau de bord web (dashboard.html) — change-le si besoin ──
 * IMPORTANT (sécurité) : ce jeton est forcément visible en clair dans le code source
 * de dashboard.html (site 100% statique, sans serveur pour le cacher). Un jeton long
 * et aléatoire comme celui-ci protège contre le brute-force/scan automatique, mais PAS
 * contre quelqu'un qui consulte directement le code source de dashboard.html — dans ce
 * cas, la seule vraie protection est de ne jamais partager/lier publiquement cette page
 * (déjà fait : Disallow: /dashboard.html dans robots.txt). */
const DASH_TOKEN = 'PlIpOz_X0sFsY1c3Z11NWRt_OzG5yLOQ';

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

/* Formule "Reste" auto (Commandes) + numérotation auto (Finances) */
function onEdit(e) {
  const sh = e.range.getSheet();
  const r = e.range.getRow(), c = e.range.getColumn();

  if (sh.getName() === SHEET_LEADS) {
    if (r >= 2 && (c === 10 || c === 11)) setRemainFormula(sh, r);
    if (r >= 2 && (c === 9 || c === 11)) maybeTransferPaymentToFinance(sh, r);
    return;
  }

  if (sh.getName() === SHEET_FIN) {
    if (r >= 2 && c !== 1 && sh.getRange(r, 1).getValue() === '') {
      const filled = sh.getRange(r, 2, 1, 7).getValues()[0].some(v => v !== '');
      if (filled) sh.getRange(r, 1).setValue(r - 1);
    }
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ─────────── Réception du formulaire du site (public, inchangé) OU actions du tableau
 * de bord (protégées par DASH_TOKEN : ajout/modif/suppression client & finance) ─────────── */
function doPost(e) {
  const p = e.parameter;

  if (p.action) {
    if (p.token !== DASH_TOKEN) return jsonOut({ ok: false, error: 'unauthorized' });
    try {
      switch (p.action) {
        case 'addLead':       return addLead(p);
        case 'updateLead':    return updateLead(p);
        case 'deleteLead':    return deleteLead(p);
        case 'addFinance':    return addFinance(p);
        case 'updateFinance': return updateFinance(p);
        case 'deleteFinance': return deleteFinance(p);
        default: return jsonOut({ ok: false, error: 'unknown action' });
      }
    } catch (err) {
      return jsonOut({ ok: false, error: String(err) });
    }
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_LEADS);
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

    return jsonOut({ ok: true, row: r });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

/* ─────────── CRUD Commandes depuis le tableau de bord ─────────── */
function addLead(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_LEADS);
  const r = sh.getLastRow() + 1;

  sh.getRange(r, 1, 1, 14).setValues([[
    r - 1,
    p.date ? new Date(p.date) : new Date(),
    p.name || '',
    p.whatsapp || '',
    p.service || '',
    p.budget || 'Non défini',
    p.desc || '',
    p.source || '🌐 Formulaire du site',
    p.status || '🆕 Nouveau',
    p.price !== undefined && p.price !== '' ? Number(p.price) : '',
    p.paid !== undefined && p.paid !== '' ? Number(p.paid) : '',
    '',
    p.delivery ? new Date(p.delivery) : '',
    p.notes || ''
  ]]);
  setRemainFormula(sh, r);
  maybeTransferPaymentToFinance(sh, r);
  return jsonOut({ ok: true, row: r });
}

function updateLead(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_LEADS);
  const r = parseInt(p.row, 10);
  if (!r || r < 2) return jsonOut({ ok: false, error: 'invalid row' });

  const fieldCols = { name: 3, whatsapp: 4, service: 5, budget: 6, desc: 7, source: 8, status: 9, price: 10, paid: 11, notes: 14 };
  Object.keys(fieldCols).forEach(key => {
    if (p[key] !== undefined) {
      let v = p[key];
      if (key === 'price' || key === 'paid') v = v === '' ? '' : Number(v);
      sh.getRange(r, fieldCols[key]).setValue(v);
    }
  });
  if (p.date !== undefined) sh.getRange(r, 2).setValue(p.date ? new Date(p.date) : '');
  if (p.delivery !== undefined) sh.getRange(r, 13).setValue(p.delivery ? new Date(p.delivery) : '');

  setRemainFormula(sh, r);
  maybeTransferPaymentToFinance(sh, r);
  return jsonOut({ ok: true, row: r });
}

function deleteLead(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_LEADS);
  const r = parseInt(p.row, 10);
  if (!r || r < 2) return jsonOut({ ok: false, error: 'invalid row' });
  sh.deleteRow(r);
  return jsonOut({ ok: true });
}

/* ─────────── CRUD Finances depuis le tableau de bord ─────────── */
function addFinance(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fin = ss.getSheetByName(SHEET_FIN);
  const r = fin.getLastRow() + 1;
  fin.getRange(r, 1, 1, 8).setValues([[
    r - 1,
    p.date ? new Date(p.date) : new Date(),
    p.type || FIN_TYPES[0],
    p.category || '',
    p.description || '',
    p.amount !== undefined && p.amount !== '' ? Number(p.amount) : 0,
    p.paymentMode || '',
    p.notes || ''
  ]]);
  return jsonOut({ ok: true, row: r });
}

function updateFinance(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fin = ss.getSheetByName(SHEET_FIN);
  const r = parseInt(p.row, 10);
  if (!r || r < 2) return jsonOut({ ok: false, error: 'invalid row' });

  const fieldCols = { type: 3, category: 4, description: 5, amount: 6, paymentMode: 7, notes: 8 };
  Object.keys(fieldCols).forEach(key => {
    if (p[key] !== undefined) {
      let v = p[key];
      if (key === 'amount') v = v === '' ? 0 : Number(v);
      fin.getRange(r, fieldCols[key]).setValue(v);
    }
  });
  if (p.date !== undefined) fin.getRange(r, 2).setValue(p.date ? new Date(p.date) : '');

  return jsonOut({ ok: true, row: r });
}

function deleteFinance(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fin = ss.getSheetByName(SHEET_FIN);
  const r = parseInt(p.row, 10);
  if (!r || r < 2) return jsonOut({ ok: false, error: 'invalid row' });
  fin.deleteRow(r);
  return jsonOut({ ok: true });
}

/* Vérification rapide (par défaut) OU données complètes du dashboard (?action=dashboard&token=...) */
function doGet(e) {
  const p = (e && e.parameter) || {};

  if (p.action === 'dashboard') {
    if (p.token !== DASH_TOKEN) {
      return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'unauthorized' }))
                           .setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify(getDashboardData()))
                         .setMimeType(ContentService.MimeType.JSON);
  }

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

/* ─────────── Extraction JSON complète (Commandes + Finances) pour dashboard.html ───────────
 * requiredKey : n'inclut que les lignes où cette colonne n'est pas vide (évite les faux positifs
 * dus aux cases à cocher qui écrivent FAUX sur des lignes par ailleurs vides). */
function sheetToObjects(sheet, keys, requiredKey) {
  if (!sheet) return [];
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const values = sheet.getRange(2, 1, last - 1, keys.length).getValues();
  const reqIdx = requiredKey ? keys.indexOf(requiredKey) : -1;
  const out = [];
  values.forEach((row, i) => {
    const keep = reqIdx >= 0 ? (row[reqIdx] !== '' && row[reqIdx] !== null) : row.some(v => v !== '' && v !== null);
    if (!keep) return;
    const obj = { row: i + 2 };
    keys.forEach((k, j) => {
      let v = row[j];
      if (v instanceof Date) v = v.toISOString();
      obj[k] = v;
    });
    out.push(obj);
  });
  return out;
}

function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const leads = sheetToObjects(ss.getSheetByName(SHEET_LEADS), [
    'num', 'date', 'name', 'whatsapp', 'service', 'budget', 'desc',
    'source', 'status', 'price', 'paid', 'remain', 'delivery', 'notes', 'transferred'
  ], 'name');
  const finances = sheetToObjects(ss.getSheetByName(SHEET_FIN), [
    'num', 'date', 'type', 'category', 'description', 'amount', 'paymentMode', 'notes'
  ], 'description');
  return { ok: true, generatedAt: new Date().toISOString(), leads, finances };
}

/* ═══════════════════════════════════════════════════════════
 *  MODULE FINANCES — v1
 *  100% additif : ne touche jamais à 📋 Commandes (données réelles).
 *  Sûr à relancer autant de fois que nécessaire.
 * ═══════════════════════════════════════════════════════════ */

/* ─────────── Création / mise à jour de la feuille 💸 Finances ─────────── */
function setupFinance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let fin = ss.getSheetByName(SHEET_FIN) || ss.insertSheet(SHEET_FIN, 1);
  fin.clear();
  fin.setRightToLeft(false);
  fin.getRange(1, 1, 1, FIN_HEADERS.length).setValues([FIN_HEADERS])
     .setBackground('#0e1b30').setFontColor('#ffffff')
     .setFontWeight('bold').setFontSize(11)
     .setHorizontalAlignment('center').setVerticalAlignment('middle');
  fin.setFrozenRows(1);
  fin.setRowHeight(1, 42);

  const widths = [45, 110, 130, 220, 300, 120, 150, 220];
  widths.forEach((w, i) => fin.setColumnWidth(i + 1, w));

  const maxRows = 1000;
  const dateRule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build();
  fin.getRange(2, 2, maxRows).setDataValidation(dateRule); // clic sur la cellule → petit calendrier natif Google Sheets
  setDropdown(fin, 3, FIN_TYPES, maxRows);
  setDropdown(fin, 4, FIN_CAT_REVENU.concat(FIN_CAT_DEPENSE), maxRows);
  setDropdown(fin, 7, FIN_PAYMENT_MODES, maxRows);

  const rowRange = fin.getRange(2, 1, maxRows, FIN_HEADERS.length);
  fin.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$C2="' + FIN_TYPES[0] + '"')
      .setBackground('#d1e7dd').setRanges([rowRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$C2="' + FIN_TYPES[1] + '"')
      .setBackground('#f8d7da').setRanges([rowRange]).build()
  ]);

  fin.getRange(2, 2, maxRows).setNumberFormat('yyyy/mm/dd');
  fin.getRange(2, 6, maxRows).setNumberFormat('#,##0 "DH"');
  fin.getRange(2, 5, maxRows).setWrap(true);
  fin.getRange(2, 1, maxRows, FIN_HEADERS.length).setVerticalAlignment('middle');

  // Ajoute les listes Finances à ⚙️ Paramètres (colonnes D-F) SANS toucher A-C (Statuts/Sources/Services)
  const conf = ss.getSheetByName(SHEET_CONF);
  if (conf) {
    conf.getRange(1, 4, 1, 3).setValues([['Types Finance', 'Catégories Finance', 'Modes de paiement']])
        .setFontWeight('bold').setBackground('#0e1b30').setFontColor('#ffffff');
    conf.getRange(2, 4, FIN_TYPES.length, 1).setValues(FIN_TYPES.map(s => [s]));
    const allCats = FIN_CAT_REVENU.concat(FIN_CAT_DEPENSE);
    conf.getRange(2, 5, allCats.length, 1).setValues(allCats.map(s => [s]));
    conf.getRange(2, 6, FIN_PAYMENT_MODES.length, 1).setValues(FIN_PAYMENT_MODES.map(s => [s]));
    conf.setColumnWidths(4, 3, 190);
  }

  try { SpreadsheetApp.getUi().alert('✅ ورقة الماليات (💸 Finances) جاهزة! سجّل فيها كل دخول ومصروف حقيقي بتاريخه.'); } catch (e) {}
}

/* ─────────── Reconstruction complète du tableau de bord (formules + graphiques) ───────────
 * Sûre à relancer : 📊 Tableau de bord est 100% calculée, aucune saisie manuelle n'y est perdue. */
function buildDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const L = "'" + SHEET_LEADS + "'";
  const F = "'" + SHEET_FIN + "'";

  // Feuille technique cachée qui alimente les graphiques
  let helper = ss.getSheetByName('_ChartData') || ss.insertSheet('_ChartData');
  helper.clear();

  helper.getRange(1, 1, 1, 2).setValues([['Catégorie', 'Montant']]);
  helper.getRange(2, 1, FIN_CAT_DEPENSE.length, 1).setValues(FIN_CAT_DEPENSE.map(c => [c]));
  helper.getRange(2, 2, FIN_CAT_DEPENSE.length, 1).setFormulas(
    FIN_CAT_DEPENSE.map(c => ['=SUMIFS(' + F + '!F:F,' + F + '!C:C,"' + FIN_TYPES[1] + '",' + F + '!D:D,"' + c + '")'])
  );

  helper.getRange(1, 4, 1, 3).setValues([['Mois', 'Revenus', 'Dépenses']]);
  for (let i = 0; i < 6; i++) {
    const monthsAgo = 5 - i;
    const r = 2 + i;
    helper.getRange(r, 4).setFormula('=TEXT(EOMONTH(TODAY(),-' + monthsAgo + ')+1,"mmm yyyy")');
    helper.getRange(r, 5).setFormula('=SUMIFS(' + F + '!F:F,' + F + '!C:C,"' + FIN_TYPES[0] + '",' + F + '!B:B,">="&EOMONTH(TODAY(),-' + (monthsAgo + 1) + ')+1,' + F + '!B:B,"<="&EOMONTH(TODAY(),-' + monthsAgo + '))');
    helper.getRange(r, 6).setFormula('=SUMIFS(' + F + '!F:F,' + F + '!C:C,"' + FIN_TYPES[1] + '",' + F + '!B:B,">="&EOMONTH(TODAY(),-' + (monthsAgo + 1) + ')+1,' + F + '!B:B,"<="&EOMONTH(TODAY(),-' + monthsAgo + '))');
  }
  helper.hideSheet();

  // Suppression puis recréation (plus sûr que clear() pour effacer fusions/graphiques résiduels)
  const oldDash = ss.getSheetByName(SHEET_DASH);
  if (oldDash) ss.deleteSheet(oldDash);
  let dash = ss.insertSheet(SHEET_DASH, 1);
  dash.setRightToLeft(false);
  try { dash.setHiddenGridlines(true); } catch (e) {}
  try { dash.setTabColor('#38bdf8'); } catch (e) {}

  dash.getRange(2, 2, 1, 8).merge().setValue('📊 لوحة التحكم الشاملة — Rachid DevWorks Pro')
      .setFontSize(18).setFontWeight('bold').setFontColor('#0e1b30')
      .setVerticalAlignment('middle');
  dash.setRowHeight(2, 44);

  // ── Cartes KPI (4 blocs colorés) ──
  const cardDefs = [
    { col: 2, label: '💰 إجمالي المداخيل',
      formula: '=SUMIF(' + F + '!C:C,"' + FIN_TYPES[0] + '",' + F + '!F:F)',
      bg: '#d1e7dd', fg: '#0f5132' },
    { col: 4, label: '💸 إجمالي المصاريف',
      formula: '=SUMIF(' + F + '!C:C,"' + FIN_TYPES[1] + '",' + F + '!F:F)',
      bg: '#f8d7da', fg: '#842029' },
    { col: 6, label: '📈 الربح الصافي',
      formula: '=SUMIF(' + F + '!C:C,"' + FIN_TYPES[0] + '",' + F + '!F:F)-SUMIF(' + F + '!C:C,"' + FIN_TYPES[1] + '",' + F + '!F:F)',
      bg: '#cfe2ff', fg: '#084298' },
    { col: 8, label: '📊 هامش الربح',
      formula: '=IFERROR((SUMIF(' + F + '!C:C,"' + FIN_TYPES[0] + '",' + F + '!F:F)-SUMIF(' + F + '!C:C,"' + FIN_TYPES[1] + '",' + F + '!F:F))/SUMIF(' + F + '!C:C,"' + FIN_TYPES[0] + '",' + F + '!F:F),0)',
      bg: '#e2d9f3', fg: '#432874', pct: true }
  ];

  cardDefs.forEach(card => {
    dash.getRange(4, card.col, 1, 2).merge().setValue(card.label)
        .setBackground(card.bg).setFontColor(card.fg)
        .setFontWeight('bold').setFontSize(11)
        .setHorizontalAlignment('center').setVerticalAlignment('middle');
    dash.getRange(5, card.col, 2, 2).merge().setFormula(card.formula)
        .setBackground(card.bg).setFontColor(card.fg)
        .setFontWeight('bold').setFontSize(20)
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setNumberFormat(card.pct ? '0.0%' : '#,##0 "DH"');
  });
  for (let c = 2; c <= 9; c++) dash.setColumnWidth(c, 115);
  dash.setRowHeight(4, 26);
  dash.setRowHeight(5, 30);
  dash.setRowHeight(6, 30);

  // ── Pipeline clients (repris de l'ancien tableau de bord) ──
  dash.getRange(9, 2, 1, 2).merge().setValue('📌 حالة المشاريع')
      .setFontSize(13).setFontWeight('bold').setFontColor('#0e1b30').setBackground('#f1f5fb');

  const pipelineLabels = ['إجمالي الطلبات', 'طلبات هاد الشهر'].concat(STATUSES);
  const pipelineFormulas = [
    '=COUNTA(' + L + '!C2:C)',
    '=COUNTIFS(' + L + '!B2:B,">="&EOMONTH(TODAY(),-1)+1)'
  ].concat(STATUSES.map(st => '=COUNTIF(' + L + '!I2:I,"' + st + '")'));

  dash.getRange(10, 2, pipelineLabels.length, 1).setValues(pipelineLabels.map(l => [l]))
      .setFontWeight('bold').setBackground('#f8fafc');
  dash.getRange(10, 3, pipelineFormulas.length, 1).setFormulas(pipelineFormulas.map(f => [f]))
      .setFontWeight('bold').setHorizontalAlignment('center').setBackground('#ffffff');
  dash.getRange(10, 2, pipelineLabels.length, 2)
      .setBorder(true, true, true, true, true, true, '#d0d7e2', SpreadsheetApp.BorderStyle.SOLID);

  // ── Graphiques financiers ──
  const chartsRow = 10 + pipelineLabels.length + 2;
  dash.getRange(chartsRow, 2, 1, 2).merge().setValue('📈 التحليل المالي')
      .setFontSize(13).setFontWeight('bold').setFontColor('#0e1b30').setBackground('#f1f5fb');

  const pieRange = helper.getRange(1, 1, FIN_CAT_DEPENSE.length + 1, 2);
  const pieChart = dash.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(pieRange)
      .setPosition(chartsRow + 2, 2, 0, 0)
      .setOption('title', 'توزيع المصاريف حسب الفئة')
      .setOption('width', 480).setOption('height', 320)
      .setOption('colors', ['#38bdf8', '#f7941d', '#a78bfa', '#34d399', '#f87171', '#fbbf24', '#60a5fa'])
      .build();
  dash.insertChart(pieChart);

  const trendRange = helper.getRange(1, 4, 7, 3);
  const trendChart = dash.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(trendRange)
      .setPosition(chartsRow + 2, 7, 0, 0)
      .setOption('title', 'المداخيل مقابل المصاريف شهرياً')
      .setOption('width', 480).setOption('height', 320)
      .setOption('series', { 0: { color: '#22c55e' }, 1: { color: '#ef4444' } })
      .build();
  dash.insertChart(trendChart);

  try { SpreadsheetApp.getUi().alert('✅ لوحة التحكم محدّثة بنجاح!'); } catch (e) {}
}

/* ─────────── Colonne de suivi (Commandes!O) : empêche les doublons de transfert ───────────
 * IMPORTANT : la validation "case à cocher" écrit FAUX dans la cellule dès qu'on l'applique,
 * même sur une cellule vide. On ne l'applique donc JAMAIS à un gros bloc de lignes vides
 * (ça décale getLastRow() très loin et casse l'ajout des nouvelles commandes) — seulement
 * ligne par ligne, via applyTransferCheckbox(), quand la ligne a réellement des données. */
function ensureTransferColumn() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const leads = ss.getSheetByName(SHEET_LEADS);
  if (!leads) return;
  if (leads.getRange(1, 15).getValue() === '') {
    leads.getRange(1, 15).setValue('Transféré Finance')
         .setBackground('#0e1b30').setFontColor('#ffffff')
         .setFontWeight('bold').setFontSize(11)
         .setHorizontalAlignment('center').setVerticalAlignment('middle');
    leads.setColumnWidth(15, 130);
  }
}

function applyTransferCheckbox(sh, r) {
  const cell = sh.getRange(r, 15);
  if (!cell.getDataValidation()) {
    cell.setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  }
}

/* ─────────── Transfert auto Commandes(Livré + Payé) → 💸 Finances (une seule fois par ligne) ─────────── */
function maybeTransferPaymentToFinance(sh, r) {
  const statut = sh.getRange(r, 9).getValue();
  if (statut !== STATUSES[4]) return; // '✅ Livré' uniquement

  const paye = sh.getRange(r, 11).getValue();
  if (!paye || paye <= 0) return;

  ensureTransferColumn();
  applyTransferCheckbox(sh, r);
  if (sh.getRange(r, 15).getValue() === true) return; // déjà transféré

  const ss = sh.getParent();
  const fin = ss.getSheetByName(SHEET_FIN);
  if (!fin) return;

  const num = sh.getRange(r, 1).getValue();
  const nom = sh.getRange(r, 3).getValue();
  const service = sh.getRange(r, 5).getValue();
  const livraison = sh.getRange(r, 13).getValue();
  const dateVal = (livraison instanceof Date) ? livraison : new Date();

  const fr = fin.getLastRow() + 1;
  fin.getRange(fr, 1, 1, 8).setValues([[
    fr - 1, dateVal, FIN_TYPES[0], FIN_CAT_REVENU[0],
    'Commande #' + num + ' — ' + nom + ' (' + service + ')',
    paye, '', 'Transfert automatique depuis 📋 Commandes'
  ]]);

  sh.getRange(r, 15).setValue(true);
}

/* ─────────── Rattrapage : transfère les commandes Livré déjà existantes (à lancer une seule fois) ─────────── */
function backfillFinanceFromCommandes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const leads = ss.getSheetByName(SHEET_LEADS);
  if (!leads) return;
  ensureTransferColumn();
  const last = leads.getLastRow();
  for (let r = 2; r <= last; r++) maybeTransferPaymentToFinance(leads, r);
  try { SpreadsheetApp.getUi().alert('✅ تم نقل كل الدفعات القديمة (المسلَّمة) لورقة 💸 Finances.'); } catch (e) {}
}

/* ─────────── Nettoyage ponctuel : répare les lignes fantômes créées par l'ancienne version
 * de ensureTransferColumn (qui appliquait la case à cocher sur 1000 lignes d'un coup, ce qui
 * écrivait FAUX partout et décalait les nouvelles commandes très loin). Sûr à relancer :
 * si tout est déjà propre, ne fait rien. ─────────── */
function fixPhantomRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const leads = ss.getSheetByName(SHEET_LEADS);
  if (!leads) return;
  const last = leads.getLastRow();
  if (last < 2) return;

  const numCols = 15; // A..O
  const allRows = leads.getRange(2, 1, last - 1, numCols).getValues();

  // Une ligne est "réelle" si au moins une colonne A..N (donc hors O, qui peut
  // légitimement contenir FAUX tout seul sur une ligne fantôme) n'est pas vide.
  // NB : on ne se fie surtout pas à une seule colonne (ex. Nom) pour détecter la fin
  // des données, car certaines lignes réelles ont le Nom vide mais un téléphone rempli.
  const realRows = allRows.filter(row => row.some((v, i) => i !== 14 && v !== '' && v !== null));

  leads.getRange(2, 1, last - 1, numCols).clearContent().clearDataValidations();
  if (realRows.length > 0) {
    leads.getRange(2, 1, realRows.length, numCols).setValues(realRows);
    for (let i = 0; i < realRows.length; i++) {
      if (realRows[i][14] !== '' && realRows[i][14] !== null) applyTransferCheckbox(leads, 2 + i);
    }
  }

  try { SpreadsheetApp.getUi().alert('✅ Nettoyage terminé. Lignes réelles : ' + realRows.length); } catch (e) {}
}

/* ─────────── Installation en un clic de tout le module finance (ne touche pas aux Commandes) ─────────── */
function setupFinanceSystem() {
  ensureTransferColumn();
  setupFinance();
  buildDashboard();
}

/* ─────────── Menu rapide à l'ouverture du classeur ─────────── */
function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('🛠️ أدوات CRM')
      .addItem('💰 تثبيت/تحديث وحدة الماليات', 'setupFinanceSystem')
      .addItem('📊 تحديث لوحة التحكم فقط', 'buildDashboard')
      .addItem('🔁 نقل الدفعات القديمة (مرة واحدة)', 'backfillFinanceFromCommandes')
      .addToUi();
  } catch (e) {}
}
