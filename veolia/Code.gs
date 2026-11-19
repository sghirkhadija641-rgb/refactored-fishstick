/**
 * Réclamations Ré-Use — version compte Veolia.
 * Deux formulaires servis par ce script, réservés aux comptes Veolia connectés :
 *  - Saisie (lien de l'application Web) : enregistre une réclamation dans la feuille « Réclamations »
 *    et donne à l'agent un code de suivi (ex. RU-K7P4Q) ;
 *  - Traitement (même lien suivi de ?f=traitement) : après l'intervention, l'agent terrain saisit le
 *    code de suivi et le traitement, enregistré dans la feuille « Traitements ».
 * Le code de suivi relie les deux feuilles. Les agents ne voient jamais les réclamations enregistrées.
 *
 * Installation : Extensions > Apps Script, coller ce code, enregistrer,
 * puis Déployer > Nouveau déploiement > Application Web
 * (Exécuter en tant que : Moi ; Qui a accès : Anyone within Veolia).
 */

const VERSION = 'Veolia 12';
const CODE_RE = /^RU-[2-9A-Z]{5}$/;
const SHEET_NAME = 'Réclamations';
const FOLDER_NAME = 'Photos réclamations Ré-Use';
const DATE_FMT = 'dd/MM/yyyy HH:mm';

const TRAITEMENTS_SHEET = 'Traitements';
const HEADERS = ['N° de réclamation', 'Code de suivi', 'Date de réception', "Date d'envoi", 'Agent', 'Matricule',
  'Nom du réclamant', 'Téléphone', 'E-mail',
  'Adresse', 'Quartier', 'Secteur', 'Commune', 'Latitude', 'Longitude', 'Précision (m)', 'Carte',
  'Nature', 'Localisation fuite', 'Importance fuite', 'Détail nature', 'Depuis', 'Période', 'CIL', 'Date/heure souhaitée',
  'Observations', 'Aperçu photo', 'Lien photo', 'Etat de la réclamation', 'ID envoi'];
const TRAITEMENT_HEADERS = ['N° de réclamation', 'Code de suivi', 'Date réception', 'Etat de la réclamation', 'Métier', 'Nature réclamation',
  'Type client', 'Nom du client', 'CIL', 'Téléphone portable 1', 'Téléphone fixe', 'Nombre de rappels',
  'Nombre de photos', 'Photos', 'Agent', 'Matricule', 'Compte Veolia', "Date d'envoi"];
const TEXT_COLS = ['N° de réclamation', 'Code de suivi', 'Matricule', 'CIL', 'Téléphone', 'E-mail', 'Téléphone portable 1', 'Téléphone fixe'];
const DATE_COLS = ['Date de réception', "Date d'envoi", 'Date réception'];
const HIDDEN_COLS = ['ID envoi'];

const NATURES = ['Réfection', "Fuite d'eau", "Qualité d'eau", 'Problème de Pression', 'Demande Coupure',
  'Demande Rebranchement', 'DFO', "Manque d'eau", 'Appel Client pour Information'];
const ETATS = ['Non encore Transmise', 'Transmise mais non encore résolue', 'Résolue Provisoirement', 'Résolue Définitivement'];
const ANCIENS_ETATS = { 'Ouverte': ETATS[0], 'En cours': ETATS[1], 'Clôturée': ETATS[3] };
const TYPES_CLIENT = ['Particulier', 'Autorité', 'Communes', 'Elus', 'Résidences Royales', 'Industriels', 'V.I.P',
  'GC-Entreprises', 'Planifié', 'GC-Administrations', 'Numéro Vert'];

function doGet(e) {
  if (e && e.parameter && e.parameter.reparer) {
    let note = '';
    try { note = ' Lignes réparées : ' + reparerAnciennesLignes() + '.'; } catch (err) { note = ' Réparation impossible : ' + err.message; }
    return ContentService.createTextOutput('Le script des réclamations Ré-Use fonctionne. (version ' + VERSION + ')' + note);
  }
  const f = e && e.parameter ? e.parameter.f : '';
  const url = ScriptApp.getService().getUrl();
  const page = f === 'traitement' ? TRAITEMENT_HTML : FORM_HTML;
  return HtmlService.createHtmlOutput(page.split('__URL__').join(url))
    .setTitle(f === 'traitement' ? 'Traitement Ré-Use' : 'Réclamations Ré-Use')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ===================== Saisie ===================== */

// Appelée par le formulaire de saisie (google.script.run).
function envoyer(d) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    const compte = Session.getActiveUser().getEmail();
    d.agent = [String(d.agent || '').trim(), compte ? '(' + compte + ')' : ''].filter(String).join(' ');
    const reception = parseDateTime(d.reception);
    if (!reception || reception > new Date(Date.now() + 864e5)) return { ok: false, error: 'date de réception manquante ou invalide' };
    for (const k of ['matricule', 'nom', 'adresse', 'quartier', 'secteur', 'ville', 'nature']) {
      if (!d[k] || !String(d[k]).trim()) return { ok: false, error: 'champ manquant (' + k + ')' };
    }
    if (NATURES.indexOf(d.nature) < 0) return { ok: false, error: 'nature inconnue' };

    const sh = getSheet();
    const c = cols(sh);

    // Même envoi reçu deux fois (réseau coupé puis rétabli) : on ne le duplique pas.
    if (d.clientId && sh.getLastRow() > 1) {
      const found = sh.getRange(2, c['ID envoi'], sh.getLastRow() - 1, 1)
        .createTextFinder(String(d.clientId)).matchEntireCell(true).findNext();
      if (found) return { ok: true, code: sh.getRange(found.getRow(), c['Code de suivi']).getValue(), doublon: true };
    }

    const code = CODE_RE.test(String(d.code || '')) ? String(d.code) : nouveauCode();
    const hasPos = d.lat !== '' && d.lat != null && d.lng !== '' && d.lng != null && isFinite(d.lat) && isFinite(d.lng);

    const values = {
      'Code de suivi': code, 'Date de réception': reception, "Date d'envoi": new Date(),
      'Agent': safe(d.agent), 'Matricule': safe(d.matricule),
      'Nom du réclamant': safe(d.nom), 'CIL': safe(d.cil), 'Téléphone': safe(d.tel), 'E-mail': safe(d.email),
      'Adresse': safe(d.adresse), 'Quartier': safe(d.quartier), 'Secteur': safe(d.secteur), 'Commune': safe(d.ville),
      'Latitude': hasPos ? Number(d.lat) : '', 'Longitude': hasPos ? Number(d.lng) : '',
      'Précision (m)': hasPos && isFinite(d.precision) ? Number(d.precision) : '',
      'Nature': d.nature, 'Localisation fuite': safe(d.localisation), 'Importance fuite': safe(d.importance),
      'Détail nature': safe(d.detail), 'Depuis': safe(d.depuis), 'Période': safe(d.periode),
      'Date/heure souhaitée': safe(d.souhaite), 'Observations': safe(d.obs),
      'Etat de la réclamation': ETATS[0], 'ID envoi': safe(d.clientId)
    };
    const line = new Array(sh.getLastColumn()).fill('');
    Object.keys(values).forEach(h => { if (c[h]) line[c[h] - 1] = values[h]; });
    sh.appendRow(line);
    const row = sh.getLastRow();

    // Lien carte : vrai lien cliquable, conservé dans le fichier Excel téléchargé.
    if (hasPos) {
      sh.getRange(row, c['Carte']).setRichTextValue(link('Ouvrir la carte',
        'https://www.google.com/maps?q=' + Number(d.lat) + ',' + Number(d.lng)));
    }
    // Photo en grand format : rangée dans Google Drive, lien dans le tableau.
    const full = jpeg(d.photo);
    if (full && full.length < 8 * 1024 * 1024) {
      const file = getFolder().createFile(Utilities.newBlob(full, 'image/jpeg', code + '.jpg'));
      sh.getRange(row, c['Lien photo']).setRichTextValue(link('Voir la photo', file.getUrl()));
    }
    // Aperçu : petite image posée dans la ligne, visible aussi dans Excel.
    const small = jpeg(d.vignette);
    if (small && small.length < 1.5 * 1024 * 1024) {
      try {
        placeThumb(sh, row, c['Aperçu photo'], Utilities.newBlob(small, 'image/jpeg', code + '-apercu.jpg'));
      } catch (e) {
        sh.getRange(row, c['Aperçu photo']).setValue('aperçu indisponible');
      }
    }
    return { ok: true, code: code };
  } catch (err) {
    return { ok: false, error: 'erreur du tableau : ' + err.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/* ===================== Traitement ===================== */

// Appelée par le formulaire de traitement (google.script.run). Ne renvoie aucune donnée du tableau.
function enregistrerTraitement(t) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    t = t || {};
    for (const k of ['code', 'agent', 'matricule', 'etat']) {
      if (!t[k] || !String(t[k]).trim()) return { ok: false, error: 'champ manquant (' + k + ')' };
    }
    if (ETATS.indexOf(t.etat) < 0) return { ok: false, error: 'état de la réclamation invalide' };
    const code = String(t.code).trim().toUpperCase();
    if (!CODE_RE.test(code)) return { ok: false, error: 'code de suivi invalide (exemple : RU-K7P4Q)' };
    const sh = getSheet(), c = cols(sh);
    const row = trouverLigne(sh, c, code, 'Code de suivi');
    if (!row) return { ok: false, error: 'code de suivi inconnu. Vérifiez le code donné lors de la saisie de la réclamation.' };
    const nature = String(sh.getRange(row, c['Nature']).getValue() || '');

    // État : mis à jour aussi dans la feuille des réclamations.
    if (c['Etat de la réclamation']) sh.getRange(row, c['Etat de la réclamation']).setValue(t.etat);

    // Une ligne par traitement dans la feuille « Traitements », reliée par le code de suivi.
    const ts = getSheetFor(TRAITEMENTS_SHEET, TRAITEMENT_HEADERS), tc = cols(ts);
    const photos = (Array.isArray(t.photos) ? t.photos : []).slice(0, 5).map(jpeg).filter(b => b && b.length < 8 * 1024 * 1024);
    const values = {
      'Code de suivi': code, 'Date réception': parseDateTime(t.dateReception) || '', 'Etat de la réclamation': t.etat,
      'Métier': 'Ré-Use', 'Nature réclamation': nature,
      'Type client': TYPES_CLIENT.indexOf(t.typeClient) >= 0 ? t.typeClient : '',
      'Nom du client': safe(t.nomClient), 'CIL': safe(t.cil), 'Téléphone portable 1': safe(t.telPortable), 'Téléphone fixe': safe(t.telFixe),
      'Nombre de rappels': Math.max(0, parseInt(t.rappels, 10) || 0), 'Nombre de photos': photos.length,
      'Agent': safe(t.agent), 'Matricule': safe(t.matricule), 'Compte Veolia': Session.getActiveUser().getEmail(),
      "Date d'envoi": new Date()
    };
    const line = new Array(ts.getLastColumn()).fill('');
    Object.keys(values).forEach(h => { if (tc[h]) line[tc[h] - 1] = values[h]; });
    ts.appendRow(line);
    const trow = ts.getLastRow();
    lierNumeroReclamation(ts, trow, trow);

    // Photos : rangées dans Google Drive, un lien par photo dans la même case.
    if (photos.length && tc['Photos']) {
      const urls = photos.map((b, i) => getFolder().createFile(Utilities.newBlob(b, 'image/jpeg', code + '-traitement-' + (i + 1) + '.jpg')).getUrl());
      const labels = urls.map((u, i) => 'Photo ' + (i + 1));
      const rt = SpreadsheetApp.newRichTextValue().setText(labels.join('  '));
      let pos = 0;
      labels.forEach((l, i) => { rt.setLinkUrl(pos, pos + l.length, urls[i]); pos += l.length + 2; });
      ts.getRange(trow, tc['Photos']).setRichTextValue(rt.build());
    }
    return { ok: true };
  } catch (err) {
    return { ok: false, error: 'erreur du tableau : ' + err.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// Formulaire de traitement : renvoie seulement la nature de la réclamation liée au code.
function natureDuCode(code) {
  try {
    code = String(code || '').trim().toUpperCase();
    if (!CODE_RE.test(code)) return { ok: false, error: 'code de suivi invalide (exemple : RU-K7P4Q)' };
    const sh = getSheet(), c = cols(sh), row = trouverLigne(sh, c, code, 'Code de suivi');
    if (!row) return { ok: false, error: 'Code de suivi inconnu : vérifiez le code donné lors de la saisie.' };
    return { ok: true, nature: String(sh.getRange(row, c['Nature']).getValue() || '') };
  } catch (err) {
    return { ok: false, error: 'vérification impossible : ' + err.message };
  }
}

// Code de suivi aléatoire, unique dans le tableau.
function nouveauCode() {
  const a = '23456789ABCDEFGHJKMNPQRSTUVWXYZ';
  const sh = getSheet(), c = cols(sh);
  for (;;) {
    let s = 'RU-';
    for (let i = 0; i < 5; i++) s += a[Math.floor(Math.random() * a.length)];
    if (!trouverLigne(sh, c, s, 'Code de suivi')) return s;
  }
}

function trouverLigne(sh, c, valeur, entete) {
  if (!valeur || !c[entete] || sh.getLastRow() < 2) return 0;
  const found = sh.getRange(2, c[entete], sh.getLastRow() - 1, 1)
    .createTextFinder(String(valeur)).matchEntireCell(true).findNext();
  return found ? found.getRow() : 0;
}

/* ===================== Tableau ===================== */

// Numéro de colonne de chaque en-tête, lu dans le tableau lui-même.
function cols(sh) {
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const m = {};
  head.forEach((h, i) => { if (h !== '' && !m[String(h)]) m[String(h)] = i + 1; });
  return m;
}

function getSheet() {
  return getSheetFor(SHEET_NAME, HEADERS);
}

function getSheetFor(name, headers) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    styleHeader(sh.getRange(1, 1, 1, headers.length));
    sh.setFrozenRows(1);
    headers.forEach((h, i) => formatColumn(sh, i + 1, h));
  } else {
    mettreAJourColonnes(sh, headers);
  }
  return sh;
}

function formatColumn(sh, col, h) {
  const body = sh.getRange(2, col, Math.max(sh.getMaxRows() - 1, 1), 1);
  if (TEXT_COLS.indexOf(h) >= 0) body.setNumberFormat('@');
  if (DATE_COLS.indexOf(h) >= 0) body.setNumberFormat(DATE_FMT);
  if (h === 'Etat de la réclamation') body.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(ETATS, true).build());
  if (h === 'Type client') body.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(TYPES_CLIENT, true).setAllowInvalid(true).build());
  if (h === 'Aperçu photo') sh.setColumnWidth(col, 160);
  if (HIDDEN_COLS.indexOf(h) >= 0) sh.hideColumns(col);
}

// Tableau créé avec une version précédente : renomme et ajoute les colonnes manquantes.
function mettreAJourColonnes(sh, headers) {
  let head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const rename = (from, to) => {
    const i = head.indexOf(from);
    if (i >= 0 && head.indexOf(to) < 0) { sh.getRange(1, i + 1).setValue(to); head[i] = to; return i + 1; }
    return 0;
  };
  // Le N° Saisie n'est plus utilisé : sa colonne est retirée.
  ['N° Saisie', 'N° réclamation'].forEach(h => {
    const i = head.indexOf(h);
    if (i >= 0 && headers === HEADERS) { sh.deleteColumn(i + 1); head.splice(i, 1); }
  });
  const etatCol = rename('Statut', 'Etat de la réclamation');
  if (etatCol) {
    if (sh.getLastRow() > 1) {
      const r = sh.getRange(2, etatCol, sh.getLastRow() - 1, 1);
      r.clearDataValidations();
      r.setValues(r.getValues().map(v => [etatDe(v[0])]));
    }
    formatColumn(sh, etatCol, 'Etat de la réclamation');
  }
  const oldPhoto = head.indexOf('Photo');
  if (oldPhoto >= 0 && head.indexOf('Aperçu photo') < 0) {
    sh.insertColumnBefore(oldPhoto + 1);
    sh.getRange(1, oldPhoto + 1, 1, 2).setValues([['Aperçu photo', 'Lien photo']]);
    styleHeader(sh.getRange(1, oldPhoto + 1, 1, 2));
    sh.setColumnWidth(oldPhoto + 1, 160);
    head.splice(oldPhoto, 1, 'Aperçu photo', 'Lien photo');
  }
  headers.forEach((h, i) => {
    if (head.indexOf(h) >= 0) return;
    // Place la colonne juste après la précédente de la liste (ou en fin de tableau).
    const prev = i > 0 ? head.indexOf(headers[i - 1]) : -1;
    const at = i === 0 ? 1 : prev >= 0 ? prev + 2 : head.length + 1;
    if (at <= head.length) sh.insertColumnBefore(at); else sh.insertColumnAfter(head.length);
    sh.getRange(1, at, sh.getMaxRows(), 1).clearDataValidations().setBackground(null).setNumberFormat('General');
    sh.showColumns(at);
    sh.getRange(1, at).setValue(h);
    styleHeader(sh.getRange(1, at));
    formatColumn(sh, at, h);
    head.splice(at - 1, 0, h);
    if (headers === TRAITEMENT_HEADERS && h === 'N° de réclamation') lierNumeroReclamation(sh, 2, sh.getLastRow());
  });
}

// Feuille « Traitements » : le N° de réclamation est repris de la feuille « Réclamations »
// par une formule (recherche du code de suivi). Il se met à jour dès que le numéro y est saisi.
function lierNumeroReclamation(ts, from, to) {
  if (to < from || to < 2) return;
  const rs = getSheet(), rc = cols(rs), tc = cols(ts);
  if (!rc['N° de réclamation'] || !rc['Code de suivi'] || !tc['N° de réclamation'] || !tc['Code de suivi']) return;
  const L = n => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  const ref = "'" + SHEET_NAME.replace(/'/g, "''") + "'!";
  const num = ref + L(rc['N° de réclamation']) + ':' + L(rc['N° de réclamation']);
  const cod = ref + L(rc['Code de suivi']) + ':' + L(rc['Code de suivi']);
  const key = L(tc['Code de suivi']);
  const f = [];
  for (let r = Math.max(from, 2); r <= to; r++) f.push(['=IFERROR(INDEX(' + num + ',MATCH(' + key + r + ',' + cod + ',0))&"","")']);
  ts.getRange(Math.max(from, 2), tc['N° de réclamation'], f.length, 1).setFormulas(f);
}

function styleHeader(range) {
  return range.setFontWeight('bold').setBackground('#0f6e86').setFontColor('#ffffff');
}

/**
 * Répare le tableau : anciens liens en #ERROR! et anciens états (Ouverte, En cours, Clôturée).
 * Se lance à l'ouverture du lien du script suivi de ?reparer=1, ou depuis l'éditeur.
 */
function reparerAnciennesLignes() {
  const sh = getSheet(), c = cols(sh), last = sh.getLastRow();
  if (last < 2) return 0;
  const n = sh.getLastColumn();
  const values = sh.getRange(2, 1, last - 1, n).getValues();
  const formulas = sh.getRange(2, 1, last - 1, n).getFormulas();
  const urlOf = f => { const m = String(f || '').match(/HYPERLINK\("([^"]+)"/i); return m ? m[1] : null; };
  let fixed = 0;
  values.forEach((v, i) => {
    const row = i + 2, f = formulas[i];
    let changed = false;
    const carte = c['Carte'] && urlOf(f[c['Carte'] - 1]);
    if (carte) { sh.getRange(row, c['Carte']).setRichTextValue(link('Ouvrir la carte', carte)); changed = true; }
    const photo = c['Lien photo'] && urlOf(f[c['Lien photo'] - 1]);
    if (photo) {
      sh.getRange(row, c['Lien photo']).setRichTextValue(link('Voir la photo', photo));
      addThumb(sh, row, c['Aperçu photo'], photo);
      changed = true;
    }
    const etat = c['Etat de la réclamation'] && v[c['Etat de la réclamation'] - 1];
    if (etat && ANCIENS_ETATS[etat]) { sh.getRange(row, c['Etat de la réclamation']).setValue(ANCIENS_ETATS[etat]); changed = true; }
    if (changed) fixed++;
  });
  return fixed;
}

function placeThumb(sh, row, col, blob) {
  const img = sh.insertImage(blob, col, row, 4, 4);
  const ratio = img.getHeight() / img.getWidth();
  img.setWidth(150).setHeight(Math.round(150 * ratio));
  sh.setRowHeight(row, Math.round(150 * ratio) + 10);
}

function addThumb(sh, row, col, driveUrl) {
  const id = (String(driveUrl).match(/[-\w]{25,}/) || [])[0];
  if (!id || !col) return;
  try {
    const file = DriveApp.getFileById(id);
    const blob = file.getThumbnail() || (file.getSize() < 1.5 * 1024 * 1024 ? file.getBlob() : null);
    if (blob) placeThumb(sh, row, col, blob);
  } catch (e) {}
}

/* ===================== Outils ===================== */

function etatDe(v) {
  const s = String(v || '');
  return ANCIENS_ETATS[s] || s || ETATS[0];
}

// Date et heure au format AAAA-MM-JJ ou AAAA-MM-JJTHH:MM.
function parseDateTime(s) {
  const m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2}))?$/);
  if (!m) return null;
  const dt = new Date(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0));
  if (isNaN(dt) || dt.getMonth() !== +m[2] - 1) return null;
  return dt;
}

function toDate(v) {
  if (v instanceof Date) return isNaN(v) ? null : v;
  const m = String(v || '').match(/^(\d{2})\/(\d{2})\/(\d{4})(?: (\d{2}):(\d{2}))?/);
  return m ? new Date(+m[3], +m[2] - 1, +m[1], +(m[4] || 0), +(m[5] || 0)) : null;
}

function fmt(v) {
  if (v instanceof Date) return isNaN(v) ? '' : Utilities.formatDate(v, Session.getScriptTimeZone(), DATE_FMT);
  return v == null ? '' : String(v);
}

function norm(s) {
  return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').trim();
}

function link(text, url) {
  return SpreadsheetApp.newRichTextValue().setText(text).setLinkUrl(url).build();
}

function jpeg(dataUrl) {
  if (typeof dataUrl !== 'string' || dataUrl.indexOf('data:image/jpeg;base64,') !== 0) return null;
  return Utilities.base64Decode(dataUrl.split(',')[1]);
}

function getFolder() {
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}

// Empêche qu'un texte saisi soit interprété comme une formule dans le tableau.
function safe(v) {
  const s = v == null ? '' : String(v).slice(0, 2000);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

// Formulaire affiché aux agents (généré depuis veolia/formulaire.html).
const FORM_HTML = "<!doctype html>\n<html lang=\"fr\">\n<head>\n<meta charset=\"utf-8\">\n<meta name=\"viewport\" content=\"width=device-width,initial-scale=1,viewport-fit=cover\">\n<meta name=\"theme-color\" content=\"#0f6e86\">\n<meta name=\"robots\" content=\"noindex\">\n<title>Réclamations Ré-Use</title>\n<style>\n:root{\n  --bg:#f3f6f7; --surface:#ffffff; --line:#d6e0e3; --fg:#14262c; --muted:#5b7078;\n  --accent:#0f6e86; --accent-ink:#ffffff; --accent-soft:#e1f0f4;\n  --ok:#2f7d4f; --ok-soft:#e3f3e9; --warn:#a76a00; --warn-soft:#fbf0d9; --bad:#b3332f; --bad-soft:#fbe5e4;\n  --f-body:system-ui,-apple-system,\"Segoe UI\",Roboto,Arial,sans-serif;\n  --f-display:\"Arial Narrow\",\"Roboto Condensed\",system-ui,sans-serif;\n  --f-mono:ui-monospace,Consolas,\"Roboto Mono\",monospace;\n  color-scheme:light;\n}\n@media (prefers-color-scheme: dark){:root{\n  --bg:#0e1a1e; --surface:#15252a; --line:#273c43; --fg:#e3eef1; --muted:#93aab1;\n  --accent:#4fb6cf; --accent-ink:#06171c; --accent-soft:#163a44;\n  --ok:#6cc792; --ok-soft:#173326; --warn:#e6b45a; --warn-soft:#3a2d12; --bad:#ef8a85; --bad-soft:#3d1c1b;\n  color-scheme:dark}}\n*{box-sizing:border-box}\n[hidden]{display:none!important}\n.code-box{border:2px solid var(--accent);border-radius:6px;padding:12px;display:flex;flex-direction:column;gap:6px;background:var(--accent-soft)}\n.code-box .code{font-family:var(--f-mono);font-size:26px;font-weight:700;letter-spacing:.08em;color:var(--accent)}\n.code-box .row-c{display:flex;align-items:center;gap:10px;flex-wrap:wrap}\nhtml,body{margin:0}\nbody{background:var(--bg);color:var(--fg);font-family:var(--f-body);font-size:15px;line-height:1.45}\n.wrap{max-width:560px;margin:0 auto;padding:16px max(14px,env(safe-area-inset-right)) 40px max(14px,env(safe-area-inset-left));display:flex;flex-direction:column;gap:14px}\nheader{border-bottom:2px solid var(--fg);padding-bottom:10px}\nh1{font-family:var(--f-display);font-weight:700;font-size:24px;margin:0}\nh1 span{color:var(--accent)}\n.sub{color:var(--muted);margin:2px 0 0;font-size:13px}\n.panel{display:flex;flex-direction:column;gap:14px}\nform{display:flex;flex-direction:column;gap:14px}\n.card{border:1px solid var(--line);border-radius:10px;padding:14px;margin:0;display:flex;flex-direction:column;gap:10px;min-width:0;background:var(--surface);box-shadow:0 1px 2px rgba(15,60,75,.06)}\n.card>.card-t{font-size:12px;text-transform:uppercase;letter-spacing:.08em;color:var(--accent);font-weight:700;padding:0 0 8px;margin:0 0 2px;border-bottom:1px solid var(--line);display:flex;align-items:center;gap:8px}\n.card>.card-t::before{content:\"\";width:4px;height:14px;border-radius:2px;background:var(--accent)}\n.lg{font-size:11px;text-transform:uppercase;letter-spacing:.08em;color:var(--accent);font-weight:700;padding:0;margin-bottom:6px}\n.row{display:grid;grid-template-columns:1fr 1fr;gap:10px}\n@media (max-width:460px){.row{grid-template-columns:1fr}}\nlabel{display:block;font-weight:600;font-size:13px}\nlabel>input,label>select,label>textarea{display:block;margin-top:4px}\n.req{color:var(--bad)}\ninput,select,textarea{font:inherit;font-size:16px;color:var(--fg);background:var(--bg);border:1px solid var(--line);border-radius:4px;padding:9px 10px;width:100%}\ntextarea{resize:vertical;min-height:64px}\ninput[type=date],input[type=time]{-webkit-appearance:none;appearance:none;display:block;min-width:0;max-width:100%;min-height:44px;text-align:left}\ninput::-webkit-date-and-time-value{text-align:left}\n.row>*{min-width:0}\ninput:focus,select:focus,textarea:focus,button:focus-visible{outline:2px solid var(--accent);outline-offset:1px}\n.radios{display:flex;flex-direction:column;gap:8px}\n.radios label{display:flex;align-items:center;gap:8px;font-weight:400;font-size:14px}\n.radios input{width:auto;margin:0}\n.fuite{background:var(--accent-soft);border-radius:4px;padding:12px}\n.geo{border:1px dashed var(--accent);border-radius:6px;padding:12px;display:flex;flex-direction:column;gap:8px}\n.geo-btn{display:flex;align-items:center;justify-content:center;gap:8px;font-size:16px;padding:12px}\n.geo-btn svg{width:20px;height:20px}\n.geo-out{font-family:var(--f-mono);font-size:13px}\n.geo-out a{color:var(--accent)}\n.det{color:var(--muted);font-size:12px;margin:0}\nbutton{font:inherit;font-weight:600;border-radius:4px;padding:10px 14px;cursor:pointer;border:1px solid var(--line);background:var(--surface);color:var(--fg)}\nbutton.primary{background:var(--accent);border-color:var(--accent);color:var(--accent-ink)}\nbutton:disabled{opacity:.55;cursor:wait}\n.actions{display:flex;gap:8px;flex-wrap:wrap}\n.actions .primary{flex:1 1 auto;font-size:16px;padding:12px}\n#preview{max-width:100%;max-height:200px;object-fit:cover;border-radius:4px;border:1px solid var(--line)}\n.msg{font-size:14px;margin:0;padding:10px 12px;border-radius:4px}\n.msg:empty{display:none}\n.msg.ok{color:var(--ok);background:var(--ok-soft)} .msg.err{color:var(--bad);background:var(--bad-soft)} .msg.info{color:var(--warn);background:var(--warn-soft)}\n.agent{display:flex;gap:8px;align-items:center;font-size:13px}\n.agent label{white-space:nowrap}\n.agent input{padding:6px 8px;font-size:15px;margin:0}\n.agent-id{display:grid;grid-template-columns:1fr 1fr;gap:10px;align-items:start}\n.agent-id label{white-space:normal}\n.agent-id input{margin-top:4px;background:var(--surface)}\n.checks{display:flex;flex-direction:column;gap:8px}\n.checks label{display:flex;align-items:center;gap:8px;font-weight:400;font-size:14px}\n.checks input{width:auto;margin:0}\n.privacy{font-size:12px;color:var(--muted);text-align:center;margin:0}\n</style>\n</head>\n<body>\n<div class=\"wrap\">\n  <header>\n    <h1>Réclamations <span>Ré-Use</span></h1>\n    <p class=\"sub\">Saisie terrain d'une réclamation client</p>\n  </header>\n\n  <section class=\"panel\">\n    <div class=\"card\">\n      <div class=\"card-t\">Agent</div>\n      <div class=\"agent agent-id\">\n        <label for=\"agent\">Nom de l'agent <span class=\"req\">*</span><input id=\"agent\" autocomplete=\"off\" placeholder=\"Votre nom\"></label>\n        <label for=\"matricule\">Matricule <span class=\"req\">*</span><input id=\"matricule\" autocomplete=\"off\" inputmode=\"text\" placeholder=\"Votre matricule\"></label>\n      </div>\n    </div>\n    <form id=\"form\" novalidate>\n      <section class=\"card\">\n        <div class=\"card-t\">Réception</div>\n        <div class=\"row\">\n          <label for=\"recDate\">Date de réception <span class=\"req\">*</span><input id=\"recDate\" type=\"date\" required></label>\n          <label for=\"recHeure\">Heure de réception<input id=\"recHeure\" type=\"time\"></label>\n        </div>\n      </section>\n\n      <section class=\"card\">\n        <div class=\"card-t\">Réclamant</div>\n        <label for=\"nom\">Nom et prénom <span class=\"req\">*</span><input id=\"nom\" autocomplete=\"off\" required></label>\n        <div class=\"row\">\n          <label for=\"tel\">Téléphone<input id=\"tel\" type=\"tel\" inputmode=\"tel\" autocomplete=\"off\" placeholder=\"facultatif\"></label>\n          <label for=\"email\">E-mail<input id=\"email\" type=\"email\" autocomplete=\"off\" placeholder=\"facultatif\"></label>\n        </div>\n      </section>\n\n      <section class=\"card\">\n        <div class=\"card-t\">Localisation GPS</div>\n        <div class=\"geo\">\n          <button type=\"button\" class=\"primary geo-btn\" id=\"btnGeo\">\n            <svg viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\" stroke-linecap=\"round\" stroke-linejoin=\"round\" aria-hidden=\"true\"><path d=\"M12 21s-7-6.2-7-11.5A7 7 0 0 1 19 9.5C19 14.8 12 21 12 21z\"/><circle cx=\"12\" cy=\"9.5\" r=\"2.5\"/></svg>\n            Enregistrer ma position\n          </button>\n          <div class=\"geo-out\" id=\"geoOut\">Aucune position enregistrée.</div>\n          <p class=\"det\" id=\"geoNote\">Placez-vous sur le lieu de la fuite puis appuyez sur le bouton.</p>\n        </div>\n      </section>\n\n      <section class=\"card\">\n        <div class=\"card-t\">Adresse</div>\n        <label for=\"adresse\">Adresse <span class=\"req\">*</span><input id=\"adresse\" autocomplete=\"off\" placeholder=\"N°, rue, lotissement\" required></label>\n        <div class=\"row\">\n          <label for=\"quartier\">Quartier <span class=\"req\">*</span><input id=\"quartier\" autocomplete=\"off\" required></label>\n          <label for=\"secteur\">Secteur <span class=\"req\">*</span><input id=\"secteur\" autocomplete=\"off\" required></label>\n        </div>\n        <label for=\"ville\">Commune <span class=\"req\">*</span><input id=\"ville\" list=\"villes\" autocomplete=\"off\" required></label>\n        <datalist id=\"villes\"><option>Rabat</option><option>Salé</option><option>Témara</option><option>Skhirat</option><option>Harhoura</option><option>Aïn Atiq</option></datalist>\n      </section>\n\n      <section class=\"card\">\n        <div class=\"card-t\">Nature de la réclamation</div>\n        <label for=\"nature\">Nature <span class=\"req\">*</span>\n          <select id=\"nature\" required>\n            <option value=\"\"></option>\n            <option>Réfection</option>\n            <option>Fuite d'eau</option>\n            <option>Qualité d'eau</option>\n            <option>Problème de Pression</option>\n            <option>Demande Coupure</option>\n            <option>Demande Rebranchement</option>\n            <option>DFO</option>\n            <option>Manque d'eau</option>\n            <option>Appel Client pour Information</option>\n          </select>\n        </label>\n        <div class=\"fuite\" id=\"fuiteBox\" hidden>\n          <div class=\"row\">\n            <div>\n              <div class=\"lg\">Localisation de la fuite <span class=\"req\">*</span></div>\n              <div class=\"radios\">\n                <label><input type=\"radio\" name=\"loc\" value=\"Fuite dans la Niche\"> Fuite dans la Niche</label>\n                <label><input type=\"radio\" name=\"loc\" value=\"Fuite dans la Rue\"> Fuite dans la Rue</label>\n                <label><input type=\"radio\" name=\"loc\" value=\"Fuite au mur de la Façade\"> Fuite au mur de la Façade</label>\n                <label><input type=\"radio\" name=\"loc\" value=\"Fuite dans la Nourrice\"> Fuite dans la Nourrice</label>\n                <label><input type=\"radio\" name=\"loc\" value=\"Compteur Sauté\"> Compteur Sauté</label>\n              </div>\n            </div>\n            <div>\n              <div class=\"lg\">Importance de la fuite</div>\n              <div class=\"radios\">\n                <label><input type=\"radio\" name=\"imp\" value=\"Faible (goutte à goutte)\"> Faible (goutte à goutte)</label>\n                <label><input type=\"radio\" name=\"imp\" value=\"Moyenne (filet d'eau)\"> Moyenne (filet d'eau)</label>\n                <label><input type=\"radio\" name=\"imp\" value=\"Importante (jet / inondation)\"> Importante (jet / inondation)</label>\n              </div>\n            </div>\n          </div>\n        </div>\n        <div class=\"fuite\" id=\"refectionBox\" hidden>\n          <div class=\"lg\">Réfection <span class=\"req\">*</span></div>\n          <div class=\"radios\">\n            <label><input type=\"radio\" name=\"refection\" value=\"Non réalisée\"> Non réalisée</label>\n            <label><input type=\"radio\" name=\"refection\" value=\"A reprendre\"> A reprendre</label>\n            <label><input type=\"radio\" name=\"refection\" value=\"Chantier non Nettoyé\"> Chantier non Nettoyé</label>\n          </div>\n        </div>\n        <div class=\"fuite\" id=\"qualiteBox\" hidden>\n          <div class=\"lg\">Nature défaut qualité <span class=\"req\">*</span></div>\n          <div class=\"checks\">\n            <label><input type=\"checkbox\" name=\"qualite\" value=\"Couleur / Turbidité\"> Couleur / Turbidité</label>\n            <label><input type=\"checkbox\" name=\"qualite\" value=\"Goût / Odeur\"> Goût / Odeur</label>\n          </div>\n          <div class=\"row\" style=\"margin-top:10px\">\n            <label for=\"qualiteDepuis\">Défaut qualité depuis<input id=\"qualiteDepuis\" autocomplete=\"off\" placeholder=\"ex. 3 jours, le 15/11\"></label>\n            <label for=\"qualitePeriode\">Période défaut qualité<input id=\"qualitePeriode\" autocomplete=\"off\" placeholder=\"ex. le matin, en continu\"></label>\n          </div>\n        </div>\n        <div class=\"fuite\" id=\"pressionBox\" hidden>\n          <div class=\"lg\">Nature défaut pression <span class=\"req\">*</span></div>\n          <div class=\"radios\">\n            <label><input type=\"radio\" name=\"pression\" value=\"Manque de pression\"> Manque de pression</label>\n            <label><input type=\"radio\" name=\"pression\" value=\"Excès de pression\"> Excès de pression</label>\n            <label><input type=\"radio\" name=\"pression\" value=\"Pas d'Eau\"> Pas d'Eau</label>\n          </div>\n          <div class=\"row\" style=\"margin-top:10px\">\n            <label for=\"pressionDepuis\">Problème de pression depuis<input id=\"pressionDepuis\" autocomplete=\"off\" placeholder=\"ex. 3 jours, le 15/11\"></label>\n            <label for=\"pressionPeriode\">Période problème pression<input id=\"pressionPeriode\" autocomplete=\"off\" placeholder=\"ex. le soir, la nuit\"></label>\n          </div>\n        </div>\n        <div class=\"fuite\" id=\"coupureBox\" hidden>\n          <div class=\"lg\">Demande de coupure</div>\n          <label for=\"coupureCil\">CIL<input id=\"coupureCil\" autocomplete=\"off\" inputmode=\"numeric\" placeholder=\"N° CIL du client\"></label>\n          <div class=\"row\" style=\"margin-top:10px\">\n            <label for=\"coupureDate\">Date souhaitée<input id=\"coupureDate\" type=\"date\"></label>\n            <label for=\"coupureHeure\">Heure souhaitée<input id=\"coupureHeure\" type=\"time\"></label>\n          </div>\n        </div>\n        <div class=\"fuite\" id=\"rebranchementBox\" hidden>\n          <div class=\"lg\">Demande de rebranchement</div>\n          <label for=\"rebranchementCil\">CIL<input id=\"rebranchementCil\" autocomplete=\"off\" inputmode=\"numeric\" placeholder=\"N° CIL du client\"></label>\n          <div class=\"row\" style=\"margin-top:10px\">\n            <label for=\"rebranchementDate\">Date souhaitée<input id=\"rebranchementDate\" type=\"date\"></label>\n            <label for=\"rebranchementHeure\">Heure souhaitée<input id=\"rebranchementHeure\" type=\"time\"></label>\n          </div>\n        </div>\n        <div class=\"fuite\" id=\"dfoBox\" hidden>\n          <div class=\"lg\">DFO à quel niveau <span class=\"req\">*</span></div>\n          <div class=\"radios\">\n            <label><input type=\"radio\" name=\"dfo\" value=\"Conduite\"> Conduite</label>\n            <label><input type=\"radio\" name=\"dfo\" value=\"Niche\"> Niche</label>\n            <label><input type=\"radio\" name=\"dfo\" value=\"Nourrice\"> Nourrice</label>\n            <label><input type=\"radio\" name=\"dfo\" value=\"Branchement\"> Branchement</label>\n          </div>\n        </div>\n        <div class=\"fuite\" id=\"manqueBox\" hidden>\n          <div class=\"lg\">Manque d'eau depuis</div>\n          <div class=\"row\">\n            <label for=\"manqueDate\">Date manque d'eau<input id=\"manqueDate\" type=\"date\"></label>\n            <label for=\"manqueHeure\">Heure manque d'eau<input id=\"manqueHeure\" type=\"time\"></label>\n          </div>\n        </div>\n        <div class=\"fuite\" id=\"appelBox\" hidden>\n          <div class=\"lg\">Raison de l'appel <span class=\"req\">*</span></div>\n          <div class=\"checks\">\n            <label><input type=\"checkbox\" name=\"appel\" value=\"Client non avisé suite coupure programmée\"> Client non avisé suite coupure programmée</label>\n            <label><input type=\"checkbox\" name=\"appel\" value=\"Problème commercial\"> Problème commercial</label>\n          </div>\n        </div>\n      </section>\n\n      <section class=\"card\">\n        <div class=\"card-t\">Photo et observations</div>\n        <label for=\"photo\">Photo de la fuite / du lieu<input id=\"photo\" type=\"file\" accept=\"image/*\" capture=\"environment\"></label>\n        <img id=\"preview\" alt=\"Aperçu de la photo\" hidden>\n        <label for=\"obs\">Observations<textarea id=\"obs\" placeholder=\"Détails donnés par le client\"></textarea></label>\n      </section>\n\n      <p class=\"msg\" id=\"msg\" role=\"status\"></p>\n      <div class=\"code-box\" id=\"codeBox\" hidden>\n        <span>Code de suivi de la réclamation :</span>\n        <div class=\"row-c\"><span class=\"code\" id=\"codeVal\"></span><button type=\"button\" id=\"btnCopy\">Copier</button></div>\n        <p class=\"det\">Notez ce code : il sera demandé dans le formulaire de traitement.</p>\n      </div>\n      <div class=\"actions\">\n        <button type=\"submit\" class=\"primary\" id=\"btnSave\">Envoyer la réclamation</button>\n        <button type=\"reset\">Effacer</button>\n      </div>\n    </form>\n  </section>\n  <p class=\"privacy\">Les réclamations envoyées ne sont pas visibles sur ce formulaire : elles sont transmises directement au responsable, avec votre compte Veolia.</p>\n  <p class=\"privacy\">Version Veolia 12</p>\n</div>\n\n<script>\n\"use strict\";\n/* Version hébergée dans le script Google Veolia : l'envoi passe par google.script.run */\nconst ENDPOINT = \"script\";\n\nconst $=id=>document.getElementById(id);\nlet pos=null, photoData='', thumbData='';\nconst QKEY='reuse-envoi-en-attente';\n\n['agent','matricule'].forEach(id=>{\n  try{$(id).value=localStorage.getItem('reuse-'+id)||''}catch(e){}\n  $(id).addEventListener('change',()=>{try{localStorage.setItem('reuse-'+id,$(id).value.trim())}catch(e){}});\n});\n\nfunction msg(t,k){$('msg').textContent=t;$('msg').className='msg '+(k||'')}\nconst mapsUrl=(lat,lng)=>'https://www.google.com/maps?q='+lat+','+lng;\n\n/* ---------- Position en un clic ---------- */\nfunction showPos(){\n  const o=$('geoOut');o.replaceChildren();\n  if(!pos){o.textContent='Aucune position enregistrée.';return}\n  const a=document.createElement('a');a.href=mapsUrl(pos.lat,pos.lng);a.target='_blank';a.rel='noopener';a.textContent='Vérifier sur la carte';\n  o.append(pos.lat.toFixed(6)+', '+pos.lng.toFixed(6)+'  (± '+Math.round(pos.acc)+' m)  ',a);\n}\n$('btnGeo').addEventListener('click',()=>{\n  const n=$('geoNote'),b=$('btnGeo');\n  if(!('geolocation' in navigator)){n.textContent=\"Ce navigateur ne donne pas accès au GPS.\";return}\n  if(!window.isSecureContext){n.textContent=\"Le GPS n'est pas disponible ici : ajoutez une photo prise avec la localisation activée, sa position sera reprise.\";return}\n  b.disabled=true;n.textContent='Recherche de la position GPS…';\n  navigator.geolocation.getCurrentPosition(p=>{\n    pos={lat:+p.coords.latitude.toFixed(6),lng:+p.coords.longitude.toFixed(6),acc:p.coords.accuracy};\n    showPos();b.disabled=false;\n    n.textContent=pos.acc>50?'Position enregistrée, mais peu précise (± '+Math.round(pos.acc)+' m). Réessayez à l’extérieur si possible.':'Position enregistrée.';\n  },err=>{\n    b.disabled=false;\n    n.textContent=err.code===1?\"Accès à la position refusé ou bloqué. Autorisez la localisation, ou ajoutez une photo prise avec la localisation activée : sa position sera reprise.\":err.code===3?\"Le GPS n'a pas répondu à temps. Réessayez à l'extérieur.\":\"Position introuvable. Vérifiez que la localisation du téléphone est activée.\";\n  },{enableHighAccuracy:true,timeout:20000,maximumAge:0});\n});\n\n/* ---------- Formulaire ---------- */\nconst BOXES={\"Fuite d'eau\":'fuiteBox','Réfection':'refectionBox',\"Qualité d'eau\":'qualiteBox','Problème de Pression':'pressionBox',\n  'Demande Coupure':'coupureBox','Demande Rebranchement':'rebranchementBox','DFO':'dfoBox',\"Manque d'eau\":'manqueBox','Appel Client pour Information':'appelBox'};\nconst frDate=s=>s?s.split('-').reverse().join('/'):'';\nconst dateHeure=(d,h)=>[frDate(d),h].filter(Boolean).join(' à ');\nfunction showBoxes(){const n=$('nature').value;Object.entries(BOXES).forEach(([k,id])=>{$(id).hidden=n!==k})}\n$('nature').addEventListener('change',showBoxes);\nfunction nowFields(){const d=new Date(),p=n=>String(n).padStart(2,'0');$('recDate').value=d.getFullYear()+'-'+p(d.getMonth()+1)+'-'+p(d.getDate());$('recHeure').value=p(d.getHours())+':'+p(d.getMinutes());$('recDate').max=$('recDate').value}\nnowFields();\n$('form').addEventListener('reset',()=>setTimeout(()=>{nowFields();showBoxes();$('preview').hidden=true;photoData='';thumbData='';pos=null;showPos();$('geoNote').textContent='Placez-vous sur le lieu de la fuite puis appuyez sur le bouton.'},0));\nfunction shrink(file,max=1280){return new Promise((res,rej)=>{const img=new Image();img.onload=()=>{const s=Math.min(1,max/Math.max(img.width,img.height)),c=document.createElement('canvas');c.width=Math.round(img.width*s);c.height=Math.round(img.height*s);c.getContext('2d').drawImage(img,0,0,c.width,c.height);URL.revokeObjectURL(img.src);res(c.toDataURL('image/jpeg',0.75))};img.onerror=()=>rej(new Error('image'));img.src=URL.createObjectURL(file)})}\n/* Position lue dans la photo (si le GPS du formulaire n'a pas pu être utilisé) */\nfunction gpsFromJpeg(buf){\n  const v=new DataView(buf);if(v.byteLength<4||v.getUint16(0)!==0xFFD8)return null;\n  let o=2;\n  while(o+4<v.byteLength){\n    const m=v.getUint16(o),len=v.getUint16(o+2);\n    if(m===0xFFE1&&v.getUint32(o+4)===0x45786966){\n      const t=o+10,le=v.getUint16(t)===0x4949;\n      const u16=p=>v.getUint16(t+p,le),u32=p=>v.getUint32(t+p,le);\n      const ifd0=u32(4),n0=u16(ifd0);let g=0;\n      for(let i=0;i<n0;i++){const en=ifd0+2+i*12;if(u16(en)===0x8825)g=u32(en+8)}\n      if(!g)return null;\n      const r={},n=u16(g);\n      for(let i=0;i<n;i++){\n        const en=g+2+i*12,tag=u16(en);\n        if(tag===1||tag===3)r[tag]=String.fromCharCode(v.getUint8(t+en+8));\n        if(tag===2||tag===4){const p=u32(en+8),q=k=>u32(p+k*8)/(u32(p+k*8+4)||1);r[tag]=q(0)+q(1)/60+q(2)/3600}\n      }\n      if(r[2]==null||r[4]==null)return null;\n      return {lat:(r[1]==='S'?-1:1)*r[2],lng:(r[3]==='W'?-1:1)*r[4]};\n    }\n    if((m&0xFF00)!==0xFF00)break;o+=2+len;\n  }\n  return null;\n}\n$('photo').addEventListener('change',async()=>{\n  const f=$('photo').files[0];photoData='';thumbData='';$('preview').hidden=true;if(!f)return;\n  if(!pos){try{const g=gpsFromJpeg(await f.arrayBuffer());\n    if(g&&isFinite(g.lat)&&isFinite(g.lng)&&(g.lat||g.lng)){pos={lat:+g.lat.toFixed(6),lng:+g.lng.toFixed(6),acc:30};showPos();$('geoNote').textContent='Position lue dans la photo.'}}catch(e){}}\n  try{photoData=await shrink(f);thumbData=await shrink(f,400);$('preview').src=photoData;$('preview').hidden=false}\n  catch(e){msg(\"Cette image n'a pas pu être lue. Essayez une autre photo.\",'err')}\n});\n\nlet lastErr='';\nfunction send(payload){\n  return new Promise((res,rej)=>{\n    google.script.run\n      .withSuccessHandler(j=>{ if(j&&j.ok)res(j); else rej(Object.assign(new Error((j&&j.error)||'refus'),{refused:true})); })\n      .withFailureHandler(err=>{ lastErr=String(err&&err.message||err); rej(err); })\n      .envoyer(payload);\n  });\n}\n\n/* File d'attente : uniquement si le réseau manque ; vidée dès l'envoi réussi */\nfunction queue(){try{return JSON.parse(localStorage.getItem(QKEY)||'[]')}catch(e){return[]}}\nfunction setQueue(q){try{q.length?localStorage.setItem(QKEY,JSON.stringify(q)):localStorage.removeItem(QKEY)}catch(e){}}\nasync function flush(){\n  if(!ENDPOINT)return;\n  let q=queue();if(!q.length)return;\n  const rest=[];\n  for(const p of q){try{await send(p)}catch(e){if(!e.refused)rest.push(p)}}\n  setQueue(rest);\n  if(q.length!==rest.length&&!rest.length)msg('Les réclamations en attente ont été envoyées.','ok');\n  else if(rest.length)msg(rest.length+' réclamation(s) en attente d’envoi : elles partiront dès que le réseau revient.','info');\n}\nwindow.addEventListener('online',flush);\n\n$('form').addEventListener('submit',async e=>{\n  e.preventDefault();\n  const v=id=>($(id).value||'').trim(),f=$('form');\n  if(!ENDPOINT)return msg(\"Ce formulaire n'est pas encore relié au tableau du responsable.\",'err');\n  const miss=[];if(!v('recDate'))miss.push('date de réception');if(!v('agent'))miss.push('nom de l’agent');if(!v('matricule'))miss.push('matricule');if(!v('nom'))miss.push('nom du réclamant');if(!v('adresse'))miss.push('adresse');if(!v('quartier'))miss.push('quartier');if(!v('secteur'))miss.push('secteur');if(!v('ville'))miss.push('commune');if(!v('nature'))miss.push('nature');\n  if(miss.length)return msg('Champ obligatoire manquant : '+miss.join(', ')+'.','err');\n  if(v('tel')&&(v('tel').match(/\\d/g)||[]).length<8)return msg('Le numéro de téléphone semble incomplet (8 chiffres minimum).','err');\n  const fuite=v('nature')===\"Fuite d'eau\";\n  if(fuite&&!f.elements.loc.value)return msg('Indiquez la localisation de la fuite.','err');\n  const nat=v('nature');let detail='',depuis='',periode='',cil='',souhaite='';\n  const checked=n=>[...f.querySelectorAll('input[name='+n+']:checked')].map(x=>x.value).join(', ');\n  if(nat==='Demande Coupure'){cil=v('coupureCil');souhaite=dateHeure(v('coupureDate'),v('coupureHeure'))}\n  if(nat==='Demande Rebranchement'){cil=v('rebranchementCil');souhaite=dateHeure(v('rebranchementDate'),v('rebranchementHeure'))}\n  if(nat==='DFO'){detail=f.elements.dfo.value;if(!detail)return msg('Indiquez à quel niveau se trouve le DFO.','err')}\n  if(nat===\"Manque d'eau\"){depuis=dateHeure(v('manqueDate'),v('manqueHeure'))}\n  if(nat==='Appel Client pour Information'){detail=checked('appel');if(!detail)return msg('Cochez la raison de l’appel.','err')}\n  if(nat==='Réfection'){detail=f.elements.refection.value;if(!detail)return msg('Indiquez le type de réfection.','err')}\n  if(nat===\"Qualité d'eau\"){detail=[...f.querySelectorAll('input[name=qualite]:checked')].map(x=>x.value).join(', ');if(!detail)return msg('Cochez la nature du défaut qualité.','err');depuis=v('qualiteDepuis');periode=v('qualitePeriode')}\n  if(nat==='Problème de Pression'){detail=f.elements.pression.value;if(!detail)return msg('Indiquez la nature du défaut de pression.','err');depuis=v('pressionDepuis');periode=v('pressionPeriode')}\n  const code=codeSuivi();\n  const payload={code,clientId:(crypto.randomUUID?crypto.randomUUID():Date.now()+'-'+Math.random()),saisieLe:new Date().toISOString(),\n    reception:v('recDate')+(v('recHeure')?'T'+v('recHeure'):''),agent:v('agent'),matricule:v('matricule'),nom:v('nom'),tel:v('tel'),email:v('email'),adresse:v('adresse'),quartier:v('quartier'),secteur:v('secteur'),ville:v('ville'),\n    lat:pos?pos.lat:'',lng:pos?pos.lng:'',precision:pos?Math.round(pos.acc):'',\n    nature:v('nature'),localisation:fuite?f.elements.loc.value:'',importance:fuite?(f.elements.imp.value||''):'',detail,depuis,periode,cil,souhaite,obs:v('obs'),photo:photoData,vignette:thumbData};\n  const b=$('btnSave');b.disabled=true;msg('Envoi en cours…','info');$('codeBox').hidden=true;\n  try{\n    const j=await send(payload);\n    f.reset();setTimeout(()=>{msg('Réclamation envoyée au responsable.','ok');showCode(code)},10);\n  }catch(err){\n    if(err.refused){msg('Réclamation refusée : '+err.message,'err')}\n    else{const q=queue();q.push(payload);setQueue(q);f.reset();\n      setTimeout(()=>showCode(code),10);\n      setTimeout(()=>msg(navigator.onLine?\"Envoi impossible pour le moment : la réclamation est gardée sur ce téléphone et sera renvoyée automatiquement à la prochaine ouverture du formulaire.\"+(lastErr?' Détail technique : '+lastErr:''):\"Pas de réseau : la réclamation est gardée sur ce téléphone et partira automatiquement dès que le réseau revient.\",'info'),10)}\n  }finally{b.disabled=false}\n});\nflush();\n\n/* Code de suivi : relie la réclamation à son traitement (ex. RU-K7P4Q) */\nfunction codeSuivi(){\n  const a='23456789ABCDEFGHJKMNPQRSTUVWXYZ',r=new Uint32Array(5);\n  (window.crypto||window.msCrypto).getRandomValues(r);\n  return 'RU-'+[...r].map(x=>a[x%a.length]).join('');\n}\nfunction showCode(c){$('codeVal').textContent=c;$('btnCopy').textContent='Copier';$('codeBox').hidden=false}\n$('btnCopy').addEventListener('click',()=>{\n  const c=$('codeVal').textContent;\n  const sel=()=>{const r=document.createRange();r.selectNodeContents($('codeVal'));const s=getSelection();s.removeAllRanges();s.addRange(r)};\n  try{navigator.clipboard.writeText(c).then(()=>{$('btnCopy').textContent='Copié'},sel)}catch(e){sel()}\n});\n</script>\n</body>\n</html>\n";

// Formulaire de traitement (généré depuis veolia/traitement.html).
const TRAITEMENT_HTML = "<!doctype html>\n<html lang=\"fr\">\n<head>\n<meta charset=\"utf-8\">\n<meta name=\"viewport\" content=\"width=device-width,initial-scale=1,viewport-fit=cover\">\n<meta name=\"theme-color\" content=\"#0f6e86\">\n<meta name=\"robots\" content=\"noindex\">\n<title>Traitement Ré-Use</title>\n<style>\n:root{\n  --bg:#f3f6f7; --surface:#ffffff; --line:#d6e0e3; --fg:#14262c; --muted:#5b7078;\n  --accent:#0f6e86; --accent-ink:#ffffff; --accent-soft:#e1f0f4;\n  --ok:#2f7d4f; --ok-soft:#e3f3e9; --warn:#a76a00; --warn-soft:#fbf0d9; --bad:#b3332f; --bad-soft:#fbe5e4;\n  --f-body:system-ui,-apple-system,\"Segoe UI\",Roboto,Arial,sans-serif;\n  --f-display:\"Arial Narrow\",\"Roboto Condensed\",system-ui,sans-serif;\n  --f-mono:ui-monospace,Consolas,\"Roboto Mono\",monospace;\n  color-scheme:light;\n}\n@media (prefers-color-scheme: dark){:root{\n  --bg:#0e1a1e; --surface:#15252a; --line:#273c43; --fg:#e3eef1; --muted:#93aab1;\n  --accent:#4fb6cf; --accent-ink:#06171c; --accent-soft:#163a44;\n  --ok:#6cc792; --ok-soft:#173326; --warn:#e6b45a; --warn-soft:#3a2d12; --bad:#ef8a85; --bad-soft:#3d1c1b;\n  color-scheme:dark}}\n*{box-sizing:border-box}\n[hidden]{display:none!important}\n.code-box{border:2px solid var(--accent);border-radius:6px;padding:12px;display:flex;flex-direction:column;gap:6px;background:var(--accent-soft)}\n.code-box .code{font-family:var(--f-mono);font-size:26px;font-weight:700;letter-spacing:.08em;color:var(--accent)}\n.code-box .row-c{display:flex;align-items:center;gap:10px;flex-wrap:wrap}\nhtml,body{margin:0}\nbody{background:var(--bg);color:var(--fg);font-family:var(--f-body);font-size:15px;line-height:1.45}\n.wrap{max-width:560px;margin:0 auto;padding:16px max(14px,env(safe-area-inset-right)) 40px max(14px,env(safe-area-inset-left));display:flex;flex-direction:column;gap:14px}\nheader{border-bottom:2px solid var(--fg);padding-bottom:10px}\nh1{font-family:var(--f-display);font-weight:700;font-size:24px;margin:0}\nh1 span{color:var(--accent)}\n.sub{color:var(--muted);margin:2px 0 0;font-size:13px}\n.panel{display:flex;flex-direction:column;gap:14px}\nform{display:flex;flex-direction:column;gap:14px}\n.card{border:1px solid var(--line);border-radius:10px;padding:14px;margin:0;display:flex;flex-direction:column;gap:10px;min-width:0;background:var(--surface);box-shadow:0 1px 2px rgba(15,60,75,.06)}\n.card>.card-t{font-size:12px;text-transform:uppercase;letter-spacing:.08em;color:var(--accent);font-weight:700;padding:0 0 8px;margin:0 0 2px;border-bottom:1px solid var(--line);display:flex;align-items:center;gap:8px}\n.card>.card-t::before{content:\"\";width:4px;height:14px;border-radius:2px;background:var(--accent)}\n.lg{font-size:11px;text-transform:uppercase;letter-spacing:.08em;color:var(--accent);font-weight:700;padding:0;margin-bottom:6px}\n.row{display:grid;grid-template-columns:1fr 1fr;gap:10px}\n@media (max-width:460px){.row{grid-template-columns:1fr}}\nlabel{display:block;font-weight:600;font-size:13px}\nlabel>input,label>select,label>textarea{display:block;margin-top:4px}\n.req{color:var(--bad)}\ninput,select,textarea{font:inherit;font-size:16px;color:var(--fg);background:var(--bg);border:1px solid var(--line);border-radius:4px;padding:9px 10px;width:100%}\ntextarea{resize:vertical;min-height:64px}\ninput[type=date],input[type=time]{-webkit-appearance:none;appearance:none;display:block;min-width:0;max-width:100%;min-height:44px;text-align:left}\ninput::-webkit-date-and-time-value{text-align:left}\n.row>*{min-width:0}\ninput:focus,select:focus,textarea:focus,button:focus-visible{outline:2px solid var(--accent);outline-offset:1px}\n.radios{display:flex;flex-direction:column;gap:8px}\n.radios label{display:flex;align-items:center;gap:8px;font-weight:400;font-size:14px}\n.radios input{width:auto;margin:0}\n.fuite{background:var(--accent-soft);border-radius:4px;padding:12px}\n.geo{border:1px dashed var(--accent);border-radius:6px;padding:12px;display:flex;flex-direction:column;gap:8px}\n.geo-btn{display:flex;align-items:center;justify-content:center;gap:8px;font-size:16px;padding:12px}\n.geo-btn svg{width:20px;height:20px}\n.geo-out{font-family:var(--f-mono);font-size:13px}\n.geo-out a{color:var(--accent)}\n.det{color:var(--muted);font-size:12px;margin:0}\nbutton{font:inherit;font-weight:600;border-radius:4px;padding:10px 14px;cursor:pointer;border:1px solid var(--line);background:var(--surface);color:var(--fg)}\nbutton.primary{background:var(--accent);border-color:var(--accent);color:var(--accent-ink)}\nbutton:disabled{opacity:.55;cursor:wait}\n.actions{display:flex;gap:8px;flex-wrap:wrap}\n.actions .primary{flex:1 1 auto;font-size:16px;padding:12px}\n#preview{max-width:100%;max-height:200px;object-fit:cover;border-radius:4px;border:1px solid var(--line)}\n.msg{font-size:14px;margin:0;padding:10px 12px;border-radius:4px}\n.msg:empty{display:none}\n.msg.ok{color:var(--ok);background:var(--ok-soft)} .msg.err{color:var(--bad);background:var(--bad-soft)} .msg.info{color:var(--warn);background:var(--warn-soft)}\n.agent{display:flex;gap:8px;align-items:center;font-size:13px}\n.agent label{white-space:nowrap}\n.agent input{padding:6px 8px;font-size:15px;margin:0}\n.agent-id{display:grid;grid-template-columns:1fr 1fr;gap:10px;align-items:start}\n.agent-id label{white-space:normal}\n.agent-id input{margin-top:4px;background:var(--surface)}\n.checks{display:flex;flex-direction:column;gap:8px}\n.checks label{display:flex;align-items:center;gap:8px;font-weight:400;font-size:14px}\n.checks input{width:auto;margin:0}\n.privacy{font-size:12px;color:var(--muted);text-align:center;margin:0}\n.fixed{padding:9px 10px;border:1px dashed var(--line);border-radius:4px;color:var(--muted);font-size:15px;margin-top:4px}\n.fixed.ok{color:var(--fg);border-style:solid;border-color:var(--accent);background:var(--accent-soft);font-weight:600}\n.thumbs{display:flex;gap:8px;flex-wrap:wrap}\n.thumbs img{width:84px;height:64px;object-fit:cover;border-radius:4px;border:1px solid var(--line)}\n</style>\n</head>\n<body>\n<div class=\"wrap\">\n  <header>\n    <h1>Traitement <span>Ré-Use</span></h1>\n    <p class=\"sub\">À remplir par l'agent terrain après l'intervention</p>\n  </header>\n\n  <section class=\"panel\">\n    <div class=\"card\">\n      <div class=\"card-t\">Agent</div>\n      <div class=\"agent agent-id\">\n        <label for=\"agent\">Nom de l'agent <span class=\"req\">*</span><input id=\"agent\" autocomplete=\"off\" placeholder=\"Votre nom\"></label>\n        <label for=\"matricule\">Matricule <span class=\"req\">*</span><input id=\"matricule\" autocomplete=\"off\" placeholder=\"Votre matricule\"></label>\n      </div>\n    </div>\n    <form id=\"form\" novalidate>\n      <section class=\"card\">\n        <div class=\"card-t\">Réclamation</div>\n        <label for=\"code\">Code de suivi <span class=\"req\">*</span><input id=\"code\" autocomplete=\"off\" autocapitalize=\"characters\" spellcheck=\"false\" placeholder=\"ex. RU-K7P4Q (donné à la saisie)\"></label>\n        <p class=\"det\" id=\"codeNote\">La nature de la réclamation s'affichera dès que le code est saisi.</p>\n        <label for=\"dateReception\">Date réception<input id=\"dateReception\" type=\"date\"></label>\n        <label for=\"etat\">Etat de la réclamation <span class=\"req\">*</span>\n          <select id=\"etat\">\n            <option value=\"\"></option>\n            <option>Résolue Définitivement</option>\n            <option>Transmise mais non encore résolue</option>\n            <option>Résolue Provisoirement</option>\n            <option>Non encore Transmise</option>\n          </select>\n        </label>\n        <div class=\"row\">\n          <div><label>Métier</label><div class=\"fixed\">Ré-Use</div></div>\n          <div><label>Nature réclamation</label><div class=\"fixed\" id=\"natureVal\">—</div></div>\n        </div>\n      </section>\n\n      <section class=\"card\">\n        <div class=\"card-t\">Client</div>\n        <div class=\"row\">\n          <label for=\"typeClient\">Type client\n            <select id=\"typeClient\">\n              <option value=\"\"></option>\n              <option>Particulier</option><option>Autorité</option><option>Communes</option><option>Elus</option>\n              <option>Résidences Royales</option><option>Industriels</option><option>V.I.P</option><option>GC-Entreprises</option>\n              <option>Planifié</option><option>GC-Administrations</option><option>Numéro Vert</option>\n            </select>\n          </label>\n          <label for=\"nomClient\">Nom du client<input id=\"nomClient\" autocomplete=\"off\"></label>\n        </div>\n        <div class=\"row\">\n          <label for=\"cil\">CIL<input id=\"cil\" autocomplete=\"off\" inputmode=\"numeric\"></label>\n          <label for=\"telPortable\">Téléphone portable 1<input id=\"telPortable\" type=\"tel\" inputmode=\"tel\" autocomplete=\"off\"></label>\n        </div>\n        <label for=\"telFixe\">Téléphone fixe<input id=\"telFixe\" type=\"tel\" inputmode=\"tel\" autocomplete=\"off\"></label>\n      </section>\n\n      <section class=\"card\">\n        <div class=\"card-t\">Suivi</div>\n        <label for=\"rappels\">Nombre de rappels<input id=\"rappels\" type=\"number\" min=\"0\" step=\"1\" inputmode=\"numeric\" value=\"0\"></label>\n        <label for=\"photos\">Photos (jusqu'à 5)<input id=\"photos\" type=\"file\" accept=\"image/*\" multiple></label>\n        <div class=\"thumbs\" id=\"thumbs\"></div>\n        <p class=\"det\">Nombre de photos : <b id=\"nbPhotos\">0</b></p>\n      </section>\n\n      <p class=\"msg\" id=\"msg\" role=\"status\"></p>\n      <div class=\"actions\">\n        <button type=\"submit\" class=\"primary\" id=\"btnSave\">Enregistrer le traitement</button>\n        <button type=\"reset\">Effacer</button>\n      </div>\n    </form>\n  </section>\n  <p class=\"privacy\">Le traitement est transmis directement au responsable, avec votre compte Veolia.</p>\n  <p class=\"privacy\">Traitement — Version Veolia 12</p>\n</div>\n\n<script>\n\"use strict\";\nconst $=id=>document.getElementById(id);\nlet photos=[];\n['agent','matricule'].forEach(id=>{\n  try{$(id).value=localStorage.getItem('reuse-'+id)||''}catch(e){}\n  $(id).addEventListener('change',()=>{try{localStorage.setItem('reuse-'+id,$(id).value.trim())}catch(e){}});\n});\nfunction msg(t,k){$('msg').textContent=t;$('msg').className='msg '+(k||'')}\nfunction shrink(file,max=1280){return new Promise((res,rej)=>{const img=new Image();img.onload=()=>{const s=Math.min(1,max/Math.max(img.width,img.height)),c=document.createElement('canvas');c.width=Math.round(img.width*s);c.height=Math.round(img.height*s);c.getContext('2d').drawImage(img,0,0,c.width,c.height);URL.revokeObjectURL(img.src);res(c.toDataURL('image/jpeg',0.75))};img.onerror=()=>rej(new Error('image'));img.src=URL.createObjectURL(file)})}\nfunction showPhotos(){const t=$('thumbs');t.replaceChildren();photos.forEach(p=>{const i=document.createElement('img');i.src=p;i.alt='Photo';t.append(i)});$('nbPhotos').textContent=photos.length}\n$('photos').addEventListener('change',async()=>{\n  const files=[...$('photos').files].slice(0,5);photos=[];\n  if($('photos').files.length>5)msg('Seules les 5 premières photos sont gardées.','info');\n  for(const f of files){try{photos.push(await shrink(f))}catch(e){msg(\"Une image n'a pas pu être lue.\",'err')}}\n  showPhotos();\n});\n$('form').addEventListener('reset',()=>setTimeout(()=>{photos=[];showPhotos();setNature('', '')},0));\n\n/* Code de suivi : dès qu'il est complet, la nature de la réclamation est reprise de la saisie */\nlet natureCode='',codeVerifie='',lookupSeq=0;\nconst normCode=s=>{let c=String(s||'').toUpperCase().replace(/\\s+/g,'');if(c&&!c.startsWith('RU-'))c='RU-'+c.replace(/^RU/,'');return c};\nfunction setNature(nature,note,kind){\n  natureCode=nature;$('natureVal').textContent=nature||'—';$('natureVal').className='fixed'+(nature?' ok':'');\n  $('codeNote').textContent=note||\"La nature de la réclamation s'affichera dès que le code est saisi.\";$('codeNote').style.color=kind==='err'?'var(--bad)':'';\n}\nfunction verifierCode(){\n  const c=normCode($('code').value);codeVerifie='';\n  if(!/^RU-[2-9A-Z]{5}$/.test(c)){setNature('',c.length>3?'Code incomplet : il ressemble à RU-K7P4Q.':'');return}\n  const seq=++lookupSeq;setNature('','Recherche de la réclamation…');\n  google.script.run\n    .withSuccessHandler(r=>{if(seq!==lookupSeq)return;\n      if(r&&r.ok){codeVerifie=c;setNature(r.nature,'Réclamation trouvée.')}\n      else setNature('',(r&&r.error)||'Code de suivi inconnu.','err')})\n    .withFailureHandler(()=>{if(seq===lookupSeq)setNature('','Vérification impossible : vérifiez le réseau.','err')})\n    .natureDuCode(c);\n}\n$('code').addEventListener('input',verifierCode);\n\nfunction send(payload){\n  return new Promise((res,rej)=>{\n    google.script.run\n      .withSuccessHandler(j=>{ if(j&&j.ok)res(j); else rej(Object.assign(new Error((j&&j.error)||'refus'),{refused:true})); })\n      .withFailureHandler(rej)\n      .enregistrerTraitement(payload);\n  });\n}\n\n$('form').addEventListener('submit',async e=>{\n  e.preventDefault();\n  const v=id=>($(id).value||'').trim();\n  const miss=[];\n  if(!v('agent'))miss.push('nom de l’agent');if(!v('matricule'))miss.push('matricule');\n  if(!v('code'))miss.push('code de suivi');if(!v('etat'))miss.push('état de la réclamation');\n  if(miss.length)return msg('Champ obligatoire manquant : '+miss.join(', ')+'.','err');\n  let code=v('code').toUpperCase().replace(/\\s+/g,'');if(code&&!code.startsWith('RU-'))code='RU-'+code.replace(/^RU/,'');\n  if(!/^RU-[2-9A-Z]{5}$/.test(code))return msg('Le code de suivi doit ressembler à RU-K7P4Q.','err');\n  if(code!==codeVerifie||!natureCode)return msg('Code de suivi inconnu : vérifiez le code donné lors de la saisie.','err');\n  const payload={code,agent:v('agent'),matricule:v('matricule'),dateReception:v('dateReception'),\n    etat:v('etat'),nature:natureCode,typeClient:v('typeClient'),nomClient:v('nomClient'),cil:v('cil'),\n    telPortable:v('telPortable'),telFixe:v('telFixe'),rappels:v('rappels'),photos};\n  const b=$('btnSave');b.disabled=true;msg('Envoi en cours…','info');\n  try{\n    const j=await send(payload);\n    $('form').reset();\n    setTimeout(()=>msg('Traitement enregistré pour la réclamation '+code+'.','ok'),10);\n  }catch(err){\n    msg(err.refused?'Traitement refusé : '+err.message:'Envoi impossible pour le moment. Vérifiez le réseau puis appuyez de nouveau sur « Enregistrer le traitement ».','err');\n  }finally{b.disabled=false}\n});\n</script>\n</body>\n</html>\n";
