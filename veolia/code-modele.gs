/**
 * Réclamations Ré-Use — version compte Veolia.
 * Un seul lien (l'application Web), réservé aux comptes Veolia connectés :
 *  - page d'accueil : l'agent choisit « Saisie d'une réclamation » ou « Traitement » ;
 *  - Saisie (?f=saisie) : enregistre une nouvelle réclamation et donne à l'agent un code de suivi (ex. RU-K7P4Q) ;
 *  - Traitement (?f=traitement) : après l'intervention, l'agent terrain saisit le code de suivi et les
 *    champs de traitement (écran WebSIGID), écrits dans les colonnes « Traitement - … » de la même ligne.
 * Les agents ne voient jamais les réclamations enregistrées.
 *
 * Installation : Extensions > Apps Script, coller ce code, enregistrer,
 * puis Déployer > Nouveau déploiement > Application Web
 * (Exécuter en tant que : Moi ; Qui a accès : Anyone within Veolia).
 */

const VERSION = 'Veolia 7';
const CODE_RE = /^RU-[2-9A-Z]{5}$/;
const SHEET_NAME = 'Réclamations';
const FOLDER_NAME = 'Photos réclamations Ré-Use';
const DATE_FMT = 'dd/MM/yyyy HH:mm';

const T = 'Traitement - ';
const TRAITEMENT_HEADERS = ['Date réception', 'Métier', 'Nature réclamation', 'Objet préliminaire', 'Type client',
  'Nom du client', 'CIL', 'Préfecture', 'Commune', 'Quartier', 'Sous-quartier', 'Rue', 'Adresse de la réclamation',
  'Téléphone portable 1', 'Téléphone fixe', 'Nombre de rappels', 'Nombre de photos', 'Photos',
  'Utilisateur saisie', "Date d'envoi"].map(h => T + h);
const HEADERS = ['N° Saisie', 'N° de réclamation', 'Code de suivi', 'Date de réception', "Date d'envoi", 'Agent', 'Matricule',
  'Nom du réclamant', 'Téléphone', 'E-mail',
  'Adresse', 'Quartier', 'Commune', 'Latitude', 'Longitude', 'Précision (m)', 'Carte',
  'Nature', 'Localisation fuite', 'Importance fuite', 'Détail nature', 'Depuis', 'Période', 'CIL', 'Date/heure souhaitée',
  'Observations', 'Aperçu photo', 'Lien photo', 'Etat de la réclamation']
  .concat(TRAITEMENT_HEADERS, ['ID envoi']);
const TEXT_COLS = ['N° de réclamation', 'Code de suivi', 'Matricule', 'CIL', 'Téléphone', 'E-mail', T + 'CIL', T + 'Téléphone portable 1', T + 'Téléphone fixe'];
const DATE_COLS = ['Date de réception', "Date d'envoi", T + 'Date réception', T + "Date d'envoi"];
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
  const page = f === 'saisie' ? FORM_HTML : f === 'traitement' ? TRAITEMENT_HTML : ACCUEIL_HTML;
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
    for (const k of ['matricule', 'nom', 'adresse', 'nature']) {
      if (!d[k] || !String(d[k]).trim()) return { ok: false, error: 'champ manquant (' + k + ')' };
    }
    if (NATURES.indexOf(d.nature) < 0) return { ok: false, error: 'nature inconnue' };

    const sh = getSheet();
    const c = cols(sh);

    // Même envoi reçu deux fois (réseau coupé puis rétabli) : on ne le duplique pas.
    if (d.clientId && sh.getLastRow() > 1) {
      const found = sh.getRange(2, c['ID envoi'], sh.getLastRow() - 1, 1)
        .createTextFinder(String(d.clientId)).matchEntireCell(true).findNext();
      if (found) return { ok: true, numero: sh.getRange(found.getRow(), c['N° Saisie']).getValue(), doublon: true };
    }

    const numero = 'RU-' + new Date().getFullYear() + '-' + String(sh.getLastRow()).padStart(5, '0');
    const code = CODE_RE.test(String(d.code || '')) ? String(d.code) : nouveauCode();
    const hasPos = d.lat !== '' && d.lat != null && d.lng !== '' && d.lng != null && isFinite(d.lat) && isFinite(d.lng);

    const values = {
      'N° Saisie': numero, 'Code de suivi': code, 'Date de réception': reception, "Date d'envoi": new Date(),
      'Agent': safe(d.agent), 'Matricule': safe(d.matricule),
      'Nom du réclamant': safe(d.nom), 'CIL': safe(d.cil), 'Téléphone': safe(d.tel), 'E-mail': safe(d.email),
      'Adresse': safe(d.adresse), 'Quartier': safe(d.quartier), 'Commune': safe(d.ville),
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
      const file = getFolder().createFile(Utilities.newBlob(full, 'image/jpeg', numero + '.jpg'));
      sh.getRange(row, c['Lien photo']).setRichTextValue(link('Voir la photo', file.getUrl()));
    }
    // Aperçu : petite image posée dans la ligne, visible aussi dans Excel.
    const small = jpeg(d.vignette);
    if (small && small.length < 1.5 * 1024 * 1024) {
      try {
        placeThumb(sh, row, c['Aperçu photo'], Utilities.newBlob(small, 'image/jpeg', numero + '-apercu.jpg'));
      } catch (e) {
        sh.getRange(row, c['Aperçu photo']).setValue('aperçu indisponible');
      }
    }
    return { ok: true, numero: numero, code: code };
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
    const num = String(t.numReclamation || '').trim();
    if (num && c['N° de réclamation']) sh.getRange(row, c['N° de réclamation']).setValue(safe(num));

    const compte = Session.getActiveUser().getEmail();
    const photos = (Array.isArray(t.photos) ? t.photos : []).slice(0, 5).map(jpeg).filter(b => b && b.length < 8 * 1024 * 1024);
    const values = {
      'Date réception': parseDateTime(t.dateReception) || '', 'Métier': 'Ré-Use', 'Nature réclamation': nature,
      'Objet préliminaire': safe(t.objet), 'Type client': TYPES_CLIENT.indexOf(t.typeClient) >= 0 ? t.typeClient : '',
      'Nom du client': safe(t.nomClient), 'CIL': safe(t.cil), 'Préfecture': safe(t.prefecture), 'Commune': safe(t.commune),
      'Quartier': safe(t.quartier), 'Sous-quartier': safe(t.sousQuartier), 'Rue': safe(t.rue),
      'Adresse de la réclamation': safe(t.adresse), 'Téléphone portable 1': safe(t.telPortable), 'Téléphone fixe': safe(t.telFixe),
      'Nombre de rappels': Math.max(0, parseInt(t.rappels, 10) || 0), 'Nombre de photos': photos.length,
      'Utilisateur saisie': safe([String(t.agent).trim(), String(t.matricule).trim(), compte ? '(' + compte + ')' : ''].filter(String).join(' ')),
      "Date d'envoi": new Date()
    };
    Object.keys(values).forEach(h => { if (c[T + h]) sh.getRange(row, c[T + h]).setValue(values[h]); });
    if (c['Etat de la réclamation']) sh.getRange(row, c['Etat de la réclamation']).setValue(t.etat);

    // Photos : rangées dans Google Drive, un lien par photo dans la même case.
    if (photos.length && c[T + 'Photos']) {
      const urls = photos.map((b, i) => getFolder().createFile(Utilities.newBlob(b, 'image/jpeg', code + '-traitement-' + (i + 1) + '.jpg')).getUrl());
      const labels = urls.map((u, i) => 'Photo ' + (i + 1));
      const text = labels.join('  ');
      const rt = SpreadsheetApp.newRichTextValue().setText(text);
      let pos = 0;
      labels.forEach((l, i) => { rt.setLinkUrl(pos, pos + l.length, urls[i]); pos += l.length + 2; });
      sh.getRange(row, c[T + 'Photos']).setRichTextValue(rt.build());
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
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    styleHeader(sh.getRange(1, 1, 1, HEADERS.length));
    sh.setFrozenRows(1);
    HEADERS.forEach((h, i) => formatColumn(sh, i + 1, h));
  } else {
    mettreAJourColonnes(sh);
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
function mettreAJourColonnes(sh) {
  let head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const rename = (from, to) => {
    const i = head.indexOf(from);
    if (i >= 0 && head.indexOf(to) < 0) { sh.getRange(1, i + 1).setValue(to); head[i] = to; return i + 1; }
    return 0;
  };
  rename('N° réclamation', 'N° Saisie');
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
  HEADERS.forEach((h, i) => {
    if (head.indexOf(h) >= 0) return;
    // Place la colonne juste après la précédente de la liste (ou en fin de tableau).
    const prev = i > 0 ? head.indexOf(HEADERS[i - 1]) : -1;
    const at = prev >= 0 ? prev + 2 : head.length + 1;
    if (at <= head.length) sh.insertColumnBefore(at); else sh.insertColumnAfter(head.length);
    sh.getRange(1, at, sh.getMaxRows(), 1).clearDataValidations().setBackground(null).setNumberFormat('General');
    sh.showColumns(at);
    sh.getRange(1, at).setValue(h);
    styleHeader(sh.getRange(1, at));
    formatColumn(sh, at, h);
    head.splice(at - 1, 0, h);
  });
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
