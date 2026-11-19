/**
 * Réclamations Ré-Use — version compte Veolia.
 * Le formulaire est servi par ce script ; seuls les comptes Veolia connectés peuvent l'ouvrir.
 *  - Onglet « Nouvelle réclamation » : tous les agents. Chaque réclamation est ajoutée au tableau,
 *    avec l'adresse Veolia de l'agent qui l'a envoyée.
 *  - Onglet « Traitement » : recherche, consultation et traitement des réclamations. Visible
 *    uniquement par la propriétaire du tableau et les adresses listées dans l'onglet
 *    « Accès traitement » du tableau.
 *
 * Installation : Extensions > Apps Script, coller ce code, enregistrer,
 * puis Déployer > Nouveau déploiement > Application Web
 * (Exécuter en tant que : Moi ; Qui a accès : Anyone within Veolia).
 * Le lien de l'application Web est le lien du formulaire.
 */

const VERSION = 'Veolia 5';
const SHEET_NAME = 'Réclamations';
const ACCESS_SHEET = 'Accès traitement';
const FOLDER_NAME = 'Photos réclamations Ré-Use';
const DATE_FMT = 'dd/MM/yyyy HH:mm';

const HEADERS = ['N° Saisie', 'N° de réclamation', 'Date de réception', "Date d'envoi", 'Agent', 'Matricule',
  'Type client', 'Nom du réclamant', 'CIL', 'Téléphone', 'E-mail',
  'Adresse', 'Quartier', 'Commune', 'Latitude', 'Longitude', 'Précision (m)', 'Carte',
  'Nature', 'Localisation fuite', 'Importance fuite', 'Détail nature', 'Depuis', 'Période', 'Date/heure souhaitée',
  'Observations', 'Aperçu photo', 'Lien photo',
  'Etat de la réclamation', 'Nombre de rappels', 'Date intervention', 'Intervenant', 'Action réalisée',
  'Photo après travaux', 'Traité par', 'Dernière mise à jour', 'ID envoi'];
const TEXT_COLS = ['N° de réclamation', 'Matricule', 'CIL', 'Téléphone', 'E-mail'];
const DATE_COLS = ['Date de réception', "Date d'envoi", 'Date intervention', 'Dernière mise à jour'];
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
  return HtmlService.createHtmlOutput(FORM_HTML)
    .setTitle('Réclamations Ré-Use')
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
    const hasPos = d.lat !== '' && d.lat != null && d.lng !== '' && d.lng != null && isFinite(d.lat) && isFinite(d.lng);

    const values = {
      'N° Saisie': numero, 'Date de réception': reception, "Date d'envoi": new Date(),
      'Agent': safe(d.agent), 'Matricule': safe(d.matricule),
      'Type client': TYPES_CLIENT.indexOf(d.typeClient) >= 0 ? d.typeClient : '',
      'Nom du réclamant': safe(d.nom), 'CIL': safe(d.cil), 'Téléphone': safe(d.tel), 'E-mail': safe(d.email),
      'Adresse': safe(d.adresse), 'Quartier': safe(d.quartier), 'Commune': safe(d.ville),
      'Latitude': hasPos ? Number(d.lat) : '', 'Longitude': hasPos ? Number(d.lng) : '',
      'Précision (m)': hasPos && isFinite(d.precision) ? Number(d.precision) : '',
      'Nature': d.nature, 'Localisation fuite': safe(d.localisation), 'Importance fuite': safe(d.importance),
      'Détail nature': safe(d.detail), 'Depuis': safe(d.depuis), 'Période': safe(d.periode),
      'Date/heure souhaitée': safe(d.souhaite), 'Observations': safe(d.obs),
      'Etat de la réclamation': ETATS[0], 'Nombre de rappels': 0, 'ID envoi': safe(d.clientId)
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
    return { ok: true, numero: numero };
  } catch (err) {
    return { ok: false, error: 'erreur du tableau : ' + err.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/* ===================== Traitement ===================== */

// Indique au formulaire si la personne connectée peut voir l'onglet Traitement.
function moi() {
  const email = Session.getActiveUser().getEmail();
  return { email: email, traitement: peutTraiter(email), etats: ETATS, typesClient: TYPES_CLIENT };
}

function peutTraiter(email) {
  email = String(email || '').trim().toLowerCase();
  if (!email) return false;
  if (email === String(Session.getEffectiveUser().getEmail()).toLowerCase()) return true;
  const sh = getAccessSheet();
  if (sh.getLastRow() < 2) return false;
  return sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues()
    .some(r => String(r[0]).trim().toLowerCase() === email);
}

function exigerTraitement() {
  if (!peutTraiter(Session.getActiveUser().getEmail())) throw new Error("accès au traitement non autorisé pour ce compte");
}

// Recherche, avec les critères de l'écran « Recherche d'une réclamation » de WebSIGID.
function rechercher(f) {
  try {
    exigerTraitement();
    f = f || {};
    const sh = getSheet(), c = cols(sh), last = sh.getLastRow();
    if (last < 2) return { ok: true, total: 0, items: [] };
    const data = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
    const g = (r, h) => c[h] ? r[c[h] - 1] : '';
    const has = (v, q) => !q || norm(v).indexOf(norm(q)) >= 0;
    const digits = s => String(s || '').replace(/\D/g, '');
    const du = parseDateTime(f.du), au = parseDateTime(f.au);
    if (au) au.setHours(23, 59, 59, 999);
    const items = data.filter(r => {
      if (!g(r, 'N° Saisie')) return false;
      if (f.num && !(has(g(r, 'N° Saisie'), f.num) || has(g(r, 'N° de réclamation'), f.num))) return false;
      const rec = toDate(g(r, 'Date de réception'));
      if (du && (!rec || rec < du)) return false;
      if (au && (!rec || rec > au)) return false;
      const etat = etatDe(g(r, 'Etat de la réclamation'));
      if (f.etat === 'non-resolues' && etat === 'Résolue Définitivement') return false;
      if (f.etat && f.etat !== 'non-resolues' && etat !== f.etat) return false;
      if (f.nature && g(r, 'Nature') !== f.nature) return false;
      if (f.typeClient && g(r, 'Type client') !== f.typeClient) return false;
      if (!has(g(r, 'Nom du réclamant'), f.nom)) return false;
      if (f.agent && !(has(g(r, 'Agent'), f.agent) || has(g(r, 'Matricule'), f.agent))) return false;
      if (!has(g(r, 'CIL'), f.cil) || !has(g(r, 'Commune'), f.commune) || !has(g(r, 'Quartier'), f.quartier)) return false;
      if (!has(g(r, 'Adresse'), f.adresse)) return false;
      if (f.tel && digits(g(r, 'Téléphone')).indexOf(digits(f.tel)) < 0) return false;
      return true;
    }).sort((a, b) => (toDate(g(b, 'Date de réception')) || 0) - (toDate(g(a, 'Date de réception')) || 0));
    return {
      ok: true, total: items.length,
      items: items.slice(0, 200).map(r => ({
        saisie: String(g(r, 'N° Saisie')), reclamation: String(g(r, 'N° de réclamation')),
        reception: fmt(g(r, 'Date de réception')), nature: String(g(r, 'Nature')),
        detail: String(g(r, 'Détail nature') || g(r, 'Localisation fuite') || ''),
        adresse: [g(r, 'Adresse'), g(r, 'Quartier'), g(r, 'Commune')].filter(String).join(', '),
        nom: String(g(r, 'Nom du réclamant')), etat: etatDe(g(r, 'Etat de la réclamation')),
        rappels: Number(g(r, 'Nombre de rappels')) || 0
      }))
    };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

// Fiche complète d'une réclamation.
function fiche(saisie) {
  try {
    exigerTraitement();
    const sh = getSheet(), c = cols(sh), row = trouverLigne(sh, c, saisie);
    if (!row) return { ok: false, error: 'réclamation introuvable' };
    const head = Object.keys(c).sort((a, b) => c[a] - c[b]);
    const vals = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
    const rich = sh.getRange(row, 1, 1, sh.getLastColumn()).getRichTextValues()[0];
    const url = h => { const rt = c[h] && rich[c[h] - 1]; return rt ? (rt.getLinkUrl() || '') : ''; };
    const skip = ['ID envoi', 'Aperçu photo', 'Carte', 'Lien photo', 'Photo après travaux'];
    const champs = head.filter(h => skip.indexOf(h) < 0)
      .map(h => [h, h === 'Etat de la réclamation' ? etatDe(vals[c[h] - 1]) : fmt(vals[c[h] - 1])])
      .filter(p => p[1] !== '');
    const interv = toDate(vals[c['Date intervention'] - 1]);
    return {
      ok: true, saisie: String(saisie), champs: champs,
      liens: { carte: url('Carte'), photo: url('Lien photo'), photoApres: url('Photo après travaux') },
      traitement: {
        etat: etatDe(vals[c['Etat de la réclamation'] - 1]),
        reclamation: String(vals[c['N° de réclamation'] - 1] || ''),
        intervention: interv ? Utilities.formatDate(interv, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm") : '',
        intervenant: String(vals[c['Intervenant'] - 1] || ''),
        action: String(vals[c['Action réalisée'] - 1] || ''),
        rappels: Number(vals[c['Nombre de rappels'] - 1]) || 0
      }
    };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

// Enregistre le traitement d'une réclamation.
function traiter(t) {
  const lock = LockService.getScriptLock();
  try {
    exigerTraitement();
    lock.waitLock(25000);
    t = t || {};
    if (ETATS.indexOf(t.etat) < 0) return { ok: false, error: 'état de la réclamation invalide' };
    const sh = getSheet(), c = cols(sh), row = trouverLigne(sh, c, t.saisie);
    if (!row) return { ok: false, error: 'réclamation introuvable' };
    const set = (h, v) => { if (c[h]) sh.getRange(row, c[h]).setValue(v); };
    set('Etat de la réclamation', t.etat);
    set('N° de réclamation', safe(t.reclamation));
    const interv = parseDateTime(t.intervention);
    set('Date intervention', interv || '');
    set('Intervenant', safe(t.intervenant));
    set('Action réalisée', safe(t.action));
    if (t.rappel) {
      const n = Number(sh.getRange(row, c['Nombre de rappels']).getValue()) || 0;
      set('Nombre de rappels', n + 1);
    }
    const apres = jpeg(t.photoApres);
    if (apres && apres.length < 8 * 1024 * 1024) {
      const saisie = sh.getRange(row, c['N° Saisie']).getValue();
      const file = getFolder().createFile(Utilities.newBlob(apres, 'image/jpeg', saisie + '-apres.jpg'));
      sh.getRange(row, c['Photo après travaux']).setRichTextValue(link('Voir la photo après', file.getUrl()));
    }
    set('Traité par', Session.getActiveUser().getEmail());
    set('Dernière mise à jour', new Date());
    return { ok: true };
  } catch (err) {
    return { ok: false, error: err.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function trouverLigne(sh, c, saisie) {
  if (!saisie || sh.getLastRow() < 2) return 0;
  const found = sh.getRange(2, c['N° Saisie'], sh.getLastRow() - 1, 1)
    .createTextFinder(String(saisie)).matchEntireCell(true).findNext();
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

function getAccessSheet() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(ACCESS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(ACCESS_SHEET);
    sh.getRange(1, 1).setValue('Adresses Veolia autorisées au traitement (une par ligne)');
    styleHeader(sh.getRange(1, 1));
    sh.setColumnWidth(1, 420);
    sh.getRange(2, 1).setValue(Session.getEffectiveUser().getEmail());
  }
  return sh;
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
  getAccessSheet();
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
const FORM_HTML = "<!doctype html>\n<html lang=\"fr\">\n<head>\n<meta charset=\"utf-8\">\n<meta name=\"viewport\" content=\"width=device-width,initial-scale=1,viewport-fit=cover\">\n<meta name=\"theme-color\" content=\"#0f6e86\">\n<meta name=\"robots\" content=\"noindex\">\n<title>Réclamations Ré-Use</title>\n<style>\n:root{\n  --bg:#f3f6f7; --surface:#ffffff; --line:#d6e0e3; --fg:#14262c; --muted:#5b7078;\n  --accent:#0f6e86; --accent-ink:#ffffff; --accent-soft:#e1f0f4;\n  --ok:#2f7d4f; --ok-soft:#e3f3e9; --warn:#a76a00; --warn-soft:#fbf0d9; --bad:#b3332f; --bad-soft:#fbe5e4;\n  --f-body:system-ui,-apple-system,\"Segoe UI\",Roboto,Arial,sans-serif;\n  --f-display:\"Arial Narrow\",\"Roboto Condensed\",system-ui,sans-serif;\n  --f-mono:ui-monospace,Consolas,\"Roboto Mono\",monospace;\n  color-scheme:light;\n}\n@media (prefers-color-scheme: dark){:root{\n  --bg:#0e1a1e; --surface:#15252a; --line:#273c43; --fg:#e3eef1; --muted:#93aab1;\n  --accent:#4fb6cf; --accent-ink:#06171c; --accent-soft:#163a44;\n  --ok:#6cc792; --ok-soft:#173326; --warn:#e6b45a; --warn-soft:#3a2d12; --bad:#ef8a85; --bad-soft:#3d1c1b;\n  color-scheme:dark}}\n*{box-sizing:border-box}\n[hidden]{display:none!important}\nhtml,body{margin:0}\nbody{background:var(--bg);color:var(--fg);font-family:var(--f-body);font-size:15px;line-height:1.45}\n.wrap{max-width:560px;margin:0 auto;padding:16px max(14px,env(safe-area-inset-right)) 40px max(14px,env(safe-area-inset-left));display:flex;flex-direction:column;gap:14px}\nheader{border-bottom:2px solid var(--fg);padding-bottom:10px}\nh1{font-family:var(--f-display);font-weight:700;font-size:24px;margin:0}\nh1 span{color:var(--accent)}\n.sub{color:var(--muted);margin:2px 0 0;font-size:13px}\n.panel{background:var(--surface);border:1px solid var(--line);border-radius:6px}\nform{padding:14px;display:flex;flex-direction:column;gap:16px}\nfieldset{border:0;padding:0;margin:0;display:flex;flex-direction:column;gap:10px;min-width:0}\nlegend,.lg{font-size:11px;text-transform:uppercase;letter-spacing:.08em;color:var(--accent);font-weight:700;padding:0;margin-bottom:6px}\n.row{display:grid;grid-template-columns:1fr 1fr;gap:10px}\n@media (max-width:460px){.row{grid-template-columns:1fr}}\nlabel{display:block;font-weight:600;font-size:13px}\nlabel>input,label>select,label>textarea{display:block;margin-top:4px}\n.req{color:var(--bad)}\ninput,select,textarea{font:inherit;font-size:16px;color:var(--fg);background:var(--bg);border:1px solid var(--line);border-radius:4px;padding:9px 10px;width:100%}\ntextarea{resize:vertical;min-height:64px}\ninput:focus,select:focus,textarea:focus,button:focus-visible{outline:2px solid var(--accent);outline-offset:1px}\n.radios{display:flex;flex-direction:column;gap:8px}\n.radios label{display:flex;align-items:center;gap:8px;font-weight:400;font-size:14px}\n.radios input{width:auto;margin:0}\n.fuite{background:var(--accent-soft);border-radius:4px;padding:12px}\n.geo{border:1px dashed var(--accent);border-radius:6px;padding:12px;display:flex;flex-direction:column;gap:8px}\n.geo-btn{display:flex;align-items:center;justify-content:center;gap:8px;font-size:16px;padding:12px}\n.geo-btn svg{width:20px;height:20px}\n.geo-out{font-family:var(--f-mono);font-size:13px}\n.geo-out a{color:var(--accent)}\n.det{color:var(--muted);font-size:12px;margin:0}\nbutton{font:inherit;font-weight:600;border-radius:4px;padding:10px 14px;cursor:pointer;border:1px solid var(--line);background:var(--surface);color:var(--fg)}\nbutton.primary{background:var(--accent);border-color:var(--accent);color:var(--accent-ink)}\nbutton:disabled{opacity:.55;cursor:wait}\n.actions{display:flex;gap:8px;flex-wrap:wrap}\n.actions .primary{flex:1 1 auto;font-size:16px;padding:12px}\n#preview{max-width:100%;max-height:200px;object-fit:cover;border-radius:4px;border:1px solid var(--line)}\n.msg{font-size:14px;margin:0;padding:10px 12px;border-radius:4px}\n.msg:empty{display:none}\n.msg.ok{color:var(--ok);background:var(--ok-soft)} .msg.err{color:var(--bad);background:var(--bad-soft)} .msg.info{color:var(--warn);background:var(--warn-soft)}\n.agent{display:flex;gap:8px;align-items:center;padding:10px 14px;background:var(--accent-soft);border-bottom:1px solid var(--line);font-size:13px}\n.agent label{white-space:nowrap}\n.agent input{padding:6px 8px;font-size:15px;margin:0}\n.agent-id{display:grid;grid-template-columns:1fr 1fr;gap:10px;align-items:start}\n.agent-id label{white-space:normal}\n.agent-id input{margin-top:4px;background:var(--surface)}\n.checks{display:flex;flex-direction:column;gap:8px}\n.checks label{display:flex;align-items:center;gap:8px;font-weight:400;font-size:14px}\n.checks input{width:auto;margin:0}\n.tabs{display:flex;gap:6px;border-bottom:2px solid var(--line)}\n.tab{border:0;border-bottom:3px solid transparent;border-radius:0;background:transparent;color:var(--muted);padding:10px 12px;margin-bottom:-2px}\n.tab.on{color:var(--accent);border-bottom-color:var(--accent)}\nbody.large .wrap{max-width:900px}\n.panel-h{display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;padding:12px 14px;border-bottom:1px solid var(--line)}\n.panel-h h2{font-size:16px;margin:0}\n.metier{font-size:12px;color:var(--muted)}\n.pad{padding:14px;display:flex;flex-direction:column;gap:12px}\ndetails{border:1px dashed var(--line);border-radius:4px;padding:8px 10px;display:flex;flex-direction:column;gap:10px}\ndetails[open]>summary{margin-bottom:8px}\nsummary{cursor:pointer;font-weight:600;color:var(--accent)}\ndetails>.row,details>label{margin-top:8px}\n.results{display:flex;flex-direction:column}\n.res{display:grid;grid-template-columns:1fr auto;gap:4px 12px;text-align:left;border:0;border-top:1px solid var(--line);border-radius:0;padding:12px 14px;background:var(--surface);width:100%;font-weight:400}\n.res:hover{background:var(--accent-soft)}\n.res b{font-family:var(--f-mono);font-size:13px}\n.res .d{color:var(--muted);font-size:12px;grid-column:1/-1}\n.etat{font-size:11px;font-weight:700;padding:2px 8px;border-radius:99px;white-space:nowrap;align-self:start}\n.e0{background:var(--bad-soft);color:var(--bad)} .e1{background:var(--warn-soft);color:var(--warn)} .e2{background:var(--accent-soft);color:var(--accent)} .e3{background:var(--ok-soft);color:var(--ok)}\n.count{padding:8px 14px;font-size:12px;color:var(--muted)}\n.champs{display:grid;grid-template-columns:minmax(120px,40%) 1fr;gap:6px 12px;margin:0;font-size:14px}\n.champs dt{color:var(--muted)} .champs dd{margin:0;overflow-wrap:anywhere}\n.liens{display:flex;gap:12px;flex-wrap:wrap}\n.liens a{color:var(--accent);font-weight:600}\n.link-btn{border:0;background:transparent;color:var(--accent);padding:4px 0}\nlabel.inline{display:flex;align-items:center;gap:8px;font-weight:400;font-size:14px}\nlabel.inline input{width:auto;margin:0}\n#tPreview{max-width:100%;max-height:200px;object-fit:cover;border-radius:4px;border:1px solid var(--line)}\n.privacy{font-size:12px;color:var(--muted);text-align:center;margin:0}\n</style>\n</head>\n<body>\n<div class=\"wrap\">\n  <header>\n    <h1>Réclamations <span>Ré-Use</span></h1>\n    <p class=\"sub\">Saisie terrain d'une réclamation client</p>\n  </header>\n\n  <nav class=\"tabs\" id=\"tabs\" hidden>\n    <button type=\"button\" class=\"tab on\" data-tab=\"saisie\">Nouvelle réclamation</button>\n    <button type=\"button\" class=\"tab\" data-tab=\"traitement\">Traitement</button>\n  </nav>\n\n  <div id=\"tabSaisie\">\n  <section class=\"panel\">\n    <div class=\"agent agent-id\">\n      <label for=\"agent\">Agent <span class=\"req\">*</span><input id=\"agent\" autocomplete=\"off\" placeholder=\"Votre nom\"></label>\n      <label for=\"matricule\">Matricule <span class=\"req\">*</span><input id=\"matricule\" autocomplete=\"off\" inputmode=\"text\" placeholder=\"Votre matricule\"></label>\n    </div>\n    <form id=\"form\" novalidate>\n      <fieldset>\n        <legend>Réclamant</legend>\n        <label for=\"nom\">Nom et prénom <span class=\"req\">*</span><input id=\"nom\" autocomplete=\"off\" required></label>\n        <div class=\"row\">\n          <label for=\"typeClient\">Type client\n            <select id=\"typeClient\">\n              <option>Particulier</option><option>Autorité</option><option>Communes</option><option>Elus</option>\n              <option>Résidences Royales</option><option>Industriels</option><option>V.I.P</option><option>GC-Entreprises</option>\n              <option>Planifié</option><option>GC-Administrations</option><option>Numéro Vert</option>\n            </select>\n          </label>\n          <label for=\"cil\">CIL<input id=\"cil\" autocomplete=\"off\" inputmode=\"numeric\" placeholder=\"facultatif\"></label>\n        </div>\n        <div class=\"row\">\n          <label for=\"tel\">Téléphone<input id=\"tel\" type=\"tel\" inputmode=\"tel\" autocomplete=\"off\" placeholder=\"facultatif\"></label>\n          <label for=\"email\">E-mail<input id=\"email\" type=\"email\" autocomplete=\"off\" placeholder=\"facultatif\"></label>\n        </div>\n      </fieldset>\n\n      <fieldset>\n        <legend>Lieu de la réclamation</legend>\n        <div class=\"geo\">\n          <button type=\"button\" class=\"primary geo-btn\" id=\"btnGeo\">\n            <svg viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\" stroke-linecap=\"round\" stroke-linejoin=\"round\" aria-hidden=\"true\"><path d=\"M12 21s-7-6.2-7-11.5A7 7 0 0 1 19 9.5C19 14.8 12 21 12 21z\"/><circle cx=\"12\" cy=\"9.5\" r=\"2.5\"/></svg>\n            Enregistrer ma position\n          </button>\n          <div class=\"geo-out\" id=\"geoOut\">Aucune position enregistrée.</div>\n          <p class=\"det\" id=\"geoNote\">Placez-vous sur le lieu de la fuite puis appuyez sur le bouton.</p>\n        </div>\n        <label for=\"adresse\">Adresse <span class=\"req\">*</span><input id=\"adresse\" autocomplete=\"off\" placeholder=\"N°, rue, lotissement\" required></label>\n        <div class=\"row\">\n          <label for=\"quartier\">Quartier / secteur<input id=\"quartier\" autocomplete=\"off\"></label>\n          <label for=\"ville\">Commune<input id=\"ville\" list=\"villes\" autocomplete=\"off\"></label>\n        </div>\n        <datalist id=\"villes\"><option>Rabat</option><option>Salé</option><option>Témara</option><option>Skhirat</option><option>Harhoura</option><option>Aïn Atiq</option></datalist>\n      </fieldset>\n\n      <fieldset>\n        <legend>Réclamation</legend>\n        <div class=\"row\">\n          <label for=\"recDate\">Date de réception <span class=\"req\">*</span><input id=\"recDate\" type=\"date\" required></label>\n          <label for=\"recHeure\">Heure de réception<input id=\"recHeure\" type=\"time\"></label>\n        </div>\n        <label for=\"nature\">Nature <span class=\"req\">*</span>\n          <select id=\"nature\" required>\n            <option value=\"\"></option>\n            <option>Réfection</option>\n            <option>Fuite d'eau</option>\n            <option>Qualité d'eau</option>\n            <option>Problème de Pression</option>\n            <option>Demande Coupure</option>\n            <option>Demande Rebranchement</option>\n            <option>DFO</option>\n            <option>Manque d'eau</option>\n            <option>Appel Client pour Information</option>\n          </select>\n        </label>\n        <div class=\"fuite\" id=\"fuiteBox\" hidden>\n          <div class=\"row\">\n            <div>\n              <div class=\"lg\">Localisation de la fuite <span class=\"req\">*</span></div>\n              <div class=\"radios\">\n                <label><input type=\"radio\" name=\"loc\" value=\"Fuite dans la Rue\"> Fuite dans la Rue</label>\n                <label><input type=\"radio\" name=\"loc\" value=\"Compteur Sauté\"> Compteur Sauté</label>\n              </div>\n            </div>\n            <div>\n              <div class=\"lg\">Importance de la fuite</div>\n              <div class=\"radios\">\n                <label><input type=\"radio\" name=\"imp\" value=\"Faible (goutte à goutte)\"> Faible (goutte à goutte)</label>\n                <label><input type=\"radio\" name=\"imp\" value=\"Moyenne (filet d'eau)\"> Moyenne (filet d'eau)</label>\n                <label><input type=\"radio\" name=\"imp\" value=\"Importante (jet / inondation)\"> Importante (jet / inondation)</label>\n              </div>\n            </div>\n          </div>\n        </div>\n        <div class=\"fuite\" id=\"refectionBox\" hidden>\n          <div class=\"lg\">Réfection <span class=\"req\">*</span></div>\n          <div class=\"radios\">\n            <label><input type=\"radio\" name=\"refection\" value=\"Non réalisée\"> Non réalisée</label>\n            <label><input type=\"radio\" name=\"refection\" value=\"A reprendre\"> A reprendre</label>\n            <label><input type=\"radio\" name=\"refection\" value=\"Chantier non Nettoyé\"> Chantier non Nettoyé</label>\n          </div>\n        </div>\n        <div class=\"fuite\" id=\"qualiteBox\" hidden>\n          <div class=\"lg\">Nature défaut qualité <span class=\"req\">*</span></div>\n          <div class=\"checks\">\n            <label><input type=\"checkbox\" name=\"qualite\" value=\"Couleur / Turbidité\"> Couleur / Turbidité</label>\n            <label><input type=\"checkbox\" name=\"qualite\" value=\"Goût / Odeur\"> Goût / Odeur</label>\n          </div>\n          <div class=\"row\" style=\"margin-top:10px\">\n            <label for=\"qualiteDepuis\">Défaut qualité depuis<input id=\"qualiteDepuis\" autocomplete=\"off\" placeholder=\"ex. 3 jours, le 15/11\"></label>\n            <label for=\"qualitePeriode\">Période défaut qualité<input id=\"qualitePeriode\" autocomplete=\"off\" placeholder=\"ex. le matin, en continu\"></label>\n          </div>\n        </div>\n        <div class=\"fuite\" id=\"pressionBox\" hidden>\n          <div class=\"lg\">Nature défaut pression <span class=\"req\">*</span></div>\n          <div class=\"radios\">\n            <label><input type=\"radio\" name=\"pression\" value=\"Manque de pression\"> Manque de pression</label>\n            <label><input type=\"radio\" name=\"pression\" value=\"Excès de pression\"> Excès de pression</label>\n            <label><input type=\"radio\" name=\"pression\" value=\"Pas d'Eau\"> Pas d'Eau</label>\n          </div>\n          <div class=\"row\" style=\"margin-top:10px\">\n            <label for=\"pressionDepuis\">Problème de pression depuis<input id=\"pressionDepuis\" autocomplete=\"off\" placeholder=\"ex. 3 jours, le 15/11\"></label>\n            <label for=\"pressionPeriode\">Période problème pression<input id=\"pressionPeriode\" autocomplete=\"off\" placeholder=\"ex. le soir, la nuit\"></label>\n          </div>\n        </div>\n        <div class=\"fuite\" id=\"coupureBox\" hidden>\n          <div class=\"lg\">Demande de coupure</div>\n          <div class=\"row\">\n            <label for=\"coupureDate\">Date souhaitée<input id=\"coupureDate\" type=\"date\"></label>\n            <label for=\"coupureHeure\">Heure souhaitée<input id=\"coupureHeure\" type=\"time\"></label>\n          </div>\n        </div>\n        <div class=\"fuite\" id=\"rebranchementBox\" hidden>\n          <div class=\"lg\">Demande de rebranchement</div>\n          <div class=\"row\">\n            <label for=\"rebranchementDate\">Date souhaitée<input id=\"rebranchementDate\" type=\"date\"></label>\n            <label for=\"rebranchementHeure\">Heure souhaitée<input id=\"rebranchementHeure\" type=\"time\"></label>\n          </div>\n        </div>\n        <div class=\"fuite\" id=\"dfoBox\" hidden>\n          <div class=\"lg\">DFO à quel niveau <span class=\"req\">*</span></div>\n          <div class=\"radios\">\n            <label><input type=\"radio\" name=\"dfo\" value=\"Conduite\"> Conduite</label>\n            <label><input type=\"radio\" name=\"dfo\" value=\"Branchement\"> Branchement</label>\n          </div>\n        </div>\n        <div class=\"fuite\" id=\"manqueBox\" hidden>\n          <div class=\"lg\">Manque d'eau depuis</div>\n          <div class=\"row\">\n            <label for=\"manqueDate\">Date manque d'eau<input id=\"manqueDate\" type=\"date\"></label>\n            <label for=\"manqueHeure\">Heure manque d'eau<input id=\"manqueHeure\" type=\"time\"></label>\n          </div>\n        </div>\n        <div class=\"fuite\" id=\"appelBox\" hidden>\n          <div class=\"lg\">Raison de l'appel <span class=\"req\">*</span></div>\n          <div class=\"checks\">\n            <label><input type=\"checkbox\" name=\"appel\" value=\"Client non avisé suite coupure programmée\"> Client non avisé suite coupure programmée</label>\n            <label><input type=\"checkbox\" name=\"appel\" value=\"Problème commercial\"> Problème commercial</label>\n          </div>\n        </div>\n        <label for=\"photo\">Photo de la fuite / du lieu<input id=\"photo\" type=\"file\" accept=\"image/*\" capture=\"environment\"></label>\n        <img id=\"preview\" alt=\"Aperçu de la photo\" hidden>\n        <label for=\"obs\">Observations<textarea id=\"obs\" placeholder=\"Détails donnés par le client\"></textarea></label>\n      </fieldset>\n\n      <p class=\"msg\" id=\"msg\" role=\"status\"></p>\n      <div class=\"actions\">\n        <button type=\"submit\" class=\"primary\" id=\"btnSave\">Envoyer la réclamation</button>\n        <button type=\"reset\">Effacer</button>\n      </div>\n    </form>\n  </section>\n  <p class=\"privacy\">Les réclamations envoyées ne sont pas visibles sur ce formulaire : elles sont transmises directement au responsable, avec votre compte Veolia.</p>\n  </div>\n\n  <div id=\"tabTraitement\" hidden>\n    <section class=\"panel\">\n      <div class=\"panel-h\"><h2>Recherche d'une réclamation</h2><span class=\"metier\">Métier : Ré-Use</span></div>\n      <form id=\"search\" class=\"pad\" novalidate>\n        <div class=\"row\">\n          <label for=\"sNum\">N° Saisie ou N° réclamation<input id=\"sNum\" autocomplete=\"off\" placeholder=\"ex. RU-2026-00012\"></label>\n          <label for=\"sEtat\">Etat de la réclamation\n            <select id=\"sEtat\">\n              <option value=\"non-resolues\">Toutes sauf résolues définitivement</option>\n              <option value=\"\">Tous les états</option>\n              <option>Non encore Transmise</option><option>Transmise mais non encore résolue</option>\n              <option>Résolue Provisoirement</option><option>Résolue Définitivement</option>\n            </select>\n          </label>\n        </div>\n        <div class=\"row\">\n          <label for=\"sDu\">Date réception du<input id=\"sDu\" type=\"date\"></label>\n          <label for=\"sAu\">au<input id=\"sAu\" type=\"date\"></label>\n        </div>\n        <div class=\"row\">\n          <label for=\"sNature\">Nature réclamation<select id=\"sNature\"><option value=\"\">Toutes</option></select></label>\n          <label for=\"sType\">Type client<select id=\"sType\"><option value=\"\">Tous</option></select></label>\n        </div>\n        <details>\n          <summary>Plus de critères</summary>\n          <div class=\"row\">\n            <label for=\"sNom\">Nom du client<input id=\"sNom\" autocomplete=\"off\"></label>\n            <label for=\"sAgent\">Utilisateur saisie (nom ou matricule)<input id=\"sAgent\" autocomplete=\"off\"></label>\n          </div>\n          <div class=\"row\">\n            <label for=\"sCil\">CIL<input id=\"sCil\" autocomplete=\"off\" inputmode=\"numeric\"></label>\n            <label for=\"sTel\">Téléphone<input id=\"sTel\" type=\"tel\" autocomplete=\"off\"></label>\n          </div>\n          <div class=\"row\">\n            <label for=\"sCommune\">Commune<input id=\"sCommune\" autocomplete=\"off\"></label>\n            <label for=\"sQuartier\">Quartier<input id=\"sQuartier\" autocomplete=\"off\"></label>\n          </div>\n          <label for=\"sAdresse\">Adresse de la réclamation<input id=\"sAdresse\" autocomplete=\"off\"></label>\n        </details>\n        <div class=\"actions\">\n          <button type=\"submit\" class=\"primary\" id=\"btnSearch\">Rechercher</button>\n          <button type=\"reset\">Effacer les critères</button>\n        </div>\n        <p class=\"msg\" id=\"sMsg\" role=\"status\"></p>\n      </form>\n      <div id=\"results\" class=\"results\"></div>\n    </section>\n\n    <section class=\"panel\" id=\"fichePanel\" hidden>\n      <div class=\"panel-h\"><h2 id=\"ficheTitre\">Fiche</h2><button type=\"button\" class=\"link-btn\" id=\"btnBack\">← Retour à la liste</button></div>\n      <div class=\"pad\">\n        <dl id=\"ficheChamps\" class=\"champs\"></dl>\n        <div class=\"liens\" id=\"ficheLiens\"></div>\n      </div>\n      <form id=\"traitement\" class=\"pad\" novalidate>\n        <fieldset>\n          <legend>Traitement</legend>\n          <label for=\"tEtat\">Etat de la réclamation <span class=\"req\">*</span><select id=\"tEtat\"></select></label>\n          <label for=\"tReclamation\">N° de réclamation (WebSIGID)<input id=\"tReclamation\" autocomplete=\"off\"></label>\n          <div class=\"row\">\n            <label for=\"tDate\">Date d'intervention<input id=\"tDate\" type=\"date\"></label>\n            <label for=\"tHeure\">Heure d'intervention<input id=\"tHeure\" type=\"time\"></label>\n          </div>\n          <label for=\"tIntervenant\">Intervenant (nom et matricule)<input id=\"tIntervenant\" autocomplete=\"off\"></label>\n          <label for=\"tAction\">Action réalisée<textarea id=\"tAction\" placeholder=\"Travaux effectués, suite donnée…\"></textarea></label>\n          <label for=\"tPhoto\">Photo après travaux<input id=\"tPhoto\" type=\"file\" accept=\"image/*\" capture=\"environment\"></label>\n          <img id=\"tPreview\" alt=\"Aperçu de la photo après travaux\" hidden>\n          <label class=\"inline\"><input type=\"checkbox\" id=\"tRappel\"> Le client a rappelé (ajoute 1 au nombre de rappels : <b id=\"tRappels\">0</b>)</label>\n        </fieldset>\n        <p class=\"msg\" id=\"tMsg\" role=\"status\"></p>\n        <div class=\"actions\"><button type=\"submit\" class=\"primary\" id=\"btnTraiter\">Enregistrer le traitement</button></div>\n      </form>\n    </section>\n  </div>\n  <p class=\"privacy\">Version Veolia 5</p>\n</div>\n\n<script>\n\"use strict\";\n/* Version hébergée dans le script Google Veolia : l'envoi passe par google.script.run */\nconst ENDPOINT = \"script\";\n\nconst $=id=>document.getElementById(id);\nlet pos=null, photoData='', thumbData='';\nconst QKEY='reuse-envoi-en-attente';\n\n['agent','matricule'].forEach(id=>{\n  try{$(id).value=localStorage.getItem('reuse-'+id)||''}catch(e){}\n  $(id).addEventListener('change',()=>{try{localStorage.setItem('reuse-'+id,$(id).value.trim())}catch(e){}});\n});\n\nfunction msg(t,k){$('msg').textContent=t;$('msg').className='msg '+(k||'')}\nconst mapsUrl=(lat,lng)=>'https://www.google.com/maps?q='+lat+','+lng;\n\n/* ---------- Position en un clic ---------- */\nfunction showPos(){\n  const o=$('geoOut');o.replaceChildren();\n  if(!pos){o.textContent='Aucune position enregistrée.';return}\n  const a=document.createElement('a');a.href=mapsUrl(pos.lat,pos.lng);a.target='_blank';a.rel='noopener';a.textContent='Vérifier sur la carte';\n  o.append(pos.lat.toFixed(6)+', '+pos.lng.toFixed(6)+'  (± '+Math.round(pos.acc)+' m)  ',a);\n}\n$('btnGeo').addEventListener('click',()=>{\n  const n=$('geoNote'),b=$('btnGeo');\n  if(!('geolocation' in navigator)){n.textContent=\"Ce navigateur ne donne pas accès au GPS.\";return}\n  if(!window.isSecureContext){n.textContent=\"Le GPS n'est pas disponible ici : ajoutez une photo prise avec la localisation activée, sa position sera reprise.\";return}\n  b.disabled=true;n.textContent='Recherche de la position GPS…';\n  navigator.geolocation.getCurrentPosition(p=>{\n    pos={lat:+p.coords.latitude.toFixed(6),lng:+p.coords.longitude.toFixed(6),acc:p.coords.accuracy};\n    showPos();b.disabled=false;\n    n.textContent=pos.acc>50?'Position enregistrée, mais peu précise (± '+Math.round(pos.acc)+' m). Réessayez à l’extérieur si possible.':'Position enregistrée.';\n  },err=>{\n    b.disabled=false;\n    n.textContent=err.code===1?\"Accès à la position refusé ou bloqué. Autorisez la localisation, ou ajoutez une photo prise avec la localisation activée : sa position sera reprise.\":err.code===3?\"Le GPS n'a pas répondu à temps. Réessayez à l'extérieur.\":\"Position introuvable. Vérifiez que la localisation du téléphone est activée.\";\n  },{enableHighAccuracy:true,timeout:20000,maximumAge:0});\n});\n\n/* ---------- Formulaire ---------- */\nconst BOXES={\"Fuite d'eau\":'fuiteBox','Réfection':'refectionBox',\"Qualité d'eau\":'qualiteBox','Problème de Pression':'pressionBox',\n  'Demande Coupure':'coupureBox','Demande Rebranchement':'rebranchementBox','DFO':'dfoBox',\"Manque d'eau\":'manqueBox','Appel Client pour Information':'appelBox'};\nconst frDate=s=>s?s.split('-').reverse().join('/'):'';\nconst dateHeure=(d,h)=>[frDate(d),h].filter(Boolean).join(' à ');\nfunction showBoxes(){const n=$('nature').value;Object.entries(BOXES).forEach(([k,id])=>{$(id).hidden=n!==k})}\n$('nature').addEventListener('change',showBoxes);\nfunction nowFields(){const d=new Date(),p=n=>String(n).padStart(2,'0');$('recDate').value=d.getFullYear()+'-'+p(d.getMonth()+1)+'-'+p(d.getDate());$('recHeure').value=p(d.getHours())+':'+p(d.getMinutes());$('recDate').max=$('recDate').value}\nnowFields();\n$('form').addEventListener('reset',()=>setTimeout(()=>{nowFields();showBoxes();$('preview').hidden=true;photoData='';thumbData='';pos=null;showPos();$('geoNote').textContent='Placez-vous sur le lieu de la fuite puis appuyez sur le bouton.'},0));\nfunction shrink(file,max=1280){return new Promise((res,rej)=>{const img=new Image();img.onload=()=>{const s=Math.min(1,max/Math.max(img.width,img.height)),c=document.createElement('canvas');c.width=Math.round(img.width*s);c.height=Math.round(img.height*s);c.getContext('2d').drawImage(img,0,0,c.width,c.height);URL.revokeObjectURL(img.src);res(c.toDataURL('image/jpeg',0.75))};img.onerror=()=>rej(new Error('image'));img.src=URL.createObjectURL(file)})}\n/* Position lue dans la photo (si le GPS du formulaire n'a pas pu être utilisé) */\nfunction gpsFromJpeg(buf){\n  const v=new DataView(buf);if(v.byteLength<4||v.getUint16(0)!==0xFFD8)return null;\n  let o=2;\n  while(o+4<v.byteLength){\n    const m=v.getUint16(o),len=v.getUint16(o+2);\n    if(m===0xFFE1&&v.getUint32(o+4)===0x45786966){\n      const t=o+10,le=v.getUint16(t)===0x4949;\n      const u16=p=>v.getUint16(t+p,le),u32=p=>v.getUint32(t+p,le);\n      const ifd0=u32(4),n0=u16(ifd0);let g=0;\n      for(let i=0;i<n0;i++){const en=ifd0+2+i*12;if(u16(en)===0x8825)g=u32(en+8)}\n      if(!g)return null;\n      const r={},n=u16(g);\n      for(let i=0;i<n;i++){\n        const en=g+2+i*12,tag=u16(en);\n        if(tag===1||tag===3)r[tag]=String.fromCharCode(v.getUint8(t+en+8));\n        if(tag===2||tag===4){const p=u32(en+8),q=k=>u32(p+k*8)/(u32(p+k*8+4)||1);r[tag]=q(0)+q(1)/60+q(2)/3600}\n      }\n      if(r[2]==null||r[4]==null)return null;\n      return {lat:(r[1]==='S'?-1:1)*r[2],lng:(r[3]==='W'?-1:1)*r[4]};\n    }\n    if((m&0xFF00)!==0xFF00)break;o+=2+len;\n  }\n  return null;\n}\n$('photo').addEventListener('change',async()=>{\n  const f=$('photo').files[0];photoData='';thumbData='';$('preview').hidden=true;if(!f)return;\n  if(!pos){try{const g=gpsFromJpeg(await f.arrayBuffer());\n    if(g&&isFinite(g.lat)&&isFinite(g.lng)&&(g.lat||g.lng)){pos={lat:+g.lat.toFixed(6),lng:+g.lng.toFixed(6),acc:30};showPos();$('geoNote').textContent='Position lue dans la photo.'}}catch(e){}}\n  try{photoData=await shrink(f);thumbData=await shrink(f,400);$('preview').src=photoData;$('preview').hidden=false}\n  catch(e){msg(\"Cette image n'a pas pu être lue. Essayez une autre photo.\",'err')}\n});\n\nlet lastErr='';\nfunction send(payload){\n  return new Promise((res,rej)=>{\n    google.script.run\n      .withSuccessHandler(j=>{ if(j&&j.ok)res(j); else rej(Object.assign(new Error((j&&j.error)||'refus'),{refused:true})); })\n      .withFailureHandler(err=>{ lastErr=String(err&&err.message||err); rej(err); })\n      .envoyer(payload);\n  });\n}\n\n/* File d'attente : uniquement si le réseau manque ; vidée dès l'envoi réussi */\nfunction queue(){try{return JSON.parse(localStorage.getItem(QKEY)||'[]')}catch(e){return[]}}\nfunction setQueue(q){try{q.length?localStorage.setItem(QKEY,JSON.stringify(q)):localStorage.removeItem(QKEY)}catch(e){}}\nasync function flush(){\n  if(!ENDPOINT)return;\n  let q=queue();if(!q.length)return;\n  const rest=[];\n  for(const p of q){try{await send(p)}catch(e){if(!e.refused)rest.push(p)}}\n  setQueue(rest);\n  if(q.length!==rest.length&&!rest.length)msg('Les réclamations en attente ont été envoyées.','ok');\n  else if(rest.length)msg(rest.length+' réclamation(s) en attente d’envoi : elles partiront dès que le réseau revient.','info');\n}\nwindow.addEventListener('online',flush);\n\n$('form').addEventListener('submit',async e=>{\n  e.preventDefault();\n  const v=id=>($(id).value||'').trim(),f=$('form');\n  if(!ENDPOINT)return msg(\"Ce formulaire n'est pas encore relié au tableau du responsable.\",'err');\n  const miss=[];if(!v('recDate'))miss.push('date de réception');if(!v('agent'))miss.push('nom de l’agent');if(!v('matricule'))miss.push('matricule');if(!v('nom'))miss.push('nom du réclamant');if(!v('adresse'))miss.push('adresse');if(!v('nature'))miss.push('nature');\n  if(miss.length)return msg('Champ obligatoire manquant : '+miss.join(', ')+'.','err');\n  if(v('tel')&&(v('tel').match(/\\d/g)||[]).length<8)return msg('Le numéro de téléphone semble incomplet (8 chiffres minimum).','err');\n  const fuite=v('nature')===\"Fuite d'eau\";\n  if(fuite&&!f.elements.loc.value)return msg('Indiquez la localisation de la fuite.','err');\n  const nat=v('nature');let detail='',depuis='',periode='',souhaite='';const cil=v('cil');\n  const checked=n=>[...f.querySelectorAll('input[name='+n+']:checked')].map(x=>x.value).join(', ');\n  if(nat==='Demande Coupure'){souhaite=dateHeure(v('coupureDate'),v('coupureHeure'))}\n  if(nat==='Demande Rebranchement'){souhaite=dateHeure(v('rebranchementDate'),v('rebranchementHeure'))}\n  if(nat==='DFO'){detail=f.elements.dfo.value;if(!detail)return msg('Indiquez à quel niveau se trouve le DFO.','err')}\n  if(nat===\"Manque d'eau\"){depuis=dateHeure(v('manqueDate'),v('manqueHeure'))}\n  if(nat==='Appel Client pour Information'){detail=checked('appel');if(!detail)return msg('Cochez la raison de l’appel.','err')}\n  if(nat==='Réfection'){detail=f.elements.refection.value;if(!detail)return msg('Indiquez le type de réfection.','err')}\n  if(nat===\"Qualité d'eau\"){detail=[...f.querySelectorAll('input[name=qualite]:checked')].map(x=>x.value).join(', ');if(!detail)return msg('Cochez la nature du défaut qualité.','err');depuis=v('qualiteDepuis');periode=v('qualitePeriode')}\n  if(nat==='Problème de Pression'){detail=f.elements.pression.value;if(!detail)return msg('Indiquez la nature du défaut de pression.','err');depuis=v('pressionDepuis');periode=v('pressionPeriode')}\n  const payload={clientId:(crypto.randomUUID?crypto.randomUUID():Date.now()+'-'+Math.random()),saisieLe:new Date().toISOString(),\n    reception:v('recDate')+(v('recHeure')?'T'+v('recHeure'):''),agent:v('agent'),matricule:v('matricule'),typeClient:v('typeClient'),nom:v('nom'),tel:v('tel'),email:v('email'),adresse:v('adresse'),quartier:v('quartier'),ville:v('ville'),\n    lat:pos?pos.lat:'',lng:pos?pos.lng:'',precision:pos?Math.round(pos.acc):'',\n    nature:v('nature'),localisation:fuite?f.elements.loc.value:'',importance:fuite?(f.elements.imp.value||''):'',detail,depuis,periode,cil,souhaite,obs:v('obs'),photo:photoData,vignette:thumbData};\n  const b=$('btnSave');b.disabled=true;msg('Envoi en cours…','info');\n  try{\n    const j=await send(payload);\n    f.reset();setTimeout(()=>msg('Réclamation envoyée au responsable.','ok'),10);\n  }catch(err){\n    if(err.refused){msg('Réclamation refusée : '+err.message,'err')}\n    else{const q=queue();q.push(payload);setQueue(q);f.reset();\n      setTimeout(()=>msg(navigator.onLine?\"Envoi impossible pour le moment : la réclamation est gardée sur ce téléphone et sera renvoyée automatiquement à la prochaine ouverture du formulaire.\"+(lastErr?' Détail technique : '+lastErr:''):\"Pas de réseau : la réclamation est gardée sur ce téléphone et partira automatiquement dès que le réseau revient.\",'info'),10)}\n  }finally{b.disabled=false}\n});\nflush();\n\n/* ===================== Traitement ===================== */\nconst run=(fn,...a)=>new Promise((res,rej)=>google.script.run.withSuccessHandler(res).withFailureHandler(rej)[fn](...a));\nconst ETAT_CLASS={'Non encore Transmise':'e0','Transmise mais non encore résolue':'e1','Résolue Provisoirement':'e2','Résolue Définitivement':'e3'};\nlet currentSaisie='', photoApres='';\nfunction tmsg(id,t,k){$(id).textContent=t;$(id).className='msg '+(k||'')}\nfunction el(tag,props,...kids){const e=document.createElement(tag);Object.assign(e,props||{});kids.forEach(k=>k!=null&&e.append(k));return e}\n[...$('nature').options].filter(o=>o.value).forEach(o=>$('sNature').add(new Option(o.value,o.value)));\n[...$('typeClient').options].forEach(o=>$('sType').add(new Option(o.value,o.value)));\n\nfunction showTab(name){\n  $('tabSaisie').hidden=name!=='saisie';$('tabTraitement').hidden=name!=='traitement';\n  document.body.classList.toggle('large',name==='traitement');\n  document.querySelectorAll('.tab').forEach(b=>b.classList.toggle('on',b.dataset.tab===name));\n  if(name==='traitement'&&!$('results').childElementCount)search();\n}\ndocument.querySelectorAll('.tab').forEach(b=>b.addEventListener('click',()=>showTab(b.dataset.tab)));\nrun('moi').then(m=>{\n  if(!m||!m.traitement)return;\n  $('tabs').hidden=false;\n  (m.etats||[]).forEach(s=>$('tEtat').add(new Option(s,s)));\n}).catch(()=>{});\n\nfunction filtres(){const v=id=>($(id).value||'').trim();\n  return {num:v('sNum'),etat:$('sEtat').value,du:v('sDu'),au:v('sAu'),nature:v('sNature'),typeClient:v('sType'),nom:v('sNom'),agent:v('sAgent'),cil:v('sCil'),tel:v('sTel'),commune:v('sCommune'),quartier:v('sQuartier'),adresse:v('sAdresse')}}\nasync function search(){\n  const b=$('btnSearch');b.disabled=true;tmsg('sMsg','Recherche en cours…','info');\n  try{\n    const r=await run('rechercher',filtres());\n    if(!r||!r.ok)return tmsg('sMsg','Recherche impossible : '+((r&&r.error)||'erreur'),'err');\n    tmsg('sMsg','');const box=$('results');box.replaceChildren();\n    box.append(el('div',{className:'count',textContent:r.total===0?'Aucune réclamation ne correspond à ces critères.':r.total+' réclamation(s)'+(r.total>r.items.length?' — les 200 plus récentes sont affichées':'')}));\n    r.items.forEach(it=>{\n      const btn=el('button',{type:'button',className:'res'},\n        el('span',{},el('b',{textContent:it.saisie}),it.reclamation?'  ·  N° '+it.reclamation:'','  ·  '+it.reception),\n        el('span',{className:'etat '+(ETAT_CLASS[it.etat]||''),textContent:it.etat}),\n        el('span',{className:'d',textContent:[it.nature+(it.detail?' ('+it.detail+')':''),it.nom,it.adresse,it.rappels?it.rappels+' rappel(s)':''].filter(Boolean).join(' — ')}));\n      btn.addEventListener('click',()=>ouvrir(it.saisie));box.append(btn);\n    });\n  }catch(e){tmsg('sMsg','Recherche impossible : '+(e&&e.message||e),'err')}\n  finally{b.disabled=false}\n}\n$('search').addEventListener('submit',e=>{e.preventDefault();search()});\n$('search').addEventListener('reset',()=>setTimeout(()=>{$('sEtat').value='non-resolues'},0));\n\nasync function ouvrir(saisie){\n  tmsg('sMsg','Ouverture de la fiche…','info');\n  try{\n    const r=await run('fiche',saisie);\n    if(!r||!r.ok)return tmsg('sMsg','Fiche impossible à ouvrir : '+((r&&r.error)||'erreur'),'err');\n    tmsg('sMsg','');currentSaisie=r.saisie;$('ficheTitre').textContent='Fiche '+r.saisie;\n    const dl=$('ficheChamps');dl.replaceChildren();\n    r.champs.forEach(([k,v])=>dl.append(el('dt',{textContent:k}),el('dd',{textContent:v})));\n    const li=$('ficheLiens');li.replaceChildren();\n    [['carte','Ouvrir la carte'],['photo','Voir la photo'],['photoApres','Voir la photo après travaux']].forEach(([k,t])=>{if(r.liens[k])li.append(el('a',{href:r.liens[k],target:'_blank',rel:'noopener',textContent:t}))});\n    const t=r.traitement;$('tEtat').value=t.etat;$('tReclamation').value=t.reclamation;\n    $('tDate').value=t.intervention?t.intervention.slice(0,10):'';$('tHeure').value=t.intervention?t.intervention.slice(11,16):'';\n    $('tIntervenant').value=t.intervenant||[$('agent').value,$('matricule').value].filter(Boolean).join(' ');\n    $('tAction').value=t.action;$('tRappels').textContent=t.rappels;$('tRappel').checked=false;\n    $('tPhoto').value='';photoApres='';$('tPreview').hidden=true;tmsg('tMsg','');\n    $('fichePanel').hidden=false;$('fichePanel').scrollIntoView({behavior:'smooth'});\n  }catch(e){tmsg('sMsg','Fiche impossible à ouvrir : '+(e&&e.message||e),'err')}\n}\n$('btnBack').addEventListener('click',()=>{$('fichePanel').hidden=true;$('results').scrollIntoView({behavior:'smooth'})});\n$('tPhoto').addEventListener('change',async()=>{\n  const f=$('tPhoto').files[0];photoApres='';$('tPreview').hidden=true;if(!f)return;\n  try{photoApres=await shrink(f);$('tPreview').src=photoApres;$('tPreview').hidden=false}catch(e){tmsg('tMsg',\"Cette image n'a pas pu être lue.\",'err')}\n});\n$('traitement').addEventListener('submit',async e=>{\n  e.preventDefault();if(!currentSaisie)return;\n  const v=id=>($(id).value||'').trim();\n  const t={saisie:currentSaisie,etat:$('tEtat').value,reclamation:v('tReclamation'),\n    intervention:v('tDate')?v('tDate')+(v('tHeure')?'T'+v('tHeure'):''):'',intervenant:v('tIntervenant'),action:v('tAction'),rappel:$('tRappel').checked,photoApres};\n  const b=$('btnTraiter');b.disabled=true;tmsg('tMsg','Enregistrement…','info');\n  try{\n    const r=await run('traiter',t);\n    if(!r||!r.ok)return tmsg('tMsg','Enregistrement impossible : '+((r&&r.error)||'erreur'),'err');\n    await ouvrir(currentSaisie);tmsg('tMsg','Traitement enregistré.','ok');search();\n  }catch(err){tmsg('tMsg','Enregistrement impossible : '+(err&&err.message||err),'err')}\n  finally{b.disabled=false}\n});\n</script>\n</body>\n</html>\n";
