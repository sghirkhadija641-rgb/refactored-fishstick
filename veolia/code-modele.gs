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
