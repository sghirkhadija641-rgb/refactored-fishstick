/**
 * Réclamations Ré-Use — version compte Veolia.
 * Le formulaire est servi par ce script ; seuls les comptes Veolia connectés peuvent l'ouvrir.
 * Chaque réclamation est ajoutée au tableau, avec l'adresse Veolia de l'agent qui l'a envoyée.
 *
 * Installation : Extensions > Apps Script, coller ce code, enregistrer,
 * puis Déployer > Nouveau déploiement > Application Web
 * (Exécuter en tant que : Moi ; Qui a accès : Anyone within Veolia).
 * Le lien de l'application Web est le lien du formulaire.
 */

const SHEET_NAME = 'Réclamations';
const FOLDER_NAME = 'Photos réclamations Ré-Use';
const HEADERS = ['N° Saisie', 'Date de réception', 'Agent', 'Matricule', 'Nom du réclamant', 'Téléphone', 'E-mail',
  'Adresse', 'Quartier', 'Commune', 'Latitude', 'Longitude', 'Précision (m)', 'Carte',
  'Nature', 'Localisation fuite', 'Importance fuite', 'Détail nature', 'Depuis', 'Période', 'CIL', 'Date/heure souhaitée',
  'Observations', 'Aperçu photo', 'Lien photo', 'Statut', 'ID envoi'];
const COL = {
  carte: HEADERS.indexOf('Carte') + 1, tel: HEADERS.indexOf('Téléphone') + 1, matricule: HEADERS.indexOf('Matricule') + 1,
  apercu: HEADERS.indexOf('Aperçu photo') + 1, lienPhoto: HEADERS.indexOf('Lien photo') + 1, statut: HEADERS.indexOf('Statut') + 1
};
const NATURES = ['Réfection', "Fuite d'eau", "Qualité d'eau", 'Problème de Pression', 'Demande Coupure',
  'Demande Rebranchement', 'DFO', "Manque d'eau", 'Appel Client pour Information'];

function doGet(e) {
  if (e && e.parameter && e.parameter.reparer) {
    let note = '';
    try { note = ' Lignes réparées : ' + reparerAnciennesLignes() + '.'; } catch (err) { note = ' Réparation impossible : ' + err.message; }
    return ContentService.createTextOutput('Le script des réclamations Ré-Use fonctionne. (version Veolia 2)' + note);
  }
  return HtmlService.createHtmlOutput(FORM_HTML)
    .setTitle('Réclamations Ré-Use')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Appelée par le formulaire (google.script.run).
function envoyer(d) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    const compte = Session.getActiveUser().getEmail();
    d.agent = [String(d.agent || '').trim(), compte ? '(' + compte + ')' : ''].filter(String).join(' ');
    for (const k of ['matricule', 'nom', 'adresse', 'nature']) {
      if (!d[k] || !String(d[k]).trim()) return ({ ok: false, error: 'champ manquant (' + k + ')' });
    }
    if (NATURES.indexOf(d.nature) < 0) return ({ ok: false, error: 'nature inconnue' });

    const sh = getSheet();

    // Même envoi reçu deux fois (réseau coupé puis rétabli) : on ne le duplique pas.
    if (d.clientId && sh.getLastRow() > 1) {
      const idCol = HEADERS.length;
      const found = sh.getRange(2, idCol, sh.getLastRow() - 1, 1).createTextFinder(String(d.clientId)).matchEntireCell(true).findNext();
      if (found) return ({ ok: true, numero: sh.getRange(found.getRow(), 1).getValue(), doublon: true });
    }

    const year = new Date().getFullYear();
    const numero = 'RU-' + year + '-' + String(sh.getLastRow()).padStart(5, '0');

    const hasPos = d.lat !== '' && d.lat != null && d.lng !== '' && d.lng != null && isFinite(d.lat) && isFinite(d.lng);

    const values = {
      'N° Saisie': numero, 'Date de réception': new Date(), 'Agent': safe(d.agent), 'Matricule': safe(d.matricule),
      'Nom du réclamant': safe(d.nom), 'Téléphone': safe(d.tel), 'E-mail': safe(d.email),
      'Adresse': safe(d.adresse), 'Quartier': safe(d.quartier), 'Commune': safe(d.ville),
      'Latitude': hasPos ? Number(d.lat) : '', 'Longitude': hasPos ? Number(d.lng) : '',
      'Précision (m)': hasPos && isFinite(d.precision) ? Number(d.precision) : '',
      'Nature': d.nature, 'Localisation fuite': safe(d.localisation), 'Importance fuite': safe(d.importance),
      'Détail nature': safe(d.detail), 'Depuis': safe(d.depuis), 'Période': safe(d.periode),
      'CIL': safe(d.cil), 'Date/heure souhaitée': safe(d.souhaite),
      'Observations': safe(d.obs), 'Statut': 'Ouverte', 'ID envoi': safe(d.clientId)
    };
    sh.appendRow(HEADERS.map(h => h in values ? values[h] : ''));
    const row = sh.getLastRow();

    // Lien carte : vrai lien cliquable, conservé dans le fichier Excel téléchargé.
    if (hasPos) {
      sh.getRange(row, COL.carte).setRichTextValue(link('Ouvrir la carte',
        'https://www.google.com/maps?q=' + Number(d.lat) + ',' + Number(d.lng)));
    }

    // Photo en grand format : rangée dans Google Drive, lien dans le tableau.
    const full = jpeg(d.photo);
    if (full && full.length < 8 * 1024 * 1024) {
      const file = getFolder().createFile(Utilities.newBlob(full, 'image/jpeg', numero + '.jpg'));
      sh.getRange(row, COL.lienPhoto).setRichTextValue(link('Voir la photo', file.getUrl()));
    }

    // Aperçu : petite image posée dans la ligne, visible aussi dans Excel.
    const small = jpeg(d.vignette);
    if (small && small.length < 1.5 * 1024 * 1024) {
      try {
        const img = sh.insertImage(Utilities.newBlob(small, 'image/jpeg', numero + '-apercu.jpg'), COL.apercu, row, 4, 4);
        const ratio = img.getHeight() / img.getWidth();
        img.setWidth(150).setHeight(Math.round(150 * ratio));
        sh.setRowHeight(row, Math.round(150 * ratio) + 10);
      } catch (e) {
        sh.getRange(row, COL.apercu).setValue('aperçu indisponible');
      }
    }
    return ({ ok: true, numero: numero });
  } catch (err) {
    return ({ ok: false, error: 'erreur du tableau : ' + err.message });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function getSheet() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    styleHeader(sh.getRange(1, 1, 1, HEADERS.length));
    sh.setFrozenRows(1);
    sh.getRange('B:B').setNumberFormat('dd/MM/yyyy HH:mm');
    sh.getRange(2, COL.tel, sh.getMaxRows() - 1, 2).setNumberFormat('@');
    sh.getRange(2, COL.matricule, sh.getMaxRows() - 1, 1).setNumberFormat('@');
    sh.getRange(2, HEADERS.indexOf('CIL') + 1, sh.getMaxRows() - 1, 1).setNumberFormat('@');
    sh.getRange(2, COL.statut, sh.getMaxRows() - 1, 1).setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(['Ouverte', 'En cours', 'Clôturée'], true).build());
    sh.setColumnWidth(COL.apercu, 160);
    sh.hideColumns(HEADERS.length);
  } else {
    mettreAJourColonnes(sh);
  }
  return sh;
}

// Tableau créé avec une version précédente : ajoute les colonnes manquantes à leur place.
function mettreAJourColonnes(sh) {
  let head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const oldNum = head.indexOf('N° réclamation');
  if (oldNum >= 0 && head.indexOf('N° Saisie') < 0) {
    sh.getRange(1, oldNum + 1).setValue('N° Saisie');
    head[oldNum] = 'N° Saisie';
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
    sh.insertColumnBefore(i + 1);
    sh.getRange(1, i + 1).setValue(h);
    styleHeader(sh.getRange(1, i + 1));
    sh.showColumns(i + 1);
    head.splice(i, 0, h);
  });
}

function styleHeader(range) {
  return range.setFontWeight('bold').setBackground('#0f6e86').setFontColor('#ffffff');
}

/**
 * Répare le tableau : anciens liens en #ERROR!, et lignes rangées dans les mauvaises colonnes
 * par l'ancienne version du script. Se lance tout seul à chaque ouverture du lien du script ;
 * on peut aussi la lancer depuis l'éditeur. Sans effet sur les lignes déjà correctes.
 */
function reparerAnciennesLignes() {
  const sh = getSheet();
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const STATUTS = ['Ouverte', 'En cours', 'Clôturée'];
  const n = HEADERS.length;
  const values = sh.getRange(2, 1, last - 1, n).getValues();
  const formulas = sh.getRange(2, 1, last - 1, n).getFormulas();
  const urlOf = f => { const m = String(f || '').match(/HYPERLINK\("([^"]+)"/i); return m ? m[1] : null; };
  let fixed = 0;
  values.forEach((v, i) => {
    const row = i + 2, f = formulas[i];
    let changed = false;
    let photoUrl = urlOf(f[COL.lienPhoto - 1]);
    // Ligne décalée d'une colonne (statut dans « Lien photo ») : on la remet en place.
    if (STATUTS.indexOf(v[COL.lienPhoto - 1]) >= 0 && STATUTS.indexOf(v[COL.statut - 1]) < 0) {
      photoUrl = urlOf(f[COL.apercu - 1]);
      sh.getRange(row, COL.statut + 1).setValue(v[COL.statut - 1]);
      sh.getRange(row, COL.statut).setValue(v[COL.lienPhoto - 1]);
      sh.getRange(row, COL.apercu, 1, 2).clearContent();
      changed = true;
    }
    const carte = urlOf(f[COL.carte - 1]);
    if (carte) { sh.getRange(row, COL.carte).setRichTextValue(link('Ouvrir la carte', carte)); changed = true; }
    if (photoUrl) {
      sh.getRange(row, COL.lienPhoto).setRichTextValue(link('Voir la photo', photoUrl));
      addThumb(sh, row, photoUrl);
      changed = true;
    }
    if (changed) fixed++;
  });
  return fixed;
}

function addThumb(sh, row, driveUrl) {
  const id = (String(driveUrl).match(/[-\w]{25,}/) || [])[0];
  if (!id) return;
  try {
    const file = DriveApp.getFileById(id);
    const blob = file.getThumbnail() || (file.getSize() < 1.5 * 1024 * 1024 ? file.getBlob() : null);
    if (!blob) return;
    const img = sh.insertImage(blob, COL.apercu, row, 4, 4);
    const ratio = img.getHeight() / img.getWidth();
    img.setWidth(150).setHeight(Math.round(150 * ratio));
    sh.setRowHeight(row, Math.round(150 * ratio) + 10);
  } catch (e) {}
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


