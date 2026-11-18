/**
 * Réclamations Ré-Use — script du tableau Google Sheets du responsable.
 * Reçoit les réclamations envoyées par le formulaire et les ajoute au tableau.
 * Les photos sont rangées dans un dossier Google Drive privé.
 *
 * Installation : Extensions > Apps Script, coller ce code, enregistrer,
 * puis Déployer > Nouveau déploiement > Application Web
 * (Exécuter en tant que : Moi ; Qui a accès : Tout le monde).
 */

const SHEET_NAME = 'Réclamations';
const FOLDER_NAME = 'Photos réclamations Ré-Use';
const HEADERS = ['N° réclamation', 'Date de réception', 'Agent', 'Nom du réclamant', 'Téléphone', 'E-mail',
  'Adresse', 'Quartier', 'Commune', 'Latitude', 'Longitude', 'Précision (m)', 'Carte',
  'Nature', 'Localisation fuite', 'Importance fuite', 'Observations', 'Aperçu photo', 'Lien photo', 'Statut', 'ID envoi'];
const COL = { carte: 13, apercu: 18, lienPhoto: 19, statut: 20 };
const NATURES = ['Réfection', "Fuite d'eau", "Qualité d'eau", 'Problème de Pression', 'Demande Coupure',
  'Demande Rebranchement', 'DFO', "Manque d'eau", 'Appel Client pour Information'];

function doGet() {
  return ContentService.createTextOutput('Le script des réclamations Ré-Use fonctionne.');
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    const d = JSON.parse(e.postData.contents);
    for (const k of ['agent', 'nom', 'tel', 'adresse', 'nature']) {
      if (!d[k] || !String(d[k]).trim()) return reply({ ok: false, error: 'champ manquant (' + k + ')' });
    }
    if (NATURES.indexOf(d.nature) < 0) return reply({ ok: false, error: 'nature inconnue' });

    const sh = getSheet();

    // Même envoi reçu deux fois (réseau coupé puis rétabli) : on ne le duplique pas.
    if (d.clientId && sh.getLastRow() > 1) {
      const idCol = HEADERS.length;
      const found = sh.getRange(2, idCol, sh.getLastRow() - 1, 1).createTextFinder(String(d.clientId)).matchEntireCell(true).findNext();
      if (found) return reply({ ok: true, numero: sh.getRange(found.getRow(), 1).getValue(), doublon: true });
    }

    const year = new Date().getFullYear();
    const numero = 'RU-' + year + '-' + String(sh.getLastRow()).padStart(5, '0');

    const hasPos = d.lat !== '' && d.lat != null && d.lng !== '' && d.lng != null && isFinite(d.lat) && isFinite(d.lng);

    sh.appendRow([
      numero, new Date(), safe(d.agent), safe(d.nom), safe(d.tel), safe(d.email),
      safe(d.adresse), safe(d.quartier), safe(d.ville),
      hasPos ? Number(d.lat) : '', hasPos ? Number(d.lng) : '', hasPos && isFinite(d.precision) ? Number(d.precision) : '', '',
      d.nature, safe(d.localisation), safe(d.importance), safe(d.obs), '', '', 'Ouverte', safe(d.clientId)
    ]);
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
    return reply({ ok: true, numero: numero });
  } catch (err) {
    return reply({ ok: false, error: 'erreur du tableau : ' + err.message });
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
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#0f6e86').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.getRange('B:B').setNumberFormat('dd/MM/yyyy HH:mm');
    sh.getRange('E:E').setNumberFormat('@');
    sh.getRange(2, COL.statut, sh.getMaxRows() - 1, 1).setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(['Ouverte', 'En cours', 'Clôturée'], true).build());
    sh.setColumnWidth(COL.apercu, 160);
    sh.hideColumns(HEADERS.length);
  } else if (sh.getRange(1, COL.apercu).getValue() === 'Photo') {
    // Tableau créé avec l'ancienne version : on ajoute la colonne « Aperçu photo ».
    sh.insertColumnBefore(COL.apercu);
    sh.getRange(1, COL.apercu, 1, 2).setValues([['Aperçu photo', 'Lien photo']])
      .setFontWeight('bold').setBackground('#0f6e86').setFontColor('#ffffff');
    sh.setColumnWidth(COL.apercu, 160);
  }
  return sh;
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

function reply(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
