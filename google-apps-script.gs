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
  'Nature', 'Localisation fuite', 'Importance fuite', 'Observations', 'Photo', 'Statut', 'ID envoi'];
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

    let photo = '';
    if (typeof d.photo === 'string' && d.photo.indexOf('data:image/jpeg;base64,') === 0) {
      const bytes = Utilities.base64Decode(d.photo.split(',')[1]);
      if (bytes.length < 8 * 1024 * 1024) {
        const file = getFolder().createFile(Utilities.newBlob(bytes, 'image/jpeg', numero + '.jpg'));
        photo = '=HYPERLINK("' + file.getUrl() + '","Voir la photo")';
      }
    }

    const hasPos = d.lat !== '' && d.lat != null && d.lng !== '' && d.lng != null && isFinite(d.lat) && isFinite(d.lng);
    const carte = hasPos ? '=HYPERLINK("https://www.google.com/maps?q=' + Number(d.lat) + ',' + Number(d.lng) + '","Ouvrir la carte")' : '';

    sh.appendRow([
      numero, new Date(), safe(d.agent), safe(d.nom), safe(d.tel), safe(d.email),
      safe(d.adresse), safe(d.quartier), safe(d.ville),
      hasPos ? Number(d.lat) : '', hasPos ? Number(d.lng) : '', hasPos && isFinite(d.precision) ? Number(d.precision) : '', carte,
      d.nature, safe(d.localisation), safe(d.importance), safe(d.obs), photo, 'Ouverte', safe(d.clientId)
    ]);
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
    sh.getRange('S2:S').setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(['Ouverte', 'En cours', 'Clôturée'], true).build());
    sh.hideColumns(HEADERS.length);
  }
  return sh;
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
