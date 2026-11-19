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

// Formulaire affiché aux agents (généré depuis veolia/formulaire.html).
const FORM_HTML = "<!doctype html>\n<html lang=\"fr\">\n<head>\n<meta charset=\"utf-8\">\n<meta name=\"viewport\" content=\"width=device-width,initial-scale=1,viewport-fit=cover\">\n<meta name=\"theme-color\" content=\"#0f6e86\">\n<meta name=\"robots\" content=\"noindex\">\n<title>Réclamations Ré-Use</title>\n<style>\n:root{\n  --bg:#f3f6f7; --surface:#ffffff; --line:#d6e0e3; --fg:#14262c; --muted:#5b7078;\n  --accent:#0f6e86; --accent-ink:#ffffff; --accent-soft:#e1f0f4;\n  --ok:#2f7d4f; --ok-soft:#e3f3e9; --warn:#a76a00; --warn-soft:#fbf0d9; --bad:#b3332f; --bad-soft:#fbe5e4;\n  --f-body:system-ui,-apple-system,\"Segoe UI\",Roboto,Arial,sans-serif;\n  --f-display:\"Arial Narrow\",\"Roboto Condensed\",system-ui,sans-serif;\n  --f-mono:ui-monospace,Consolas,\"Roboto Mono\",monospace;\n  color-scheme:light;\n}\n@media (prefers-color-scheme: dark){:root{\n  --bg:#0e1a1e; --surface:#15252a; --line:#273c43; --fg:#e3eef1; --muted:#93aab1;\n  --accent:#4fb6cf; --accent-ink:#06171c; --accent-soft:#163a44;\n  --ok:#6cc792; --ok-soft:#173326; --warn:#e6b45a; --warn-soft:#3a2d12; --bad:#ef8a85; --bad-soft:#3d1c1b;\n  color-scheme:dark}}\n*{box-sizing:border-box}\nhtml,body{margin:0}\nbody{background:var(--bg);color:var(--fg);font-family:var(--f-body);font-size:15px;line-height:1.45}\n.wrap{max-width:560px;margin:0 auto;padding:16px max(14px,env(safe-area-inset-right)) 40px max(14px,env(safe-area-inset-left));display:flex;flex-direction:column;gap:14px}\nheader{border-bottom:2px solid var(--fg);padding-bottom:10px}\nh1{font-family:var(--f-display);font-weight:700;font-size:24px;margin:0}\nh1 span{color:var(--accent)}\n.sub{color:var(--muted);margin:2px 0 0;font-size:13px}\n.panel{background:var(--surface);border:1px solid var(--line);border-radius:6px}\nform{padding:14px;display:flex;flex-direction:column;gap:16px}\nfieldset{border:0;padding:0;margin:0;display:flex;flex-direction:column;gap:10px;min-width:0}\nlegend,.lg{font-size:11px;text-transform:uppercase;letter-spacing:.08em;color:var(--accent);font-weight:700;padding:0;margin-bottom:6px}\n.row{display:grid;grid-template-columns:1fr 1fr;gap:10px}\n@media (max-width:460px){.row{grid-template-columns:1fr}}\nlabel{display:block;font-weight:600;font-size:13px}\nlabel>input,label>select,label>textarea{display:block;margin-top:4px}\n.req{color:var(--bad)}\ninput,select,textarea{font:inherit;font-size:16px;color:var(--fg);background:var(--bg);border:1px solid var(--line);border-radius:4px;padding:9px 10px;width:100%}\ntextarea{resize:vertical;min-height:64px}\ninput:focus,select:focus,textarea:focus,button:focus-visible{outline:2px solid var(--accent);outline-offset:1px}\n.radios{display:flex;flex-direction:column;gap:8px}\n.radios label{display:flex;align-items:center;gap:8px;font-weight:400;font-size:14px}\n.radios input{width:auto;margin:0}\n.fuite{background:var(--accent-soft);border-radius:4px;padding:12px}\n.geo{border:1px dashed var(--accent);border-radius:6px;padding:12px;display:flex;flex-direction:column;gap:8px}\n.geo-btn{display:flex;align-items:center;justify-content:center;gap:8px;font-size:16px;padding:12px}\n.geo-btn svg{width:20px;height:20px}\n.geo-out{font-family:var(--f-mono);font-size:13px}\n.geo-out a{color:var(--accent)}\n.det{color:var(--muted);font-size:12px;margin:0}\nbutton{font:inherit;font-weight:600;border-radius:4px;padding:10px 14px;cursor:pointer;border:1px solid var(--line);background:var(--surface);color:var(--fg)}\nbutton.primary{background:var(--accent);border-color:var(--accent);color:var(--accent-ink)}\nbutton:disabled{opacity:.55;cursor:wait}\n.actions{display:flex;gap:8px;flex-wrap:wrap}\n.actions .primary{flex:1 1 auto;font-size:16px;padding:12px}\n#preview{max-width:100%;max-height:200px;object-fit:cover;border-radius:4px;border:1px solid var(--line)}\n.msg{font-size:14px;margin:0;padding:10px 12px;border-radius:4px}\n.msg:empty{display:none}\n.msg.ok{color:var(--ok);background:var(--ok-soft)} .msg.err{color:var(--bad);background:var(--bad-soft)} .msg.info{color:var(--warn);background:var(--warn-soft)}\n.agent{display:flex;gap:8px;align-items:center;padding:10px 14px;background:var(--accent-soft);border-bottom:1px solid var(--line);font-size:13px}\n.agent label{white-space:nowrap}\n.agent input{padding:6px 8px;font-size:15px;margin:0}\n.agent-id{display:grid;grid-template-columns:1fr 1fr;gap:10px;align-items:start}\n.agent-id label{white-space:normal}\n.agent-id input{margin-top:4px;background:var(--surface)}\n.checks{display:flex;flex-direction:column;gap:8px}\n.checks label{display:flex;align-items:center;gap:8px;font-weight:400;font-size:14px}\n.checks input{width:auto;margin:0}\n.privacy{font-size:12px;color:var(--muted);text-align:center;margin:0}\n</style>\n</head>\n<body>\n<div class=\"wrap\">\n  <header>\n    <h1>Réclamations <span>Ré-Use</span></h1>\n    <p class=\"sub\">Saisie terrain d'une réclamation client</p>\n  </header>\n\n  <section class=\"panel\">\n    <div class=\"agent agent-id\">\n      <label for=\"agent\">Agent <span class=\"req\">*</span><input id=\"agent\" autocomplete=\"off\" placeholder=\"Votre nom\"></label>\n      <label for=\"matricule\">Matricule <span class=\"req\">*</span><input id=\"matricule\" autocomplete=\"off\" inputmode=\"text\" placeholder=\"Votre matricule\"></label>\n    </div>\n    <form id=\"form\" novalidate>\n      <fieldset>\n        <legend>Réclamant</legend>\n        <label for=\"nom\">Nom et prénom <span class=\"req\">*</span><input id=\"nom\" autocomplete=\"off\" required></label>\n        <div class=\"row\">\n          <label for=\"tel\">Téléphone<input id=\"tel\" type=\"tel\" inputmode=\"tel\" autocomplete=\"off\" placeholder=\"facultatif\"></label>\n          <label for=\"email\">E-mail<input id=\"email\" type=\"email\" autocomplete=\"off\" placeholder=\"facultatif\"></label>\n        </div>\n      </fieldset>\n\n      <fieldset>\n        <legend>Lieu de la réclamation</legend>\n        <div class=\"geo\">\n          <button type=\"button\" class=\"primary geo-btn\" id=\"btnGeo\">\n            <svg viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\" stroke-linecap=\"round\" stroke-linejoin=\"round\" aria-hidden=\"true\"><path d=\"M12 21s-7-6.2-7-11.5A7 7 0 0 1 19 9.5C19 14.8 12 21 12 21z\"/><circle cx=\"12\" cy=\"9.5\" r=\"2.5\"/></svg>\n            Enregistrer ma position\n          </button>\n          <div class=\"geo-out\" id=\"geoOut\">Aucune position enregistrée.</div>\n          <p class=\"det\" id=\"geoNote\">Placez-vous sur le lieu de la fuite puis appuyez sur le bouton.</p>\n        </div>\n        <label for=\"adresse\">Adresse <span class=\"req\">*</span><input id=\"adresse\" autocomplete=\"off\" placeholder=\"N°, rue, lotissement\" required></label>\n        <div class=\"row\">\n          <label for=\"quartier\">Quartier / secteur<input id=\"quartier\" autocomplete=\"off\"></label>\n          <label for=\"ville\">Commune<input id=\"ville\" list=\"villes\" autocomplete=\"off\"></label>\n        </div>\n        <datalist id=\"villes\"><option>Rabat</option><option>Salé</option><option>Témara</option><option>Skhirat</option><option>Harhoura</option><option>Aïn Atiq</option></datalist>\n      </fieldset>\n\n      <fieldset>\n        <legend>Réclamation</legend>\n        <label for=\"nature\">Nature <span class=\"req\">*</span>\n          <select id=\"nature\" required>\n            <option value=\"\"></option>\n            <option>Réfection</option>\n            <option>Fuite d'eau</option>\n            <option>Qualité d'eau</option>\n            <option>Problème de Pression</option>\n            <option>Demande Coupure</option>\n            <option>Demande Rebranchement</option>\n            <option>DFO</option>\n            <option>Manque d'eau</option>\n            <option>Appel Client pour Information</option>\n          </select>\n        </label>\n        <div class=\"fuite\" id=\"fuiteBox\" hidden>\n          <div class=\"row\">\n            <div>\n              <div class=\"lg\">Localisation de la fuite <span class=\"req\">*</span></div>\n              <div class=\"radios\">\n                <label><input type=\"radio\" name=\"loc\" value=\"Fuite dans la Niche\"> Fuite dans la Niche</label>\n                <label><input type=\"radio\" name=\"loc\" value=\"Fuite dans la Rue\"> Fuite dans la Rue</label>\n                <label><input type=\"radio\" name=\"loc\" value=\"Fuite au mur de la Façade\"> Fuite au mur de la Façade</label>\n                <label><input type=\"radio\" name=\"loc\" value=\"Fuite dans la Nourrice\"> Fuite dans la Nourrice</label>\n                <label><input type=\"radio\" name=\"loc\" value=\"Compteur Sauté\"> Compteur Sauté</label>\n              </div>\n            </div>\n            <div>\n              <div class=\"lg\">Importance de la fuite</div>\n              <div class=\"radios\">\n                <label><input type=\"radio\" name=\"imp\" value=\"Faible (goutte à goutte)\"> Faible (goutte à goutte)</label>\n                <label><input type=\"radio\" name=\"imp\" value=\"Moyenne (filet d'eau)\"> Moyenne (filet d'eau)</label>\n                <label><input type=\"radio\" name=\"imp\" value=\"Importante (jet / inondation)\"> Importante (jet / inondation)</label>\n              </div>\n            </div>\n          </div>\n        </div>\n        <div class=\"fuite\" id=\"refectionBox\" hidden>\n          <div class=\"lg\">Réfection <span class=\"req\">*</span></div>\n          <div class=\"radios\">\n            <label><input type=\"radio\" name=\"refection\" value=\"Non réalisée\"> Non réalisée</label>\n            <label><input type=\"radio\" name=\"refection\" value=\"A reprendre\"> A reprendre</label>\n            <label><input type=\"radio\" name=\"refection\" value=\"Chantier non Nettoyé\"> Chantier non Nettoyé</label>\n          </div>\n        </div>\n        <div class=\"fuite\" id=\"qualiteBox\" hidden>\n          <div class=\"lg\">Nature défaut qualité <span class=\"req\">*</span></div>\n          <div class=\"checks\">\n            <label><input type=\"checkbox\" name=\"qualite\" value=\"Couleur / Turbidité\"> Couleur / Turbidité</label>\n            <label><input type=\"checkbox\" name=\"qualite\" value=\"Goût / Odeur\"> Goût / Odeur</label>\n          </div>\n          <div class=\"row\" style=\"margin-top:10px\">\n            <label for=\"qualiteDepuis\">Défaut qualité depuis<input id=\"qualiteDepuis\" autocomplete=\"off\" placeholder=\"ex. 3 jours, le 15/11\"></label>\n            <label for=\"qualitePeriode\">Période défaut qualité<input id=\"qualitePeriode\" autocomplete=\"off\" placeholder=\"ex. le matin, en continu\"></label>\n          </div>\n        </div>\n        <div class=\"fuite\" id=\"pressionBox\" hidden>\n          <div class=\"lg\">Nature défaut pression <span class=\"req\">*</span></div>\n          <div class=\"radios\">\n            <label><input type=\"radio\" name=\"pression\" value=\"Manque de pression\"> Manque de pression</label>\n            <label><input type=\"radio\" name=\"pression\" value=\"Excès de pression\"> Excès de pression</label>\n            <label><input type=\"radio\" name=\"pression\" value=\"Pas d'Eau\"> Pas d'Eau</label>\n          </div>\n          <div class=\"row\" style=\"margin-top:10px\">\n            <label for=\"pressionDepuis\">Problème de pression depuis<input id=\"pressionDepuis\" autocomplete=\"off\" placeholder=\"ex. 3 jours, le 15/11\"></label>\n            <label for=\"pressionPeriode\">Période problème pression<input id=\"pressionPeriode\" autocomplete=\"off\" placeholder=\"ex. le soir, la nuit\"></label>\n          </div>\n        </div>\n        <div class=\"fuite\" id=\"coupureBox\" hidden>\n          <div class=\"lg\">Demande de coupure</div>\n          <label for=\"coupureCil\">CIL<input id=\"coupureCil\" autocomplete=\"off\" inputmode=\"numeric\" placeholder=\"N° CIL du client\"></label>\n          <div class=\"row\" style=\"margin-top:10px\">\n            <label for=\"coupureDate\">Date souhaitée<input id=\"coupureDate\" type=\"date\"></label>\n            <label for=\"coupureHeure\">Heure souhaitée<input id=\"coupureHeure\" type=\"time\"></label>\n          </div>\n        </div>\n        <div class=\"fuite\" id=\"rebranchementBox\" hidden>\n          <div class=\"lg\">Demande de rebranchement</div>\n          <label for=\"rebranchementCil\">CIL<input id=\"rebranchementCil\" autocomplete=\"off\" inputmode=\"numeric\" placeholder=\"N° CIL du client\"></label>\n          <div class=\"row\" style=\"margin-top:10px\">\n            <label for=\"rebranchementDate\">Date souhaitée<input id=\"rebranchementDate\" type=\"date\"></label>\n            <label for=\"rebranchementHeure\">Heure souhaitée<input id=\"rebranchementHeure\" type=\"time\"></label>\n          </div>\n        </div>\n        <div class=\"fuite\" id=\"dfoBox\" hidden>\n          <div class=\"lg\">DFO à quel niveau <span class=\"req\">*</span></div>\n          <div class=\"radios\">\n            <label><input type=\"radio\" name=\"dfo\" value=\"Conduite\"> Conduite</label>\n            <label><input type=\"radio\" name=\"dfo\" value=\"Niche\"> Niche</label>\n            <label><input type=\"radio\" name=\"dfo\" value=\"Nourrice\"> Nourrice</label>\n            <label><input type=\"radio\" name=\"dfo\" value=\"Branchement\"> Branchement</label>\n          </div>\n        </div>\n        <div class=\"fuite\" id=\"manqueBox\" hidden>\n          <div class=\"lg\">Manque d'eau depuis</div>\n          <div class=\"row\">\n            <label for=\"manqueDate\">Date manque d'eau<input id=\"manqueDate\" type=\"date\"></label>\n            <label for=\"manqueHeure\">Heure manque d'eau<input id=\"manqueHeure\" type=\"time\"></label>\n          </div>\n        </div>\n        <div class=\"fuite\" id=\"appelBox\" hidden>\n          <div class=\"lg\">Raison de l'appel <span class=\"req\">*</span></div>\n          <div class=\"checks\">\n            <label><input type=\"checkbox\" name=\"appel\" value=\"Client non avisé suite coupure programmée\"> Client non avisé suite coupure programmée</label>\n            <label><input type=\"checkbox\" name=\"appel\" value=\"Problème commercial\"> Problème commercial</label>\n          </div>\n        </div>\n        <label for=\"photo\">Photo de la fuite / du lieu<input id=\"photo\" type=\"file\" accept=\"image/*\" capture=\"environment\"></label>\n        <img id=\"preview\" alt=\"Aperçu de la photo\" hidden>\n        <label for=\"obs\">Observations<textarea id=\"obs\" placeholder=\"Détails donnés par le client\"></textarea></label>\n      </fieldset>\n\n      <p class=\"msg\" id=\"msg\" role=\"status\"></p>\n      <div class=\"actions\">\n        <button type=\"submit\" class=\"primary\" id=\"btnSave\">Envoyer la réclamation</button>\n        <button type=\"reset\">Effacer</button>\n      </div>\n    </form>\n  </section>\n  <p class=\"privacy\">Les réclamations envoyées ne sont pas visibles sur ce formulaire : elles sont transmises directement au responsable, avec votre compte Veolia.</p>\n  <p class=\"privacy\">Version Veolia 2</p>\n</div>\n\n<script>\n\"use strict\";\n/* Version hébergée dans le script Google Veolia : l'envoi passe par google.script.run */\nconst ENDPOINT = \"script\";\n\nconst $=id=>document.getElementById(id);\nlet pos=null, photoData='', thumbData='';\nconst QKEY='reuse-envoi-en-attente';\n\n['agent','matricule'].forEach(id=>{\n  try{$(id).value=localStorage.getItem('reuse-'+id)||''}catch(e){}\n  $(id).addEventListener('change',()=>{try{localStorage.setItem('reuse-'+id,$(id).value.trim())}catch(e){}});\n});\n\nfunction msg(t,k){$('msg').textContent=t;$('msg').className='msg '+(k||'')}\nconst mapsUrl=(lat,lng)=>'https://www.google.com/maps?q='+lat+','+lng;\n\n/* ---------- Position en un clic ---------- */\nfunction showPos(){\n  const o=$('geoOut');o.replaceChildren();\n  if(!pos){o.textContent='Aucune position enregistrée.';return}\n  const a=document.createElement('a');a.href=mapsUrl(pos.lat,pos.lng);a.target='_blank';a.rel='noopener';a.textContent='Vérifier sur la carte';\n  o.append(pos.lat.toFixed(6)+', '+pos.lng.toFixed(6)+'  (± '+Math.round(pos.acc)+' m)  ',a);\n}\n$('btnGeo').addEventListener('click',()=>{\n  const n=$('geoNote'),b=$('btnGeo');\n  if(!('geolocation' in navigator)){n.textContent=\"Ce navigateur ne donne pas accès au GPS.\";return}\n  if(!window.isSecureContext){n.textContent=\"Le GPS n'est pas disponible ici : ajoutez une photo prise avec la localisation activée, sa position sera reprise.\";return}\n  b.disabled=true;n.textContent='Recherche de la position GPS…';\n  navigator.geolocation.getCurrentPosition(p=>{\n    pos={lat:+p.coords.latitude.toFixed(6),lng:+p.coords.longitude.toFixed(6),acc:p.coords.accuracy};\n    showPos();b.disabled=false;\n    n.textContent=pos.acc>50?'Position enregistrée, mais peu précise (± '+Math.round(pos.acc)+' m). Réessayez à l’extérieur si possible.':'Position enregistrée.';\n  },err=>{\n    b.disabled=false;\n    n.textContent=err.code===1?\"Accès à la position refusé ou bloqué. Autorisez la localisation, ou ajoutez une photo prise avec la localisation activée : sa position sera reprise.\":err.code===3?\"Le GPS n'a pas répondu à temps. Réessayez à l'extérieur.\":\"Position introuvable. Vérifiez que la localisation du téléphone est activée.\";\n  },{enableHighAccuracy:true,timeout:20000,maximumAge:0});\n});\n\n/* ---------- Formulaire ---------- */\nconst BOXES={\"Fuite d'eau\":'fuiteBox','Réfection':'refectionBox',\"Qualité d'eau\":'qualiteBox','Problème de Pression':'pressionBox',\n  'Demande Coupure':'coupureBox','Demande Rebranchement':'rebranchementBox','DFO':'dfoBox',\"Manque d'eau\":'manqueBox','Appel Client pour Information':'appelBox'};\nconst frDate=s=>s?s.split('-').reverse().join('/'):'';\nconst dateHeure=(d,h)=>[frDate(d),h].filter(Boolean).join(' à ');\nfunction showBoxes(){const n=$('nature').value;Object.entries(BOXES).forEach(([k,id])=>{$(id).hidden=n!==k})}\n$('nature').addEventListener('change',showBoxes);\n$('form').addEventListener('reset',()=>setTimeout(()=>{showBoxes();$('preview').hidden=true;photoData='';thumbData='';pos=null;showPos();$('geoNote').textContent='Placez-vous sur le lieu de la fuite puis appuyez sur le bouton.'},0));\nfunction shrink(file,max=1280){return new Promise((res,rej)=>{const img=new Image();img.onload=()=>{const s=Math.min(1,max/Math.max(img.width,img.height)),c=document.createElement('canvas');c.width=Math.round(img.width*s);c.height=Math.round(img.height*s);c.getContext('2d').drawImage(img,0,0,c.width,c.height);URL.revokeObjectURL(img.src);res(c.toDataURL('image/jpeg',0.75))};img.onerror=()=>rej(new Error('image'));img.src=URL.createObjectURL(file)})}\n/* Position lue dans la photo (si le GPS du formulaire n'a pas pu être utilisé) */\nfunction gpsFromJpeg(buf){\n  const v=new DataView(buf);if(v.byteLength<4||v.getUint16(0)!==0xFFD8)return null;\n  let o=2;\n  while(o+4<v.byteLength){\n    const m=v.getUint16(o),len=v.getUint16(o+2);\n    if(m===0xFFE1&&v.getUint32(o+4)===0x45786966){\n      const t=o+10,le=v.getUint16(t)===0x4949;\n      const u16=p=>v.getUint16(t+p,le),u32=p=>v.getUint32(t+p,le);\n      const ifd0=u32(4),n0=u16(ifd0);let g=0;\n      for(let i=0;i<n0;i++){const en=ifd0+2+i*12;if(u16(en)===0x8825)g=u32(en+8)}\n      if(!g)return null;\n      const r={},n=u16(g);\n      for(let i=0;i<n;i++){\n        const en=g+2+i*12,tag=u16(en);\n        if(tag===1||tag===3)r[tag]=String.fromCharCode(v.getUint8(t+en+8));\n        if(tag===2||tag===4){const p=u32(en+8),q=k=>u32(p+k*8)/(u32(p+k*8+4)||1);r[tag]=q(0)+q(1)/60+q(2)/3600}\n      }\n      if(r[2]==null||r[4]==null)return null;\n      return {lat:(r[1]==='S'?-1:1)*r[2],lng:(r[3]==='W'?-1:1)*r[4]};\n    }\n    if((m&0xFF00)!==0xFF00)break;o+=2+len;\n  }\n  return null;\n}\n$('photo').addEventListener('change',async()=>{\n  const f=$('photo').files[0];photoData='';thumbData='';$('preview').hidden=true;if(!f)return;\n  if(!pos){try{const g=gpsFromJpeg(await f.arrayBuffer());\n    if(g&&isFinite(g.lat)&&isFinite(g.lng)&&(g.lat||g.lng)){pos={lat:+g.lat.toFixed(6),lng:+g.lng.toFixed(6),acc:30};showPos();$('geoNote').textContent='Position lue dans la photo.'}}catch(e){}}\n  try{photoData=await shrink(f);thumbData=await shrink(f,400);$('preview').src=photoData;$('preview').hidden=false}\n  catch(e){msg(\"Cette image n'a pas pu être lue. Essayez une autre photo.\",'err')}\n});\n\nlet lastErr='';\nfunction send(payload){\n  return new Promise((res,rej)=>{\n    google.script.run\n      .withSuccessHandler(j=>{ if(j&&j.ok)res(j); else rej(Object.assign(new Error((j&&j.error)||'refus'),{refused:true})); })\n      .withFailureHandler(err=>{ lastErr=String(err&&err.message||err); rej(err); })\n      .envoyer(payload);\n  });\n}\n\n/* File d'attente : uniquement si le réseau manque ; vidée dès l'envoi réussi */\nfunction queue(){try{return JSON.parse(localStorage.getItem(QKEY)||'[]')}catch(e){return[]}}\nfunction setQueue(q){try{q.length?localStorage.setItem(QKEY,JSON.stringify(q)):localStorage.removeItem(QKEY)}catch(e){}}\nasync function flush(){\n  if(!ENDPOINT)return;\n  let q=queue();if(!q.length)return;\n  const rest=[];\n  for(const p of q){try{await send(p)}catch(e){if(!e.refused)rest.push(p)}}\n  setQueue(rest);\n  if(q.length!==rest.length&&!rest.length)msg('Les réclamations en attente ont été envoyées.','ok');\n  else if(rest.length)msg(rest.length+' réclamation(s) en attente d’envoi : elles partiront dès que le réseau revient.','info');\n}\nwindow.addEventListener('online',flush);\n\n$('form').addEventListener('submit',async e=>{\n  e.preventDefault();\n  const v=id=>($(id).value||'').trim(),f=$('form');\n  if(!ENDPOINT)return msg(\"Ce formulaire n'est pas encore relié au tableau du responsable.\",'err');\n  const miss=[];if(!v('agent'))miss.push('nom de l’agent');if(!v('matricule'))miss.push('matricule');if(!v('nom'))miss.push('nom du réclamant');if(!v('adresse'))miss.push('adresse');if(!v('nature'))miss.push('nature');\n  if(miss.length)return msg('Champ obligatoire manquant : '+miss.join(', ')+'.','err');\n  if(v('tel')&&(v('tel').match(/\\d/g)||[]).length<8)return msg('Le numéro de téléphone semble incomplet (8 chiffres minimum).','err');\n  const fuite=v('nature')===\"Fuite d'eau\";\n  if(fuite&&!f.elements.loc.value)return msg('Indiquez la localisation de la fuite.','err');\n  const nat=v('nature');let detail='',depuis='',periode='',cil='',souhaite='';\n  const checked=n=>[...f.querySelectorAll('input[name='+n+']:checked')].map(x=>x.value).join(', ');\n  if(nat==='Demande Coupure'){cil=v('coupureCil');souhaite=dateHeure(v('coupureDate'),v('coupureHeure'))}\n  if(nat==='Demande Rebranchement'){cil=v('rebranchementCil');souhaite=dateHeure(v('rebranchementDate'),v('rebranchementHeure'))}\n  if(nat==='DFO'){detail=f.elements.dfo.value;if(!detail)return msg('Indiquez à quel niveau se trouve le DFO.','err')}\n  if(nat===\"Manque d'eau\"){depuis=dateHeure(v('manqueDate'),v('manqueHeure'))}\n  if(nat==='Appel Client pour Information'){detail=checked('appel');if(!detail)return msg('Cochez la raison de l’appel.','err')}\n  if(nat==='Réfection'){detail=f.elements.refection.value;if(!detail)return msg('Indiquez le type de réfection.','err')}\n  if(nat===\"Qualité d'eau\"){detail=[...f.querySelectorAll('input[name=qualite]:checked')].map(x=>x.value).join(', ');if(!detail)return msg('Cochez la nature du défaut qualité.','err');depuis=v('qualiteDepuis');periode=v('qualitePeriode')}\n  if(nat==='Problème de Pression'){detail=f.elements.pression.value;if(!detail)return msg('Indiquez la nature du défaut de pression.','err');depuis=v('pressionDepuis');periode=v('pressionPeriode')}\n  const payload={clientId:(crypto.randomUUID?crypto.randomUUID():Date.now()+'-'+Math.random()),saisieLe:new Date().toISOString(),\n    agent:v('agent'),matricule:v('matricule'),nom:v('nom'),tel:v('tel'),email:v('email'),adresse:v('adresse'),quartier:v('quartier'),ville:v('ville'),\n    lat:pos?pos.lat:'',lng:pos?pos.lng:'',precision:pos?Math.round(pos.acc):'',\n    nature:v('nature'),localisation:fuite?f.elements.loc.value:'',importance:fuite?(f.elements.imp.value||''):'',detail,depuis,periode,cil,souhaite,obs:v('obs'),photo:photoData,vignette:thumbData};\n  const b=$('btnSave');b.disabled=true;msg('Envoi en cours…','info');\n  try{\n    const j=await send(payload);\n    f.reset();setTimeout(()=>msg(j.numero?'Réclamation envoyée. N° Saisie : '+j.numero:'Réclamation envoyée au responsable.','ok'),10);\n  }catch(err){\n    if(err.refused){msg('Réclamation refusée : '+err.message,'err')}\n    else{const q=queue();q.push(payload);setQueue(q);f.reset();\n      setTimeout(()=>msg(navigator.onLine?\"Envoi impossible pour le moment : la réclamation est gardée sur ce téléphone et sera renvoyée automatiquement à la prochaine ouverture du formulaire.\"+(lastErr?' Détail technique : '+lastErr:''):\"Pas de réseau : la réclamation est gardée sur ce téléphone et partira automatiquement dès que le réseau revient.\",'info'),10)}\n  }finally{b.disabled=false}\n});\nflush();\n</script>\n</body>\n</html>\n";
