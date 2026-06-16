// --- GetFit Challenge - Google Apps Script Backend (Version Ultra-Sécurisée) ---
const ACCESS_PIN = '48960'; // Ton code d'accès privé
const STATE_SHEET = 'STATE_HISTORY';
const AUDIT_SHEET = 'AUDIT_LOG';
const BACKUP_SHEET = 'BACKUPS_QUOTIDIENS';

// ─── LECTURE DES DONNÉES (doGet) ───
function doGet(e) {
  if (e.parameter.pin !== ACCESS_PIN) {
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: 'Unauthorized'}))
                         .setMimeType(ContentService.MimeType.JSON);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = getOrCreateSheet(ss, STATE_SHEET);
  const lastRow = sh.getLastRow();

  let val = '{}';
  if (lastRow > 1) {
    val = sh.getRange(lastRow, 2).getValue();
  }

  return ContentService.createTextOutput(val || '{}')
                       .setMimeType(ContentService.MimeType.JSON);
}

// ─── ÉCRITURE DES DONNÉES (doPost) ───
function doPost(e) {
  // Verrou : sérialise toutes les écritures concurrentes (3 iPhones en même temps)
  // pour qu'aucune ne se base sur une lecture périmée et n'efface le travail d'un autre.
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (err) {
    return errorResponse('Serveur occupé, réessaie dans un instant.');
  }

  try {
    const payload = JSON.parse(e.postData.contents);

    // 1. Vérification de sécurité
    if (!payload || (payload.pin !== ACCESS_PIN && e.parameter.pin !== ACCESS_PIN)) {
      return errorResponse('Unauthorized');
    }

    if (!payload.state || !payload.state.data || !payload.state.config) {
      return errorResponse('Payload corrompu : données manquantes. Rejeté par le serveur.');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const stateSheet = getOrCreateSheet(ss, STATE_SHEET);
    const auditSheet = getOrCreateSheet(ss, AUDIT_SHEET, ['Horodatage', 'Action', 'Mois concerné', 'Personne', 'Poids (kg)', 'Calories (kcal/j)', 'Gain/Perte (€)', 'Validation', 'Note/Source']);
    const backupSheet = getOrCreateSheet(ss, BACKUP_SHEET, ['Date de Backup', 'Snapshot JSON complet']);

    const now = new Date();
    const nowStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
    const todayStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy');

    // 2. Lecture de l'état VRAIMENT le plus récent, sous verrou.
    const lastRow = stateSheet.getLastRow();
    const lastStateStr = lastRow > 1 ? stateSheet.getRange(lastRow, 2).getValue() : '';
    const base = lastStateStr ? JSON.parse(lastStateStr) : {config: {}, data: {}, manual: {}, trips: []};
    if (!base.data) base.data = {};
    if (!base.manual) base.manual = {};
    if (!base.trips) base.trips = [];

    // 3. Fusion atomique : on n'applique JAMAIS l'état complet envoyé par un téléphone
    //    (qui peut être basé sur une copie locale périmée). On applique seulement :
    //    a) les changements explicites par personne/mois (changes[])
    //    b) les sections (trips/config) si elles diffèrent réellement de la base actuelle
    const changes = payload.changes || [];
    changes.forEach(function (ch) {
      if (!base.data[ch.month]) base.data[ch.month] = {};
      base.data[ch.month][ch.person] = {kg: ch.kg, kcal: ch.kcal, amount: ch.amount};
      if (ch.note === 'manual override') {
        if (!base.manual[ch.month]) base.manual[ch.month] = {};
        base.manual[ch.month][ch.person] = ch.amount;
      }
    });

    const incomingTripsStr = JSON.stringify(payload.state.trips || []);
    if (incomingTripsStr !== JSON.stringify(base.trips)) {
      base.trips = payload.state.trips || [];
    }
    const incomingConfigStr = JSON.stringify(payload.state.config || {});
    if (incomingConfigStr !== JSON.stringify(base.config)) {
      base.config = payload.state.config;
    }

    const mergedStateStr = JSON.stringify(base);

    // 4. Bouclier anti-corruption, appliqué sur le résultat de la fusion (pas sur l'envoi brut)
    if (lastStateStr && lastStateStr.length > 500 && mergedStateStr.length < (lastStateStr.length * 0.7)) {
      return errorResponse('Alerte de corruption : Les nouvelles données sont anormalement réduites par rapport à l\'historique. Écriture bloquée.');
    }

    // 5. SAUVEGARDE EN MODE "APPEND-ONLY" (On n'écrase jamais)
    if (lastRow === 0) { stateSheet.appendRow(['Horodatage', 'Etat Global (JSON)']); stateSheet.getRange('1:1').setFontWeight('bold').setBackground('#e0e0e0'); }
    stateSheet.appendRow([nowStr, mergedStateStr]);

    // 6. AUDIT LOG HUMAINEMENT LISIBLE
    if (changes.length > 0) {
      const logData = changes.map(function (ch) {
        return [
          nowStr,
          'Saisie / Edition',
          ch.month,
          ch.person ? ch.person.toUpperCase() : 'Inconnu',
          ch.kg != null ? ch.kg : '—',
          ch.kcal != null ? ch.kcal : '—',
          ch.amount != null ? ch.amount + ' €' : '—',
          ch.pass === true ? '✓ Succès' : (ch.pass === false ? '✗ Echec' : 'En attente'),
          ch.note || 'saisie manuelle'
        ];
      });
      auditSheet.getRange(auditSheet.getLastRow() + 1, 1, logData.length, logData[0].length).setValues(logData);
    }

    // 7. SNAPSHOTS QUOTIDIENS (Backups)
    const lastBackupRow = backupSheet.getLastRow();
    let lastBackupDate = '';
    if (lastBackupRow > 1) {
      lastBackupDate = backupSheet.getRange(lastBackupRow, 1).getValue();
    }
    if (lastBackupDate !== todayStr) {
      backupSheet.appendRow([todayStr, mergedStateStr]);
    }

    return ContentService.createTextOutput(JSON.stringify({status: 'success', timestamp: nowStr}))
                         .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return errorResponse('Erreur serveur : ' + err.toString());
  } finally {
    lock.releaseLock();
  }
}

// ─── FONCTIONS UTILITAIRES ───

function getOrCreateSheet(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (headers) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e0e0e0');
      sheet.setFrozenRows(1);
      for (let i = 1; i <= headers.length; i++) {
        sheet.setColumnWidth(i, 150);
      }
    }
  }
  return sheet;
}

function errorResponse(msg) {
  return ContentService.createTextOutput(JSON.stringify({status: 'error', message: msg}))
                       .setMimeType(ContentService.MimeType.JSON);
}
