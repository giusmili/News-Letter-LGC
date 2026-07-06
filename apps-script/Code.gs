/**
 * Newsletter La Grande Classe — envoi via Google Apps Script + Gmail.
 *
 * Prérequis :
 * 1. Ce script doit être lié à une Google Sheet ("Extensions > Apps Script")
 *    contenant un onglet nommé "Abonnés" avec les colonnes :
 *    A: Email | B: Prénom | C: Date d'envoi (laissée vide au départ)
 * 2. news.html doit être poussé sur la branche "master" du repo GitHub public
 *    (giusmili/News-Letter-LGC) pour que RAW_HTML_URL reste à jour automatiquement.
 
 */

const RAW_HTML_URL =
  'https://raw.githubusercontent.com/giusmili/News-Letter-LGC/master/news.html';
const SUJET = "Newsletter La Grande Classe - Octobre 2025";

function envoyerNewsletter() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Abonnés');
  const data = sheet.getDataRange().getValues();
  const htmlBody = UrlFetchApp.fetch(RAW_HTML_URL, { muteHttpExceptions: true }).getContentText();

  for (let i = 1; i < data.length; i++) {
    const [email, prenom, dejaEnvoye] = data[i];
    if (!email || dejaEnvoye) continue;

    MailApp.sendEmail({
      to: email,
      subject: SUJET,
      htmlBody: htmlBody,
      name: 'La Grande Classe',
    });

    sheet.getRange(i + 1, 3).setValue(new Date());
    Utilities.sleep(1000); // marge de sécurité vis-à-vis des quotas d'envoi
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Newsletter')
    .addItem('Envoyer la newsletter', 'envoyerNewsletter')
    .addToUi();
}
