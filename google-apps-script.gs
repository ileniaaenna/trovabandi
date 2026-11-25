/**
 * Trova bandi – raccolta richieste in Google Sheets
 * Incolla questo codice in: Foglio Google > Estensioni > Apps Script
 * Poi: Esegui il deploy > Nuovo deployment > App web
 *      Esegui come: Me · Chi ha accesso: Chiunque
 */

const EMAIL_AVVISO = 'ileniaaenna@gmail.com';   // dove ricevi l'avviso di ogni nuova richiesta
const NOME_FOGLIO = 'Richieste';

const COLONNE = ['Data', 'Nome', 'Email', 'Telefono', 'Azienda', 'Messaggio',
  'Tipo soggetto', 'Dimensione', 'Anni', 'Regione', 'Settore', 'Investimento €',
  'Domanda entro', 'Spese avviate', 'Spese', 'Forma preferita',
  'Bandi compatibili', 'Bandi con condizioni', 'Consenso privacy', 'Pagina'];

function foglio_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(NOME_FOGLIO);
  if (!sh) {
    sh = ss.insertSheet(NOME_FOGLIO);
    sh.appendRow(COLONNE);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, COLONNE.length).setFontWeight('bold');
  }
  return sh;
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents || '{}');
    if (d.sito) return risposta_({ ok: true });             // campo trappola anti-spam
    if (!d.email || !d.consenso) return risposta_({ ok: false, errore: 'dati mancanti' });

    const t = v => String(v == null ? '' : v).slice(0, 2000);
    const riga = [new Date(), t(d.nome), t(d.email), t(d.telefono), t(d.azienda), t(d.messaggio),
      t(d.soggetto), t(d.dimensione), t(d.anni), t(d.regione), t(d.settore), t(d.importo),
      t(d.entro), d.avviato ? 'Sì' : 'No', t(d.spese), t(d.forme),
      t(d.compatibili), t(d.condizioni), d.consenso ? 'Sì' : 'No', t(d.pagina)];

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    foglio_().appendRow(riga);
    lock.releaseLock();

    MailApp.sendEmail({
      to: EMAIL_AVVISO,
      replyTo: t(d.email),
      subject: 'Trova bandi – nuova richiesta da ' + (t(d.nome) || t(d.email)),
      body: COLONNE.map((c, i) => c + ': ' + (riga[i] instanceof Date ? riga[i].toLocaleString('it-IT') : riga[i])).join('\n') +
        '\n\nApri il foglio: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl()
    });
    return risposta_({ ok: true });
  } catch (err) {
    return risposta_({ ok: false, errore: String(err) });
  }
}

function doGet() { return risposta_({ ok: true, servizio: 'Trova bandi' }); }

function risposta_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/** Esegui una volta a mano per creare il foglio e autorizzare l'invio email. */
function prova() {
  foglio_();
  MailApp.sendEmail(EMAIL_AVVISO, 'Trova bandi – collegamento riuscito', 'Il foglio Richieste è pronto.');
}
