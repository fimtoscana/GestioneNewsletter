/**
 * Configurazione per la gestione delle richieste di newsletter.
 */
const SHEET_ID = '1LDqX7WScxsP4LM6mzSjmcgX73i1ZPKVZOVjaQI1e9sU';
const SHEET_NAME = 'Mail';
const MANAGEMENT_QUERY =
  'is:unread (subject:"Sospendi Newsletter email" OR subject:"Riattiva Newsletter email" OR subject:"Cancella Newsletter email" OR subject:"Modifica frequenza Newsletter email")';

const LOG_SHEET_ID = '16sJ4NTmvJqZo6P19wUR1tLg9AF80KKF0vU6UY631JHw';
const LOG_SHEET_NAME = 'logml';
const FEEDBACK_ALIAS = 'rassegna@fimtoscana.it';
const FIM_LOGO_URL =
  'https://ci3.googleusercontent.com/meips/ADKq_NYLN_GXFoAhBFQyQ4YoHUUcoSZV8ECPvMmqXcLLrdtJ5ATxyRV4Cc4L0rnzRSdjF8kVGSOdDCrfynnntiR1D78quNnReMeskQ0zS54bHahNVkzO-Tbv6zGz30ADuwFG3X_5ch_W3D9XasACyw=s0-d-e1-ft#https://www.fim-cisl.it/wp-content/uploads/2015/01/Anteprima-Logo-Fim-nazionale-FB.jpg';

const STATUS = {
  ACTIVE: 'attivo',
  SUSPENDED: 'sospeso',
  CANCELLED: 'cancellato',
};

const FREQUENCIES = ['Giornaliero', 'Settimanale'];

const SUBJECT_TEMPLATES = {
  suspend: /^Sospendi Newsletter email\s+(.+)$/i,
  resume: /^Riattiva Newsletter email\s+(.+)$/i,
  cancel: /^Cancella Newsletter email\s+(.+)$/i,
  changeFrequency: /^Modifica frequenza Newsletter email\s+(.+)\s+da\s+(Giornaliero|Settimanale)\s+a\s+(Giornaliero|Settimanale)$/i,
};

/**
 * Entry point richiamato da un trigger a tempo (es. ogni 5 minuti).
 */
function processNewsletterManagementRequests() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    Logger.log('Impossibile ottenere il lock per la gestione delle richieste: uscita senza elaborazione.');
    return;
  }

  let retryCount = 0;

  try {
    const threads = fetchAllMatchingThreads(MANAGEMENT_QUERY);
    threads.forEach((thread) => {
      let threadHasErrors = false;
      let processedMessages = false;

      thread.getMessages().forEach((message) => {
        if (!message.isUnread()) return;

        try {
          handleMessage(message);
          message.markRead();
          processedMessages = true;
        } catch (err) {
          threadHasErrors = true;
          retryCount += 1;
          const sender = extractEmailAddress(message.getFrom());
          Logger.log(
            'Errore nel processare il thread ' +
              thread.getId() +
              ' "' +
              message.getSubject() +
              '": ' +
              err.message +
              '. retryPending=' +
              retryCount
          );
          if (sender) {
            notifyError(
              sender,
              'Si è verificato un problema nella gestione della tua richiesta. Il sistema ritenterà automaticamente.'
            );
          }
        }
      });

      if (processedMessages && !threadHasErrors) {
        thread.moveToArchive();
      }
    });

    if (retryCount > 0) {
      Logger.log('Messaggi che richiederanno un nuovo tentativo: ' + retryCount);
    }

  } finally {
    lock.releaseLock();
  }
}

/**
 * Recupera tutti i thread che soddisfano la query Gmail indicata.
 */
function fetchAllMatchingThreads(query) {
  const threads = [];
  const batchSize = 100;
  let start = 0;

  while (true) {
    const batch = GmailApp.search(query, start, batchSize);
    if (!batch.length) {
      break;
    }

    batch.forEach((thread) => threads.push(thread));

    if (batch.length < batchSize) {
      break;
    }

    start += batchSize;
  }

  return threads;
}

/**
 * Analizza il messaggio e applica l'azione richiesta.
 */
function handleMessage(message) {
  const subject = (message.getSubject() || '').trim();
  const sender = extractEmailAddress(message.getFrom());
  const request = parseRequest(subject);

  if (!request) {
    notifyError(sender, 'Il sistema non ha riconosciuto la richiesta. Usa i pulsanti nel footer della newsletter.');
    return;
  }

  if (request.email.toLowerCase() !== sender.toLowerCase()) {
    notifyError(sender, 'L’indirizzo indicato nel messaggio non coincide con il mittente.');
    return;
  }

  const sheet = getSheet(SHEET_ID, SHEET_NAME);
  const record = findSubscriberRow(sheet, request.email);
  if (!record) {
    notifyError(sender, 'Il tuo indirizzo non risulta tra gli iscritti. Contatta l’associazione per assistenza.');
    return;
  }

  const previousStatus = record.status;
  const previousFrequency = record.frequency;

  switch (request.action) {
    case 'suspend':
      sheet.getRange(record.row, 2).setValue(STATUS.SUSPENDED);
      logChange(
        request.email,
        'suspend',
        previousStatus,
        STATUS.SUSPENDED,
        previousFrequency,
        previousFrequency
      );
      notifySuccess(sender, 'Invio sospeso correttamente.');
      break;

    case 'resume':
      sheet.getRange(record.row, 2).setValue(STATUS.ACTIVE);
      logChange(
        request.email,
        'resume',
        previousStatus,
        STATUS.ACTIVE,
        previousFrequency,
        previousFrequency
      );
      notifySuccess(sender, 'Invio riattivato correttamente.');
      break;

    case 'cancel': {
      sheet.deleteRow(record.row);
      logChange(
        request.email,
        'cancel',
        previousStatus,
        STATUS.CANCELLED,
        previousFrequency,
        ''
      );
      notifySuccess(sender, 'Iscrizione cancellata definitivamente.');
      break;
    }

    case 'changeFrequency': {
      const currentFrequency = (record.frequency || '').toString().trim();
      const normalizedCurrent = currentFrequency.toLowerCase();
      const normalizedRequested = request.newFrequency.toLowerCase();

      if (normalizedCurrent === normalizedRequested && currentFrequency !== '') {
        logChange(
          request.email,
          'changeFrequency',
          previousStatus,
          previousStatus,
          previousFrequency,
          previousFrequency
        );
        notifySuccess(
          sender,
          'La frequenza è già impostata su ' + request.newFrequency + '. Nessuna modifica necessaria.'
        );
        break;
      }

      sheet.getRange(record.row, 3).setValue(request.newFrequency);
      logChange(
        request.email,
        'changeFrequency',
        previousStatus,
        previousStatus,
        previousFrequency,
        request.newFrequency
      );
      notifySuccess(sender, 'Frequenza aggiornata a ' + request.newFrequency + '.');
      break;
    }

    default:
      notifyError(sender, 'Richiesta non supportata.');
  }
}

/**
 * Ritorna il foglio indicato e lancia un errore se non esiste.
 */
function getSheet(id, name) {
  const spreadsheet = SpreadsheetApp.openById(id);
  const sheet = spreadsheet.getSheetByName(name);
  if (!sheet) {
    throw new Error('Il foglio "' + name + '" non è stato trovato nel file ' + id + '.');
  }
  return sheet;
}

/**
 * Cerca il record (dalla riga 2 in poi) corrispondente a un indirizzo email.
 */
function findSubscriberRow(sheet, email) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  const target = email.toLowerCase();

  for (let i = 0; i < data.length; i++) {
    const rowMail = (data[i][0] || '').toString().trim().toLowerCase();
    if (rowMail === target) {
      return {
        row: i + 2,
        status: (data[i][1] || '').toString().trim(),
        frequency: (data[i][2] || '').toString().trim(),
      };
    }
  }

  return null;
}

/**
 * Effettua il parsing del subject per riconoscere l’azione richiesta.
 */
function parseRequest(subject) {
  let match;

  match = SUBJECT_TEMPLATES.suspend.exec(subject);
  if (match) return { action: 'suspend', email: match[1].trim() };

  match = SUBJECT_TEMPLATES.resume.exec(subject);
  if (match) return { action: 'resume', email: match[1].trim() };

  match = SUBJECT_TEMPLATES.cancel.exec(subject);
  if (match) return { action: 'cancel', email: match[1].trim() };

  match = SUBJECT_TEMPLATES.changeFrequency.exec(subject);
  if (match) {
    const requested = match[3].trim().toLowerCase();
    const normalizedFrequencies = FREQUENCIES.map((frequency) => frequency.toLowerCase());
    const frequencyIndex = normalizedFrequencies.findIndex((frequency) => frequency === requested);
    if (frequencyIndex === -1) return null;
    const canonicalFrequency = FREQUENCIES[frequencyIndex];
    return {
      action: 'changeFrequency',
      email: match[1].trim(),
      currentFrequency: match[2].trim(),
      newFrequency: canonicalFrequency,
    };
  }

  return null;
}

/**
 * Estrae l’indirizzo email dal campo “From”.
 */
function extractEmailAddress(from) {
  const match = /<([^>]+)>/.exec(from);
  return (match ? match[1] : from).trim();
}

/**
 * Invia un’email di conferma al mittente.
 */
function buildFeedbackHtml(primaryMessage, secondaryMessage) {
  let html =
    '<div style="font-family:Arial,sans-serif;line-height:1.4;color:#333;">' +
    '<div style="margin-bottom:16px;">' +
    '<img src="' +
    FIM_LOGO_URL +
    '" alt="Logo FIM" style="max-width:220px;height:auto;" />' +
    '</div>' +
    '<p style="margin:0 0 12px 0;">' +
    primaryMessage +
    '</p>';

  if (secondaryMessage) {
    html += '<p style="margin:0;">' + secondaryMessage + '</p>';
  }

  html += '</div>';
  return html;
}

function notifySuccess(recipient, message) {
  const htmlBody = buildFeedbackHtml(message, 'Grazie per averci contattato.');

  GmailApp.sendEmail(recipient, 'Gestione Newsletter - conferma', message, {
    from: FEEDBACK_ALIAS,
    htmlBody: htmlBody,
  });
}

/**
 * Invia un’email di errore al mittente.
 */
function notifyError(recipient, message) {
  const fallbackMessage = message + '\n\nSe il problema persiste, rispondi a questa email.';
  const htmlBody = buildFeedbackHtml(message, 'Se il problema persiste, rispondi a questa email.');

  GmailApp.sendEmail(recipient, 'Gestione Newsletter - errore', fallbackMessage, {
    from: FEEDBACK_ALIAS,
    htmlBody: htmlBody,
  });
}

/**
 * Registra la modifica effettuata sul foglio di log.
 */
function logChange(email, action, oldStatus, newStatus, oldFrequency, newFrequency) {
  const sheet = getSheet(LOG_SHEET_ID, LOG_SHEET_NAME);
  const timestamp = new Date();
  sheet.appendRow([
    timestamp,
    email,
    action,
    oldStatus || '',
    newStatus || '',
    oldFrequency || '',
    newFrequency || '',
  ]);
}
