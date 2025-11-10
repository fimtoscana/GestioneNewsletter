# Gestione Newsletter

Questo repository contiene lo script Google Apps Script utilizzato per gestire le richieste ricevute via email dagli iscritti alla newsletter di FIM Toscana.

## File principali
- `Code.gs`: logica completa per leggere le email non lette con oggetto corrispondente alle azioni supportate, interpretare le richieste di sospensione, riattivazione, cancellazione e cambio frequenza, aggiornare il foglio di lavoro `Mail` (ID `1LDqX7WScxsP4LM6mzSjmcgX73i1ZPKVZOVjaQI1e9sU`) e registrare le modifiche sul foglio di log `logml` del file `Log Rassegna` (ID `16sJ4NTmvJqZo6P19wUR1tLg9AF80KKF0vU6UY631JHw`).

## Configurazione
1. In Apps Script creare un trigger a tempo (es. ogni 5 minuti) che invochi `processNewsletterManagementRequests`.
2. Verificare che i fogli `Mail` e `logml` esistano e abbiano rispettivamente le colonne `Mail | Status | Scheduling | Tipologia` e `Data | Mail | Azione | Status precedente | Status nuovo | Frequenza precedente | Frequenza nuova`.
3. Accertarsi che i pulsanti inseriti nelle newsletter generino email con oggetto nel formato previsto (es. `Sospendi Newsletter email indirizzo@example.com`). Lo script cercherà automaticamente queste email non lette, le processerà, le segnerà come lette e archivierà i thread gestiti correttamente.

## Log delle operazioni
Ogni evento genera una riga nel foglio `logml` con i seguenti campi:
- **Data**: timestamp dell’operazione.
- **Mail**: indirizzo dell’iscritto coinvolto.
- **Azione**: tipo di richiesta gestita (`suspend`, `resume`, `cancel`, `changeFrequency`).
- **Status precedente**: stato registrato prima dell’elaborazione (`attivo`, `sospeso`, `cancellato` o vuoto).
- **Status nuovo**: stato risultante dopo l’operazione.
- **Frequenza precedente**: frequenza pianificata prima della richiesta (`Giornaliero`, `Settimanale` o vuoto).
- **Frequenza nuova**: frequenza impostata dopo la gestione (vuota quando non applicabile o dopo una cancellazione).
