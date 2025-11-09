# Gestione Newsletter

Questo repository contiene lo script Google Apps Script utilizzato per gestire le richieste ricevute via email dagli iscritti alla newsletter di FIM Toscana.

## File principali
- `Code.gs`: logica completa per leggere le email con etichetta `newsletter-manage`, interpretare le richieste di sospensione, riattivazione, cancellazione e cambio frequenza, aggiornare il foglio di lavoro `Mail` (ID `1LDqX7WScxsP4LM6mzSjmcgX73i1ZPKVZOVjaQI1e9sU`) e registrare le modifiche sul foglio di log `logml` del file `Log Rassegna` (ID `16sJ4NTmvJqZo6P19wUR1tLg9AF80KKF0vU6UY631JHw`).

## Configurazione
1. Creare in Gmail un filtro che applichi l'etichetta `newsletter-manage` ai messaggi generati dai pulsanti di gestione presenti nelle newsletter.
2. In Apps Script creare un trigger a tempo (es. ogni 5 minuti) che invochi `processNewsletterManagementRequests`.
3. Verificare che i fogli `Mail` e `logml` esistano e abbiano rispettivamente le colonne `Mail | Status | Scheduling | Tipologia` e `Data | Mail | Status | Scheduling`.

## Log delle operazioni
Ogni modifica va a registrare una riga nel foglio `logml` con:
- **Data**: timestamp dell’operazione
- **Mail**: indirizzo dell’iscritto coinvolto
- **Status**: valore aggiornato dopo l’operazione (`attivo`, `sospeso` o `cancellato`)
- **Scheduling**: frequenza attuale (`Giornaliero`, `Settimanale` o vuoto se non applicabile)
