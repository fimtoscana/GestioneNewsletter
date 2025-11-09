# Linee guida per GestioneNewsletter

- Mantieni la documentazione in italiano e aggiorna il `README.md` quando introduci nuove funzionalità o comandi.
- Se aggiungi codice Python, formatta i file con [Black](https://black.readthedocs.io/en/stable/) (line length 88) e ordina gli import con [isort](https://pycqa.github.io/isort/).
- Quando implementi nuova logica applicativa, aggiungi test automatici usando `pytest` quando possibile; se non puoi, spiega il motivo nel riepilogo finale.
- Elenca le nuove dipendenze Python nel file `requirements.txt` e includi eventuali istruzioni di setup nel `README.md`.
- Mantieni i testi destinati all'utente finale in italiano, salvo differenti specifiche del ticket.
