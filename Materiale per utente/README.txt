Sunrise + Astra Input Builder — istruzioni per l'utente
=========================================================

Requisiti
---------
• Windows 10 o 11 (64 bit).
• Connessione internet SOLO al primo avvio (per scaricare Python
  e le librerie necessarie, circa 100 MB).
• Nessuna installazione preliminare: né Python, né uv, né altro.

Come si avvia
-------------
1. Scompattare la cartella ricevuta in un percorso qualsiasi
   (per esempio C:\Users\<utente>\Documents\excel-pipeline).
2. Doppio click su launch.bat.
3. Al primo avvio comparirà una finestra "prompt dei comandi" che:
   – scarica Python 3.11.9 nella sotto-cartella python\;
   – installa pip e le dipendenze (streamlit, openpyxl, pandas).
   Il tutto richiede 1–2 minuti la prima volta.
4. Terminata la preparazione, si apre automaticamente il browser
   di default sull'indirizzo http://localhost:8501, con la pagina
   Sunrise pronta all'uso.
5. Per navigare fra Sunrise e Astra usare il menu di sinistra.

Come si chiude
--------------
Chiudere la finestra "prompt dei comandi" del launcher. La porta
locale viene liberata automaticamente.

Aggiornamenti
-------------
La versione fornita è definitiva; non richiede aggiornamenti.
Se in futuro venisse rilasciata una nuova versione, ne verrà
consegnata una nuova cartella da usare al posto di questa.

Struttura della cartella (dopo il primo avvio)
-----------------------------------------------
  app.py                    entry point dell'app (routing multi-pagina)
  excel_pipeline/           logica di elaborazione (pura, deterministica)
  pages/                    UI Streamlit (pagine Sunrise e Astra)
  pyproject.toml            elenco delle dipendenze
  python\                   interprete Python 3.11.9 embedded (creato al 1° avvio)
  runs\                     una sotto-cartella per ogni esecuzione (creata a runtime)
  user_prefs.json           preferenze utente persistenti (percorsi di output)
  launch.bat                lo script da lanciare
  README.txt                questo file

Note operative
--------------
• Ogni volta che si preme "▶ Run pipeline" viene creata una nuova
  sotto-cartella dentro runs\<YYYY-MM-DD_HHMMSS>\ contenente i file
  di input caricati, gli output prodotti, e il report di validazione.
• I percorsi indicati per il "salvataggio massivo" nelle sezioni 4
  di ciascuna pagina vengono ricordati fra una sessione e l'altra.

Risoluzione problemi
--------------------
• Se il browser non si apre automaticamente: aprire manualmente
  http://localhost:8501 (o aggiornare la pagina qualora fosse aperta).
• Se il primo avvio fallisce sul download di Python: verificare la
  connessione internet e rilanciare launch.bat.
• Se rimane bloccato all'avvio con una finestra nera aperta, chiuderla
  e rilanciare (il setup è idempotente: riprende dallo step giusto).
