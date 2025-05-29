NeuraDox 1.0 - Documentazione

NeuraDox 1.0 è un'applicazione dotata di interfaccia grafica, realizzata con Tkinter, che automatizza il processo di copia e rinominazione dei file (in questo caso, file PDF) sulla base di un file Excel. L'applicazione offre anche la possibilità di creare un archivio ZIP dei file elaborati e include una guida operativa completa.

1. Prerequisiti
1.1 Python
Versione consigliata: Python 3.6 o superiore.

Nota su Tkinter:

Su Windows e macOS, Tkinter è generalmente incluso nell'installazione di Python.

Su Linux (ad esempio Debian/Ubuntu) potrebbe essere necessario installare Tkinter separatamente:

bash
sudo apt-get install python3-tk
1.2 Dipendenze Esterne
Il codice utilizza alcune librerie non incluse nella libreria standard di Python:

Pandas: Per gestire la lettura e manipolazione dei file Excel.

Openpyxl: Per leggere file Excel in formato .xlsx.

Xlrd: Per leggere file Excel in formato .xls. Nota: Per i file .xls è consigliata la versione 1.2.0 poiché le versioni successive non supportano questo formato.

Assicurati di installare anche queste librerie.

2. Installazione delle Dipendenze
Per installare le dipendenze richieste, utilizza pip. Esegui il seguente comando in una console o terminale:

bash
pip install pandas openpyxl xlrd==1.2.0
Questo comando installerà:

pandas: per le operazioni sui DataFrame e per l'importazione dei dati dal file Excel.

openpyxl: per leggere file .xlsx.

xlrd (versione 1.2.0): per leggere file .xls.

3. Setup ed Esecuzione dell'Applicazione
3.1 Scarica il Codice
Salva il codice in un file, ad esempio neuradox.py.

3.2 Esecuzione
Per avviare l'applicazione, esegui il seguente comando da terminale o prompt dei comandi:

bash
python neuradox.py
Se sul tuo sistema il comando corretto è python3, utilizza:

bash
python3 neuradox.py
In questo modo verrà avviata la GUI dell'applicazione.

4. Guida Operativa
4.1 File Excel
Il file Excel deve contenere almeno 5 colonne.

Note importanti:

Colonna 2: Si assume contenga il nome originale dei file (senza estensione).

Colonna 5: Dovrà contenere il nuovo nome che il file deve assumere (senza estensione).

Nota: Lo script aggiunge automaticamente l'estensione .pdf sia al file originale che a quello rinominato.

4.2 Selezione Directory e File
All'interno dell'interfaccia grafica sono presenti le seguenti sezioni:

Sezione File Excel: Consente di selezionare il file Excel tramite un file dialog.

Sezione File da Elaborare: Permette di scegliere la cartella contenente i file (es. file PDF) da processare.

Sezione File Elaborati: Specifica la cartella in cui verranno copiati e rinominati i file.

4.3 Funzionalità Aggiuntive
Avvia Elaborazione: Avvia il processo che, per ogni riga del file Excel, copia il file (se presente) dalla cartella di origine a quella di destinazione rinominandolo. Una barra di avanzamento e un'etichetta di stato mostrano il progresso dell'operazione.

Scarica ZIP: Permette di creare un archivio ZIP della cartella contenente i file elaborati.

Guida: Visualizza una guida rapida con le istruzioni d'uso, informazioni sulla licenza (GNU GPLv3) e i contatti dello sviluppatore.

4.4 Gestione degli Errori
Se il file Excel non è valido, la cartella di origine o quella di destinazione non esiste, oppure mancano file necessari all'operazione, l'applicazione mostrerà messaggi di errore tramite finestre di dialogo (utilizzando messagebox). L'aggiornamento del progresso e gli errori (ad esempio, file mancanti) vengono inoltre visualizzati in una finestra di log (stato).

5. Distribuzione
Se intendi distribuire l'applicazione a utenti che non dispongono di Python o delle relative dipendenze, puoi convertire lo script in un eseguibile. Uno strumento molto utile a questo scopo è PyInstaller.

5.1 Creazione di un Eseguibile
Installa PyInstaller (se non è già presente):

bash
pip install pyinstaller
Converti lo script in un eseguibile con il seguente comando:

bash
pyinstaller --onefile neuradox.py
Questo comando creerà un file eseguibile all'interno della cartella dist/, che potrai distribuire agli utenti.

6. Conclusioni
Questa documentazione descrive in dettaglio come configurare l'ambiente, installare le dipendenze e avviare NeuraDox 1.0. Seguendo i passaggi illustrati, dovresti essere in grado di utilizzare l'applicazione senza problemi.

Suggerimenti per possibili sviluppi futuri:

Implementare un sistema di logging più dettagliato per una diagnostica avanzata.

Gestire formati di file Excel più complessi o espandere il mapping dei dati.

Rifinire l'impacchettamento dell'applicativo per semplificare ulteriormente la distribuzione.