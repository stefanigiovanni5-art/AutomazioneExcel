# Smart Sales Report

Smart Sales Report è un'applicazione Python progettata per automatizzare l'analisi dei dati di vendita e la creazione di report Excel professionali.

Il programma importa file CSV o Excel, controlla e pulisce i dati, calcola automaticamente i principali KPI aziendali e genera un report Excel completo con dashboard, statistiche e grafici.

## Funzionalità

- Importazione di file CSV e XLSX
- Pulizia e validazione automatica dei dati
- Rimozione dei record duplicati
- Analisi dei valori mancanti
- Riconoscimento delle colonne necessarie all'analisi
- Calcolo automatico del fatturato
- Calcolo dei principali KPI di vendita
- Analisi delle vendite per prodotto
- Analisi delle vendite per categoria
- Creazione automatica di una dashboard Excel
- Generazione automatica di grafici
- Statistiche descrittive
- Controllo della qualità dei dati
- Esportazione del report finale in formato XLSX
- Apertura automatica del report generato

## KPI calcolati

Il programma calcola automaticamente:

- Fatturato totale
- Quantità totale venduta
- Numero di operazioni
- Valore medio per operazione
- Prodotto con il maggior fatturato

## Report generato

Il file Excel prodotto contiene diversi fogli dedicati all'analisi:

- Dashboard
- Dati Puliti
- Statistiche
- Qualità Dati
- Informazioni
- Vendite per Prodotto
- Vendite per Categoria

La dashboard include grafici per facilitare la lettura dei risultati e l'individuazione dei prodotti e delle categorie con il maggiore fatturato.

## Tecnologie utilizzate

- Python
- Pandas
- OpenPyXL
- Tkinter
- Microsoft Excel
- Git
- GitHub

## Installazione

Clonare il repository:

```bash
git clone https://github.com/stefanigiovanni5-art/AutomazioneExcel.git
```

Entrare nella cartella del progetto:

```bash
cd AutomazioneExcel
```

Installare le dipendenze:

```bash
python -m pip install -r requirements.txt
```

## Utilizzo

Avviare il programma con:

```bash
python report.py
```

Selezionare il file di vendita da analizzare quando richiesto dal programma.

Al termine dell'elaborazione verrà generato automaticamente il report Excel con i risultati dell'analisi.

## Struttura del progetto

```text
AutomazioneExcel/
├── report.py
├── requirements.txt
├── vendite.csv
├── README.md
└── .gitignore
```

`vendite.csv` contiene dati dimostrativi utilizzabili per provare il programma.

## Privacy dei dati

I dati reali dei clienti non devono essere pubblicati nel repository.

Prima di utilizzare il progetto con dati aziendali reali, verificare che file contenenti informazioni riservate, personali o commercialmente sensibili siano esclusi dal controllo versione.

## Possibili sviluppi futuri

- Supporto a ulteriori formati di dati
- Dashboard più avanzate
- Analisi temporali delle vendite
- Filtri e configurazioni personalizzabili
- Generazione di ulteriori KPI
- Miglioramento dell'interfaccia grafica

## Autore

Progetto sviluppato come soluzione Python per l'automazione dell'analisi delle vendite e della reportistica Excel.