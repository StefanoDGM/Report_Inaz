# Report Commesse

Strumento Windows per elaborare il file di stampa commesse mensile.

## Cosa fa

Il programma legge il file `.xlsx` piu recente presente nella cartella `input` e crea un file elaborato nella cartella `output`.

Il file generato contiene 2 fogli:

1. `Stampa Commesse Dipendente`
   - copia il foglio sorgente
   - converte le ore da formato `hh:mm` in formato centesimale con 5 decimali
   - esempio: `2:30` diventa `2.50000`

2. `Riepilogo Viaggi`
   - raggruppa per dipendente e giorno
   - mostra solo le colonne utili alla lettura: `Reparto`, `Cod. Progetto`, `Progetto`, `Cod. Argomento`, `Argomento`, `Codice dipendente`, `Nominativo`, `Data`
   - calcola le ore lavorate totali
   - per le righe con `Descr.Reparto = MANUTENTORI` somma le ore di `COMMESSA` e `CHIUSURA`
   - per i `MANUTENTORI` calcola `Ore viaggio lorde` e `Ore viaggio nette`
   - per i `MANUTENTORI` calcola anche `% viaggio lorde` e `% viaggio nette`
   - per tutti gli altri reparti lascia queste colonne a `0`
   - puoi filtrare manualmente la colonna `% viaggio nette` o `% viaggio lorde` in Excel

## Struttura cartelle

- `input/` contiene il file sorgente del mese
- `output/` contiene il file elaborato
- `src/` contiene il codice Python

Quando usi l'eseguibile compilato, queste cartelle devono stare nella stessa cartella del `.exe`.

## Uso

1. Copia nella cartella `input` il file Excel del mese.
2. Avvia il programma.
3. Premi `Elabora file` oppure usa l'eseguibile.
4. Troverai il risultato nella cartella `output`.
5. Apri il foglio `Riepilogo Viaggi` e filtra la colonna `% viaggio nette` o `% viaggio lorde` in Excel come preferisci.

## Avvio

### Versione eseguibile

Se e presente il file `.exe`, basta aprirlo con doppio click. Il programma elabora il file piu recente in `input` e genera il riepilogo pronto per il filtro manuale in Excel.

### Versione Python

Serve Python 3.11 o superiore e il pacchetto `openpyxl`.

Esempio:

```powershell
python .\src\report_commesse.py
```

Oppure da riga di comando:

```powershell
python .\src\report_commesse.py --cli
```

## Build

Per generare l'eseguibile usa:

```powershell
.\build_exe.ps1
```

## Note

- Il foglio sorgente atteso si chiama `Stampa Commesse Dipendente`.
- Il programma prende il file `.xlsx` piu recente nella cartella `input`.
- I file temporanei di Excel con prefisso `~$` vengono ignorati.
- Se vuoi cambiare il criterio di calcolo delle ore viaggio, la logica e nel file `src/report_commesse.py`.
