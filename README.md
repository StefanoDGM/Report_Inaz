# Report Commesse

Strumento Windows per elaborare il file di stampa commesse mensile.

## Cosa fa

Il programma legge il file `.xlsx` piu recente presente nella cartella `input` e crea due file elaborati nella cartella `output`.

I file generati sono 2:

1. `ore_analitica.xls`
   - copia il foglio sorgente
   - converte le ore da formato `hh:mm` in formato centesimale con 5 decimali
   - per i `MANUTENTORI` toglie 1 ora di pausa pranzo con priorita: `COMMESSA`, poi `CHIUSURA`, poi le altre commesse del giorno
   - esempio: `2:30` diventa `2.50000`
   - viene salvato nel vecchio formato Excel 97-2003 `.xls`

2. `*_riepilogo.xlsx`
   - raggruppa per dipendente e giorno
   - mostra solo le colonne utili alla lettura: `Reparto`, `Codice dipendente`, `Nominativo`, `Data`
   - calcola le ore lavorate totali
   - calcola le `Ore sede ufficio` sommando i progetti che contengono `Sede Ufficio`
   - calcola anche la `% sede ufficio` rispetto al totale della giornata
   - per le righe con `Descr.Reparto = MANUTENTORI` somma le ore di `COMMESSA` e `CHIUSURA`
   - per i `MANUTENTORI` calcola `Ore viaggio lorde` e `Ore viaggio nette`
   - per i `MANUTENTORI` calcola anche `% viaggio lorde` e `% viaggio nette`
   - per tutti gli altri reparti lascia queste colonne a `0`
   - include anche il foglio `Probabili errori` con i casi da controllare
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
4. Troverai i due risultati nella cartella `output`.
5. Apri il file `*_riepilogo.xlsx` e filtra la colonna `% viaggio lorde` o `% viaggio nette` in Excel come preferisci.

## Avvio

### Versione eseguibile

Se e presente il file `.exe`, basta aprirlo con doppio click. Il programma elabora il file piu recente in `input` e genera entrambi i file pronti in `output`.

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
