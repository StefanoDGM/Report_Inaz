# Report Commesse

Strumento Windows per elaborare la stampa commesse mensile e confrontarla con il riepilogo giornaliero INAZ.

## Cosa fa

Il programma legge dalla cartella `input`:

- il file `.xlsx` piu recente il cui nome non contiene `Riepilogo Giornaliero`, usato come stampa commesse;
- il file `.xlsx` piu recente il cui nome contiene `Riepilogo Giornaliero`, usato per i controlli sulle ore INAZ, se presente.

I file temporanei Excel con prefisso `~$` vengono ignorati. I risultati vengono creati nella cartella `output`.

I file generati sono 2:

1. `ore_analitica.xls`
   - copia il foglio sorgente
   - converte le ore da formato `hh:mm` in formato centesimale con 5 decimali
   - per i `MANUTENTORI` toglie 1 ora di pausa pranzo con priorita: `COMMESSA`, poi `CHIUSURA`, poi le altre commesse del giorno
   - esempio: `2:30` diventa `2.50000`
   - viene salvato nel vecchio formato Excel 5.0/95 `.xls`

2. `*_riepilogo.xlsx`
   - raggruppa per dipendente e giorno
   - mostra solo le colonne utili alla lettura: `Reparto`, `Codice dipendente`, `Nominativo`, `Data`
   - calcola le `Ore lorde lavorate`, che mantengono la pausa pranzo presente nell'input
   - calcola le `Ore nette lavorate` applicando ai `MANUTENTORI` la stessa sottrazione della pausa usata in `ore_analitica.xls`
   - calcola le `Ore sede ufficio` sommando i progetti che contengono `Sede Ufficio`, escludendo le righe con argomento `CHIUSURA`
   - calcola `% sede ufficio lorde` rispetto alle ore lavorate lorde
   - calcola `% sede ufficio nette` rispetto alle ore lavorate nette
   - per le righe con `Descr.Reparto = MANUTENTORI` somma le ore di `COMMESSA` e `CHIUSURA`
   - per i `MANUTENTORI` calcola `Ore viaggio lorde` e `Ore viaggio nette`
   - per i `MANUTENTORI` calcola `% viaggio lorde` sulle ore lavorate lorde e `% viaggio nette` sulle ore lavorate nette
   - per tutti gli altri reparti lascia queste colonne a `0`
   - include anche il foglio `Probabili errori`, con numeri riferiti direttamente alle righe di `ore_analitica.xls`
   - aggiunge le colonne `Controllo ore INAZ` e `Delta ore INAZ` al riepilogo
   - confronta sempre valori netti: `ORD + causali STR*` del `Riepilogo Giornaliero` contro le ore nette rendicontate
   - per i `MANUTENTORI` le ore nette rendicontate sono ottenute sottraendo un'ora di pausa pranzo nei feriali, prima dalle righe di viaggio (`COMMESSA`, poi `CHIUSURA`) e, se non bastano, dalle altre commesse della giornata
   - per i `MANUTENTORI`, se la prima timbratura e precedente alle 06:00, l'intero anticipo rispetto alle 06:00 viene confrontato con l'eccesso delle ore nette rendicontate; il caso e accettato solo se i due valori differiscono al massimo di 15 minuti
   - se il valore della colonna `Orario` termina con `N`, un'eccedenza delle ore nette rendicontate rispetto a `ORD + STR*` non viene segnalata: la `N` indica che gli straordinari sono esclusi dalle ore riconosciute da INAZ
   - segnala, per tutti i reparti, solo differenze superiori a 15 minuti
   - puoi filtrare manualmente la colonna `% viaggio nette` o `% viaggio lorde` in Excel

## Controlli nel foglio Probabili errori

Il foglio segnala i casi seguenti:

- ore di viaggio pari al 100% delle ore lavorate della giornata;
- differenza superiore a 15 minuti tra ore nette rendicontate e `ORD + STR*` di INAZ;
- progetto e argomento entrambi mancanti;
- argomento mancante per un progetto diverso da `COMMESSA`;
- progetto `Costi per formazione aziendale` con argomento diverso da `ZZCOSTO`, `COSTO` o `CHIUSURA`;
- progetto `Sede Ufficio` dei `MANUTENTORI` con argomento diverso da `ZZSEDE`, `ATTIVITA TECN IN SEDE` o `CHIUSURA`.

Il confronto con le ore INAZ non viene segnalato nelle giornate che riportano ferie, ROL, smart working, maternita o malattia. Se il `Riepilogo Giornaliero` non e presente, il programma genera comunque i due output, ma non esegue questo confronto.

## Struttura cartelle

- `input/` contiene la stampa commesse e, facoltativamente, il riepilogo giornaliero dello stesso periodo
- `output/` contiene il file elaborato
- `src/` contiene il codice Python

Quando usi l'eseguibile compilato, queste cartelle devono stare nella stessa cartella del `.exe`.

## Uso

1. Copia nella cartella `input` la stampa commesse del mese.
2. Copia nella stessa cartella anche il `Riepilogo Giornaliero` del periodo per abilitare il controllo delle ore INAZ.
3. Avvia il programma.
4. Premi `Elabora file` oppure usa l'eseguibile.
5. Troverai i due risultati nella cartella `output`.
6. Apri il file `*_riepilogo.xlsx` e controlla il foglio `Probabili errori`.

## Avvio

### Versione eseguibile

Se e presente il file `.exe`, basta aprirlo con doppio click. Il programma elabora il file piu recente in `input` e genera entrambi i file pronti in `output`.

### Versione Python

Serve Python 3.11 o superiore, il pacchetto `openpyxl` e Microsoft Excel installato. Excel viene usato in background per salvare `ore_analitica.xls` nel formato Excel 5.0/95.

Installa le dipendenze:

```powershell
python -m pip install -r .\requirements.txt
```

Esempio:

```powershell
python .\src\report_commesse.py
```

Oppure da riga di comando:

```powershell
python .\src\report_commesse.py --cli
```

Per indicare esplicitamente la stampa commesse:

```powershell
python .\src\report_commesse.py --cli --input ".\input\Stampa Commesse Dipendente.xlsx"
```

In modalita CLI, se l'elaborazione fallisce, i dettagli tecnici vengono salvati in `output/ReportCommesse_error.log`.

## Build

Per generare l'eseguibile usa:

```powershell
.\build_exe.ps1
```

## Note

- Il foglio sorgente atteso si chiama `Stampa Commesse Dipendente`.
- Il programma associa i dati dei due file tramite nominativo e data.
- Nei giorni feriali la pausa dei `MANUTENTORI` viene sottratta; nel fine settimana non viene applicata.
- I file temporanei di Excel con prefisso `~$` vengono ignorati.
- Se vuoi cambiare il criterio di calcolo delle ore viaggio, la logica e nel file `src/report_commesse.py`.
