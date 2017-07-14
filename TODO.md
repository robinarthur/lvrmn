TODO:

### Erstellen

- Liste mit Aktien
- Liste mit User Agents
- Funktion für Crawl URLs
- Funktion für das Speichern/ Öffnen der CSVs
- Funktion für das auswerten und abspeichern der Daten aus den CSVs

### nachschreiben

- Funktion AktienUebertragenDatenHolen Zeile 277 - 340

- Function AktienUebertragenDatenHolen(sh As Worksheet, abZeile As Long, nurFehler As Boolean, neuHolen As Boolean)
    Call DatenZurISINHolen(isin, zeile, shA, shQ, sh, neuHolen)
    Call LeereQuerySheet(shQ)
    Call LeereQuerySheet(shQ)

- Funktion DatenZurISINHolen
- Function DatenZurISINHolen(isin As String, zeile As Long, shA As Worksheet, shQ As Worksheet, sh As Worksheet, Optional neuHolen As Boolean)

    Call LeereStatusMeldung(zeile, sh)
    Call GrunddatenZurISIN(isin, zeile, quellzeile, shA, sh)
    Call AktuellerTerminUndKursZurISIN(isin, zeile, quellzeile, shA, shQ, sh)
    Call OnVistaDaten(isin, zeile, quellzeile, shA, shQ, sh, shVorher, zeileVorher)
    Call Marktkapitalisierung(isin, zeile, quellzeile, shA, shQ, sh)
    Call Quartalszahlen(isin, zeile, quellzeile, shA, shQ, sh, shVorher, parameterZeileVorher)
    Call AnalystenMeinungen(isin, zeile, quellzeile, shA, shQ, sh)
    Call GewinnRevisionen(isin, zeile, sh, shVorher, zeileVorher)
    Call HistorischeKurse(isin, zeile, quellzeile, shA, shQ, sh)
    Call DreiMonatsReversal(isin, zeile, quellzeile, shA, shQ, sh, shVorher, parameterZeileVorher)
    Call Bemerkungen(isin, zeile, sh, shVorher, zeileVorher)
