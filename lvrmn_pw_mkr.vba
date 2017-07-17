Option Explicit

'--------------------------------------------------------
' Tool-Code
'--------------------------------------------------------

'Spaltennummern im Tabellenblatt "Aktien" und "Vorlage" (also auch in den Bewertungsblättern)
Public Const SPALTE_NAME = 1
Public Const SPALTE_ISIN = 2
Public Const SPALTE_GROESSE = 3
Public Const SPALTE_ART = 4

'Spaltennummern im Tabellenblatt "Aktien"
Public Const SPALTE_WAEHRUNG = 5                   'ie Originalwährung der Aktie, wird für QZ und 3MonRev verwendet
Public Const SPALTE_FINANZEN_NET = 6               'hier steht der URL-Teil für finanzen.net
Public Const SPALTE_ONVISTA = 7                    'hier steht der URL-Teil für onvista.de
Public Const SPALTE_HIST_ONVISTA_ORIG = 8          'URL-Teil zur Abfrage historischer Kurse bei OnVista in Originalwährung
Public Const SPALTE_HIST_ONVISTA = 9               'URL-Teil zur Abfrage historischer Kurse bei OnVista in Euro
Public Const SPALTE_BENCHMARK_NAME = 10            'Bezeichnung der Benchmark (des Vergleichsindex)
Public Const SPALTE_BENCHMARK_HIST_ONVISTA = 11    'Teil der Onvista-URL zur Abfrage historischer Kurse für die Benchmark
Public Const SPALTE_4TRADERS = 12                  'URL-Teil für de.4-traders.com
Public Const SPALTE_NUR_MANUELLE_TERMINE = 13      'hier wird angegeben, ob nur die in den folgenden Spalten stehenden Termine verwendet werden sollen (ohne Web)
Public Const SPALTE_TERMINE = 14                   'Ab dieser Spalte können Quartalszahlentermine eingetragen sein.

'Spalten im Tabellenblatt "Vorlage" bzw. in den daraus erzeugten Bewertungsblättern
'In diese Spalten werden jaweils die aus dem Web ausgelesenen Daten geschrieben.
'--------------------------------------------------------------------------------------------------
Public Const SPALTE_DATUM = 5
Public Const SPALTE_KURS = 6
'--------------------------------------------------------------------------------------------------
Public Const SPALTE_LJ = 7             'letztes Geschäftsjahr
Public Const SPALTE_ROE = 8
Public Const SPALTE_EBITMARGE = 9
Public Const SPALTE_EKQUOTE = 10
'--------------------------------------------------------------------------------------------------
'Spalten für EPS (earnings per share = Gewinn pro Aktie) über 5 Geschäftsjahre
'vom vorvorletzten bis zum nächsten Geschäftsjahr
Public Const SPALTE_EPSLJ2 = 11
Public Const SPALTE_EPSLJ1 = 12
Public Const SPALTE_EPSLJ = 13
Public Const SPALTE_EPSAJ = 14
Public Const SPALTE_EPSNJ = 15
'--------------------------------------------------------------------------------------------------
Public Const SPALTE_ANALYSTENANZAHL = 20
Public Const SPALTE_ANALYSTENMEINUNG = 21
'--------------------------------------------------------------------------------------------------
'Spalten für Daten zur Berechnung der Reaktion auf Quartalszahlen
Public Const SPALTE_MARKTKAP = 22          'nur informativ - nicht zur Berechnung
Public Const SPALTE_BENCHMARK = 23
Public Const SPALTE_DATUMZAHLEN = 24
Public Const SPALTE_DATUMVORTAG = 25
Public Const SPALTE_KURSZAHLEN = 26
Public Const SPALTE_KURSVORTAG = 27
Public Const SPALTE_BENCHMARKKURS = 28
Public Const SPALTE_BENCHMARKVORTAG = 29
'--------------------------------------------------------------------------------------------------
'Spalten für die Berechnung der Gewinnrevisionen sind jeweils die 4 direkt links von diesen
Public Const SPALTE_EPSAJ_WDH = 37
Public Const SPALTE_EPSNJ_WDH = 42
'--------------------------------------------------------------------------------------------------
'Spalten zur Berechnung der Kursentwicklung über 6 Monate bzw. 1 Jahr
Public Const SPALTE_DATUM6MON = 45
Public Const SPALTE_KURS6MON = 46
Public Const SPALTE_DATUM1JAHR = 49
Public Const SPALTE_KURS1JAHR = 50
'--------------------------------------------------------------------------------------------------
'Spalten zur Berechnung des Dreimonatsreversals (nur für Large Caps)
Public Const SPALTE_DATUM3MON = 54
Public Const SPALTE_KURS3MON = 55
Public Const SPALTE_BENCHMARK3MON = 56
Public Const SPALTE_DATUM2MON = 57
Public Const SPALTE_KURS2MON = 58
Public Const SPALTE_BENCHMARK2MON = 59
Public Const SPALTE_DATUM1MON = 62
Public Const SPALTE_KURS1MON = 63
Public Const SPALTE_BENCHMARK1MON = 64
Public Const SPALTE_DATUM0MON = 67
Public Const SPALTE_KURS0MON = 68
Public Const SPALTE_BENCHMARK0MON = 69
Public Const SPALTE_PKT_3MONREV = 86
'--------------------------------------------------------------------------------------------------
'Bemerkungen
Public Const SPALTE_BEMERKUNGEN = 90
'--------------------------------------------------------------------------------------------------
'KGV-Spalten für die zweite Berechnungsvariante des KGV-Kriteriums
'vom vorvorletzten bis zum aktuellen Geschäftsjahr
Public Const SPALTE_KGVLJ2 = 91
Public Const SPALTE_KGVLJ1 = 92
Public Const SPALTE_KGVLJ = 93
Public Const SPALTE_KGVAJ = 94
'--------------------------------------------------------------------------------------------------

Public Const ANZAHL_VERSUCHE_WEBZUGRIFF = 2    'sooft wird das Laden einer URL maximal versucht

'--------------------------------------------------------------------------------------------------

'für Status-Notizen zum Programmablauf (Fehlermeldungen, Warnungen)
Public Const SPALTE_STATUS = 102

'Status-Farben:
Dim STATUS_FEHLER As Long, STATUS_WARNUNG As Long

'Zum Merken, welches der letzte Börsentag ist zu Onvista-Id eines Index plus Datum, für historische Kursabfragen
Dim BTIdDatum() As String         'Werte haben Form <OnvistaHistID>#JJJJ-MM-TT    zu diesem Index und diesem Datum wird der letzte Börsentag bis zu diesem Datum gesucht
Dim BT() As String                'Werte haben die Form JJJJ-MM-TT                das ist der gesuchte Börsentag zum Wert es vorigen Arrays an gleicher Position
Dim BTU As Long                   'Ubound dieser beiden Arrays

Function SetzeEinstellungen()
    'Farben für die Fehlermeldungen und Warnungen
    STATUS_FEHLER = RGB(255, 200, 170)  'rosa
    STATUS_WARNUNG = RGB(255, 255, 0)   'gelb
    'Arrays zum Merken der Börsentage initialisieren
    ReDim BTIdDatum(0)
    ReDim BT(0)
    BTU = -1
End Function


Public Function PruefeAktienliste() As Boolean
    'Prüft, ob das Tabellenblatt "Aktien" den richtigen Aufbau einer Aktienliste hat.
    'Wird vor jeder Berechnung und Aktualisierung benutzt.

    'Rückgabewert: True, wenn Blatt "Aktien" die richtigen Spalten hat, sonst False

    Dim shA As Worksheet
    Dim meldung As String
    Dim spalten As String
    Dim V As Variant
    Dim K As Integer
    Dim kopf As String

    Set shA = Sheets("Aktien")

    spalten = "Name#ISIN#Größe#Art#Währung#URLTeil finanzen.net#URL-Teil OnVista" + _
    "#URL-Teil hist. OnVista Landeswährung#URL-Teil hist. OnVista Euro#Benchmark" + _
    "#Benchm. URL-Teil hist. OnVista#URL-Teil 4-traders#nur manuelle Termine#Termine"

    meldung = ""

    V = Split(spalten, "#")
    For K = 0 To UBound(V)
        kopf = CStr(shA.Cells(1, K + 1).Value)
        If InStr(Trim(LCase(kopf)), Trim(LCase(V(K)))) <> 1 Then
            meldung = meldung + "Spalte " + Chr(65 + K) + ": " + V(K) + Chr(13)
        End If
    Next

    If meldung = "" Then
        PruefeAktienliste = True
    Else
        MsgBox "Das Tabellenblatt ""Aktien"" hat nicht den richtigen Aufbau. Erwartet wird:" + Chr(13) + Chr(13) + meldung, vbExclamation + vbOKOnly, "Aktienliste"
        PruefeAktienliste = False
    End If

End Function


'================================================================
'Die Makros
'================================================================
Sub A_alle_Bewertungen_erzeugen()
    'Legt ein neues Tabellenblatt mit aktuellem Datum der Form JJJJ-MM-TT an.
    'Erzeugt zu jeder Zeile das Blattes "Aktien" eine Zeile in diesem neuen Blatt und holt jeweils die Daten für die Bewertung aus dem Internet.

    Dim sh As Worksheet
    Dim neuHolen As Boolean

    If Not PruefeAktienliste() Then Exit Sub
    neuHolen = DialogDatenNeuHolen()

    Set sh = NeuesBlattAnlegen()
    If Not (sh Is Nothing) Then
        Call AktienUebertragenDatenHolen(sh, 0, False, neuHolen)
        Sheets(Sheets.Count).Activate
        Application.StatusBar = "OK"
    End If

End Sub


Sub B_Bewertungen_ab_dieser_Zeile_erzeugen()
    'Wird ausgehend von einer Zeile in einem Bewertungsblatt gestartet.
    'Überschreibt im aktuellen Bewertungsblatt ab dieser Zeile mit neu geladenen Daten.

    Dim sh As Worksheet
    Dim zeile As Long
    Dim isin As String
    Dim neuHolen As Boolean

    Set sh = ActiveSheet
    If Not (sh.Name Like "####-##-##") Then
        'kein Bewertungsblatt
        Exit Sub
    End If
    zeile = ActiveCell.Row
    If zeile = 1 Then
        'Überschriftenzeile
        Exit Sub
    End If
    isin = CStr(sh.Cells(zeile, SPALTE_ISIN).Value)
    If isin = "" Then
        'keine Aktien-Zeile
        Exit Sub
    End If

    If Not PruefeAktienliste() Then Exit Sub
    neuHolen = DialogDatenNeuHolen()

    Call AktienUebertragenDatenHolen(sh, zeile, False, neuHolen)
    sh.Activate
    Application.StatusBar = "OK"

End Sub


Sub C_fehlerhafte_Bewertungen_wiederholen()
    'Wird ausgehend von einem Bewertungsblatt gestartet.
    'Berechnet alle Zeilen, die ganz hinten eine Fehlermarkierung haben, noch einmal

    Dim sh As Worksheet
    Dim zeile As Long
    Dim neuHolen As Boolean

    Set sh = ActiveSheet
    If Not (sh.Name Like "####-##-##") Then
        'kein Bewertungsblatt
        Exit Sub
    End If

    If Not PruefeAktienliste() Then Exit Sub
    neuHolen = DialogDatenNeuHolen()

    Call AktienUebertragenDatenHolen(sh, 2, True, neuHolen)
    sh.Activate
    Application.StatusBar = "OK"

End Sub


Sub D_diese_Bewertungszeile_aktualisieren()
    'Wird ausgehend von einer Zeile im Bewertungsblatt gestartet.
    'Überschreibt diese Zeile mit neu geladenen Daten.

    Dim sh As Worksheet
    Dim zeile As Long
    Dim isin As String
    Dim shA As Worksheet
    Dim shQ As Worksheet
    Dim neuHolen As Boolean

    Set sh = ActiveSheet
    If Not (sh.Name Like "####-##-##") Then
        'kein Bewertungsblatt
        Exit Sub
    End If
    zeile = ActiveCell.Row
    If zeile = 1 Then
        'Überschriftenzeile
        Exit Sub
    End If
    isin = CStr(sh.Cells(zeile, SPALTE_ISIN).Value)
    If isin = "" Then
        'keine Aktien-Zeile
        Exit Sub
    End If

    If Not PruefeAktienliste() Then Exit Sub
    neuHolen = DialogDatenNeuHolen()

    Call SetzeEinstellungen
    Set shA = Sheets("Aktien")
    Set shQ = Sheets("Query")

    Application.StatusBar = "aktuelle Zeile"
    Call DatenZurISINHolen(isin, zeile, shA, shQ, sh, neuHolen)
    Call LeereQuerySheet(shQ)

    Call sh.Activate
    Application.StatusBar = "OK"

End Sub
'================================================================


Function DialogDatenNeuHolen() As Boolean
    'Abfrage, welcher Daten-Modus verwendet werden soll.
    'Das bezieht sich auf die Daten zur Raktion auf Quartalszahlen und für Large-Caps zum Dreimonatsreversal

    'Wenn False geantwortet wird, werden Daten zu QZ und 3MonRev aus dem vorigen Bewertungsblatt übernommen,
    'sofern Datum bzw. Benchmark gleich geblieben sind. (spart Laufzeit)
    'Wenn True geantwortet wird, werden die Daten alle aktuell aus dem Web gezogen.
    'Das ist bei Änderung von OnVista Hist. Ids sinnvoll. (wird genauer)

    Dim dialog As New frmModus
    dialog.Show
    DialogDatenNeuHolen = dialog.Tag
End Function


Function NeuesBlattAnlegen() As Worksheet
    'Legt ein neues Tabellenblatt als Kopie von "Vorlage" an und benennt es mit aktuellem Datum der Form JJJJ-MM-TT.
    'Überträgt die Berechnungsformeln aus der Vorlage auf den späteren gesamten Datenbereich des neuen Bewertungsblattes.

    'wird benutzt in Makro A_alle_Bewertungen_erzeugen

    Dim blattName As String
    Dim sh As Worksheet
    Dim shV As Worksheet
    Dim shA As Worksheet
    Dim maxSpalte As Long
    Dim maxZeile As Long
    Dim rng As Range
    Dim rngVoll As Range

    blattName = Format(Now, "yyyy-mm-dd")

    'Es darf nur ein Tabellenblatt dieses Namens geben:
    On Error Resume Next
    Set sh = Sheets(blattName)
    On Error GoTo 0
    If Not (sh Is Nothing) Then
        Call sh.Select
        Call MsgBox("Tabellenblatt """ + blattName + """ gibt es schon.", vbOKOnly + vbExclamation, "Blatt schon vorhanden")
        Call sh.Delete
        'Prüfen, ob wirklich gelöscht wurde:
        Set sh = Nothing
        On Error Resume Next
        Set sh = Sheets(blattName)
        On Error GoTo 0
        If Not (sh Is Nothing) Then
            'Im Löschen-Dialog wurde "Abbrechen" geklickt.
            Exit Function
        End If
    End If

    'Ermitteln, bis zu welcher Zeile bzw. Spalte das neue Bewertungsblatt am Ende ausgefüllt sein wird. ( maxZeile bzw. maxSpalte )
    Set shV = Sheets("Vorlage")
    maxSpalte = 3
    Do Until CStr(shV.Cells(1, maxSpalte).Value) = ""
        maxSpalte = maxSpalte + 1
    Loop
    maxSpalte = maxSpalte - 1
    Set shA = Sheets("Aktien")
    maxZeile = 2
    Do Until CStr(shA.Cells(maxZeile, 1).Value) = ""
        maxZeile = maxZeile + 1
    Loop
    maxZeile = maxZeile - 1

    'Anlegen des neuen Bewertungsblattes als letztes Blatt der Arbeitsmappe:
    Set sh = Sheets(Sheets.Count)
    Call shV.Copy(, sh)
    Set sh = Sheets(Sheets.Count)
    sh.Name = blattName

    'Übertragen der Berechnungsformeln aus der Vorlage auf den gesamten Datenbereich des neuen Bewertungsblattes:
    Set rng = sh.Range(sh.Cells(2, SPALTE_ISIN + 1), sh.Cells(2, maxSpalte))
    Set rngVoll = sh.Range(sh.Cells(2, SPALTE_ISIN + 1), sh.Cells(maxZeile, maxSpalte))
    Call rng.AutoFill(rngVoll)

    Set NeuesBlattAnlegen = sh      'das vorbereitete neue Bewertungsblatt

End Function


Function AktienUebertragenDatenHolen(sh As Worksheet, abZeile As Long, nurFehler As Boolean, neuHolen As Boolean)
    'Überträgt die Aktien (jeweils Name und ISIN) aus dem Tabellenblatt "Aktien" in das Bewertungsblatt (sh), sofern nötig.
    'Ruft die benötigten Daten aus dem Internet ab und trägt sie ein.

    'Parameter:     sh = Ziel-Bewertungsblatt
    '               abZeile = ab welcher Zeile des Bewertungsblattes mit dem Holen der Daten negonnen wird
    '               nurFehler = Wenn das True ist, werden nur die Daten zu den mit Fehler markierten Zeilen geholt
    '               neuHolen = Wenn das True ist, werden QZ und 3MonRev immer aus dem Web abgerufen, ansonsten vom vorigen Bewertungsblatt wenn möglich

    'wird benutzt in
    '           Makro A_alle_Bewertungen_erzeugen             --> abZeile = 0
    '           Makro B_Bewertungen_ab_dieser_Zeile_erzeugen  --> abZeile = Nummer der Zeile, ab der das passieren soll
    '           Makro C_fehlerhafte_Bewertungen_wiederholen   --> abZeile = 0

    Dim shA As Worksheet
    Dim shQ As Worksheet

    Dim zeile As Long
    Dim aktie As String
    Dim isin As String
    Dim K As Long

    Call SetzeEinstellungen
    Set shA = Sheets("Aktien")
    Set shQ = Sheets("Query")

    If abZeile = 0 Then
        'Name, ISIN und Größe zur Aktie übertragen, denn das Bewertungsblatt wurde gerade neu angelegt
        zeile = 2
        Do Until CStr(shA.Cells(zeile, SPALTE_NAME).Value) = ""
            sh.Cells(zeile, SPALTE_NAME).Value = shA.Cells(zeile, SPALTE_NAME).Value
            sh.Cells(zeile, SPALTE_ISIN).Value = shA.Cells(zeile, SPALTE_ISIN).Value
            sh.Cells(zeile, SPALTE_GROESSE).Value = shA.Cells(zeile, SPALTE_GROESSE).Value
            sh.Cells(zeile, SPALTE_ART).Value = shA.Cells(zeile, SPALTE_ART).Value
            zeile = zeile + 1
        Loop
        zeile = 2
    Else
        zeile = abZeile
    End If

    'Daten zu den Aktien holen und eintragen
    Do Until CStr(sh.Cells(zeile, SPALTE_NAME).Value) = ""
        aktie = sh.Cells(zeile, SPALTE_NAME).Value
        Application.StatusBar = "Zeile " + CStr(zeile) + " - " + aktie
        DoEvents

        If (nurFehler = False) Or HatFehler(zeile, sh) Then

            isin = sh.Cells(zeile, SPALTE_ISIN).Value
            Call DatenZurISINHolen(isin, zeile, shA, shQ, sh, neuHolen)

            'Damit nicht alles verloren ist, falls sich Excel mal aufhängt:
            If zeile Mod 50 = 0 Then
                Call LeereQuerySheet(shQ)
                sh.Parent.Save
            End If

        End If

        zeile = zeile + 1
    Loop

    'Aufräumen:
    Call LeereQuerySheet(shQ)

End Function


Function DatenZurISINHolen(isin As String, zeile As Long, shA As Worksheet, shQ As Worksheet, sh As Worksheet, Optional neuHolen As Boolean)
    'Fragt zu gegebener ISIN einer Aktie alle benötigten Daten aus dem Web ab und trägt diese in die zugehörige Zeile des aktuellen Bewertungsblattes ein.

    'Parameter:     isin    zu dieser ISIN werden die Daten geholt
    '               zeile   in diese Zeile des Bewertungsblattes werden die Daten geschrieben
    '               shA     Tabellenblatt "Aktien"
    '               shQ     Tabellenblatt "Query" - Hilfsblatt für die Web-Abfragen
    '               sh      Ziel-Bewertungsblatt
    '               neuHolen = True bewirkt, dass Kurse für Reaktion auf QZ und für Dreimonatsreversal aktuell aus dem Web gezogen werden,
    '                                        auch wenn sie vom vorigen Blatt genommen werden könnten

    'wird verwendet in
    '               AktienUebertragenDatenHolen
    '               D_diese_Bewertungszeile_aktualisieren

    Dim quellzeile As Long
    Dim K As Long
    Dim shVorher As Worksheet
    Dim zeileVorher As Long
    Dim parameterZeileVorher As Long

    Call LeereStatusMeldung(zeile, sh)

    If CStr(shA.Cells(zeile, SPALTE_ISIN).Value) = isin Then
        'Die Aktie steht im Bewertungsblatt in der gleichen Zeile wie im Blatt "Aktien".
        quellzeile = zeile
    Else
        'Suchen die Zeile mit der ISIN im Tabellenblatt "Aktien":
        quellzeile = 2
        Do Until CStr(shA.Cells(quellzeile, SPALTE_ISIN).Value) = ""
            If CStr(shA.Cells(quellzeile, SPALTE_ISIN).Value) = isin Then
                GoTo gefunden
            End If
            quellzeile = quellzeile + 1
        Loop
        Exit Function       'keine passende Zeile gefunden
gefunden:
    End If

    'Tabellenblatt mit der vorigen Bewertung holen, sofern vorhanden:
    For K = 1 To Sheets.Count
        If Sheets(K).Name = sh.Name Then
            If K > 1 Then
                If Sheets(K - 1).Name Like "####-##-##" Then
                    Set shVorher = Sheets(K - 1)
                    Exit For
                End If
            End If
        End If
    Next

    'passende Zeile im vorigen Bewertungsblatt suchen:
    If Not (shVorher Is Nothing) Then
        If CStr(shVorher.Cells(zeile, SPALTE_ISIN).Value) = isin Then
            zeileVorher = zeile
            GoTo gefundenVorher
        End If
        K = 2
        Do Until CStr(shVorher.Cells(K, SPALTE_ISIN).Value) = ""
            If CStr(shVorher.Cells(K, SPALTE_ISIN).Value) = isin Then
                zeileVorher = K
                GoTo gefundenVorher
            End If
            K = K + 1
        Loop
gefundenVorher:
    End If

    If neuHolen Then
        parameterZeileVorher = 0
    Else
        parameterZeileVorher = zeileVorher
    End If

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

End Function


Function GrunddatenZurISIN(isin As String, zeile As Long, quellzeile As Long, shA As Worksheet, sh As Worksheet)
    'Überträgt Name, Größe und Art vom Blatt Aktien ins Bewertungsblatt

    'Parameter:     isin    zu dieser ISIN werden die Daten geholt
    '               zeile   in diese Zeile des Bewertungsblattes werden die Daten geschrieben
    '               quellzeile  passende Zeile im Tabellenblatt "Aktien"
    '               shA     Tabellenblatt "Aktien"
    '               sh      Ziel-Bewertungsblatt

    'wird verwendet in DatenZurISINHolen

    sh.Cells(zeile, SPALTE_NAME).Value = shA.Cells(quellzeile, SPALTE_NAME).Value
    sh.Cells(zeile, SPALTE_GROESSE).Value = shA.Cells(quellzeile, SPALTE_GROESSE).Value
    sh.Cells(zeile, SPALTE_ART).Value = shA.Cells(quellzeile, SPALTE_ART).Value

End Function


Function AktuellerTerminUndKursZurISIN(isin As String, zeile As Long, quellzeile As Long, shA As Worksheet, shQ As Worksheet, sh As Worksheet)
    'Holt den letzten Kurs und das Datum, von wenn dieser ist, trägt das ins Ziel-Bewertungsblatt in die Zeile zur ISIN ein.
    'benutzt dazu finanzen.net

    'Parameter:     isin    zu dieser ISIN werden die Daten geholt
    '               zeile   in diese Zeile des Bewertungsblattes werden die Daten geschrieben
    '               quellZeile  in dieser Zeile des Tabellenblattes "Aktien" stehen benötigte Hilfsdaten zur Aktie (URL-Teile für die Abfragen usw.)
    '               shA     Tabellenblatt "Aktien"
    '               shQ     Tabellenblatt "Query" - Hilfsblatt für die Web-Abfragen
    '               sh      Ziel-Bewertungsblatt

    'wird verwendet in DatenZurISINHolen

    Const URL_ANFANG = "URL;http://www.finanzen.net/boersenplaetze/"

    Dim urlTeil As String
    Dim url As String
    Dim rng As Range
    Dim Z As Long

    Dim termin As Variant
    Dim Kurs As Double

    Dim inhalt As String

    sh.Cells(zeile, SPALTE_DATUM).Value = ""
    sh.Cells(zeile, SPALTE_KURS).Value = ""

    termin = Empty

    'Seite aus dem Web abrufen:
    urlTeil = shA.Cells(quellzeile, SPALTE_FINANZEN_NET).Value
    If urlTeil = "" Then
        Exit Function
    End If
    url = URL_ANFANG + urlTeil

    If LadeURL(ANZAHL_VERSUCHE_WEBZUGRIFF, 1, shQ, url) = 0 Then

        Call StatusMeldung(zeile, sh, "Lade-F.", Replace(url, "URL;", ""), STATUS_FEHLER)

    Else

        'benötigte Daten extrahieren:
        Set rng = shQ.Cells.Find("Deutsche Börsenübersicht", shQ.Cells(1, 1))
        If rng Is Nothing Then

            Call StatusMeldung(zeile, sh, "Daten-F.", Replace(url, "URL;", ""), STATUS_FEHLER)

        Else

            Z = rng.Row + 1
            If CStr(shQ.Cells(Z, 1).Value) = "Börse" Then
                If CStr(shQ.Cells(Z + 1, 1).Value) <> "" Then
                    inhalt = CStr(shQ.Cells(Z + 1, 2).Value)
                    If inhalt Like "* EUR" Then
                        inhalt = MeinSystemFormat(Replace(inhalt, " EUR", ""))
                        If IsNumeric(inhalt) Then
                            Kurs = CDbl(inhalt)
                        End If
                        inhalt = CStr(shQ.Cells(Z + 1, 9).Value)
                        termin = DatumsWert(inhalt)
                    End If
                End If
            End If

        End If

        'Werte in Bewertungstabelle eintragen:
        If Not (IsEmpty(termin)) And (Kurs > 0) Then
            sh.Cells(zeile, SPALTE_DATUM).Value = termin
            sh.Cells(zeile, SPALTE_KURS).Value = Kurs
        End If

    End If

End Function


Function OnVistaDaten(isin As String, zeile As Long, quellzeile As Long, shA As Worksheet, shQ As Worksheet, sh As Worksheet, shVorher As Worksheet, zeileVorher As Long)
    'Holt RoE, EBIT-Marge, Eigenkapitalquote des letzten Geschäftsjahres,
    'sowie EPS über fünf Geschäftsjahre (vorvorletztes, vorletztes, letztes, aktuelles geschätzt, nächstes geschätzt).
    'weiterhin KGV des vorvorletzten, des vorletzten und des letzten und (geschätzt) des aktuellen Geschäftsjahres
    'Trägt die Daten ins Ziel-Bewertungsblatt in die Zeile zur ISIN ein.
    'benutzt dazu onvista.de
    'wenn etwas bei onvista nicht vorhanden ist, wird es aus der passenden Zeile aus dem vorigen Bewertungsblatt übernommen,
    'um manuelle Einträge bestehen zu lassen

    'Parameter:     isin    zu dieser ISIN werden die Daten geholt
    '               zeile   in diese Zeile des Bewertungsblattes werden die Daten geschrieben
    '               quellZeile  in dieser Zeile des Tabellenblattes "Aktien" stehen benötigte Hilfsdaten zur Aktie (URL-Teile für die Abfragen usw.)
    '               shA     Tabellenblatt "Aktien"
    '               shQ     Tabellenblatt "Query" - Hilfsblatt für die Web-Abfragen
    '               sh      Ziel-Bewertungsblatt
    '               shVorher    voriges Bewertungsblatt
    '               zeileVorher Zeile dieser Aktie im vorigen Bewertungsblatt

    'wird verwendet in DatenZurISINHolen

    Const URL_ANFANG = "URL;http://www.onvista.de/aktien/fundamental/"

    Dim urlTeil As String
    Dim url As String
    Dim rng As Range
    Dim zeileJahr As Long
    Dim spalteLJ As Long
    Dim K As Long
    Dim wert As Variant
    Dim inhalt As String
    Dim LJ As String
    Dim gleichesLJ As Boolean
    Dim zeileRent As Long
    Dim spalteRent As Long
    Dim zeileEK As Long
    Dim spalteEK As Long
    Dim zeileEPS As Long
    Dim zeileKGV As Long

    sh.Cells(zeile, SPALTE_LJ).Value = ""
    sh.Cells(zeile, SPALTE_ROE).Value = ""
    sh.Cells(zeile, SPALTE_EBITMARGE).Value = ""
    sh.Cells(zeile, SPALTE_EKQUOTE).Value = ""
    sh.Cells(zeile, SPALTE_EPSLJ2).Value = ""
    sh.Cells(zeile, SPALTE_EPSLJ1).Value = ""
    sh.Cells(zeile, SPALTE_EPSLJ).Value = ""
    sh.Cells(zeile, SPALTE_EPSAJ).Value = ""
    sh.Cells(zeile, SPALTE_EPSNJ).Value = ""

    sh.Cells(zeile, SPALTE_KGVLJ2).Value = ""
    sh.Cells(zeile, SPALTE_KGVLJ1).Value = ""
    sh.Cells(zeile, SPALTE_KGVLJ).Value = ""
    sh.Cells(zeile, SPALTE_KGVAJ).Value = ""

    'Seite aus dem Internet laden
    urlTeil = shA.Cells(quellzeile, SPALTE_ONVISTA).Value
    If urlTeil = "" Then
        Exit Function
    End If
    url = URL_ANFANG + urlTeil

    If LadeURL(ANZAHL_VERSUCHE_WEBZUGRIFF, 1, shQ, url) = 0 Then

        Call StatusMeldung(zeile, sh, "Lade-F.", Replace(url, "URL;", ""), STATUS_FEHLER)

    Else

        'Letztes Geschäftsjahr heraussuchen:
        Set rng = shQ.Cells.Find("Gewinn", shQ.Cells(1, 1))
        If rng Is Nothing Then

            Call StatusMeldung(zeile, sh, "Daten-F.", Replace(url, "URL;", ""), STATUS_FEHLER)

        Else

            zeileJahr = rng.Row
            K = 2
            Do Until CStr(shQ.Cells(zeileJahr, K).Value) = ""
                inhalt = CStr(shQ.Cells(zeileJahr, K).Value)
                If (inhalt Like "####") Or (inhalt Like "##/##") Then
                    spalteLJ = K
                    LJ = inhalt
                    Exit Do
                End If
                K = K + 1
            Loop
            sh.Cells(zeile, SPALTE_LJ).Value = LJ
        End If

        If LJ <> "" Then
            'schauen, ob es eine vorige Bewertung mit gleichem Geschäftsjahr gibt
            gleichesLJ = False
            If zeileVorher > 1 Then
                If CStr(shVorher.Cells(zeileVorher, SPALTE_LJ).Value) = LJ Then
                    gleichesLJ = True
                End If
            End If

            'RoE und EBIT-Marge:
            Set rng = shQ.Cells.Find("Rentabilität Mehr", shQ.Cells(1, 1))
            If Not rng Is Nothing Then
                zeileRent = rng.Row
                For K = 2 To 10
                    If CStr(shQ.Cells(zeileRent, K).Value) = LJ Then
                        spalteRent = K
                        Exit For
                    End If
                Next
                If spalteRent > 0 Then
                    For K = 20 To 25
                        If shQ.Cells(zeileRent + K, 1).Value = "Eigenkapitalrendite" Then
                            wert = MeinProzentWert(shQ.Cells(zeileRent + K, spalteRent).Value)
                            If IsNumeric(wert) And (CStr(wert) <> "") Then
                                sh.Cells(zeile, SPALTE_ROE).Value = wert
                            Else
                                If gleichesLJ Then
                                    'übernehmen evt. manuell eingepflegten Wert aus der vorigen Bewertung
                                    sh.Cells(zeile, SPALTE_ROE).Value = shVorher.Cells(zeileVorher, SPALTE_ROE).Value
                                End If
                            End If
                        Else
                            If shQ.Cells(zeileRent + K, 1).Value = "EBIT-Marge" Then
                                wert = MeinProzentWert(shQ.Cells(zeileRent + K, spalteRent).Value)
                                If IsNumeric(wert) And (CStr(wert) <> "") Then
                                    sh.Cells(zeile, SPALTE_EBITMARGE).Value = wert
                                Else
                                    If gleichesLJ Then
                                        'übernehmen evt. manuell eingepflegten Wert aus der vorigen Bewertung
                                        sh.Cells(zeile, SPALTE_EBITMARGE).Value = shVorher.Cells(zeileVorher, SPALTE_EBITMARGE).Value
                                    End If
                                End If
                            End If
                        End If
                    Next
                Else
                    'keine Angaben zum aktuellen Geschäftsjahr - vielleicht manuell im vorigen Blatt
                    If gleichesLJ Then
                        sh.Cells(zeile, SPALTE_ROE).Value = shVorher.Cells(zeileVorher, SPALTE_ROE).Value
                        sh.Cells(zeile, SPALTE_EBITMARGE).Value = shVorher.Cells(zeileVorher, SPALTE_EBITMARGE).Value
                    End If
                End If
            End If

            'EK-Quote
            Set rng = shQ.Cells.Find("Bilanz Mehr", shQ.Cells(1, 1))
            If Not rng Is Nothing Then
                zeileEK = rng.Row
                For K = 2 To 10
                    If shQ.Cells(zeileEK, K).Value = LJ Then
                        spalteEK = K
                        Exit For
                    End If
                Next
                If spalteEK > 0 Then
                    For K = 13 To 20
                        If shQ.Cells(zeileEK + K, 1).Value = "Eigenkapitalquote" Then
                            wert = MeinProzentWert(shQ.Cells(zeileEK + K, spalteEK).Value)
                            If IsNumeric(wert) And (CStr(wert) <> "") Then
                                sh.Cells(zeile, SPALTE_EKQUOTE).Value = wert
                            Else
                                If gleichesLJ Then
                                    'übernehmen evt. manuell eingepflegten Wert aus der vorigen Bewertung
                                    sh.Cells(zeile, SPALTE_EKQUOTE).Value = shVorher.Cells(zeileVorher, SPALTE_EKQUOTE).Value
                                End If
                            End If
                        End If
                    Next
                Else
                    'keine Angaben zum aktuellen Geschäftsjahr - vielleicht manuell im vorigen Blatt
                    If gleichesLJ Then
                        sh.Cells(zeile, SPALTE_EKQUOTE).Value = shVorher.Cells(zeileVorher, SPALTE_EKQUOTE).Value
                    End If
                End If
            End If

            'EPS-Werte und (historische) KGV:
            For K = 7 To 15
                If shQ.Cells(zeileJahr + K, 1).Value = "Gewinn pro Aktie in EUR" Then
                    zeileEPS = zeileJahr + K
                    zeileKGV = zeileEPS + 1
                    Exit For
                End If
            Next
            If zeileEPS > 0 Then
                'EPS-Werte:
                inhalt = MeinSystemFormat(shQ.Cells(zeileEPS, spalteLJ + 2).Value)
                If IsNumeric(inhalt) And (inhalt <> "") Then
                    sh.Cells(zeile, SPALTE_EPSLJ2).Value = CDbl(inhalt)
                Else
                    If gleichesLJ Then
                        'übernehmen evt. manuell eingepflegten Wert aus der vorigen Bewertung
                        sh.Cells(zeile, SPALTE_EPSLJ2).Value = shVorher.Cells(zeileVorher, SPALTE_EPSLJ2).Value
                    End If
                End If
                inhalt = MeinSystemFormat(shQ.Cells(zeileEPS, spalteLJ + 1).Value)
                If IsNumeric(inhalt) And (inhalt <> "") Then
                    sh.Cells(zeile, SPALTE_EPSLJ1).Value = CDbl(inhalt)
                Else
                    If gleichesLJ Then
                        'übernehmen evt. manuell eingepflegten Wert aus der vorigen Bewertung
                        sh.Cells(zeile, SPALTE_EPSLJ1).Value = shVorher.Cells(zeileVorher, SPALTE_EPSLJ1).Value
                    End If
                End If
                inhalt = MeinSystemFormat(shQ.Cells(zeileEPS, spalteLJ).Value)
                If IsNumeric(inhalt) And (inhalt <> "") Then
                    sh.Cells(zeile, SPALTE_EPSLJ).Value = CDbl(inhalt)
                Else
                    If gleichesLJ Then
                        'übernehmen evt. manuell eingepflegten Wert aus der vorigen Bewertung
                        sh.Cells(zeile, SPALTE_EPSLJ).Value = shVorher.Cells(zeileVorher, SPALTE_EPSLJ).Value
                    End If
                End If
                If spalteLJ > 2 Then
                    inhalt = MeinSystemFormat(shQ.Cells(zeileEPS, spalteLJ - 1).Value)
                    If IsNumeric(inhalt) Then
                        sh.Cells(zeile, SPALTE_EPSAJ).Value = CDbl(inhalt)
                    End If
                End If
                If spalteLJ > 3 Then
                    inhalt = MeinSystemFormat(shQ.Cells(zeileEPS, spalteLJ - 2).Value)
                    If IsNumeric(inhalt) Then
                        sh.Cells(zeile, SPALTE_EPSNJ).Value = CDbl(inhalt)
                    End If
                End If
                'Die Nullen von rechts entfernen - Schwäche von OnVista - bei EPS steht oftmals 0, wenn kein Wert bekannt bzw. geschätzt ist.
                'Die Wahrscheinlichkeit, dass 0 hier korrekt wäre, ist sehr klein.
                For K = SPALTE_EPSNJ To SPALTE_EPSLJ2 Step -1
                    If IsNumeric(sh.Cells(zeile, K).Value) Then
                        If sh.Cells(zeile, K).Value = 0 Then
                            sh.Cells(zeile, K).Value = ""
                        Else
                            Exit For
                        End If
                    End If
                Next
                'KGV:
                inhalt = MeinSystemFormat(shQ.Cells(zeileKGV, spalteLJ + 2).Value)
                If IsNumeric(inhalt) And (inhalt <> "") Then
                    sh.Cells(zeile, SPALTE_KGVLJ2).Value = CDbl(inhalt)
                Else
                    If gleichesLJ Then
                        'übernehmen evt. manuell eingepflegten Wert aus der vorigen Bewertung
                        sh.Cells(zeile, SPALTE_KGVLJ2).Value = shVorher.Cells(zeileVorher, SPALTE_KGVLJ2).Value
                    End If
                End If
                inhalt = MeinSystemFormat(shQ.Cells(zeileKGV, spalteLJ + 1).Value)
                If IsNumeric(inhalt) And (inhalt <> "") Then
                    sh.Cells(zeile, SPALTE_KGVLJ1).Value = CDbl(inhalt)
                Else
                    If gleichesLJ Then
                        'übernehmen evt. manuell eingepflegten Wert aus der vorigen Bewertung
                        sh.Cells(zeile, SPALTE_KGVLJ1).Value = shVorher.Cells(zeileVorher, SPALTE_KGVLJ1).Value
                    End If
                End If
                inhalt = MeinSystemFormat(shQ.Cells(zeileKGV, spalteLJ).Value)
                If IsNumeric(inhalt) And (inhalt <> "") Then
                    sh.Cells(zeile, SPALTE_KGVLJ).Value = CDbl(inhalt)
                Else
                    If gleichesLJ Then
                        'übernehmen evt. manuell eingepflegten Wert aus der vorigen Bewertung
                        sh.Cells(zeile, SPALTE_KGVLJ).Value = shVorher.Cells(zeileVorher, SPALTE_KGVLJ).Value
                    End If
                End If
                If spalteLJ > 2 Then
                    inhalt = MeinSystemFormat(shQ.Cells(zeileKGV, spalteLJ - 1).Value)
                    If IsNumeric(inhalt) Then
                        sh.Cells(zeile, SPALTE_KGVAJ).Value = CDbl(inhalt)
                    End If
                End If
            End If

        End If

    End If

End Function


Function AnalystenMeinungen(isin As String, zeile As Long, quellzeile As Long, shA As Worksheet, shQ As Worksheet, sh As Worksheet)
    'Holt die Anzahl der Analysten und deren Meinung, trägt sie ins Ziel-Bewertungsblatt in die Zeile zur ISIN ein.
    'benutzt dazu de.4-traders.com

    'Parameter:     isin    zu dieser ISIN werden die Daten geholt
    '               zeile   in diese Zeile des Bewertungsblattes werden die Daten geschrieben
    '               quellZeile  in dieser Zeile des Tabellenblattes "Aktien" stehen benötigte Hilfsdaten zur Aktie (URL-Teile für die Abfragen usw.)
    '               shA     Tabellenblatt "Aktien"
    '               shQ     Tabellenblatt "Query" - Hilfsblatt für die Web-Abfragen
    '               sh      Ziel-Bewertungsblatt

    'wird verwendet in DatenZurISINHolen

    Const URL_ANFANG = "URL;http://de.4-traders.com/"
    Const URL_ENDE = "/analystenerwartungen/"

    Dim urlTeil As String
    Dim url As String
    Dim rng As Range
    Dim Z As Long
    Dim S As Long
    Dim inhalt As String
    Dim ergebnis As Integer
    Dim anzahl As Integer

    sh.Cells(zeile, SPALTE_ANALYSTENANZAHL).Value = ""
    sh.Cells(zeile, SPALTE_ANALYSTENMEINUNG).Value = ""

    'Internet-Seite laden
    urlTeil = shA.Cells(quellzeile, SPALTE_4TRADERS).Value   'URL-Teil für 4-traders
    If (Len(urlTeil) > 1) And (Right(urlTeil, 1) = "/") Then
        urlTeil = Left(urlTeil, Len(urlTeil) - 1)
    End If
    If urlTeil = "" Then
        Exit Function
    End If
    url = URL_ANFANG + urlTeil + URL_ENDE

    If LadeURL(ANZAHL_VERSUCHE_WEBZUGRIFF, 1, shQ, url) = 0 Then

        Call StatusMeldung(zeile, sh, "Lade-F.", Replace(url, "URL;", ""), STATUS_FEHLER)

    Else

        'Daten extrahieren und ins Ziel-Bewertungsblatt eintragen
        sh.Cells(zeile, SPALTE_ANALYSTENANZAHL).Value = 0
        Set rng = shQ.Cells.Find("Durchschnittl. Empfehlung", shQ.Cells(1, 1))
        If Not rng Is Nothing Then

            Z = rng.Row
            S = rng.Column + 1
            inhalt = CStr(shQ.Cells(Z, S).Value)
            Select Case inhalt
                Case "KAUFEN": ergebnis = 1
                Case "AUFSTOCKEN": ergebnis = 2
                Case "HALTEN": ergebnis = 3
                Case "REDUZIEREN": ergebnis = 4
                Case "VERKAUFEN": ergebnis = 5
            End Select
            If ergebnis > 0 Then
                sh.Cells(zeile, SPALTE_ANALYSTENMEINUNG).Value = ergebnis
                inhalt = CStr(shQ.Cells(Z + 1, S).Value)
                If IsNumeric(inhalt) Then
                    anzahl = CInt(inhalt)
                    sh.Cells(zeile, SPALTE_ANALYSTENANZAHL).Value = anzahl
                End If
            End If
        End If

    End If

End Function


Function Marktkapitalisierung(isin As String, zeile As Long, quellzeile, shA As Worksheet, shQ As Worksheet, sh As Worksheet)
    'Holt die aktuelle Marktkapitalisierung, trägt sie ins Ziel-Bewertungsblatt in die Zeile zur ISIN ein.
    'benutzt dazu finanzen.net

    'Parameter:     isin    zu dieser ISIN werden die Daten geholt
    '               zeile   in diese Zeile des Bewertungsblattes werden die Daten geschrieben
    '               quellZeile  in dieser Zeile des Tabellenblattes "Aktien" stehen benötigte Hilfsdaten zur Aktie (URL-Teile für die Abfragen usw.)
    '               shA     Tabellenblatt "Aktien"
    '               shQ     Tabellenblatt "Query" - Hilfsblatt für die Web-Abfragen
    '               sh      Ziel-Bewertungsblatt

    'wird verwendet in DatenZurISINHolen

    Const URL_ANFANG = "URL;http://www.finanzen.net/aktien/"

    Dim urlTeil As String
    Dim url As String
    Dim rng As Range
    Dim K As Integer

    sh.Cells(zeile, SPALTE_MARKTKAP).Value = ""

    'Seite aus dem Web abrufen:
    urlTeil = shA.Cells(quellzeile, SPALTE_FINANZEN_NET).Value
    If urlTeil = "" Then
        Exit Function
    End If
    url = URL_ANFANG + urlTeil

    If LadeURL(ANZAHL_VERSUCHE_WEBZUGRIFF, 1, shQ, url) = 0 Then

        Call StatusMeldung(zeile, sh, "Lade-F.", Replace(url, "URL;", ""), STATUS_FEHLER)

    Else

        'Marktkapitalisierung auslesen und in Bewertungsblatt eintragen:
        Set rng = shQ.Cells.Find("Marktkapitalisierung (EUR)", shQ.Cells(1, 1))
        If Not rng Is Nothing Then
            For K = 1 To 5
                If CStr(shQ.Cells(rng.Row, rng.Column + K).Value) Like "* M??." Then
                    sh.Cells(zeile, SPALTE_MARKTKAP).Value = shQ.Cells(rng.Row, rng.Column + K).Value
                    Exit For
                End If
            Next
        End If

    End If

End Function


Function Quartalszahlen(isin As String, zeile As Long, quellzeile As Long, shA As Worksheet, shQ As Worksheet, sh As Worksheet, shVorher As Worksheet, zeileVorher As Long)
    'Ermittelt das Datum der letzten Zahlen, sowie die entsprechenden historischen Kurse der Aktie und der Benchmark (Vergleichsindex)
    'benutzt dazu finanzen.net (Datum der letzten Zahlen) und onvista.de zur Ermittlung der historischen Kurse
    'wenn das Datum der letzten Zahlen seit der vorigen Bewertung unverändert ist, werden die dazugehörigen Zahlen aus der vorigen Bewertung überträgen (spart Web-Zugriffe)

    'Parameter:     isin    zu dieser ISIN werden die Daten geholt
    '               zeile   in diese Zeile des Bewertungsblattes werden die Daten geschrieben
    '               quellZeile  in dieser Zeile des Tabellenblattes "Aktien" stehen benötigte Hilfsdaten zur Aktie (URL-Teile für die Abfragen usw.)
    '               shA     Tabellenblatt "Aktien"
    '               shQ     Tabellenblatt "Query" - Hilfsblatt für die Web-Abfragen
    '               sh      Ziel-Bewertungsblatt
    '               shVorher voriges Bewertungsblatt
    '               zeileVorher Zeile dieser Aktie im vorigen Bewertungsblatt

    'wird verwendet in DatenZurISINHolen

    Dim urlTeil As String
    Dim url As String
    Dim rng As Range
    Dim nurManuelleTermine As Boolean
    Dim zeileTermine As Long
    Dim zeileDatum As Long
    Dim Datum As Variant
    Dim DatumVortag As Variant
    Dim datumWeb As Variant
    Dim inhalt As String
    Dim heute As Variant
    Dim datumStamm As Variant
    Dim aktuellWarnung As Boolean
    Dim K As Long
    Dim idOnvista As String
    Dim idBOnvista As String
    Dim datumKurs As Variant
    Dim datumKursVortag As Variant
    Dim datumKursB As Variant
    Dim datumKursVortagB As Variant
    Dim gleichesDatum As Boolean
    Dim gleicheBenchmark As Boolean

    Const URL_ANFANG = "URL;http://www.finanzen.net/termine/"

    sh.Cells(zeile, SPALTE_BENCHMARK).Value = ""
    sh.Cells(zeile, SPALTE_DATUMZAHLEN).Value = ""
    sh.Cells(zeile, SPALTE_DATUMVORTAG).Value = ""
    sh.Cells(zeile, SPALTE_KURSZAHLEN).Value = ""
    sh.Cells(zeile, SPALTE_KURSVORTAG).Value = ""
    sh.Cells(zeile, SPALTE_BENCHMARKKURS).Value = ""
    sh.Cells(zeile, SPALTE_BENCHMARKVORTAG).Value = ""

    'Datum der letzten Zahlen herausfinden
    '------------------------------------------------------------------------------

    aktuellWarnung = True

    'aus den Stammdaten (Tabellenblatt "Aktien")
    datumStamm = Empty
    heute = DateValue(Now)          'DateValue ist OK, weil Now vom System kommt
    K = SPALTE_TERMINE
    Do Until CStr(shA.Cells(quellzeile, K).Value) = ""
        Datum = shA.Cells(quellzeile, K).Value
        If TypeName(Datum) <> "Date" Then
            Datum = Empty
        End If
        If Not IsEmpty(Datum) Then
            If Datum > heute Then
                aktuellWarnung = False      'die Stammdaten sind für die Zukunft gepflegt, so ist das verwendete Datum höchstwahrscheinlich aktuell
                Exit Do
            End If
            datumStamm = Datum
        End If
        K = K + 1
    Loop

    nurManuelleTermine = (CStr(shA.Cells(quellzeile, SPALTE_NUR_MANUELLE_TERMINE).Value) <> "")
    If nurManuelleTermine Then GoTo datumWebUeberspringen

    'von der Web-Seite
    urlTeil = shA.Cells(quellzeile, SPALTE_FINANZEN_NET).Value         'finanzen.net
    If urlTeil = "" Then
        GoTo datumWebUeberspringen
    End If
    url = URL_ANFANG + urlTeil
    If LadeURL(ANZAHL_VERSUCHE_WEBZUGRIFF, 1, shQ, url) = 0 Then
        Call StatusMeldung(zeile, sh, "Lade-F.", Replace(url, "URL;", ""), STATUS_FEHLER)
    Else
        'Termin der letzten Zahlen extrahieren
        datumWeb = Empty
        Set rng = shQ.Cells.Find("vergangene Termine", shQ.Cells(1, 1))
        If rng Is Nothing Then
            Call StatusMeldung(zeile, sh, "Daten-F.", Replace(url, "URL;", ""), STATUS_FEHLER)
        Else
            zeileTermine = rng.Row
            Set rng = shQ.Cells.Find("Terminart", shQ.Cells(zeileTermine, 1))
            If Not rng Is Nothing Then
                zeileTermine = rng.Row
                For K = 1 To 15
                    If (shQ.Cells(zeileTermine + K, 1) = "Quartalszahlen") Or (shQ.Cells(zeileTermine + K, 1).Value = "Jahresabschluss") Then
                        zeileDatum = zeileTermine + K
                        inhalt = CStr(shQ.Cells(zeileDatum, 4).Value)
                        datumWeb = DatumsWert(inhalt)
                        Exit For
                    End If
                Next
            End If
        End If
    End If
datumWebUeberspringen:

    Datum = datumWeb
    If IsEmpty(Datum) Then
        'kein Datum im Web gefunden oder gleich keins gesucht
        Datum = datumStamm
    Else
        'Datum im Web gefunden, steht in dem Stammdaten ein aktuelleres?
        If Not IsEmpty(datumStamm) Then
            If datumStamm > Datum Then
                Datum = datumStamm
            Else
                aktuellWarnung = False    'Datum aus der Terminliste im Web ist neuer, ist höchstwahrscheinlich aktuell
            End If
        Else
            aktuellWarnung = False  'Datum wurde ausschließlich im Web gefunden, ist höchstwahrscheinlich aktuell
        End If
    End If

    'Historische Kurse zum (Quartalszahlen-)Datum und Vortag für Aktie und Benchmark ermitteln:
    '------------------------------------------------------------------------------

    'Benchmark und QZ-Datum eintragen:
    sh.Cells(zeile, SPALTE_BENCHMARK).Value = shA.Cells(quellzeile, SPALTE_BENCHMARK_NAME).Value
    If Not IsEmpty(Datum) Then

        sh.Cells(zeile, SPALTE_DATUMZAHLEN).Value = Datum
        If aktuellWarnung Then
            Call StatusMeldung(zeile, sh, "QZ aktuell?", "", STATUS_WARNUNG)
        End If

        gleichesDatum = False
        gleicheBenchmark = False
        If zeileVorher > 1 Then
            If CStr(shVorher.Cells(zeileVorher, SPALTE_DATUMZAHLEN).Value) = CStr(Datum) Then
                gleichesDatum = True
                If CStr(sh.Cells(zeile, SPALTE_BENCHMARK).Value) = CStr(shVorher.Cells(zeileVorher, SPALTE_BENCHMARK).Value) Then
                    gleicheBenchmark = True
                End If
            End If
        End If

        'Historische Kurse für das QZ-Datum und den Vortag:
        'zuerst für die Benchmark
        idBOnvista = shA.Cells(quellzeile, SPALTE_BENCHMARK_HIST_ONVISTA).Value
        If gleicheBenchmark Then
            'aus dem vorigen Tabellenblatt übernehmen
            sh.Cells(zeile, SPALTE_DATUMVORTAG).Value = shVorher.Cells(zeileVorher, SPALTE_DATUMVORTAG).Value
            sh.Cells(zeile, SPALTE_BENCHMARKKURS).Value = shVorher.Cells(zeileVorher, SPALTE_BENCHMARKKURS).Value
            sh.Cells(zeile, SPALTE_BENCHMARKVORTAG).Value = shVorher.Cells(zeileVorher, SPALTE_BENCHMARKVORTAG).Value
        Else
            'aktuell aus dem Web holen
            If (idBOnvista = "") Then
                Exit Function
            End If
            datumKursB = DatumKursHistorisch(idBOnvista, CDate(Datum), shQ, zeile, sh, True)
            If datumKursB(0) <> "" Then
                Datum = DatumsWert(datumKursB(0))
                If Datum < sh.Cells(zeile, SPALTE_DATUMZAHLEN).Value Then
                    Call StatusMeldung(zeile, sh, "QZ: " + Format(sh.Cells(zeile, SPALTE_DATUMZAHLEN).Value, "dd.mm.yy"), "", STATUS_WARNUNG)  'der historische Kurs stammt nicht vom QZ-Datum
                    sh.Cells(zeile, SPALTE_DATUMZAHLEN).Value = Datum
                    gleichesDatum = False
                End If
                If IsNumeric(datumKursB(1)) Then
                    sh.Cells(zeile, SPALTE_BENCHMARKKURS).Value = CDbl(datumKursB(1))
                End If
                'Vortag
                DatumVortag = DateAdd("d", -1, Datum)
                datumKursVortagB = DatumKursHistorisch(idBOnvista, CDate(DatumVortag), shQ, zeile, sh, True)
                If datumKursVortagB(0) <> "" Then
                    DatumVortag = DatumsWert(datumKursVortagB(0))
                    sh.Cells(zeile, SPALTE_DATUMVORTAG).Value = DatumVortag
                    If IsNumeric(datumKursVortagB(1)) Then
                        sh.Cells(zeile, SPALTE_BENCHMARKVORTAG).Value = CDbl(datumKursVortagB(1))
                    End If
                End If
            End If
        End If

        'historische Kurse für die Aktie
        If gleichesDatum Then
            'aus dem vorigen Bewertungsblatt übernehmen
            sh.Cells(zeile, SPALTE_KURSZAHLEN).Value = shVorher.Cells(zeileVorher, SPALTE_KURSZAHLEN).Value
            sh.Cells(zeile, SPALTE_DATUMVORTAG).Value = shVorher.Cells(zeileVorher, SPALTE_DATUMVORTAG).Value
            sh.Cells(zeile, SPALTE_KURSVORTAG).Value = shVorher.Cells(zeileVorher, SPALTE_KURSVORTAG).Value
        Else
            'aktuell aus dem Web holen
            idOnvista = shA.Cells(quellzeile, SPALTE_HIST_ONVISTA_ORIG).Value   'Abfrage in anderer Währung als Euro
            If idOnvista = "" Then
                idOnvista = shA.Cells(quellzeile, SPALTE_HIST_ONVISTA).Value
            End If
            If (idOnvista = "") Then
                Exit Function
            End If
            datumKurs = DatumKursHistorisch(idOnvista, CDate(Datum), shQ, zeile, sh, False, idBOnvista)
            If datumKurs(0) <> "" Then
                If datumKurs(0) = Format(sh.Cells(zeile, SPALTE_DATUMZAHLEN).Value, "yyyy-mm-dd") Then
                    If IsNumeric(datumKurs(1)) Then
                        sh.Cells(zeile, SPALTE_KURSZAHLEN).Value = CDbl(datumKurs(1))
                    End If
                    'Vortag
                    DatumVortag = DateAdd("d", -1, Datum)
                    datumKursVortag = DatumKursHistorisch(idOnvista, CDate(DatumVortag), shQ, zeile, sh, False, idBOnvista)
                    If datumKursVortag(0) <> "" Then
                        If datumKursVortag(0) = Format(sh.Cells(zeile, SPALTE_DATUMVORTAG).Value, "yyyy-mm-dd") Then
                            If IsNumeric(datumKursVortag(1)) Then
                                sh.Cells(zeile, SPALTE_KURSVORTAG).Value = CDbl(datumKursVortag(1))
                            End If
                        End If
                    End If
                End If
            End If

        End If

    End If

End Function


Function GewinnRevisionen(isin As String, zeile As Long, sh As Worksheet, shVorher As Worksheet, zeileVorher As Long)
    'Übertragt die geschätzten EPS-Werte für das aktuelle und das nächste Jahr aus der vorigen Bewertung in diese Ziel-Bewertung

    'Parameter: isin    es geht um die Aktie mit dieser ISIN
    '           zeile   Zeilennummer der Aktie im aktuellen Bewertungsblatt
    '           sh      aktuelles Bewertungsblatt
    '           shVorher    voriges Bewertungsblatt
    '           zeileVorher Zeile dieser Aktie im vorigen Bewertungsblatt

    'wird verwendet in DatenZurISINHolen

    Dim K As Integer

    For K = 4 To 1 Step -1
            sh.Cells(zeile, SPALTE_EPSAJ_WDH - K).Value = ""
            sh.Cells(zeile, SPALTE_EPSNJ_WDH - K).Value = ""
    Next

    If zeileVorher > 1 Then

        'EPS-Daten übertragen:
        If CStr(sh.Cells(zeile, SPALTE_LJ).Value) = CStr(shVorher.Cells(zeileVorher, SPALTE_LJ).Value) Then
            'kein Geschäftsjahreswechsel
            For K = 4 To 1 Step -1
                sh.Cells(zeile, SPALTE_EPSAJ_WDH - K).Value = shVorher.Cells(zeileVorher, SPALTE_EPSAJ_WDH - K + 1).Value
                sh.Cells(zeile, SPALTE_EPSNJ_WDH - K).Value = shVorher.Cells(zeileVorher, SPALTE_EPSNJ_WDH - K + 1).Value
            Next
        Else
            'Geschäftsjahreswechsel - aus den Schätzungen für das nächste Jahr werden nun Schätzungen für das aktuelle Jahr
            For K = 4 To 1 Step -1
                sh.Cells(zeile, SPALTE_EPSAJ_WDH - K).Value = shVorher.Cells(zeileVorher, SPALTE_EPSNJ_WDH - K + 1).Value
            Next
        End If

    End If

End Function


Function HistorischeKurse(isin As String, zeile As Long, quellzeile As Long, shA As Worksheet, shQ As Worksheet, sh As Worksheet)
    'Ermittelt jeweils Datum und Kurs der Aktie von vor 6 Monaten bzw. 1 Jahr.

    'Parameter:     isin    zu dieser ISIN werden die Daten geholt
    '               zeile   in diese Zeile des Bewertungsblattes werden die Daten geschrieben
    '               quellZeile  in dieser Zeile des Tabellenblattes "Aktien" stehen benötigte Hilfsdaten zur Aktie (URL-Teile für die Abfragen usw.)
    '               shA     Tabellenblatt "Aktien"
    '               shQ     Tabellenblatt "Query" - Hilfsblatt für die Web-Abfragen
    '               sh      Ziel-Bewertungsblatt

    'wird verwendet in DatenZurISINHolen

    Dim idOnvista As String
    Dim idOnvistaB As String
    Dim inhalt As String
    Dim Datum As Date
    Dim DatumHist As Date
    Dim datumKurs As Variant

    sh.Cells(zeile, SPALTE_DATUM6MON).Value = ""
    sh.Cells(zeile, SPALTE_KURS6MON).Value = ""
    sh.Cells(zeile, SPALTE_DATUM1JAHR).Value = ""
    sh.Cells(zeile, SPALTE_KURS1JAHR).Value = ""

    idOnvista = shA.Cells(quellzeile, SPALTE_HIST_ONVISTA).Value
    If (idOnvista = "") Then
        Exit Function
    End If
    idOnvistaB = shA.Cells(quellzeile, SPALTE_BENCHMARK_HIST_ONVISTA).Value

    inhalt = CStr(sh.Cells(zeile, SPALTE_DATUM).Value)
    If inhalt = "" Then
        Exit Function
    End If
    Datum = CDate(inhalt)

    'vor 6 Monaten:
    DatumHist = DateAdd("m", -6, Datum)
    datumKurs = DatumKursHistorisch(idOnvista, DatumHist, shQ, zeile, sh, False, idOnvistaB)
    If datumKurs(0) <> "" Then
        sh.Cells(zeile, SPALTE_DATUM6MON).Value = DatumsWert(datumKurs(0))
        If IsNumeric(datumKurs(1)) Then
            sh.Cells(zeile, SPALTE_KURS6MON).Value = CDbl(datumKurs(1))
        End If
    End If

    'vor 1 Jahr:
    DatumHist = DateAdd("yyyy", -1, Datum)
    datumKurs = DatumKursHistorisch(idOnvista, DatumHist, shQ, zeile, sh, False, idOnvistaB)
    If datumKurs(0) <> "" Then
        sh.Cells(zeile, SPALTE_DATUM1JAHR).Value = DatumsWert(datumKurs(0))
        If IsNumeric(datumKurs(1)) Then
            sh.Cells(zeile, SPALTE_KURS1JAHR).Value = CDbl(datumKurs(1))
        End If
    End If

End Function


Function DreiMonatsReversal(isin As String, zeile As Long, quellzeile As Long, shA As Worksheet, shQ As Worksheet, sh As Worksheet, shVorher As Worksheet, zeileVorher As Long)
    'Ermittelt für Aktie und Benchmark die Schlusskurse der letzten vier Monate, damit danach die letzten drei Monatsentwicklungen verglichen werden können.
    'nur für Large Caps
    'wenn sich nichts geändert hat (Datum, Benchmark), werden die Daten aus der vorigen Bewertung übernommen, aber nur, wenn sie dort vollständig vorhanden waren

    'Parameter:     isin    zu dieser ISIN werden die Daten geholt
    '               zeile   in diese Zeile des Bewertungsblattes werden die Daten geschrieben
    '               quellZeile  in dieser Zeile des Tabellenblattes "Aktien" stehen benötigte Hilfsdaten zur Aktie (URL-Teile für die Abfragen usw.)
    '               shA     Tabellenblatt "Aktien"
    '               shQ     Tabellenblatt "Query" - Hilfsblatt für die Web-Abfragen
    '               sh      Ziel-Bewertungsblatt
    '               shVorher    Tabellenblatt zur vorigen Bewertung
    '               zeileVorher Zeile dieser Aktie in der vorigen Bewertung

    'wird verwendet in DatenZurISINHolen

    Dim idOnvista As String
    Dim idBOnvista As String
    Dim Datum As Date
    Dim web0 As Boolean, web1 As Boolean, web2 As Boolean, web3 As Boolean
    Dim datumKurs As Variant
    Dim K As Integer
    Dim gleicheMon As Boolean
    Dim kurseVonVorher As Boolean
    Dim gleicheBenchmark As Boolean
    Dim datumVorher As Variant

    sh.Cells(zeile, SPALTE_DATUM3MON).Value = ""
    sh.Cells(zeile, SPALTE_KURS3MON).Value = ""
    sh.Cells(zeile, SPALTE_BENCHMARK3MON).Value = ""
    sh.Cells(zeile, SPALTE_DATUM2MON).Value = ""
    sh.Cells(zeile, SPALTE_KURS2MON).Value = ""
    sh.Cells(zeile, SPALTE_BENCHMARK2MON).Value = ""
    sh.Cells(zeile, SPALTE_DATUM1MON).Value = ""
    sh.Cells(zeile, SPALTE_KURS1MON).Value = ""
    sh.Cells(zeile, SPALTE_BENCHMARK1MON).Value = ""
    sh.Cells(zeile, SPALTE_DATUM0MON).Value = ""
    sh.Cells(zeile, SPALTE_KURS0MON).Value = ""
    sh.Cells(zeile, SPALTE_BENCHMARK0MON).Value = ""

    'nur für Large Caps:
    If CStr(sh.Cells(zeile, SPALTE_GROESSE).Value) <> "L" Then
        Exit Function
    End If

    'Ids für historische Kursabfrage bei OnVista:
    idOnvista = CStr(shA.Cells(quellzeile, SPALTE_HIST_ONVISTA_ORIG).Value)
    If idOnvista = "" Then
        idOnvista = CStr(shA.Cells(quellzeile, SPALTE_HIST_ONVISTA).Value)
    End If
    idBOnvista = CStr(shA.Cells(quellzeile, SPALTE_BENCHMARK_HIST_ONVISTA).Value)

    'letztes Monatsende vor bzw. zum Kursdatum ermitteln, Kurs der Aktie und Benchmark am letzten Börsentag dazu abfragen
    If CStr(sh.Cells(zeile, SPALTE_DATUM).Value) = "" Then
        Exit Function
    End If
    Datum = CDate(CStr(sh.Cells(zeile, SPALTE_DATUM).Value))
    Datum = DateAdd("d", 1, Datum)
    Datum = CDate(DatumsWert("01." + Right("0" + CStr(Month(Datum)), 2) + "." + CStr(Year(Datum))))
    Datum = DateAdd("d", -1, Datum)

    'zum Nachsehen, ob das beim letzten Lauf schon genauso ausgewertet wurde
    gleicheMon = False
    gleicheBenchmark = False
    kurseVonVorher = False

    If zeileVorher > 1 Then
        If CStr(shVorher.Cells(zeileVorher, SPALTE_GROESSE).Value) = "L" Then
            datumVorher = shVorher.Cells(zeileVorher, SPALTE_DATUM0MON).Value
            If CStr(datumVorher) <> "" Then
                If IsDate(datumVorher) Then
                    If (Year(datumVorher) = Year(Datum)) And (Month(datumVorher) = Month(Datum)) Then
                        gleicheMon = True
                    End If
                End If
            End If
        End If
    End If

    'Kurse für die Benchmark am Monatsende könnten schon für eine der Aktien davor geholt worden sein:
    For K = zeile - 1 To 2 Step -1
        If (CStr(sh.Cells(K, SPALTE_GROESSE).Value) = "L") And (CStr(sh.Cells(K, SPALTE_BENCHMARK).Value) = CStr(shA.Cells(quellzeile, SPALTE_BENCHMARK_NAME).Value)) And IsNumeric(sh.Cells(K, SPALTE_PKT_3MONREV).Value) Then
            If IsDate(sh.Cells(K, SPALTE_DATUM0MON).Value) Then
                If CStr(sh.Cells(K, SPALTE_DATUM0MON).Value) = CStr(sh.Cells(zeile, SPALTE_DATUM0MON).Value) Then
                    'Datumswerte und Kurse für die Benchmark können aus dieser Zeile übertragen werden:
                    sh.Cells(zeile, SPALTE_DATUM3MON).Value = sh.Cells(K, SPALTE_DATUM3MON).Value
                    sh.Cells(zeile, SPALTE_BENCHMARK3MON).Value = sh.Cells(K, SPALTE_BENCHMARK3MON).Value
                    sh.Cells(zeile, SPALTE_DATUM2MON).Value = sh.Cells(K, SPALTE_DATUM2MON).Value
                    sh.Cells(zeile, SPALTE_BENCHMARK2MON).Value = sh.Cells(K, SPALTE_BENCHMARK2MON).Value
                    sh.Cells(zeile, SPALTE_DATUM1MON).Value = sh.Cells(K, SPALTE_DATUM1MON).Value
                    sh.Cells(zeile, SPALTE_BENCHMARK1MON).Value = sh.Cells(K, SPALTE_BENCHMARK1MON).Value
                    sh.Cells(zeile, SPALTE_DATUM0MON).Value = sh.Cells(K, SPALTE_DATUM0MON).Value
                    sh.Cells(zeile, SPALTE_BENCHMARK0MON).Value = sh.Cells(K, SPALTE_BENCHMARK0MON).Value
                    GoTo benchmarkErledigt
                End If
            End If
        End If
    Next

    'schauen, ob die Kurse für die Benchmark vom vorigen Bewertungsblatt geholt werden können
    If gleicheMon Then
        If CStr(shVorher.Cells(zeileVorher, SPALTE_BENCHMARK).Value) = CStr(shA.Cells(quellzeile, SPALTE_BENCHMARK_NAME).Value) Then
            If IsNumeric(shVorher.Cells(zeileVorher, SPALTE_PKT_3MONREV).Value) Then
                gleicheBenchmark = True
            End If
        End If
    End If

    If gleicheBenchmark Then

        'Benchmark-Vergleichskurs vom vorigen Bewertungsblatt holen
        sh.Cells(zeile, SPALTE_DATUM3MON).Value = shVorher.Cells(zeileVorher, SPALTE_DATUM3MON).Value
        sh.Cells(zeile, SPALTE_BENCHMARK3MON).Value = shVorher.Cells(zeileVorher, SPALTE_BENCHMARK3MON).Value
        sh.Cells(zeile, SPALTE_DATUM2MON).Value = shVorher.Cells(zeileVorher, SPALTE_DATUM2MON).Value
        sh.Cells(zeile, SPALTE_BENCHMARK2MON).Value = shVorher.Cells(zeileVorher, SPALTE_BENCHMARK2MON).Value
        sh.Cells(zeile, SPALTE_DATUM1MON).Value = shVorher.Cells(zeileVorher, SPALTE_DATUM1MON).Value
        sh.Cells(zeile, SPALTE_BENCHMARK1MON).Value = shVorher.Cells(zeileVorher, SPALTE_BENCHMARK1MON).Value
        sh.Cells(zeile, SPALTE_DATUM0MON).Value = shVorher.Cells(zeileVorher, SPALTE_DATUM0MON).Value
        sh.Cells(zeile, SPALTE_BENCHMARK0MON).Value = shVorher.Cells(zeileVorher, SPALTE_BENCHMARK0MON).Value
        GoTo benchmarkErledigt

    Else

        'Kurse für die Benchmark müssen aus dem Internet geladen werden:
        datumKurs = DatumKursHistorisch(idBOnvista, Datum, shQ, zeile, sh, True)
        If (datumKurs(0) <> "") And (datumKurs(1) <> "") Then
            sh.Cells(zeile, SPALTE_DATUM0MON).Value = DatumsWert(datumKurs(0))
            sh.Cells(zeile, SPALTE_BENCHMARK0MON).Value = CDbl(datumKurs(1))
        End If
        'einen Monat weiter zurück:
        Datum = CDate(DatumsWert("01." + Right("0" + CStr(Month(Datum)), 2) + "." + CStr(Year(Datum))))
        Datum = DateAdd("d", -1, Datum)
        datumKurs = DatumKursHistorisch(idBOnvista, Datum, shQ, zeile, sh, True)
        If (datumKurs(0) <> "") And (datumKurs(1) <> "") Then
            sh.Cells(zeile, SPALTE_DATUM1MON).Value = DatumsWert(datumKurs(0))
            sh.Cells(zeile, SPALTE_BENCHMARK1MON).Value = CDbl(datumKurs(1))
        End If
        'einen zweiten Monat weiter zurück:
        Datum = CDate(DatumsWert("01." + Right("0" + CStr(Month(Datum)), 2) + "." + CStr(Year(Datum))))
        Datum = DateAdd("d", -1, Datum)
        datumKurs = DatumKursHistorisch(idBOnvista, Datum, shQ, zeile, sh, True)
        If (datumKurs(0) <> "") And (datumKurs(1) <> "") Then
            sh.Cells(zeile, SPALTE_DATUM2MON).Value = DatumsWert(datumKurs(0))
            sh.Cells(zeile, SPALTE_BENCHMARK2MON).Value = CDbl(datumKurs(1))
        End If
        'einen dritten Monat weiter zurück:
        Datum = CDate(DatumsWert("01." + Right("0" + CStr(Month(Datum)), 2) + "." + CStr(Year(Datum))))
        Datum = DateAdd("d", -1, Datum)
        datumKurs = DatumKursHistorisch(idBOnvista, Datum, shQ, zeile, sh, True)
        If (datumKurs(0) <> "") And (datumKurs(1) <> "") Then
            sh.Cells(zeile, SPALTE_DATUM3MON).Value = DatumsWert(datumKurs(0))
            sh.Cells(zeile, SPALTE_BENCHMARK3MON).Value = CDbl(datumKurs(1))
        End If

    End If

benchmarkErledigt:

    If gleicheMon Then
        If IsNumeric(shVorher.Cells(zeileVorher, SPALTE_PKT_3MONREV).Value) Then
            kurseVonVorher = True
        End If
    End If

    If kurseVonVorher Then
        If CStr(sh.Cells(zeile, SPALTE_DATUM3MON).Value) = CStr(shVorher.Cells(zeileVorher, SPALTE_DATUM3MON).Value) Then
            sh.Cells(zeile, SPALTE_KURS3MON).Value = shVorher.Cells(zeileVorher, SPALTE_KURS3MON).Value
        Else
            web3 = True  'doch aus dem Web holen, da abweichendes Datum
        End If
        If CStr(sh.Cells(zeile, SPALTE_DATUM2MON).Value) = CStr(shVorher.Cells(zeileVorher, SPALTE_DATUM2MON).Value) Then
            sh.Cells(zeile, SPALTE_KURS2MON).Value = shVorher.Cells(zeileVorher, SPALTE_KURS2MON).Value
        Else
            web2 = True
        End If
        If CStr(sh.Cells(zeile, SPALTE_DATUM1MON).Value) = CStr(shVorher.Cells(zeileVorher, SPALTE_DATUM1MON).Value) Then
            sh.Cells(zeile, SPALTE_KURS1MON).Value = shVorher.Cells(zeileVorher, SPALTE_KURS1MON).Value
        Else
            web1 = True
        End If
        If CStr(sh.Cells(zeile, SPALTE_DATUM0MON).Value) = CStr(shVorher.Cells(zeileVorher, SPALTE_DATUM0MON).Value) Then
            sh.Cells(zeile, SPALTE_KURS0MON).Value = shVorher.Cells(zeileVorher, SPALTE_KURS0MON).Value
        Else
            web0 = True
        End If

    Else

        'alles aus dem Web holen
        web0 = True
        web1 = True
        web2 = True
        web3 = True

    End If

    'Daten aus dem Web holen
    If (idOnvista = "") Or (idBOnvista = "") Then
        Exit Function
    End If
    If web0 And (CStr(sh.Cells(zeile, SPALTE_DATUM0MON).Value) <> "") Then
        datumKurs = DatumKursHistorisch(idOnvista, sh.Cells(zeile, SPALTE_DATUM0MON).Value, shQ, zeile, sh, False, idBOnvista)
        If (datumKurs(0) = Format(sh.Cells(zeile, SPALTE_DATUM0MON).Value, "yyyy-mm-dd")) And (datumKurs(1) <> "") Then
            sh.Cells(zeile, SPALTE_KURS0MON).Value = CDbl(datumKurs(1))
        End If
    End If
    If web1 And (CStr(sh.Cells(zeile, SPALTE_DATUM1MON).Value) <> "") Then
        datumKurs = DatumKursHistorisch(idOnvista, sh.Cells(zeile, SPALTE_DATUM1MON).Value, shQ, zeile, sh, False, idBOnvista)
        If (datumKurs(0) = Format(sh.Cells(zeile, SPALTE_DATUM1MON).Value, "yyyy-mm-dd")) And (datumKurs(1) <> "") Then
            sh.Cells(zeile, SPALTE_KURS1MON).Value = CDbl(datumKurs(1))
        End If
    End If
    If web2 And (CStr(sh.Cells(zeile, SPALTE_DATUM2MON).Value) <> "") Then
        datumKurs = DatumKursHistorisch(idOnvista, sh.Cells(zeile, SPALTE_DATUM2MON).Value, shQ, zeile, sh, False, idBOnvista)
        If (datumKurs(0) = Format(sh.Cells(zeile, SPALTE_DATUM2MON).Value, "yyyy-mm-dd")) And (datumKurs(1) <> "") Then
            sh.Cells(zeile, SPALTE_KURS2MON).Value = CDbl(datumKurs(1))
        End If
    End If
    If web3 And (CStr(sh.Cells(zeile, SPALTE_DATUM3MON).Value) <> "") Then
        datumKurs = DatumKursHistorisch(idOnvista, sh.Cells(zeile, SPALTE_DATUM3MON).Value, shQ, zeile, sh, False, idBOnvista)
        If (datumKurs(0) = Format(sh.Cells(zeile, SPALTE_DATUM3MON).Value, "yyyy-mm-dd")) And (datumKurs(1) <> "") Then
            sh.Cells(zeile, SPALTE_KURS3MON).Value = CDbl(datumKurs(1))
        End If
    End If

End Function


Function Bemerkungen(isin As String, zeile As Long, sh As Worksheet, shVorher As Worksheet, zeileVorher As Long)
    'Übertragt die Bemerkungen soweit vorhanden aus dem vorigen Blatt ins aktuelle mit davorstehendem Datum,
    'um zu zeigen, aus welcher Bewertung diese stammen

    'Parameter: isin    es geht um die Aktie mit dieser ISIN
    '           zeile   Zeilennummer der Aktie im aktuellen Bewertungsblatt
    '           sh      aktuelles Bewertungsblatt
    '           shVorher    voriges Bewertungsblatt
    '           zeileVorher Zeile dieser Aktie im vorigen Bewertungsblatt

    'wird verwendet in DatenZurISINHolen

    Dim bemerkung As String

    sh.Cells(zeile, SPALTE_BEMERKUNGEN).Value = ""

    If zeileVorher > 1 Then
        bemerkung = CStr(shVorher.Cells(zeileVorher, SPALTE_BEMERKUNGEN).Value)
        If bemerkung <> "" Then
            If bemerkung Like "####-##-##: *" Then
                sh.Cells(zeile, SPALTE_BEMERKUNGEN).Value = bemerkung
            Else
                sh.Cells(zeile, SPALTE_BEMERKUNGEN).Value = shVorher.Name + ": " + bemerkung
            End If
        End If
    End If

End Function


Function DatumKursHistorisch(Id As String, Datum As Date, shQ As Worksheet, zeile As Long, sh As Worksheet, fuerBenchmark As Boolean, Optional benchmarkId As String) As Variant
    'Holt zu einer Aktie und einem Datum den letzten Kurs zu bzw vor diesem Datum. (Vor diesem Datum, falls das kein Börsentag war.)
    'Verwendet dazu OnVista.

    'Parameter:     id      interne OnVista-Id zum Auffinden der Aktie bzw. des Index
    '                       Mehrfachwerte durch Komma getrennt gehen auch. In Klammern darf jeweils der Börsenplatz dahinter stehen
    '               datum   an bzw. vor diesem Datum soll der Kurs ermittelt werden
    '               shQ     Tabellenblatt "Query" - Hilfsblatt für die Web-Abfragen
    '               zeile   aktuelle Zeile des Bewertungsblattes
    '               sh      Bewertungsblatt
    '               fuerBenchmark  True, wenn es sich um Kursabfrage für eine Benchmark (Index) handelt
    '               benchmarkId Id der zugehörigen Benchmark (für Aktie), um zu überprüfen, ob der richtige Börsentag herausgekommen ist

    'Rückgabewert:  String-Array der Länge 2: bei Index 0 steht das Datum, bei Index 1 steht der Kurs als yyyy-mm-dd

    'verwendet in: Quartalszahlen,Dreimonatsreversal, HistorischeKurse

    Dim datumKurs(1) As String
    Dim datumKursDefault(1) As String
    Dim Ids As Variant
    Dim eineId As String
    Dim datumVon As Date
    Dim url As String
    Dim inhalt As String
    Dim K As Integer, J As Integer
    Dim V As Variant
    Dim datumStr As String
    Dim kursDatum As Date
    Dim testDatum As Variant
    Dim benchmarkBT As String

    'Web-Abfrage - Kurse um das gewünschte Datum herum (ab 14 Tage zurück für einen Monat):
    datumVon = DateAdd("d", -14, Datum)

    'in id können mehrere Onvista-hist-Ids stehen, die durch Komme getrennt sind. Hinter jeder Id kann in Klammern ein Börsenplatz stehen.
    'die Angabe der Börsenplatzes im Klartext dient nur der besseren Lesbarkeit für den Nutzer.
    Ids = Split(Id, ",")
    For J = 0 To UBound(Ids)

        V = Split(Ids(J), "(")
        eineId = Trim(V(0))
        url = "URL;http://www.onvista.de/onvista/boxes/historicalquote/export.csv?notationId=" + eineId + "&dateStart=" + Format(datumVon, "dd.mm.yyyy") + "&interval=M1"

        If LadeURL(ANZAHL_VERSUCHE_WEBZUGRIFF, 1, shQ, url) = 0 Then

            If Not (sh Is Nothing) Then
                Call StatusMeldung(zeile, sh, "Lade-F.", Replace(url, "URL;", ""), STATUS_FEHLER)
            End If

        Else

            'Datum und Kurs heraussuchen:
            For K = 30 To 2 Step -1
                inhalt = CStr(shQ.Cells(K, 1).Value)
                If inhalt <> "" Then
                    V = Split(inhalt, ";")
                    datumStr = Trim(V(0))
                    If UBound(V) >= 4 Then
                        testDatum = DatumsWert(datumStr)
                        If Not IsEmpty(testDatum) Then
                            kursDatum = CDate(testDatum)
                            If kursDatum <= Datum Then
                                datumKurs(0) = Format(kursDatum, "yyyy-mm-dd")
                                datumKurs(1) = MeinSystemFormat(V(4))
                                GoTo datumKursGefunden
                            End If
                        End If
                    End If
                End If
            Next
datumKursGefunden:

            If datumKurs(0) <> "" Then
                'heben das erste beste Datum-Kurs-Paar auf, falls kein genau passendes gefunden wird
                If datumKursDefault(0) = "" Then
                    datumKursDefault(0) = datumKurs(0)
                    datumKursDefault(1) = datumKurs(1)
                End If

                If fuerBenchmark Then
                    'für Benchmark wird das immer als richtig angenommen
                    'Kombination aus Id und Datum und das Börsentagsdatum merken
                    Call MerkeBoersentag(eineId, Datum, datumKurs(0))
                    GoTo datumKursOK
                Else
                    'anhand der Benchmark-Id in Kombination mit dem Datum prüfen, ob das richtige Börsentagsdatum herausgekommen ist.
                    If benchmarkId <> "" Then
                        benchmarkBT = Boersentag(benchmarkId, Datum, shQ)
                        If benchmarkBT = datumKurs(0) Then
                            'es ist der richtige Börsentag gefunden worden
                            GoTo datumKursOK
                        End If
                    Else
                        'richtiger Börsentag ist nicht nachprüfbar
                        GoTo datumKursOK
                    End If
                End If
            End If

        End If

    Next
    datumKurs(0) = datumKursDefault(0)
    datumKurs(1) = datumKursDefault(1)

    'Das Datum könnte falsch sein.
    If Not (sh Is Nothing) Then
        If datumKurs(0) <> "" Then
            V = Split(datumKurs(0), "-")
            Call StatusMeldung(zeile, sh, "K " + V(2) + "." + V(1) + "." + Right(V(0), 2) + " ?", "", STATUS_WARNUNG)
        End If
    End If

datumKursOK:
    DatumKursHistorisch = datumKurs

End Function


Function MerkeBoersentag(benchmarkId As String, Datum As Date, BTDatum As String)
    'Merkt sich in globalen Array-Variablen, dass bei Abfrage von Datum zur Onvista-hist-Id der Benchmark benchmarkId
    'BTDatum als Börsentag herauskommt

    'Parameter:     benchmarkId ist die Onvista-hist-Id eines Index, also ein paar Ziffern
    '               Datum ist ein Datumswert, zu dem ein Kurs abgefragt wurde
    '               BTDatum ist der Börsentag, der zu diesem Datumswert passt in der Form yyyy-mm-dd

    Dim idDatum As String
    Dim K As Integer

    idDatum = benchmarkId + "#" + Format(Datum, "yyyy-mm-dd")
    For K = 0 To BTU
        If BTIdDatum(K) = idDatum Then
            'ist schon da
            Exit Function
        End If
    Next
    BTU = BTU + 1
    ReDim Preserve BTIdDatum(BTU)
    ReDim Preserve BT(BTU)
    BTIdDatum(BTU) = idDatum
    BT(BTU) = BTDatum

End Function


Function Boersentag(benchmarkId As String, Datum As Date, shQ As Worksheet) As String
    'Ermittelt den passenden Boersentag anhand der OnVista hist-Id der Benchmark und Datum
    'schaut zunächst in den globalen BT-Arrays nach, wenn es dort nicht steht, wird Abfrage im Web ausgeführt

    'Parameter:     benchmarkId OnVista-hist-Id der Benchmark
    '               Datum
    '               shQ = Hilfsblatt für Web-Abfrage

    'Rückgabewert hat die Form yyyy-mm-dd

    Dim idDatum As String
    Dim K As Integer
    Dim datumKurs As Variant

    idDatum = benchmarkId + "#" + Format(Datum, "yyyy-mm-dd")
    'in den globalen BT-Arrays nachsehen
    For K = 0 To BTU
        If BTIdDatum(K) = idDatum Then
            Boersentag = BT(K)
            Exit Function
        End If
    Next

    'Web-Abfrage ist nötig
    datumKurs = DatumKursHistorisch(benchmarkId, Datum, shQ, 0, Nothing, True)
    Boersentag = datumKurs(0)

End Function


Public Function LadeURL(maxVersuche As Integer, aktVersuch As Integer, shQ As Worksheet, url As String, Optional mitWebFormatierung As Boolean) As Integer
    'lädt den Inhalt der Seite mit der url ins Hilfsblatt shQ.
    'Falls ein Fehler auftritt, wird es neu versucht (rekursiver Aufruf)

    'Parameter:     maxVersuche = maximale Anzahl der Versuche
    '               aktVersuch  = Nummer des aktuellen Versuches
    '               shQ  = Hilfsblatt "Query" zum Durchführen der Web-Abfragen
    '               url  = URL die abgerufen werden soll

    'Rückgabewert:  beim wievielten Versuch es geklappt hat, 0 wenn es nicht geklappt hat

    Dim hatGeklappt As Integer
    Dim qt As QueryTable

    'Es ist zwar kein guter Stil, bei jedem Abruf einer URL zum Auslesen von Daten eine neue QueryTable anzulegen,
    'antatt das einmal am Anfang zu tun und dann wiederzuverwenden. Allerdings werden in dem Fall die Abfragen
    'mit zunehmender Anzahl immer langsamer.
    'Wenn man jedoch, wie hier umgesetzt, vor jedem Abrufen einer URL die QueryTable ganz neu erzeugt, läuft es schneller.

    'Blatt "Query" muss leer sein, Reste entfernen:
    Call LeereQuerySheet(shQ)

    'Abfrage vorbereiten:
    Set qt = shQ.QueryTables.Add(url, shQ.Cells(1, 1))
    With qt
        .Name = "Aktie"
        .BackgroundQuery = False                    'die Abfragen müssen immer synchron laufen
        .RefreshStyle = xlInsertDeleteCells
        .WebSelectionType = xlEntirePage
        If mitWebFormatierung Then
            .WebFormatting = xlWebFormattingAll
        Else
            .WebFormatting = xlWebFormattingNone
        End If
        .WebDisableDateRecognition = True
    End With
    shQ.Cells.NumberFormat = "@"

    'Abfrage auslösen
    DoEvents
    On Error Resume Next
    qt.Refresh (False)
    If Err.Number <> 0 Then
        Err.Clear
        If aktVersuch < maxVersuche Then
            aktVersuch = aktVersuch + 1
            Application.Wait (Now + TimeValue("0:00:03"))
            hatGeklappt = LadeURL(maxVersuche, aktVersuch, shQ, url)
        End If
    Else
        hatGeklappt = aktVersuch
    End If
    On Error GoTo 0
    LadeURL = hatGeklappt

End Function


Public Function LeereQuerySheet(shQ As Worksheet)
    'Query-Blatt aufräumen, Reste entfernen (QueryTables)

    'Parameter:     shQ    = Tabellenblatt "Query" - Hilfsblatt für die Web-Abfragen

    Dim qt As QueryTable

    shQ.Cells.Clear
    For Each qt In shQ.QueryTables
        qt.Delete
    Next

End Function


Function StatusMeldung(zeile As Long, sh As Worksheet, meldung As String, url As String, farbe As Long)
    'Rückmeldung ganz hinten im Bewertungsblatt, falls ein Fehler aufgetreten ist oder über eine Besonderheit gewarnt wird

    'Parameter:     zeile   aktuelle Zeile im Bewertungsblatt
    '               sh      Bewertungsblatt
    '               meldung Text, der als Meldung erscheint
    '               url     Webadresse, zu der die Statusmeldung verlinkt (kann auch "" sein)
    '               farbe   Farbe, die die Zelle bekommt (z.B. Rot-Ton für Fehler, Gelb-Ton für Warnung)

    'wird in diversen Functions verwendet, in denen Daten aus dem Internet gezogen werden

    Dim K As Integer

    'nächste freie Spalte finden
    K = SPALTE_STATUS
    Do Until CStr(sh.Cells(zeile, K).Value) = ""
        K = K + 1
    Loop

    'Meldung, Link und Farbe:
    sh.Cells(zeile, K).Value = meldung
    If url <> "" Then
        Call sh.Cells(zeile, K).Hyperlinks.Add(sh.Cells(zeile, K), url)
    End If
    sh.Cells(zeile, K).Interior.Color = farbe

End Function


Function LeereStatusMeldung(zeile As Long, sh As Worksheet)
    'leert die Status-Spalten in der aktuellen Zeile im Bewertungsblatt

    'Parameter:  zeile = aktuelle Zeile
    '            sh = aktuelles Bewertungsblatt

    Dim K As Integer
    Dim link As Hyperlink

    K = SPALTE_STATUS
    For K = SPALTE_STATUS To SPALTE_STATUS + 10
        For Each link In sh.Cells(zeile, K).Hyperlinks
            link.Delete
        Next
        sh.Cells(zeile, K).Value = ""
        sh.Cells(zeile, K).Interior.ColorIndex = xlColorIndexNone
    Next

End Function


Function HatFehler(zeile As Long, sh As Worksheet) As Boolean
    'Gibt an, ob nach einem Bewertungslauf eine Fehlermeldung in der betreffenden Zeile eines Bewertungsblattes steht.

    'Parameter  zeile   Zeile im Bewertungsblatt
    '           sh      Bewertungsblatt

    'Rückgabewert: True, wenn ein Fehler da steht, sonst False

    Dim K As Integer

    For K = 0 To 10
        If sh.Cells(zeile, SPALTE_STATUS + K).Cells.Interior.Color = STATUS_FEHLER Then
            HatFehler = True
            Exit Function
        End If
    Next

    HatFehler = False

End Function


Function DatumsWert(wert As Variant) As Variant
    'Versucht, den Wert wert in einen Datumswert umzuwandeln. Das ist unabhängig vom Datumsformat des Systems
    'Wenn das nicht klappt, wird Empty zurückgegeben
    'Ansonsten wird ein Wert vom Typ Date zurückgegeben.
    Dim wertS As String
    Dim V As Variant

    wertS = CStr(wert)
    If (wertS Like "##.##.####") Or (wertS Like "##.##.##") Then    'dd.mm.yyyy oder dd.mm.yy im 21. Jh.
        V = Split(wertS, ".")
        If Len(V(2)) = 2 Then
            V(2) = "20" + V(2)
        End If
        DatumsWert = DateSerial(V(2), V(1), V(0))
        Exit Function
    End If
    If (wertS Like "####-##-##") Or (wertS Like "##-##-##") Then    'yyyy-mm-dd oder yy-mm-dd im 21. Jh.
        V = Split(wertS, "-")
        If Len(V(0)) = 2 Then
            V(0) = "20" + V(0)
        End If
        DatumsWert = DateSerial(V(0), V(1), V(2))
        Exit Function
    End If

    DatumsWert = Empty

End Function


Function MeinSystemFormat(wert As Variant) As String
    'Wandelt wert (deutsches Zahlenformat) in einen String um mit zum System passenden Dezimal- bzw. Tausendertrenner
    'Damit funktioniert die Umwandlung in Double (über Cdbl) dann auf jedem System korrekt

    Dim dezi As String
    Dim tsd As String
    Dim wertS As String
    Dim V As Variant
    Dim K As Integer

    MeinSystemFormat = ""

    wertS = CStr(wert)
    If wertS = "" Then
        Exit Function
    End If

    dezi = Application.DecimalSeparator
    tsd = Application.ThousandsSeparator

    V = Split(wertS, ",")
    For K = 0 To UBound(V)
        V(K) = Replace(V(K), ".", tsd)
    Next

    MeinSystemFormat = Join(V, dezi)

End Function


Function MeinProzentWert(wert As Variant) As Variant
    'Rechnet die Prozentangabe von OnVista in eine Zahl um, beachtet systemspezifische Trenner

    Dim inhalt As String

    MeinProzentWert = ""
    inhalt = MeinSystemFormat(wert)
    If InStr(inhalt, "%") > 1 Then
        inhalt = Replace(inhalt, "%", "")
        If IsNumeric(inhalt) Then
            MeinProzentWert = CDbl(inhalt) / 100
        End If
    Else
        If IsNumeric(inhalt) Then
            MeinProzentWert = CDbl(inhalt)
        End If
    End If

End Function
