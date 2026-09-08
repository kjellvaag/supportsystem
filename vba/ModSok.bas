Attribute VB_Name = "ModSok"
Option Explicit

' ============================================================================
' ModSok – sanntidssøk i kunnskapsbasen og tidligere saker.
' Standardmodul. Kalles fra arkmodulene til "Søk KB" (Sheet3) og "Søk saker"
' (Sheet4) via HåndterSøkKBEndring / HåndterSøkSakerEndring (debouce via
' Application.OnTime).
'
' Ark-kodenavn (kontrakt.CODE_NAMES):
'   Sheet1 = Dashboard, Sheet2 = Saker, Sheet3 = Søk KB,
'   Sheet4 = Søk saker, Sheet5 = Kunnskapsbase.
'
' Scoring delegeres til ModFelles.ScoreDok (matcher kontrakt.score).
' Søket erstatter INDEX/MATCH-hjelpekolonnene (Saker!Y, Kunnskapsbase!L),
' som IKKE brukes i makroversjonen.
' ============================================================================

' ----- Debounce (modulnivå-tidspunkter) -------------------------------------
Private kbTimer As Date
Private sakTimer As Date
Private Const DEBOUNCE_DAGER As Double = 300# / 86400000#   ' 300 ms

' ----- Høydepunktfarge for treff (lys gul, samme som FFEB9C) ----------------
Private Const FARG_TREFF As Long = 10284031   ' RGB(255, 235, 156)

' ----- Kapasiteter (kontrakt) ----------------------------------------------
Private Const MAX_K As Long = 301   ' datarader i Kunnskapsbase 2..301
Private Const MAX_S As Long = 501   ' datarader i Saker 2..501
Private Const N_KB As Long = 25     ' maks KB-treff som vises
Private Const N_SAK As Long = 30    ' maks sakstreff som vises

' ----- Treff-kandidat (for rangering) ---------------------------------------
Private Type Treff
    Score As Long   ' score fra ScoreDok (høyere = bedre)
    Dato As Double  ' dato for tiebreak (nyere = bedre; 0 = ukjent)
    Rad As Long     ' indeks inn i data-arrayet (lavere rad = først ved likt)
End Type


' ============================================================================
' HJELEPERE
' ============================================================================

' Konverterer celleverdi til tekst; tom/feil/Null -> "".
Private Function SomTekst(ByVal v As Variant) As String
    If IsNull(v) Or IsEmpty(v) Or IsError(v) Then
        SomTekst = ""
    Else
        SomTekst = CStr(v)
    End If
End Function

' Konverterer celleverdi til datotall (for tiebreak); 0 = ukjent/ikke dato.
Private Function DatoSomTall(ByVal v As Variant) As Double
    If IsNumeric(v) Then
        DatoSomTall = CDbl(v)
    ElseIf IsDate(v) Then
        DatoSomTall = CDbl(CDate(v))
    Else
        DatoSomTall = 0
    End If
End Function

' Returnerer True hvis a skal sorteres FØR b (score høyest, deretter nyest
' dato, deretter lavest rad).
Private Function ErBedre(ByVal a As Treff, ByVal b As Treff) As Boolean
    If a.Score <> b.Score Then
        ErBedre = (a.Score > b.Score)
    ElseIf a.Dato <> b.Dato Then
        ErBedre = (a.Dato > b.Dato)
    Else
        ErBedre = (a.Rad < b.Rad)
    End If
End Function

' Stabil innstikk-sortering synkende (best først).
Private Sub SorterTreff(ByRef treff() As Treff)
    Dim i As Long, j As Long
    Dim nokkel As Treff

    If UBound(treff) <= LBound(treff) Then Exit Sub

    For i = LBound(treff) + 1 To UBound(treff)
        nokkel = treff(i)
        j = i - 1
        Do While j >= LBound(treff)
            If ErBedre(nokkel, treff(j)) Then
                treff(j + 1) = treff(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        treff(j + 1) = nokkel
    Next i
End Sub

' Returnerer True hvis status finnes i lista over tillatte statuser.
Private Function ErTillatt(ByVal status As String, ByRef tillatt() As String, ByVal ant As Long) As Boolean
    Dim m As Long
    ErTillatt = False
    For m = 1 To ant
        If status = tillatt(m) Then
            ErTillatt = True
            Exit Function
        End If
    Next m
End Function

' Høydepunkt (helcelle-markering): lys gul fyll + fet på tittelcellen.
Private Sub MarkerTreff(ws As Worksheet, rad As Long, kol As Long)
    With ws.Cells(rad, kol)
        .Interior.Color = FARG_TREFF
        .Font.Bold = True
    End With
End Sub


' ============================================================================
' SØK KB
' ============================================================================
Public Sub SøkKB()
    ' Kildekolonner i Kunnskapsbase (1-basert; kontrakt.KB_SEARCH_FIELDS):
    Const KOL_TITTEL As Long = 2     ' B
    Const KOL_KATEGORI As Long = 3   ' C (filter)
    Const KOL_PROBLEM As Long = 4    ' D
    Const KOL_LOSNING As Long = 5    ' E
    Const KOL_NOKKEL As Long = 6     ' F
    Const KOL_DATO As Long = 8       ' H (tiebreak)
    ' Resultatområde (kontrakt): header rad 8, resultatrader 9..33, kol A..K

    Dim arr As Variant
    Dim treff() As Treff
    Dim sok As String
    Dim kat As String
    Dim antall As Long
    Dim vist As Long
    Dim i As Long, j As Long, c As Long
    Dim score As Long
    Dim datoRå As Variant

    On Error GoTo Feil
    Application.EnableEvents = False

    sok = Normaliser(CStr(Sheet3.Range("B4").Value))
    kat = Trim$(CStr(Sheet3.Range("B5").Value))

    ' Tøm resultatområdet uansett (også ved tomt søk / ren modus)
    Sheet3.Range("A9:K33").Clear

    If sok = "" Then
        Sheet3.Range("A7").Value = "Skriv søkeord ovenfor og trykk «SØK I KB " & ChrW(8594) & "»."
        GoTo Ut
    End If

    ' Last hele dataområdet på én gang (raskt – unngå celle-for-celle)
    arr = Sheet5.Range("A2:K" & MAX_K).Value2

    antall = 0
    ReDim treff(1 To 1)
    For i = 1 To MAX_K - 1
        If SomTekst(arr(i, KOL_TITTEL)) <> "" Then
            If kat = "" Or SomTekst(arr(i, KOL_KATEGORI)) = kat Then
                score = ScoreDok(sok, SomTekst(arr(i, KOL_TITTEL)), _
                                 SomTekst(arr(i, KOL_NOKKEL)), _
                                 SomTekst(arr(i, KOL_PROBLEM)), _
                                 SomTekst(arr(i, KOL_LOSNING)))
                If score > 0 Then
                    datoRå = arr(i, KOL_DATO)
                    antall = antall + 1
                    ReDim Preserve treff(1 To antall)
                    treff(antall).Score = score
                    treff(antall).Dato = DatoSomTall(datoRå)
                    treff(antall).Rad = i
                End If
            End If
        End If
    Next i

    If antall > 0 Then SorterTreff treff

    ' Skriv inntil N_KB treff (kolonner A..K = kildekolonner 1..11)
    vist = 0
    For j = 1 To antall
        If vist >= N_KB Then Exit For
        For c = 1 To 11
            Sheet3.Cells(9 + vist, c).Value = arr(treff(j).Rad, c)
        Next c
        Sheet3.Cells(9 + vist, 8).NumberFormat = "DD.MM.YYYY"
        MarkerTreff Sheet3, 9 + vist, 2   ' tittel (B)
        vist = vist + 1
    Next j

    If antall = 0 Then
        Sheet3.Range("A7").Value = "Ingen treff – prøv andre søkeord."
    Else
        Sheet3.Range("A7").Value = antall & " treff i kunnskapsbasen:"
    End If

Ut:
    Application.EnableEvents = True
    Exit Sub

Feil:
    Application.EnableEvents = True
    Resume Ut
End Sub


' ============================================================================
' SØK SAKER
' ============================================================================
Public Sub SøkSaker()
    ' Kildekolonner i Saker (1-basert; kontrakt.SAKER_SEARCH_FIELDS):
    Const KOL_SAKSNR As Long = 1     ' A
    Const KOL_OPPRETTET As Long = 2  ' B (tiebreak)
    Const KOL_KATEGORI As Long = 6   ' F
    Const KOL_TITTEL As Long = 8     ' H
    Const KOL_BESKRIVELSE As Long = 9 ' I (problem)
    Const KOL_STATUS As Long = 13    ' M (filter mot avkryssingsboksene)
    Const KOL_LOSNING As Long = 21   ' U
    Const KOL_KBREF As Long = 22     ' V
    ' Resultatområde (kontrakt): header rad 7, resultatrader 8..37,
    ' kolonner A,B,F,H,U,V (kontrakt.SOK_SAK_COLS).

    Dim arr As Variant
    Dim statusArr As Variant
    Dim tillatt() As String
    Dim treff() As Treff
    Dim sok As String
    Dim antTillatt As Long
    Dim antall As Long
    Dim vist As Long
    Dim i As Long, j As Long
    Dim score As Long
    Dim datoRå As Variant

    On Error GoTo Feil
    Application.EnableEvents = False

    sok = Normaliser(CStr(Sheet4.Range("B4").Value))

    ' Tøm resultatområdet (kolonner A,B,F,H,U,V) uansett
    Sheet4.Range("A8:A37,B8:B37,F8:F37,H8:H37,U8:U37,V8:V37").Clear

    ' Les status-filteret (skjult speil-tabell AA5:AB14; AA = statusnavn,
    ' AB = SANN/USANN fra avkryssingsboksene). kontrakt.SOK_SAK_STATUS.
    statusArr = Sheet4.Range("AA5:AB14").Value2
    antTillatt = 0
    ReDim tillatt(1 To 10)
    For i = 1 To 10
        If SomTekst(statusArr(i, 1)) <> "" Then
            If statusArr(i, 2) = True Then
                antTillatt = antTillatt + 1
                tillatt(antTillatt) = SomTekst(statusArr(i, 1))
            End If
        End If
    Next i

    If sok = "" Then
        Sheet4.Range("A6").Value = "Skriv søkeord ovenfor og trykk «SØK I SAKER " & ChrW(8594) & "»."
        GoTo Ut
    End If

    ' Last hele dataområdet på én gang
    arr = Sheet2.Range("A2:W" & MAX_S).Value2

    antall = 0
    ReDim treff(1 To 1)
    For i = 1 To MAX_S - 1
        If SomTekst(arr(i, KOL_TITTEL)) <> "" Then
            ' Status-filter: ta kun med saker i avkryssede statuser
            If ErTillatt(SomTekst(arr(i, KOL_STATUS)), tillatt, antTillatt) Then
                score = ScoreDok(sok, SomTekst(arr(i, KOL_TITTEL)), "", _
                                 SomTekst(arr(i, KOL_BESKRIVELSE)), _
                                 SomTekst(arr(i, KOL_LOSNING)))
                If score > 0 Then
                    datoRå = arr(i, KOL_OPPRETTET)
                    antall = antall + 1
                    ReDim Preserve treff(1 To antall)
                    treff(antall).Score = score
                    treff(antall).Dato = DatoSomTall(datoRå)
                    treff(antall).Rad = i
                End If
            End If
        End If
    Next i

    If antall > 0 Then SorterTreff treff

    ' Skriv inntil N_SAK treff (kolonner A,B,F,H,U,V)
    vist = 0
    For j = 1 To antall
        If vist >= N_SAK Then Exit For
        Sheet4.Cells(8 + vist, 1).Value = arr(treff(j).Rad, KOL_SAKSNR)
        Sheet4.Cells(8 + vist, 2).Value = arr(treff(j).Rad, KOL_OPPRETTET)
        Sheet4.Cells(8 + vist, 6).Value = arr(treff(j).Rad, KOL_KATEGORI)
        Sheet4.Cells(8 + vist, 8).Value = arr(treff(j).Rad, KOL_TITTEL)
        Sheet4.Cells(8 + vist, 21).Value = arr(treff(j).Rad, KOL_LOSNING)
        Sheet4.Cells(8 + vist, 22).Value = arr(treff(j).Rad, KOL_KBREF)
        Sheet4.Cells(8 + vist, 2).NumberFormat = "DD.MM.YYYY HH:MM"
        MarkerTreff Sheet4, 8 + vist, 8   ' tittel (H)
        vist = vist + 1
    Next j

    If antall = 0 Then
        Sheet4.Range("A6").Value = "Ingen saker med valgte statuser matcher søket."
    Else
        Sheet4.Range("A6").Value = antall & " treff:"
    End If

Ut:
    Application.EnableEvents = True
    Exit Sub

Feil:
    Application.EnableEvents = True
    Resume Ut
End Sub


' ============================================================================
' DEBOUNCE (Application.OnTime, 300 ms)
' ============================================================================

Public Sub PlanleggSøkKB()
    ' Avbryt ev. tidligere planlagt kjøring, og planlegg ny om 300 ms.
    On Error Resume Next
    If kbTimer <> 0 Then
        Application.OnTime EarliestTime:=kbTimer, Procedure:="ModSok.SøkKB", Schedule:=False
    End If
    On Error GoTo 0
    kbTimer = Now + DEBOUNCE_DAGER
    Application.OnTime EarliestTime:=kbTimer, Procedure:="ModSok.SøkKB", Schedule:=True
End Sub

Public Sub PlanleggSøkSaker()
    On Error Resume Next
    If sakTimer <> 0 Then
        Application.OnTime EarliestTime:=sakTimer, Procedure:="ModSok.SøkSaker", Schedule:=False
    End If
    On Error GoTo 0
    sakTimer = Now + DEBOUNCE_DAGER
    Application.OnTime EarliestTime:=sakTimer, Procedure:="ModSok.SøkSaker", Schedule:=True
End Sub


' ============================================================================
' INNGANGSPUNKTER FRA ARKMODULENE
' ============================================================================

' Kalles fra Worksheet_Change i "Søk KB" (Sheet3). Utløser søk når søkeord
' (B4) eller kategori (B5) endres.
Public Sub HåndterSøkKBEndring(ByVal Target As Range)
    If Not Intersect(Target, Target.Worksheet.Range("B4:B5")) Is Nothing Then
        PlanleggSøkKB
    End If
End Sub

' Kalles fra Worksheet_Change i "Søk saker" (Sheet4). Utløser søk når
' søkeord (B4) eller status-avkryssingene (AB5:AB14) endres.
Public Sub HåndterSøkSakerEndring(ByVal Target As Range)
    If Not Intersect(Target, Target.Worksheet.Range("B4")) Is Nothing Then
        PlanleggSøkSaker
    ElseIf Not Intersect(Target, Target.Worksheet.Range("AB5:AB14")) Is Nothing Then
        PlanleggSøkSaker
    End If
End Sub


' ============================================================================
' HENDELSESBEHANDLERE – legges i ARKMODULENE (ikke i denne standardmodulen)
' ----------------------------------------------------------------------------
' Worksheet_Change kan IKKE ligge i en standardmodul – derfor delegeres den
' via de offentlige Håndter*-subene ovenfor.
'
' 1) I ARKMODULEN for "Søk KB" (Sheet3):
'
'     Private Sub Worksheet_Change(ByVal Target As Range)
'         ModSok.HåndterSøkKBEndring Target
'     End Sub
'
' 2) I ARKMODULEN for "Søk saker" (Sheet4):
'
'     Private Sub Worksheet_Change(ByVal Target As Range)
'         ModSok.HåndterSøkSakerEndring Target
'     End Sub
' ============================================================================
