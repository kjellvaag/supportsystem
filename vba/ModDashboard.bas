Attribute VB_Name = "ModDashboard"
Option Explicit

' ============================================================================
' ModDashboard — oppdaterer dashboard-hjelpetabellene og diagrammene.
' Standardmodul.
'
' Hjelpetabeller (kontrakt):
'   DASH_STATUS_RANGE    = A27:B32   (status   | antall)
'   DASH_KATEGORI_RANGE  = A35:B44   (kategori | antall, topp-10 sortert synkende)
'   DASH_ANSVARLIG_RANGE = C27:D36   (ansvarlig | antall, topp-10 sortert synkende)
'
' Kandidatlistene hentes fra Oppsett (Kategorier = E4:E50, Ansatte = A4:A50) slik
' at topplisten følger Oppsett automatisk og utelater «– ikke tildelt –».
'
' Diagrammer (opprettelsesrekkefølge i generatoren):
'   1 = kake    "Saker per status"     — verdier B27:B32, kategorier A27:A32
'   2 = stolpe  "Topp saker pr kategori"  — verdier B35:B44, kategorier A35:A44
'   3 = stolpe  "Topp saker pr ansvarlig" — verdier D27:D36, kategorier C27:C36
'
' Diagrammene flyttes ALDRI — kun datakildereferansene bygges på nytt (og
' eventuelt krympes til faktisk antall rader). Høyde holdes uendret for å
' unngå å ødelegge oppsettet.
' ============================================================================

Private Const DASH_TOPP_N As Long = 10

Public Sub OppdaterDashboard()
    ' Saker-kolonner (1-basert; kontrakt.SAKER_HEADERS):
    Const KOL_KATEGORI As Long = 6    ' F
    Const KOL_ANSVARLIG As Long = 12  ' L
    Const KOL_STATUS As Long = 13     ' M
    Const MAX_S As Long = 501         ' kontrakt.MAX_S

    Dim arr As Variant, katArr As Variant, ansArr As Variant
    Dim r As Long, nKat As Long, nAns As Long
    Dim sistStatus As Long, sistKategori As Long, sistAnsvarlig As Long

    On Error GoTo Feil
    Application.EnableEvents = False

    ' Last hele Saker-området og kandidatlistene på én gang (raskt)
    arr = Sheet2.Range("A2:W" & MAX_S).Value2
    katArr = Sheet6.Range("E4:E50").Value2
    ansArr = Sheet6.Range("A4:A50").Value2

    ' --- Saker per status (A27:A32 / B27:B32) -------------------------------
    Sheet1.Range("B27:B32").ClearContents
    sistStatus = 27
    For r = 27 To 32
        If Trim$(CStr(Sheet1.Cells(r, 1).Value)) <> "" Then
            Sheet1.Cells(r, 2).Value = TellSaker(arr, KOL_STATUS, CStr(Sheet1.Cells(r, 1).Value), MAX_S - 1)
            sistStatus = r
        End If
    Next r

    ' --- Topp saker pr kategori (A34 overskrift, A35:B44 data) ---------------
    Sheet1.Range("A35:B44").ClearContents
    nKat = SkrivToppListe(arr, katArr, KOL_KATEGORI, 35, 1)
    Sheet1.Cells(34, 1).Value = "Topp " & nKat & " saker pr kategori"
    sistKategori = 35 + IIf(nKat > 1, nKat - 1, 0)

    ' --- Topp saker pr ansvarlig (C26 overskrift, C27:D36 data) --------------
    Sheet1.Range("C27:D36").ClearContents
    nAns = SkrivToppListe(arr, ansArr, KOL_ANSVARLIG, 27, 3)
    Sheet1.Cells(26, 3).Value = "Topp " & nAns & " saker pr ansvarlig"
    sistAnsvarlig = 27 + IIf(nAns > 1, nAns - 1, 0)

    ' --- Bygg om diagrammenes datakilder (IKKE flytt diagrammene) -----------
    OppdaterDiagram Sheet1.ChartObjects(1).Chart, 27, sistStatus, 2, 1    ' kake: B-verdier, A-kategorier
    OppdaterDiagram Sheet1.ChartObjects(2).Chart, 35, sistKategori, 2, 1  ' stolpe kategori
    OppdaterDiagram Sheet1.ChartObjects(3).Chart, 27, sistAnsvarlig, 4, 3 ' stolpe ansvarlig

Ut:
    Application.EnableEvents = True
    Exit Sub

Feil:
    Application.EnableEvents = True
    Resume Ut
End Sub


' Tell antall saker der kolonnen `kolonne` er lik `verdi`.
' Store/små bokstaver ignoreres (vbTextCompare) — samme oppførsel som COUNTIFS.
Private Function TellSaker(arr As Variant, kolonne As Long, verdi As String, antRader As Long) As Long
    Dim i As Long
    Dim n As Long
    n = 0
    For i = 1 To antRader
        If StrComp(CStr(arr(i, kolonne)), verdi, vbTextCompare) = 0 Then n = n + 1
    Next i
    TellSaker = n
End Function


' Teller forekomster av kolonnen `kol` per kandidat i `kandidater` (2D-kolonne),
' sorterer synkende etter antall (stabil innsettingssortering) og skriver inntil
' DASH_TOPP_N rader (navn + antall) fra `startRad` i kolonnene `navnKol`/`navnKol+1`.
' Returnerer antall rader som ble skrevet (0..DASH_TOPP_N).
Private Function SkrivToppListe(arr As Variant, kandidater As Variant, _
                                kol As Long, startRad As Long, navnKol As Long) As Long
    Dim antKand As Long, antRader As Long
    antKand = UBound(kandidater, 1)
    antRader = UBound(arr, 1)

    Dim navn() As String
    Dim antall() As Long
    ReDim navn(1 To antKand)
    ReDim antall(1 To antKand)

    Dim i As Long, j As Long, n As Long, cnt As Long
    Dim etikett As String
    n = 0
    For i = 1 To antKand
        etikett = Trim$(CStr(kandidater(i, 1)))
        If etikett <> "" Then
            cnt = TellSaker(arr, kol, etikett, antRader)
            If cnt > 0 Then
                n = n + 1
                navn(n) = etikett
                antall(n) = cnt
            End If
        End If
    Next i

    ' Stabil innsettingssortering: synkende etter antall, lik antall beholder rekkefølge.
    ' NB: VBA kortslutter IKKE `And`, så j>=1 sjekkes i Do While før antall(j) leses.
    Dim tmpN As String, tmpA As Long
    For i = 2 To n
        tmpN = navn(i): tmpA = antall(i)
        j = i - 1
        Do While j >= 1
            If antall(j) < tmpA Then
                navn(j + 1) = navn(j)
                antall(j + 1) = antall(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        navn(j + 1) = tmpN
        antall(j + 1) = tmpA
    Next i

    ' Skriv inntil DASH_TOPP_N rader.
    Dim skrevet As Long
    skrevet = 0
    For i = 1 To n
        If skrevet >= DASH_TOPP_N Then Exit For
        skrevet = skrevet + 1
        Sheet1.Cells(startRad + skrevet - 1, navnKol).Value = navn(i)
        Sheet1.Cells(startRad + skrevet - 1, navnKol + 1).Value = antall(i)
    Next i

    SkrivToppListe = skrevet
End Function


' Setter diagrammets verdier (kolonne `verdiKol`) og kategorier (kolonne
' `katKol`) til radene `startRad`..`sluttRad`. Bruker SeriesCollection(1)
' slik at eksisterende formatering (farger, tegnforklaring, akser) beholdes.
Private Sub OppdaterDiagram(ch As Chart, startRad As Long, sluttRad As Long, _
                            verdiKol As Long, katKol As Long)
    If ch.SeriesCollection.Count < 1 Then Exit Sub
    With ch.SeriesCollection(1)
        .Values = Sheet1.Range(Sheet1.Cells(startRad, verdiKol), Sheet1.Cells(sluttRad, verdiKol))
        .XValues = Sheet1.Range(Sheet1.Cells(startRad, katKol), Sheet1.Cells(sluttRad, katKol))
    End With
End Sub


' ============================================================================
' AKTIVERING — legges i ThisWorkbook-modulen (ikke her):
'
'   Private Sub Workbook_SheetActivate(ByVal Sh As Object)
'       If Sh.Name = "Dashboard" Then ModDashboard.OppdaterDashboard
'   End Sub
'
' Alternativt kan OppdaterDashboard kalles manuelt (f.eks. via en knapp).
' ============================================================================
