' ============================================================================
' ModDashboard — oppdaterer dashboard-hjelpetabellene.
' Standardmodul (StarBasic for LibreOffice Calc / Apache OpenOffice Calc).
'
' Hjelpetabeller (kontrakt):
'   DASH_STATUS_RANGE    = A27:B32   (status   | antall)
'   DASH_KATEGORI_RANGE  = A35:B44   (kategori | antall, topp-10 sortert synkende)
'   DASH_ANSVARLIG_RANGE = C27:D36   (ansvarlig | antall, topp-10 sortert synkende)
'
' Kandidatlistene hentes fra Oppsett (Kategorier = E4:E50, Ansatte = A4:A50) slik
' at topplisten følger Oppsett automatisk og utelater «– ikke tildelt –».
'
' Diagrammene i .ods-en er "levende" Calc-diagrammer som peker på disse
' cellene — når vi skriver nye tall/navn oppdateres diagrammene automatisk.
' Derfor trenger vi IKKE å flytte eller bygge om diagramobjektene (det er også
' en hard begrensning fra AGENTS.md).
'
' Aktivering: knapp-triggeret (LO har ingen Workbook_SheetActivate som i VBA),
' eller via dokumenthendelsen "Ark aktivert". OppdaterDashboard settes som
' makro på en knapp på Dashboard-arket.
' ============================================================================

Sub OppdaterDashboard()
    ' Saker-kolonner (0-basert; kontrakt.SAKER_HEADERS):
    Const KOL_KATEGORI As Long = 5    ' F
    Const KOL_ANSVARLIG As Long = 11  ' L
    Const KOL_STATUS As Long = 12     ' M
    Const MAX_S As Long = 501         ' kontrakt.MAX_S

    Dim oDoc As Object
    Dim oDash As Object, oSaker As Object, oOppsett As Object
    Dim arr As Variant, katArr As Variant, ansArr As Variant
    Dim i As Long, antRader As Long
    Dim etikett As String
    Dim nKat As Long, nAns As Long

    oDoc = ThisComponent
    oDash = oDoc.Sheets.getByName("Dashboard")
    oSaker = oDoc.Sheets.getByName("Saker")
    oOppsett = oDoc.Sheets.getByName("Oppsett")

    ' Last hele Saker-området og kandidatlistene på én gang (raskt)
    arr = oSaker.getCellRangeByName("A2:W" & MAX_S).getDataArray()
    antRader = UBound(arr, 1)         ' 0..499
    katArr = oOppsett.getCellRangeByName("E4:E50").getDataArray()
    ansArr = oOppsett.getCellRangeByName("A4:A50").getDataArray()

    ' --- Saker per status (A27:A32 / B27:B32) -------------------------------
    oDash.getCellRangeByName("B27:B32").clearContents(7)
    For i = 26 To 31                  ' 0-basert rad 26..31 = rad 27..32
        etikett = Trim(oDash.getCellByPosition(0, i).getString())
        If etikett <> "" Then
            oDash.getCellByPosition(1, i).setValue(TellSaker(arr, KOL_STATUS, etikett, antRader))
        End If
    Next i

    ' --- Topp saker pr kategori (A34 overskrift, A35:B44 data) ---------------
    oDash.getCellRangeByName("A35:B44").clearContents(7)
    nKat = SkrivToppListe(oDash, arr, katArr, KOL_KATEGORI, 34, 0)
    oDash.getCellByPosition(0, 33).setString("Topp " & nKat & " saker pr kategori")

    ' --- Topp saker pr ansvarlig (C26 overskrift, C27:D36 data) --------------
    oDash.getCellRangeByName("C27:D36").clearContents(7)
    nAns = SkrivToppListe(oDash, arr, ansArr, KOL_ANSVARLIG, 26, 2)
    oDash.getCellByPosition(2, 25).setString("Topp " & nAns & " saker pr ansvarlig")

    ' Diagrammene leser direkte fra hjelpetabellene ovenfor og oppdateres
    ' automatisk når cellene endres — ingen UNO-diagram-API nødvendig.
End Sub


' Tell antall saker der kolonnen `kol0` er lik `verdi`.
' Store/små bokstaver ignoreres (StrComp med compare=1) — samme oppførsel
' som COUNTIFS.
Function TellSaker(arr As Variant, kol0 As Long, verdi As String, antRader As Long) As Long
    Dim i As Long
    Dim n As Long
    n = 0
    For i = 0 To antRader
        If StrComp(CStr(arr(i, kol0)), verdi, 1) = 0 Then n = n + 1
    Next i
    TellSaker = n
End Function


' Teller forekomster av kolonnen `kol0` per kandidat i `kandidater` (2D-kolonne),
' sorterer synkende etter antall (stabil innsettingssortering) og skriver inntil
' 10 rader (navn + antall) fra `startRad0` i kolonnene `navnKol0`/`navnKol0+1`.
' Alle posisjoner er 0-baserte. Returnerer antall rader som ble skrevet (0..10).
Function SkrivToppListe(oDash As Object, arr As Variant, kandidater As Variant, _
                        kol0 As Long, startRad0 As Long, navnKol0 As Long) As Long
    Dim antKand As Long, antRader As Long
    antKand = UBound(kandidater, 1)
    antRader = UBound(arr, 1)

    Dim navn() As String
    Dim antall() As Long
    ReDim navn(0 To antKand)
    ReDim antall(0 To antKand)

    Dim i As Long, j As Long, n As Long, cnt As Long
    Dim etikett As String
    n = 0
    For i = 0 To antKand
        etikett = Trim(CStr(kandidater(i, 0)))
        If etikett <> "" Then
            cnt = TellSaker(arr, kol0, etikett, antRader)
            If cnt > 0 Then
                navn(n) = etikett
                antall(n) = cnt
                n = n + 1
            End If
        End If
    Next i

    ' Stabil innsettingssortering: synkende etter antall, lik antall beholder rekkefølge.
    ' NB: Basic kortslutter IKKE `And`, så j>=0 sjekkes i Do While før antall(j) leses.
    Dim tmpN As String, tmpA As Long
    For i = 1 To n - 1
        tmpN = navn(i): tmpA = antall(i)
        j = i - 1
        Do While j >= 0
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

    ' Skriv inntil 10 rader.
    Dim skrevet As Long
    skrevet = 0
    For i = 0 To n - 1
        If skrevet >= 10 Then Exit For
        oDash.getCellByPosition(navnKol0, startRad0 + skrevet).setString(navn(i))
        oDash.getCellByPosition(navnKol0 + 1, startRad0 + skrevet).setValue(antall(i))
        skrevet = skrevet + 1
    Next i
    SkrivToppListe = skrevet
End Function
