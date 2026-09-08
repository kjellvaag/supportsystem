' ============================================================================
' ModSok — søk i kunnskapsbasen og tidligere saker (knapp-triggeret).
' Standardmodul (StarBasic for LibreOffice Calc / Apache OpenOffice Calc).
'
' LO/OO har ingen Worksheet_Change-hendelse, så søket er KNAPP-triggeret:
'   - SokKB     → knyttes til "SØK I KB →" (Søk KB, celle D4) via Calc-UI.
'   - SokSaker  → knyttes til "SØK I SAKER →" (Søk saker, celle D4) via Calc-UI.
' Knappetilordning gjøres i Calc (Skjema > Trykknapp, knytt hendelsen "Utfør
' handling" til makroen) — se basic/README.md. Se også "BEST-EFFORT LISTENER"
' nederst i filen for sanntidsøk-alternativet.
'
' Scoring delegeres til ModFelles.ScoreDok (matcher kontrakt.score). Søket
' erstatter INDEX/MATCH-hjelpekolonnene (Saker!Y, Kunnskapsbase!L), som IKKE
' brukes i makroversjonen.
'
' All lesing skjer via getDataArray() (én bulk-lesing — mye raskere enn
' celle-for-celle, som er tregt i LO Basic).
'
' Indeksering: getDataArray() og getCellByPosition() er 0-BASERT (kolonne OG
' rad). Alle KOL_*-konstantene under er derfor 0-baserte.
' ============================================================================

' ----- Høydepunktfarge for treff (lys gul = FFEB9C, samme som arbeidsbokens
'      YELLOW_FILL). 0xFFEB9C = 16771996. UNO CellBackColor bruker 0xRRGGBB,
'      og StarBasic RGB(red,green,blue) = red*65536+green*256+blue = 0xRRGGBB,
'      dvs. UNO-formatet direkte (motsatt av VBA, der RGB er 0xBBGGRR).
Const FARG_TREFF As Long = 16771996   ' = RGB(255, 235, 156) = FFEB9C

' ----- Kapasiteter (kontrakt) ----------------------------------------------
Const MAX_K As Long = 301   ' datarader i Kunnskapsbase 2..301
Const MAX_S As Long = 501   ' datarader i Saker 2..501
Const N_KB As Long = 25     ' maks KB-treff som vises
Const N_SAK As Long = 30    ' maks sakstreff som vises

' ----- Treff-kandidat (for rangering) ---------------------------------------
Type Treff
    Score As Long   ' score fra ScoreDok (høyere = bedre)
    Dato As Double  ' dato for tiebreak (nyere = bedre; 0 = ukjent)
    Rad As Long     ' radindeks inn i data-arrayet (lavere rad = først ved likt)
End Type


' ============================================================================
' HJELEPERE
' ============================================================================

' Konverterer celleverdi til tekst; tom/feil/Null -> "".
Function SomTekst(v As Variant) As String
    If IsNull(v) Or IsEmpty(v) Or IsError(v) Then
        SomTekst = ""
    Else
        SomTekst = CStr(v)
    End If
End Function

' Konverterer celleverdi til datotall (for tiebreak); 0 = ukjent/ikke dato.
' Datoer kommer fra getDataArray() som Double (internt serienummer).
Function DatoSomTall(v As Variant) As Double
    DatoSomTall = 0
    If IsNull(v) Or IsEmpty(v) Or IsError(v) Then Exit Function
    If IsNumeric(v) Then
        DatoSomTall = CDbl(v)
    ElseIf IsDate(v) Then
        On Error Resume Next
        DatoSomTall = CDbl(CDate(v))
        On Error GoTo 0
    End If
End Function

' Returnerer True hvis a skal sorteres FØR b (score høyest, deretter nyest
' dato, deretter lavest rad).
Function ErBedre(a As Treff, b As Treff) As Boolean
    If a.Score <> b.Score Then
        ErBedre = (a.Score > b.Score)
    ElseIf a.Dato <> b.Dato Then
        ErBedre = (a.Dato > b.Dato)
    Else
        ErBedre = (a.Rad < b.Rad)
    End If
End Function

' Stabil innstikk-sortering synkende (best først).
Sub SorterTreff(treff() As Treff)
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
Function ErTillatt(status As String, tillatt() As String, ant As Long) As Boolean
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
' rad0/kol0 er 0-baserte.
Sub MarkerTreff(oSheet As Object, rad0 As Long, kol0 As Long)
    Dim oCell As Object
    oCell = oSheet.getCellByPosition(kol0, rad0)
    oCell.CellBackColor = FARG_TREFF
    oCell.CharWeight = 150   ' com.sun.star.awt.FontWeight.BOLD
End Sub

' Skriver en vilkårlig celleverdi: tom -> tømmes, tall/dato -> setValue,
' annet -> setString.
Sub SkrivCelle(oCell As Object, v As Variant)
    If IsEmpty(v) Or IsNull(v) Then
        oCell.clearContents(7)
    ElseIf IsNumeric(v) Or IsDate(v) Then
        oCell.setValue(CDbl(v))
    Else
        oCell.setString(CStr(v))
    End If
End Sub

' Kopierer nummerformatet fra en referansecelle (som allerede har riktig
' format fra generatoren) til målcellen — slipper å lage Locale-struktur og
' slå opp i NumberFormats.
Sub KopierDatoformat(oMalCell As Object, oRefCell As Object)
    On Error Resume Next
    oMalCell.NumberFormat = oRefCell.NumberFormat
    On Error GoTo 0
End Sub


' ============================================================================
' SØK KB
' ============================================================================
Sub SokKB()
    ' 0-baserte kildekolonner i Kunnskapsbase (kontrakt.KB_SEARCH_FIELDS):
    Const KOL_TITTEL As Long = 1     ' B
    Const KOL_KATEGORI As Long = 2   ' C (filter)
    Const KOL_PROBLEM As Long = 3    ' D
    Const KOL_LOSNING As Long = 4    ' E
    Const KOL_NOKKEL As Long = 5     ' F
    Const KOL_DATO As Long = 7       ' H (tiebreak)
    ' Resultatområde (kontrakt): header rad 8, resultatrader 9..33, kol A..K

    Dim oDoc As Object
    Dim oSoek As Object, oKB As Object
    Dim arr As Variant
    Dim treff() As Treff
    Dim ut() As Variant
    Dim sok As String, kat As String
    Dim antall As Long, vist As Long
    Dim i As Long, j As Long, c As Long
    Dim score As Long
    Dim rad As Long

    oDoc = ThisComponent
    oSoek = oDoc.Sheets.getByName("Søk KB")
    oKB = oDoc.Sheets.getByName("Kunnskapsbase")

    ' Inndata: søkeord i B4, valgfri kategori i B5 (B4 = col 1, rad 3 i 0-basert)
    sok = Normaliser(oSoek.getCellByPosition(1, 3).getString())
    kat = Trim(oSoek.getCellByPosition(1, 4).getString())

    ' Tøm resultatområdet uansett (også ved tomt søk / ren modus)
    oSoek.getCellRangeByName("A9:K33").clearContents(511)

    If sok = "" Then
        oSoek.getCellRangeByName("A7").setString("Skriv søkeord ovenfor og trykk «SØK I KB →».")
        Exit Sub
    End If

    ' Last hele dataområdet A2:K301 på én gang (raskt)
    arr = oKB.getCellRangeByName("A2:K" & MAX_K).getDataArray()

    antall = 0
    ReDim treff(1 To 1)
    For i = 0 To MAX_K - 2
        If SomTekst(arr(i, KOL_TITTEL)) <> "" Then
            If kat = "" Or SomTekst(arr(i, KOL_KATEGORI)) = kat Then
                score = ScoreDok(sok, SomTekst(arr(i, KOL_TITTEL)), _
                                 SomTekst(arr(i, KOL_NOKKEL)), _
                                 SomTekst(arr(i, KOL_PROBLEM)), _
                                 SomTekst(arr(i, KOL_LOSNING)))
                If score > 0 Then
                    antall = antall + 1
                    ReDim Preserve treff(1 To antall)
                    treff(antall).Score = score
                    treff(antall).Dato = DatoSomTall(arr(i, KOL_DATO))
                    treff(antall).Rad = i
                End If
            End If
        End If
    Next i

    If antall > 0 Then SorterTreff(treff)

    ' Skriv inntil N_KB treff (kolonner A..K = kildekolonner 0..10, samme rekkefølge)
    vist = antall
    If vist > N_KB Then vist = N_KB
    If vist > 0 Then
        ReDim ut(0 To vist - 1, 0 To 10)
        For j = 0 To vist - 1
            rad = treff(j + 1).Rad
            For c = 0 To 10
                ut(j, c) = arr(rad, c)
            Next c
        Next j
        ' Skriv alle treff i ett bulk-kall (rad 9 = 0-basert rad 8)
        oSoek.getCellRangeByPosition(0, 8, 10, 8 + vist - 1).setDataArray(ut)
        ' Marker tittel (B) og sett datoformat (H) per treffrad
        For j = 0 To vist - 1
            MarkerTreff(oSoek, 8 + j, 1)                                            ' B (tittel)
            KopierDatoformat(oSoek.getCellByPosition(7, 8 + j), oKB.getCellByPosition(7, 1))  ' H (dato), ref KB!H2
        Next j
    End If

    If antall = 0 Then
        oSoek.getCellRangeByName("A7").setString("Ingen treff — prøv andre søkeord.")
    Else
        oSoek.getCellRangeByName("A7").setString(antall & " treff i kunnskapsbasen:")
    End If
End Sub


' ============================================================================
' SØK SAKER
' ============================================================================
Sub SokSaker()
    ' 0-baserte kildekolonner i Saker (kontrakt.SAKER_SEARCH_FIELDS):
    Const KOL_SAKSNR As Long = 0      ' A
    Const KOL_OPPRETTET As Long = 1   ' B (tiebreak)
    Const KOL_KATEGORI As Long = 5    ' F
    Const KOL_TITTEL As Long = 7      ' H
    Const KOL_BESKRIVELSE As Long = 8 ' I (problem)
    Const KOL_STATUS As Long = 12     ' M (filter mot avkryssingsboksene)
    Const KOL_LOSNING As Long = 20    ' U
    Const KOL_KBREF As Long = 21      ' V
    ' Resultatområde (kontrakt): header rad 7, resultatrader 8..37,
    ' kolonner A,B,F,H,U,V (kontrakt.SOK_SAK_COLS).

    Dim oDoc As Object
    Dim oSoek As Object, oSaker As Object
    Dim arr As Variant
    Dim statusArr As Variant
    Dim tillatt() As String
    Dim treff() As Treff
    Dim sok As String
    Dim antTillatt As Long
    Dim antall As Long, vist As Long
    Dim i As Long, j As Long
    Dim score As Long
    Dim rad As Long, rr As Long

    oDoc = ThisComponent
    oSoek = oDoc.Sheets.getByName("Søk saker")
    oSaker = oDoc.Sheets.getByName("Saker")

    sok = Normaliser(oSoek.getCellByPosition(1, 3).getString())   ' B4

    ' Tøm resultatområdet (kolonner A,B,F,H,U,V) uansett
    oSoek.getCellRangeByName("A8:A37").clearContents(511)
    oSoek.getCellRangeByName("B8:B37").clearContents(511)
    oSoek.getCellRangeByName("F8:F37").clearContents(511)
    oSoek.getCellRangeByName("H8:H37").clearContents(511)
    oSoek.getCellRangeByName("U8:U37").clearContents(511)
    oSoek.getCellRangeByName("V8:V37").clearContents(511)

    ' Les status-filteret (skjult speil-tabell AA5:AB14; AA = statusnavn,
    ' AB = SANN/USANN fra avkryssingsboksene). kontrakt.SOK_SAK_STATUS.
    statusArr = oSoek.getCellRangeByName("AA5:AB14").getDataArray()
    antTillatt = 0
    ReDim tillatt(1 To 10)
    For i = 0 To 9
        If SomTekst(statusArr(i, 0)) <> "" Then
            If statusArr(i, 1) = True Then
                antTillatt = antTillatt + 1
                tillatt(antTillatt) = SomTekst(statusArr(i, 0))
            End If
        End If
    Next i

    If sok = "" Then
        oSoek.getCellRangeByName("A6").setString("Skriv søkeord ovenfor og trykk «SØK I SAKER →».")
        Exit Sub
    End If

    ' Last hele dataområdet A2:W501 på én gang
    arr = oSaker.getCellRangeByName("A2:W" & MAX_S).getDataArray()

    antall = 0
    ReDim treff(1 To 1)
    For i = 0 To MAX_S - 2
        If SomTekst(arr(i, KOL_TITTEL)) <> "" Then
            ' Status-filter: ta kun med saker i avkryssede statuser
            If ErTillatt(SomTekst(arr(i, KOL_STATUS)), tillatt, antTillatt) Then
                score = ScoreDok(sok, SomTekst(arr(i, KOL_TITTEL)), "", _
                                 SomTekst(arr(i, KOL_BESKRIVELSE)), _
                                 SomTekst(arr(i, KOL_LOSNING)))
                If score > 0 Then
                    antall = antall + 1
                    ReDim Preserve treff(1 To antall)
                    treff(antall).Score = score
                    treff(antall).Dato = DatoSomTall(arr(i, KOL_OPPRETTET))
                    treff(antall).Rad = i
                End If
            End If
        End If
    Next i

    If antall > 0 Then SorterTreff(treff)

    ' Skriv inntil N_SAK treff (kolonner A,B,F,H,U,V) — celle-for-celle siden
    ' kolonnene ikke er sammenhengende (setDataArray krever sammenhengende område).
    vist = antall
    If vist > N_SAK Then vist = N_SAK
    For j = 1 To vist
        rad = treff(j).Rad
        rr = 7 + (j - 1)   ' 0-basert resultatrad (rad 8 = 0-basert 7)
        SkrivCelle(oSoek.getCellByPosition(0, rr), arr(rad, KOL_SAKSNR))     ' A
        SkrivCelle(oSoek.getCellByPosition(1, rr), arr(rad, KOL_OPPRETTET))  ' B
        SkrivCelle(oSoek.getCellByPosition(5, rr), arr(rad, KOL_KATEGORI))   ' F
        SkrivCelle(oSoek.getCellByPosition(7, rr), arr(rad, KOL_TITTEL))     ' H
        SkrivCelle(oSoek.getCellByPosition(20, rr), arr(rad, KOL_LOSNING))   ' U
        SkrivCelle(oSoek.getCellByPosition(21, rr), arr(rad, KOL_KBREF))     ' V
        ' Datoformat på Opprettet (B) og markering av tittel (H)
        KopierDatoformat(oSoek.getCellByPosition(1, rr), oSaker.getCellByPosition(1, 1))  ' ref Saker!B2
        MarkerTreff(oSoek, rr, 7)                                            ' H (tittel)
    Next j

    If antall = 0 Then
        oSoek.getCellRangeByName("A6").setString("Ingen saker med valgte statuser matcher søket.")
    Else
        oSoek.getCellRangeByName("A6").setString(antall & " treff:")
    End If
End Sub


' ============================================================================
' BEST-EFFORT LISTENER (valgfritt — krever manuell tilordning)
' ----------------------------------------------------------------------------
' LO/OO har ingen Worksheet_Change. Ønsker dere sanntidssøk i stedet for
' knappetrykk, kan dere tilordne SokKB/SokSaker til dokumenthendelsen
' "Dokumentinnhold endret" (Verktøy > Tilpass > Hendelser > Lagre i:
' <dokumentet> > "Innhold endret" → pek til ModSok.SokKB). Da kjøres søket
' ved hver endring — merk at dette kan bli tregt på store filer (derav knapp
' som standard). En programmatisk XModifyListener kan også registreres ved
' åpning, f.eks. i en AutoOpen som T11/T12 kobler opp mot en konkret hendelse.
'
' Dette er bevisst ikke ferdig implementert her; T11/T12 kan koble det opp mot
' konkrete knapper/hendelser når .ods-pakkingen er på plass.
' ============================================================================
