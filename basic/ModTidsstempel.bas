' ============================================================================
' ModTidsstempel — tidsstempel for arkfanen "Saker".
' Standardmodul (StarBasic for LibreOffice Calc / Apache OpenOffice Calc).
'
' LO/OO har ingen Worksheet_Change-hendelse, så tidsstempelet er KNAPP-triggeret
' (eller kan tilordnes en dokumenthendelse, se basic/README.md). Sub-en
' StemplTidsstempel settes som makro på en knapp, f.eks. en trykknapp på
' Saker-arket.
'
' Erstatter sirkulær-referanse-formelen i Saker!B. Makroversjonen har INGEN
' formel i kolonne B — tidsstempelet skrives som en vanlig verdi (Now).
'
' Kolonnen B har allerede formatet "DD.MM.YYYY HH:MM" fra generatoren
' (build_common skriver number_format på B for alle profiler), så vi trenger
' bare å skrive verdien.
'
' Kolonner (0-basert, getDataArray/getCellByPosition):
'   Opprettet = 1 (B), Innmelder = 2 (C), Tittel = 7 (H)
' Datarader: 2..501 (kontrakt.MAX_S = 501).
' ============================================================================

' ----------------------------------------------------------------------------
' Fyller "Opprettet" (B) i Saker med nåværende dato/klokkeslett for rader der
' Innmelder (C) ELLER Tittel (H) er fylt ut, og Opprettet (B) er tom. Tømmer
' også Opprettet hvis BÅDE Innmelder og Tittel senere er fjernet — speiler
' formelen IF(C&H="","",IF(B="",NOW(),B)).
' ----------------------------------------------------------------------------
Sub StemplTidsstempel()
    Const KOL_OPPRETTET As Long = 1    ' B
    Const KOL_INNMELDER As Long = 2    ' C
    Const KOL_TITTEL As Long = 7       ' H
    Const MAX_S As Long = 501          ' kontrakt.MAX_S

    Dim oDoc As Object
    Dim oSheet As Object
    Dim arr As Variant
    Dim oCell As Object
    Dim i As Long
    Dim sisteRad As Long
    Dim innmelder As String, tittel As String, opprettet As String

    oDoc = ThisComponent
    oSheet = oDoc.Sheets.getByName("Saker")

    ' Les hele dataområdet A2:W501 på én gang (raskt; 0-basert 2D-array)
    arr = oSheet.getCellRangeByName("A2:W" & MAX_S).getDataArray()
    sisteRad = UBound(arr, 1)          ' 0..499 for 500 datarader

    For i = 0 To sisteRad
        innmelder = Trim(CStr(arr(i, KOL_INNMELDER)))
        tittel = Trim(CStr(arr(i, KOL_TITTEL)))
        opprettet = Trim(CStr(arr(i, KOL_OPPRETTET)))

        If innmelder = "" And tittel = "" Then
            ' Brukeren har fjernet både Innmelder og Tittel — tøm tidsstempelet
            If opprettet <> "" Then
                oSheet.getCellByPosition(KOL_OPPRETTET, i + 1).setString("")
            End If
        ElseIf opprettet = "" Then
            ' Fyll kun hvis Opprettet er tom — låser dermed tidsstempelet
            oCell = oSheet.getCellByPosition(KOL_OPPRETTET, i + 1)
            oCell.setValue(Now())
        End If
    Next i
End Sub
