Attribute VB_Name = "ModTidsstempel"
Option Explicit

' ============================================================================
' ModTidsstempel – automatisk tidsstempel for arkfanen "Saker".
' Standardmodul. Logikken ligger i TidsstempelEndring; selve Worksheet_Change
' er en énlinjes hendelsesbehandler som MÅ ligge i arkmodulen til "Saker"
' (Sheet2) – se nederst i filen.
'
' Erstatter sirkulær-referanse-formelen i Saker!B. Makroversjonen har INGEN
' formel i kolonne B – tidsstempelet skrives som en vanlig verdi (Now).
'
' Kolonner (kontrakt.SAKER_HEADERS, 1-basert):
'   B = Opprettet (2), C = Innmelder (3), H = Tittel (8)
' Datarader: 2..MAX_S (kontrakt.MAX_S = 501).
' ============================================================================

' ----------------------------------------------------------------------------
' Selve logikken: fyller "Opprettet" (B) når Innmelder (C) eller Tittel (H)
' skrives, og tømmer B hvis BÅDE C og H senere fjernes (speiler formelen
' IF(C&H="","",IF(B="",NOW(),B))).
' Robust også ved flercelle-klipp/lim: begrenses til kolonne C og H.
' Ark identifiseres via Target.Worksheet – derfor kallerbar fra arkmodulen.
' ----------------------------------------------------------------------------
Public Sub TidsstempelEndring(ByVal Target As Range)
    Const KOL_OPPRETTET As Long = 2   ' B
    Const KOL_INNMELDER As Long = 3   ' C
    Const KOL_TITTEL As Long = 8      ' H
    Const MAX_S As Long = 501         ' kontrakt.MAX_S

    Dim ws As Worksheet
    Dim berort As Range
    Dim kolonneC As Range
    Dim kolonneH As Range
    Dim c As Range
    Dim r As Long

    On Error GoTo Feil
    Application.EnableEvents = False

    Set ws = Target.Worksheet

    ' Finn hvilke endrede celler som ligger i kolonne C eller H (datarader)
    Set kolonneC = Intersect(Target, ws.Range(ws.Cells(2, KOL_INNMELDER), ws.Cells(MAX_S, KOL_INNMELDER)))
    Set kolonneH = Intersect(Target, ws.Range(ws.Cells(2, KOL_TITTEL), ws.Cells(MAX_S, KOL_TITTEL)))

    If Not kolonneC Is Nothing Then Set berort = kolonneC
    If Not kolonneH Is Nothing Then
        If berort Is Nothing Then
            Set berort = kolonneH
        Else
            Set berort = Union(berort, kolonneH)
        End If
    End If

    If Not berort Is Nothing Then
        For Each c In berort
            r = c.Row
            If Trim$(CStr(ws.Cells(r, KOL_INNMELDER).Value)) = "" And _
               Trim$(CStr(ws.Cells(r, KOL_TITTEL).Value)) = "" Then
                ' Brukeren har fjernet både Innmelder og Tittel – tøm tidsstempelet
                ws.Cells(r, KOL_OPPRETTET).ClearContents
            ElseIf Trim$(CStr(ws.Cells(r, KOL_OPPRETTET).Value)) = "" Then
                ' Fyll kun hvis Opprettet er tom – låser dermed tidsstempelet
                With ws.Cells(r, KOL_OPPRETTET)
                    .Value = Now
                    .NumberFormat = "DD.MM.YYYY HH:MM"
                End With
            End If
        Next c
    End If

Ut:
    Application.EnableEvents = True
    Exit Sub

Feil:
    ' Alltid re-aktiver hendelsesbehandling – ellers stopper ALLE hendelser i
    ' arbeidsboken (også andre makroer). Deretter avslutt kontrollert.
    Application.EnableEvents = True
    Resume Ut
End Sub


' ============================================================================
' HENDELSESBEHANDLER – legges i ARKMODULEN for "Saker" (Sheet2), ikke her:
'
'   Private Sub Worksheet_Change(ByVal Target As Range)
'       ModTidsstempel.TidsstempelEndring Target
'   End Sub
'
' Worksheet_Change kan IKKE ligge i en standardmodul – derfor delegeres den
' til den offentlige TidsstempelEndring ovenfor.
' ============================================================================
