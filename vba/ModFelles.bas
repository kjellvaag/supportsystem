Attribute VB_Name = "ModFelles"
Option Explicit

' ============================================================================
' ModFelles – felles hjelpefunksjoner for søk og tidsstempel.
' Standardmodul. Funksjonene er rene (ingen sideneffekter) og kan kalles fra
' ModSok, ModDashboard og andre moduler.
'
' Scoring matcher kontrakt.score EKSAKT (se kontrakt.py):
'   vekter: tittel=4, nøkkelord=3, problem=2, løsning=1
'   fuzzy (Levenshtein <= 2) brukes KUN når det ikke finnes eksakt delstreng-treff.
'
' Store/små bokstaver ignoreres ved at all tekst kjøres gjennom LCase i
' Normaliser – derfor ingen Option Compare Text her (og dermed ingen
' overraskelser med norske tegn i andre moduler).
' ============================================================================

' ----------------------------------------------------------------------------
' Normaliser: små bokstaver, bindestrek FJERNES (ikke mellomrom!), og løp av
' whitespace kollapses til ett mellomrom.
' KRAV (kontrakt.normalize):  Normaliser("Wi-Fi nede") = "wifi nede"
' NB: bindestreken FJERNES med Replace(t, "-", "") – IKKE erstattes med mellomrom.
' ----------------------------------------------------------------------------
Public Function Normaliser(ByVal t As String) As String
    Dim s As String
    Dim resultat As String
    Dim i As Long
    Dim tegn As String
    Dim sisteVarMellomrom As Boolean

    s = LCase(t)
    s = Replace(s, "-", "")          ' bindestrek FJERNES (ikke mellomrom!)

    ' Kollaps alle whitespace-tegn (mellomrom, tabulator, linjeskift) til ett mellomrom
    resultat = ""
    sisteVarMellomrom = True
    For i = 1 To Len(s)
        tegn = Mid$(s, i, 1)
        If tegn = " " Or tegn = vbTab Or tegn = vbLf Or tegn = vbCr Then
            If Not sisteVarMellomrom Then
                resultat = resultat & " "
                sisteVarMellomrom = True
            End If
        Else
            resultat = resultat & tegn
            sisteVarMellomrom = False
        End If
    Next i

    Normaliser = Trim$(resultat)
End Function

' ----------------------------------------------------------------------------
' Levenshtein: standard dynamisk-programmering redigeringsavstand (to rader).
' Returnerer Long. Matcher kontrakt.levenshtein.
' ----------------------------------------------------------------------------
Public Function Levenshtein(ByVal a As String, ByVal b As String) As Long
    Dim la As Long, lb As Long
    Dim i As Long, j As Long
    Dim kostnad As Long
    Dim forrige() As Long
    Dim gjeldende() As Long
    Dim tmp As String

    If a = b Then
        Levenshtein = 0
        Exit Function
    End If

    la = Len(a)
    lb = Len(b)
    If la < lb Then
        ' Sørg for at a er den lengste (som kontrakt.levenshtein)
        tmp = a: a = b: b = tmp
        la = Len(a): lb = Len(b)
    End If
    If lb = 0 Then
        Levenshtein = la
        Exit Function
    End If

    ReDim forrige(0 To lb)
    ReDim gjeldende(0 To lb)
    For j = 0 To lb
        forrige(j) = j
    Next j

    For i = 1 To la
        gjeldende(0) = i
        For j = 1 To lb
            If Mid$(a, i, 1) = Mid$(b, j, 1) Then kostnad = 0 Else kostnad = 1
            ' gjeldende(j) = min(gjeldende(j-1)+1, forrige(j)+1, forrige(j-1)+kostnad)
            gjeldende(j) = gjeldende(j - 1) + 1
            If forrige(j) + 1 < gjeldende(j) Then gjeldende(j) = forrige(j) + 1
            If forrige(j - 1) + kostnad < gjeldende(j) Then gjeldende(j) = forrige(j - 1) + kostnad
        Next j
        For j = 0 To lb
            forrige(j) = gjeldende(j)
        Next j
    Next i

    Levenshtein = forrige(lb)
End Function

' ----------------------------------------------------------------------------
' ScoreDok: scorer et dokument mot et søk.
' Vekter (kontrakt.FELT_VEKTER): tittel=4, nøkkelord=3, problem=2, løsning=1.
' 1) For hvert søkeord: gi feltets vekt hvis ordet er delstreng i feltet.
' 2) Hvis ingen delstreng-treff totalt: fuzzy fallback der et ord i feltet må
'    ha Levenshtein-avstand <= 2 fra søkeordet (Exit For bryter kun ord-løkka,
'    slik at flere felt kan gi vekt – samme som kontrakt.score sin "break").
' ----------------------------------------------------------------------------
Public Function ScoreDok(ByVal query As String, ByVal tittel As String, ByVal nokkelord As String, _
                         ByVal problem As String, ByVal losning As String) As Long
    Const V_TITTEL As Long = 4
    Const V_NOKKEL As Long = 3
    Const V_PROBLEM As Long = 2
    Const V_LOSNING As Long = 1

    Dim term() As String
    Dim felt(0 To 3) As String
    Dim vekt(0 To 3) As Long
    Dim total As Long
    Dim i As Long, k As Long, j As Long
    Dim ord() As String
    Dim nq As String

    ScoreDok = 0
    nq = Normaliser(query)
    If nq = "" Then Exit Function

    ' Felter i samme rekkefølge som FELT_VEKTER (tittel, nøkkelord, problem, løsning)
    felt(0) = Normaliser(tittel)
    vekt(0) = V_TITTEL
    felt(1) = Normaliser(nokkelord)
    vekt(1) = V_NOKKEL
    felt(2) = Normaliser(problem)
    vekt(2) = V_PROBLEM
    felt(3) = Normaliser(losning)
    vekt(3) = V_LOSNING

    term = Split(nq, " ")

    ' 1) Eksakt delstreng-treff
    total = 0
    For i = LBound(term) To UBound(term)
        If term(i) <> "" Then
            For k = 0 To 3
                If felt(k) <> "" Then
                    If InStr(1, felt(k), term(i), vbTextCompare) > 0 Then
                        total = total + vekt(k)
                    End If
                End If
            Next k
        End If
    Next i

    ' 2) Fuzzy fallback – kun når ingen eksakt delstreng-treff
    If total = 0 Then
        For i = LBound(term) To UBound(term)
            If term(i) <> "" Then
                For k = 0 To 3
                    If felt(k) <> "" Then
                        ord = Split(felt(k), " ")
                        For j = LBound(ord) To UBound(ord)
                            If ord(j) <> "" Then
                                If Levenshtein(term(i), ord(j)) <= 2 Then
                                    total = total + vekt(k)
                                    Exit For   ' gå videre til neste felt
                                End If
                            End If
                        Next j
                    End If
                Next k
            End If
        Next i
    End If

    ScoreDok = total
End Function
