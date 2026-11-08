Attribute VB_Name = "ModulAPI"
Option Explicit

' =====================================================================
'  BET PRO v35 - date gratuite de pe football-data.co.uk
'  22 de ligi europene, cu goluri la pauza si la final, cornere, cartonase si cote Betfair
'  Exchange (Bet365 unde lipsesc).
'  Nu are nevoie de cheie si nu are limita de cereri.
'
'  Butoane:
'    Meciurile zilei ....... meciurile programate din cele 22 de ligi
'    Actualizeaza istoric .. rezultatele sezonului curent + celui trecut
'                            (apasa-l de 1-2 ori pe saptamana)
'    Download (in lista) ... umple foile de calcul pentru meciul ales
'    Backtest .............. alege singur cele mai bune setari, joaca botul pe hartie pe tot
'                            istoricul (foaia Backtest) si corecteaza procentele
'    Recomandari ........... ce ar juca botul azi + text gata de copiat (foaia Recomandari)
'    Trimite email ......... trimite recomandarile pe email (setari in foaia Email)
'    RuleazaAutomat ........ tot lantul + email; il porneste zilnic BetPro_Automat.vbs
'    LIVE / Stop LIVE ...... scor si minut live, sanse recalculate, biletele de azi live (foaia Live)
'    Cote reale ............ la Recomandari: cele mai bune cote de la toate casele (API-Football Pro)
'    Scanare value ......... O0.5/O1.5 HT, GG HT, O1.5, O2.5, O3.5, GG, HT/FT, scor corect si
'                            cornere, fiecare cu pragul lui (foaia Scanare),
'                            selectiile se noteaza in foaia Jurnal si se verifica
'                            automat la Actualizeaza istoric
'
'  Umple aceleasi foi ca inainte (Listing, Home, HomeHome, HomeAway,
'  Away, AwayAway, AwayHome, Odds), deci calculele autorului raman.
' =====================================================================

Private Const FD_BASE As String = "https://www.football-data.co.uk/"
Private Const LIGI As String = _
    "E0=ANGLIA: Premier League;E1=ANGLIA: Championship;E2=ANGLIA: League One;" & _
    "E3=ANGLIA: League Two;EC=ANGLIA: National League;" & _
    "SC0=SCOTIA: Premiership;SC1=SCOTIA: Championship;SC2=SCOTIA: League One;SC3=SCOTIA: League Two;" & _
    "D1=GERMANIA: Bundesliga;D2=GERMANIA: 2. Bundesliga;" & _
    "I1=ITALIA: Serie A;I2=ITALIA: Serie B;" & _
    "SP1=SPANIA: LaLiga;SP2=SPANIA: Segunda Division;" & _
    "F1=FRANTA: Ligue 1;F2=FRANTA: Ligue 2;" & _
    "N1=OLANDA: Eredivisie;B1=BELGIA: Pro League;P1=PORTUGALIA: Liga Portugal;" & _
    "T1=TURCIA: Super Lig;G1=GRECIA: Super League"
Private Const SH_IST As String = "Istoric"
Private Const N_FORMA As Long = 10     ' cate meciuri recente se iau pentru fiecare echipa la scanare
Private Const J_START As Long = 20     ' primul rand cu meciuri in foaia Jurnal
Private Const NPM As Long = 12         ' numarul de piete - 1 (pietele 0..12)
Private Const MIN_FORMA As Long = 5    ' sub atatea meciuri in istoric, meciul e sarit la scanare
Private gNMax As Long                  ' cate meciuri (pe teren) intra in forta unei echipe
Private gDecay As Double               ' cat de repede scade greutatea meciurilor vechi
Private gK As Double                   ' "meciuri imaginare" de medie a ligii (trage spre medie)
Private mEroare As String        ' motivul ultimei descarcari esuate
Private gTacut As Boolean        ' True cand ruleaza automat (fara ferestre)
Private gLog As String           ' mesajele adunate in modul automat
Private gLiveAuto As Boolean     ' LIVE auto pornit
Private gLiveNext As Date         ' urmatoarea actualizare LIVE
Private mNextB As Long           ' folosit de cititorul JSON
Private mLiveN As Long           ' meciurile live pregatite pentru cautare
Private mLiveH() As String
Private mLiveA() As String
Private mLiveF As Collection
Private mCornHT As Object        ' cornerele de la pauza, tinute minte cat e Excel deschis
Private Const API_DEF As String = ""   ' cheia se pune in foaia Email, B13

' ============================ MECIURILE ZILEI ============================
Public Sub getMS()
    Dim ws As Worksheet, lines As Variant, hdr As Object, f As Variant
    Dim d As Date, tzh As Long, n As Long, r As Long, dt As Double
    Dim iDiv As Long, iD As Long, iT As Long, iH As Long, iA As Long
    Dim outv() As Variant, cod As String, sr As String, sr1 As String

    On Error GoTo Esec
    Set ws = Sheets("Listing")
    If gTacut Then
        d = Date                          ' rularea automata ia mereu ziua de azi
    Else
        d = CDate(ws.Range("Z3").Value)
    End If
    tzh = ws.Range("X2").Value - 1

    StartLucru
    If ws.FilterMode Then ws.ShowAllData
    ws.Range("I2:I5000").ClearHyperlinks
    clearParsed
    clearMain

    If Not DescarcaCsv(FD_BASE & "fixtures.csv", lines, hdr) Then
        Mesaj "Nu am putut descarca programul meciurilor de pe football-data.co.uk." & vbLf & _
            mEroare, vbExclamation, "BET PRO"
        GoTo Final
    End If
    iDiv = Col(hdr, "Div"): iD = Col(hdr, "Date"): iT = Col(hdr, "Time")
    iH = Col(hdr, "HomeTeam"): iA = Col(hdr, "AwayTeam")
    If iD < 0 Or iH < 0 Or iA < 0 Then
        Mesaj "Fisierul de programari are alt format decat ma asteptam.", vbExclamation, "BET PRO"
        GoTo Final
    End If

    ReDim outv(1 To UBound(lines) + 1, 1 To 22)
    n = 0
    For r = 1 To UBound(lines)
        If Len(lines(r)) > 5 Then
            f = Split(lines(r), ",")
            If UBound(f) >= iA Then
                dt = CitesteData(f(iD), Camp(f, iT))
                If dt > 0 And Int(dt + tzh / 24#) = CLng(d) Then
                    n = n + 1
                    cod = Camp(f, iDiv)
                    outv(n, 1) = NumeLiga(cod)
                    outv(n, 2) = Format(dt + tzh / 24#, "dd.mm.yyyy hh:mm")
                    outv(n, 3) = Trim$(f(iH))
                    outv(n, 4) = Trim$(f(iA))
                    outv(n, 5) = "M" & n
                    outv(n, 6) = cod
                    outv(n, 9) = "Download"
                    outv(n, 15) = CotaSursa(f, hdr, "BFE>2.5", "B365>2.5", "Avg>2.5", "BbAv>2.5", sr)
                    outv(n, 14) = sr
                    outv(n, 13) = CotaSursa(f, hdr, "BFE<2.5", "B365<2.5", "Avg<2.5", "BbAv<2.5", sr1)
                    outv(n, 16) = dt
                    outv(n, 17) = CotaSursa(f, hdr, "BFEH", "B365H", "AvgH", "BbAvH", sr1)
                    outv(n, 18) = CotaSursa(f, hdr, "BFED", "B365D", "AvgD", "BbAvD", sr1)
                    outv(n, 19) = CotaSursa(f, hdr, "BFEA", "B365A", "AvgA", "BbAvA", sr1)
                    outv(n, 20) = outv(n, 17): outv(n, 21) = outv(n, 18): outv(n, 22) = outv(n, 19)
                End If
            End If
        End If
    Next r

    If n = 0 Then
        Mesaj "Nu sunt meciuri din cele 22 de ligi pe " & Format(d, "dd.mm.yyyy") & "." & vbLf & _
            "Programul de pe site acopera doar zilele urmatoare si se actualizeaza de cateva ori pe saptamana.", _
            vbInformation, "BET PRO"
        GoTo Final
    End If

    ws.Range("B2:B5000,J2:J5000").NumberFormat = "@"
    ws.Range("A2").Resize(n, 22).Value = outv
    ws.Range("A1:V" & (n + 1)).Sort Key1:=ws.Range("B1"), Order1:=xlAscending, Header:=xlYes

    ws.Hyperlinks.Add Anchor:=ws.Range("I2:I" & (n + 1)), Address:="", TextToDisplay:="Download"
    With ws.Range("I2:I" & (n + 1))
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(31, 56, 140)
        .Font.Size = 11
        .Font.Bold = True
        .Font.Underline = xlUnderlineStyleSingle
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    ws.Columns(5).ColumnWidth = 0.05
    ws.Columns("M:P").Hidden = True
    ws.Activate
    ws.Range("A1").Select

    If RanduriIstoric() < 2 Then
        Mesaj "Lista e gata. Istoricul e inca gol: apasa o data 'Actualizeaza istoric' " & _
            "inainte de primul Download.", vbInformation, "BET PRO"
    End If

Final:
    GataLucru
    Exit Sub
Esec:
    Mesaj "A aparut o eroare la meciurile zilei:" & vbLf & Err.Description, vbExclamation, "BET PRO"
    Resume Final
End Sub

' ============================ ISTORIC ============================
Public Sub ActualizeazaIstoric()
    Dim ws As Worksheet, lines As Variant, hdr As Object, f As Variant
    Dim coduri As Variant, cod As Variant, off As Long, sez As String
    Dim iD As Long, iT As Long, iH As Long, iA As Long, iFH As Long, iFA As Long, iHH As Long, iHA As Long
    Dim outv() As Variant, k As Long, r As Long, okF As Long, esuate As String, dt As Double, last As Long
    Dim motivEsec As String, iHC As Long, iAC As Long, iHY As Long, iAY As Long, sr As String, sr2 As String

    On Error GoTo Esec
    StartLucru
    Set ws = FoaieIstoric()
    coduri = CoduriLigi()
    ReDim outv(1 To 60000, 1 To 21)

    For off = 0 To 1
        sez = CodSezon(off)
        For Each cod In coduri
            Application.StatusBar = "BET PRO: descarc " & NumeLiga(CStr(cod)) & " (" & sez & ")..."
            DoEvents
            If DescarcaCsv(FD_BASE & "mmz4281/" & sez & "/" & cod & ".csv", lines, hdr) Then
                okF = okF + 1
                iD = Col(hdr, "Date"): iT = Col(hdr, "Time")
                iH = Col(hdr, "HomeTeam"): If iH < 0 Then iH = Col(hdr, "HT")
                iA = Col(hdr, "AwayTeam"): If iA < 0 Then iA = Col(hdr, "AT")
                iFH = Col(hdr, "FTHG"): If iFH < 0 Then iFH = Col(hdr, "HG")
                iFA = Col(hdr, "FTAG"): If iFA < 0 Then iFA = Col(hdr, "AG")
                iHH = Col(hdr, "HTHG"): iHA = Col(hdr, "HTAG")
                iHC = Col(hdr, "HC"): iAC = Col(hdr, "AC")
                iHY = Col(hdr, "HY"): iAY = Col(hdr, "AY")
                If iD >= 0 And iH >= 0 And iA >= 0 And iFH >= 0 And iFA >= 0 Then
                    For r = 1 To UBound(lines)
                        If Len(lines(r)) > 5 And k < 60000 Then
                            f = Split(lines(r), ",")
                            If UBound(f) >= iFA And UBound(f) >= iA Then
                                dt = CitesteData(f(iD), Camp(f, iT))
                                If dt > 0 And f(iFH) <> "" And f(iFA) <> "" Then
                                    k = k + 1
                                    outv(k, 1) = cod
                                    outv(k, 2) = NumeLiga(CStr(cod))
                                    outv(k, 3) = dt
                                    outv(k, 4) = Trim$(f(iH))
                                    outv(k, 5) = Trim$(f(iA))
                                    If Camp(f, iHH) <> "" And Camp(f, iHA) <> "" Then
                                        outv(k, 6) = Val(f(iHH)): outv(k, 7) = Val(f(iHA))
                                    End If
                                    outv(k, 8) = Val(f(iFH)): outv(k, 9) = Val(f(iFA))
                                    If Camp(f, iHC) <> "" And Camp(f, iAC) <> "" Then
                                        outv(k, 10) = Val(f(iHC)): outv(k, 11) = Val(f(iAC))
                                    End If
                                    ' cote Over/Under 2.5: Betfair Exchange daca exista, altfel Bet365, altfel media
                                    outv(k, 12) = CotaSursa(f, hdr, "BFE>2.5", "B365>2.5", "Avg>2.5", "BbAv>2.5", sr)
                                    If sr = "BF" Then
                                        outv(k, 13) = Cota(CotaDin(f, hdr, "BFE<2.5"))
                                        outv(k, 14) = Cota(CotaDin(f, hdr, "BFEC>2.5"))
                                        outv(k, 15) = Cota(CotaDin(f, hdr, "BFEC<2.5"))
                                    Else
                                        outv(k, 13) = CotaSursa(f, hdr, "B365<2.5", "B365<2.5", "Avg<2.5", "BbAv<2.5", sr2)
                                        outv(k, 14) = Cota(CotaDin(f, hdr, "B365C>2.5", "AvgC>2.5", "PC>2.5"))
                                        outv(k, 15) = Cota(CotaDin(f, hdr, "B365C<2.5", "AvgC<2.5", "PC<2.5"))
                                    End If
                                    outv(k, 16) = sr
                                    If Camp(f, iHY) <> "" And Camp(f, iAY) <> "" Then
                                        outv(k, 17) = Val(f(iHY)): outv(k, 18) = Val(f(iAY))
                                    End If
                                    outv(k, 19) = CotaSursa(f, hdr, "BFEH", "B365H", "AvgH", "BbAvH", sr2)
                                    outv(k, 20) = CotaSursa(f, hdr, "BFED", "B365D", "AvgD", "BbAvD", sr2)
                                    outv(k, 21) = CotaSursa(f, hdr, "BFEA", "B365A", "AvgA", "BbAvA", sr2)
                                End If
                            End If
                        End If
                    Next r
                End If
            Else
                esuate = esuate & cod & " " & sez & ", "
                If motivEsec = "" Then motivEsec = mEroare
                If okF = 0 And off = 0 And cod = coduri(0) Then
                    Mesaj "Nu pot descarca nici primul fisier:" & vbLf & vbLf & mEroare & vbLf & vbLf & _
                        "Trimite-mi o poza cu acest mesaj.", vbExclamation, "BET PRO"
                    GoTo Final
                End If
            End If
        Next cod
    Next off

    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If last >= 2 Then ws.Range("A2:U" & last).ClearContents
    If k > 0 Then
        ws.Range("C2:C60001").NumberFormat = "yyyy.mm.dd hh:mm"
        ws.Range("A2").Resize(k, 21).Value = outv
    End If
    ws.Range("J1:R1").Value = Array("Cornere gazde", "Cornere oaspeti", "Cota O2.5", "Cota U2.5", _
        "Cota O2.5 inchidere", "Cota U2.5 inchidere", "Sursa cota", "Galbene gazde", "Galbene oaspeti")
    ws.Range("S1:U1").Value = Array("Cota 1", "Cota X", "Cota 2")
    ws.Range("T1").Value = "Cota X"
    ws.Range("W1").Value = "Actualizat"
    ws.Range("W2").Value = Now

    Dim verificate As Long
    verificate = VerificaJurnal()
    Dim bilV As Long
    bilV = VerificaBilete()
    If esuate <> "" Then esuate = vbLf & vbLf & "Nu am putut descarca: " & Left$(esuate, Len(esuate) - 2) & _
        vbLf & vbLf & "Motiv: " & motivEsec
    Mesaj "Istoric actualizat: " & k & " meciuri din " & okF & " fisiere." & vbLf & _
        "Meciuri verificate in Jurnal: " & verificate & ". Bilete verificate: " & bilV & "." & esuate, vbInformation, "BET PRO"

Final:
    GataLucru
    Exit Sub
Esec:
    Mesaj "A aparut o eroare la actualizarea istoricului:" & vbLf & Err.Description, vbExclamation, "BET PRO"
    Resume Final
End Sub

' butonul vechi "Rezultate" trimite aici
Public Sub getResults()
    ActualizeazaIstoric
End Sub

' butonul vechi "Test cheie" trimite aici
Public Sub TestCheie()
    Dim lines As Variant, hdr As Object
    If DescarcaCsv(FD_BASE & "mmz4281/" & CodSezon(0) & "/E0.csv", lines, hdr) Then
        Mesaj "Conexiunea merge. Premier League " & CodSezon(0) & ": " & UBound(lines) & " randuri.", vbInformation, "BET PRO"
    Else
        Mesaj "Nu ma pot conecta la football-data.co.uk." & vbLf & vbLf & mEroare, vbExclamation, "BET PRO"
    End If
End Sub

' ============================ DOWNLOAD MECI ============================
Public Sub getGame(id As String)
    Dim ws As Worksheet, rw As Long, acasa As String, oaspete As String
    Dim ts As Double, tzh As Long, nH As Long, nA As Long, o25 As Variant

    On Error GoTo Esec
    Set ws = Sheets("Listing")
    rw = RandMeci(id)
    If rw = 0 Then
        Mesaj "Nu gasesc meciul in lista. Apasa intai 'Meciurile zilei'.", vbExclamation, "BET PRO"
        Exit Sub
    End If
    acasa = Trim$(CStr(ws.Cells(rw, "C").Value))
    oaspete = Trim$(CStr(ws.Cells(rw, "D").Value))
    If IsNumeric(ws.Cells(rw, "P").Value2) Then ts = CDbl(ws.Cells(rw, "P").Value2)
    tzh = ws.Range("X2").Value - 1
    o25 = ws.Cells(rw, "O").Value
    ws.Range("Y1").Value = id

    StartLucru
    clearParsed

    With Sheets("Odds")
        .Range("A2:C7").Value = Array2D( _
            "FT 1", ws.Cells(rw, "Q").Value, ws.Cells(rw, "Q").Value, _
            "FT X", ws.Cells(rw, "R").Value, ws.Cells(rw, "R").Value, _
            "FT 2", ws.Cells(rw, "S").Value, ws.Cells(rw, "S").Value, _
            "O0.5 HT", Empty, Empty, _
            "O1.5 HT", Empty, Empty, _
            "O2.5 FT", o25, o25)
        .Range("E1").Value = "Cote Betfair Exchange (sau Bet365 unde lipsesc), de pe football-data.co.uk. Cotele pentru prima repriza nu sunt in fisier."
    End With

    nH = IstoricLocal(acasa, ts, tzh, "Home", "HomeHome", "HomeAway")
    nA = IstoricLocal(oaspete, ts, tzh, "Away", "AwayHome", "AwayAway")

    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    Sheets("BazaDate").Select

    If nH = 0 Or nA = 0 Then
        Mesaj "Istoric gasit: " & acasa & " - " & nH & " meciuri, " & oaspete & " - " & nA & " meciuri." & vbLf & _
            "Daca e 0, apasa 'Actualizeaza istoric' si incearca din nou.", vbExclamation, "BET PRO"
    End If

Final:
    GataLucru
    Exit Sub
Esec:
    Mesaj "A aparut o eroare la descarcarea meciului:" & vbLf & Err.Description, vbExclamation, "BET PRO"
    Resume Final
End Sub

' Istoricul unei echipe din foaia ascunsa "Istoric"
Private Function IstoricLocal(ByVal echipa As String, ByVal ts As Double, ByVal tzh As Long, _
        ByVal shAll As String, ByVal shAcasa As String, ByVal shDeplasare As String) As Long
    Dim ws As Worksheet, vals As Variant, last As Long, i As Long, cnt As Long
    Dim v(1 To 11) As Variant, wasHome As Boolean, nume As Variant
    Dim nAll As Long, nA As Long, nD As Long

    Set ws = FoaieIstoric()
    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If last < 2 Then Exit Function
    vals = ws.Range("A2:I" & last).Value2   ' Value2: datele vin ca numere

    For Each nume In Array(shAll, shAcasa, shDeplasare)
        Sheets(nume).Range("B2:B200").NumberFormat = "@"
    Next nume
    nAll = 1: nA = 1: nD = 1

    For i = 1 To UBound(vals, 1)
        If (Trim$(CStr(vals(i, 4))) = echipa Or Trim$(CStr(vals(i, 5))) = echipa) And IsNumeric(vals(i, 3)) Then
            If ts = 0 Or CDbl(vals(i, 3)) < ts Then
                wasHome = (Trim$(CStr(vals(i, 4))) = echipa)
                v(1) = vals(i, 2)
                v(2) = Format(CDbl(vals(i, 3)) + tzh / 24#, "yyyy.mm.dd hh:mm")
                v(3) = vals(i, 4)
                v(4) = vals(i, 5)
                If Gol(vals(i, 6)) Or Gol(vals(i, 7)) Then
                    v(5) = Empty: v(6) = Empty: v(7) = Empty: v(8) = Empty
                Else
                    v(5) = vals(i, 6): v(6) = vals(i, 7)
                    v(7) = vals(i, 8) - vals(i, 6): v(8) = vals(i, 9) - vals(i, 7)
                End If
                v(9) = IIf(wasHome, "Home", "Away")
                v(10) = vals(i, 8): v(11) = vals(i, 9)

                ScrieRand Sheets(shAll), nAll, v
                If wasHome Then
                    ScrieRand Sheets(shAcasa), nA, v
                Else
                    ScrieRand Sheets(shDeplasare), nD, v
                End If
                cnt = cnt + 1
            End If
        End If
    Next i

    For Each nume In Array(shAll, shAcasa, shDeplasare)
        With Sheets(nume)
            .Range("A1:K200").Sort Key1:=.Range("B1"), Order1:=xlDescending, Header:=xlYes
        End With
    Next nume
    IstoricLocal = cnt
End Function

Private Sub ScrieRand(sh As Worksheet, ByRef n As Long, v() As Variant)
    If n >= 200 Then Exit Sub
    n = n + 1
    sh.Range("A" & n).Resize(1, 11).Value = v
End Sub


' ============================ SCANARE VALUE ============================
' Trece prin toate meciurile din Listing si calculeaza din istoric (ultimele
' N_FORMA meciuri: gazdele acasa, oaspetii in deplasare) sansele pentru
' Over 0.5 HT, Over 1.5 HT si Over 2.5 FT. Pragul se schimba in foaia Scanare (B2).
Public Sub ScaneazaValue()
    Dim wl As Worksheet, ws As Worksheet, hist As Worksheet
    Dim L As Variant, H As Variant, lastL As Long, lastH As Long, i As Long, r As Long, m As Long
    Dim pr As Variant, outv() As Variant, n As Long, sarite As Long, nr As String
    Dim nh As Long, na As Long
    Dim lH As Double, lA As Double, lFT As Double, lHT As Double, lHTh As Double, lHTa As Double, lC As Double, lY As Double
    Dim p(0 To NPM) As Double, lab(0 To NPM) As String, liniaC As Double, liniaY As Double, k As Long, oricare As Boolean
    Dim cota25 As Double, cotaEf As Double, sursa As String, sel As String, ts As Double, c As Variant, colP As Variant, nume As Variant
    Dim jr() As Variant, nj As Long, adaugate As Long, com As Double, wPiata As Double, primaL As Boolean
    Dim liste As Object, ligiStat As Object, ord() As Long, cheie() As Double, t As Long, nTot As Long
    Dim okC As Boolean, lCh As Double, lCa As Double, okY As Boolean, lYh As Double, lYa As Double, pRaw As Double
    Dim calX(0 To NPM, 0 To 9) As Double, calY(0 To NPM, 0 To 9) As Double, calOk(0 To NPM, 0 To 9) As Boolean, areCal As Boolean

    On Error GoTo Esec
    Set wl = Sheets("Listing")
    lastL = wl.Cells(wl.Rows.Count, "E").End(xlUp).Row
    If lastL < 2 Then
        Set ws = FoaieScanare()
        ws.Range("A5:AM3000").Clear
        ws.Range("A3").Value = "Nu sunt meciuri in lista (" & Format(Now, "dd.mm.yyyy hh:mm") & ")."
        Mesaj "Lista e goala. Apasa intai 'Meciurile zilei'.", vbInformation, "BET PRO"
        Exit Sub
    End If
    Set hist = FoaieIstoric()
    lastH = hist.Cells(hist.Rows.Count, 1).End(xlUp).Row
    If lastH < 3 Then
        Mesaj "Istoricul e gol. Apasa 'Actualizeaza istoric'.", vbInformation, "BET PRO"
        Exit Sub
    End If

    StartLucru
    IncarcaSetari
    Set ws = FoaieScanare()
    pr = Praguri(ws)
    com = ComisionBF()
    wPiata = GreutatePiata()
    nume = PieteNume()

    L = wl.Range("A1:V" & lastL).Value2
    H = hist.Range("A2:U" & lastH).Value2

    ' forta echipelor si mediile ligilor, din tot istoricul, in ordinea datei
    nTot = UBound(H, 1)
    ReDim ord(1 To nTot): ReDim cheie(1 To nTot)
    For i = 1 To nTot
        ord(i) = i
        If IsNumeric(H(i, 3)) Then cheie(i) = CDbl(H(i, 3))
    Next i
    SorteazaIndex cheie, ord, 1, nTot
    Set liste = CreateObject("Scripting.Dictionary")
    Set ligiStat = CreateObject("Scripting.Dictionary")
    For t = 1 To nTot
        If cheie(ord(t)) > 0 Then AdaugaMeci H, ord(t), liste, ligiStat
    Next t
    areCal = IncarcaCalibrare(calX, calY, calOk)

    ReDim outv(1 To lastL, 1 To 39)
    ReDim jr(1 To lastL, 1 To 17)

    For i = 2 To lastL
        If Trim$(CStr(L(i, 3))) <> "" And Trim$(CStr(L(i, 4))) <> "" Then
            ts = 0
            If IsNumeric(L(i, 16)) Then ts = CDbl(L(i, 16))
            If ts > 0 And ts < CDbl(Now) Then GoTo Urmatorul   ' meci deja inceput
            If Not Lambdas(H, liste, ligiStat, CStr(L(i, 6)), Trim$(CStr(L(i, 3))), Trim$(CStr(L(i, 4))), _
                    nh, na, lH, lA, lHTh, lHTa, okC, lCh, lCa, okY, lYh, lYa) Then
                sarite = sarite + 1
                GoTo Urmatorul
            End If
            n = n + 1
            AjusteazaLaPiata lH, lA, lHTh, lHTa, NumarCota(L(i, 17)), NumarCota(L(i, 18)), NumarCota(L(i, 19)), _
                NumarCota(L(i, 15)), NumarCota(L(i, 13)), wPiata
            SalveazaLambda Trim$(CStr(L(i, 3))), Trim$(CStr(L(i, 4))), lH, lA, lHTh, lHTa, Not primaL, _
                IIf(okC, lCh + lCa, 0)
            primaL = True
            lFT = lH + lA: lHT = lHTh + lHTa

            For m = 0 To NPM: p(m) = 0: lab(m) = "": Next m
            p(0) = PoisPeste(0, lHT)
            p(1) = PoisPeste(1, lHT)
            p(2) = (1 - Exp(-lHTh)) * (1 - Exp(-lHTa))
            p(3) = PoisPeste(1, lFT)
            p(4) = PoisPeste(2, lFT)
            p(5) = PoisPeste(3, lFT)
            p(6) = (1 - Exp(-lH)) * (1 - Exp(-lA))
            CelMaiProbabilHTFT lHTh, lHTa, lH, lA, lab(7), p(7)
            p(7) = LimitaHTFT(p(7), lab(7), lH, lA)
            CelMaiProbabilScor lH, lA, lab(8), p(8)
            p(10) = PoisPeste(1, lH)
            p(11) = PoisPeste(1, lA)
            ' procentele corectate dupa ultimul backtest (daca exista)
            If areCal Then
                For m = 0 To NPM
                    If m <> 9 And m <> 12 Then p(m) = Calibreaza(m, p(m), calX, calY, calOk)
                Next m
            End If

            ' cornere: cea mai mare linie (7.5 - 11.5) peste prag
            lC = 0: liniaC = 0
            If okC Then
                lC = lCh + lCa
                For k = 11 To 7 Step -1
                    pRaw = PoisPeste(k, lC)
                    If areCal Then pRaw = Calibreaza(9, pRaw, calX, calY, calOk)
                    If pRaw >= pr(9) Then liniaC = k + 0.5: Exit For
                Next k
                If liniaC = 0 Then liniaC = 8.5
                p(9) = PoisPeste(CLng(liniaC - 0.5), lC)
                If areCal Then p(9) = Calibreaza(9, p(9), calX, calY, calOk)
                lab(9) = "O" & Format(liniaC, "0.0")
            End If
            ' cartonase galbene: cea mai mare linie (2.5 - 6.5) peste prag
            lY = 0: liniaY = 0
            If okY Then
                lY = lYh + lYa
                For k = 6 To 2 Step -1
                    pRaw = PoisPeste(k, lY)
                    If areCal Then pRaw = Calibreaza(12, pRaw, calX, calY, calOk)
                    If pRaw >= pr(12) Then liniaY = k + 0.5: Exit For
                Next k
                If liniaY = 0 Then liniaY = 3.5
                p(12) = PoisPeste(CLng(liniaY - 0.5), lY)
                If areCal Then p(12) = Calibreaza(12, p(12), calX, calY, calOk)
                lab(12) = "O" & Format(liniaY, "0.0")
            End If

            cota25 = 0: sursa = ""
            If IsNumeric(L(i, 15)) And Not IsEmpty(L(i, 15)) Then cota25 = CDbl(L(i, 15))
            sursa = CStr(L(i, 14))
            cotaEf = CotaEfectiva(cota25, sursa, com)

            outv(n, 1) = L(i, 1)
            outv(n, 2) = Right$(CStr(L(i, 2)), 5)
            outv(n, 3) = Trim$(CStr(L(i, 3))) & " - " & Trim$(CStr(L(i, 4)))
            outv(n, 4) = "'" & nh & " / " & na
            outv(n, 5) = p(0): outv(n, 6) = CotaMin(p(0))
            outv(n, 7) = p(1): outv(n, 8) = CotaMin(p(1))
            outv(n, 9) = p(2): outv(n, 10) = CotaMin(p(2))
            outv(n, 11) = p(3): outv(n, 12) = CotaMin(p(3))
            outv(n, 13) = p(4): outv(n, 14) = CotaMin(p(4))
            If cota25 > 1 Then outv(n, 15) = cota25: outv(n, 16) = p(4) * cotaEf - 1
            outv(n, 17) = p(5): outv(n, 18) = CotaMin(p(5))
            outv(n, 19) = p(6): outv(n, 20) = CotaMin(p(6))
            If lab(9) <> "" Then
                outv(n, 21) = lC: outv(n, 22) = lab(9)
                outv(n, 23) = p(9): outv(n, 24) = CotaMin(p(9))
            End If
            outv(n, 25) = "'" & lab(7): outv(n, 26) = p(7)
            outv(n, 27) = "'" & lab(8): outv(n, 28) = p(8)
            outv(n, 29) = p(10): outv(n, 30) = CotaMin(p(10))
            outv(n, 31) = p(11): outv(n, 32) = CotaMin(p(11))
            If lab(12) <> "" Then
                outv(n, 33) = lY: outv(n, 34) = lab(12)
                outv(n, 35) = p(12): outv(n, 36) = CotaMin(p(12))
            End If
            If cota25 > 1 Then outv(n, 37) = NumeSursa(sursa)

            sel = "": oricare = False
            For m = 0 To NPM
                If p(m) > 0 And p(m) >= pr(m) Then
                    oricare = True
                    Select Case m
                        Case 4
                            If cota25 > 1 And p(4) * cotaEf - 1 > 0 Then
                                sel = sel & "O2.5 VALUE la " & Format(cota25, "0.00") & " " & NumeSursa(sursa) & "; "
                            Else
                                sel = sel & "O2.5 de la " & Format(CotaMin(p(4)), "0.00") & "; "
                            End If
                        Case 7, 8
                            sel = sel & nume(m) & " " & lab(m) & " (" & Format(p(m), "0%") & "); "
                        Case 9, 12
                            sel = sel & nume(m) & " " & lab(m) & " de la " & Format(CotaMin(p(m)), "0.00") & "; "
                        Case Else
                            sel = sel & nume(m) & " de la " & Format(CotaMin(p(m)), "0.00") & "; "
                    End Select
                End If
            Next m
            outv(n, 38) = sel
            outv(n, 39) = p(0)
            If p(4) > outv(n, 39) Then outv(n, 39) = p(4)
            If p(6) > outv(n, 39) Then outv(n, 39) = p(6)

            ' Jurnal: meciurile cu cel putin o piata peste prag
            If oricare Then
                nj = nj + 1
                jr(nj, 1) = ts
                jr(nj, 2) = L(i, 1)
                jr(nj, 3) = Trim$(CStr(L(i, 3)))
                jr(nj, 4) = Trim$(CStr(L(i, 4)))
                For m = 0 To 6
                    jr(nj, 5 + m) = Recomandare(p(m), pr(m))
                Next m
                jr(nj, 12) = lab(7) & " (" & Format(p(7), "0%") & ")"     ' HT/FT: mereu
                jr(nj, 13) = lab(8) & " (" & Format(p(8), "0%") & ")"     ' Scor corect: mereu
                If lab(9) <> "" And p(9) >= pr(9) Then jr(nj, 14) = lab(9) & " (" & Format(p(9), "0%") & ")"
                jr(nj, 15) = Recomandare(p(10), pr(10))
                jr(nj, 16) = Recomandare(p(11), pr(11))
                If lab(12) <> "" And p(12) >= pr(12) Then jr(nj, 17) = lab(12) & " (" & Format(p(12), "0%") & ")"
            End If
        End If
Urmatorul:
    Next i

    ws.Range("A5:AM3000").Clear
    ws.Range("A3").Value = "Analizate: " & n & " meciuri care n-au inceput. Sarite (sub " & MIN_FORMA & _
        " meciuri in istoric): " & sarite & ". Scanat la " & Format(Now, "dd.mm.yyyy hh:mm") & "."
    If n > 0 Then
        nr = CStr(n + 4)
        ws.Range("A5").Resize(n, 39).Value = outv
        ws.Range("A4:AM" & nr).Sort Key1:=ws.Range("AM4"), Order1:=xlDescending, Header:=xlYes
        For Each c In Array("E", "G", "I", "K", "M", "Q", "S", "W", "Z", "AB", "AC", "AE", "AI", "AM")
            ws.Range(c & "5:" & c & nr).NumberFormat = "0%"
        Next c
        For Each c In Array("F", "H", "J", "L", "N", "O", "R", "T", "X", "AD", "AF", "AJ")
            ws.Range(c & "5:" & c & nr).NumberFormat = "0.00"
        Next c
        ws.Range("U5:U" & nr).NumberFormat = "0.0"
        ws.Range("AG5:AG" & nr).NumberFormat = "0.0"
        ws.Range("P5:P" & nr).NumberFormat = "+0%;-0%;0%"
        colP = Array("E", "G", "I", "K", "M", "Q", "S", "Z", "AB", "W", "AC", "AE", "AI")
        For r = 5 To n + 4
            For m = 0 To NPM
                If Not IsEmpty(ws.Range(colP(m) & r).Value) Then
                    If ws.Range(colP(m) & r).Value >= pr(m) Then
                        ws.Range(colP(m) & r).Interior.Color = RGB(198, 239, 206)
                        ws.Range(colP(m) & r).Font.Bold = True
                    End If
                End If
            Next m
            If Not IsEmpty(ws.Range("O" & r).Value) Then
                If ws.Range("P" & r).Value > 0 Then
                    ws.Range("P" & r).Interior.Color = RGB(198, 239, 206)
                Else
                    ws.Range("P" & r).Interior.Color = RGB(255, 214, 214)
                End If
            End If
            If CStr(ws.Range("AL" & r).Value) <> "" Then ws.Range("A" & r & ":D" & r & ",AL" & r).Font.Bold = True
        Next r
        ws.Range("A4:AL" & nr).Borders.Color = RGB(200, 205, 215)
        ws.Range("A4:AL" & nr).Columns.AutoFit
    End If
    ws.Columns("A:AL").Hidden = False
    ws.Columns("AM").Hidden = True
    If nj > 0 Then adaugate = AdaugaInJurnal(jr, nj)
    ws.Range("A3").Value = ws.Range("A3").Value & " Meciuri noi trecute in Jurnal: " & adaugate & "." & _
        IIf(areCal, " Procente corectate dupa ultimul Backtest.", " (Ruleaza Backtest o data ca procentele sa fie corectate.)")
    ws.Activate
    ws.Range("A1").Select

Final:
    GataLucru
    Exit Sub
Esec:
    Mesaj "A aparut o eroare la scanare:" & vbLf & Err.Description, vbExclamation, "BET PRO"
    Resume Final
End Sub

' ordinea pietelor folosita peste tot (Scanare, praguri, Jurnal)
Private Function PieteNume() As Variant
    PieteNume = Array("O0.5 HT", "O1.5 HT", "GG HT", "O1.5", "O2.5", "O3.5", "GG", "HT/FT", "Scor corect", "Cornere", _
        "Gazde O1.5", "Oaspeti O1.5", "Cartonase")
End Function

Private Function PraguriImplicite() As Variant
    ' O0.5 HT, O1.5 HT, GG HT, O1.5, O2.5, O3.5, GG, HT/FT, Scor corect, Cornere, Gazde O1.5, Oaspeti O1.5, Cartonase
    PraguriImplicite = Array(0.75, 0.75, 0.25, 0.75, 0.75, 0.5, 0.75, 0.4, 0.15, 0.75, 0.6, 0.5, 0.75)
End Function

' citeste pragurile din tabelul AF5:AG14 al foii Scanare
Private Function Praguri(ws As Worksheet) As Variant
    Dim r(0 To NPM) As Double, m As Long, v As Variant, impl As Variant
    impl = PraguriImplicite()
    For m = 0 To NPM
        v = ws.Range("AP" & (5 + m)).Value2
        r(m) = impl(m)
        If IsNumeric(v) And Not IsEmpty(v) Then
            r(m) = CDbl(v)
            If r(m) > 1 Then r(m) = r(m) / 100
            If r(m) <= 0 Or r(m) >= 1 Then r(m) = impl(m)
        End If
        ws.Range("AP" & (5 + m)).Value = r(m)
    Next m
    Praguri = r
End Function

Private Function CotaSursa(f As Variant, hdr As Object, ByVal cBF As String, ByVal cB365 As String, _
        ByVal cAvg As String, ByVal cAvg2 As String, ByRef sursa As String) As Variant
    Dim x As Double
    sursa = ""
    x = CotaDin(f, hdr, cBF)
    If x > 1 Then sursa = "BF": CotaSursa = x: Exit Function
    x = CotaDin(f, hdr, cB365)
    If x > 1 Then sursa = "B365": CotaSursa = x: Exit Function
    x = CotaDin(f, hdr, cAvg, cAvg2)
    If x > 1 Then sursa = "AVG": CotaSursa = x: Exit Function
    CotaSursa = Empty
End Function

Private Function NumeSursa(ByVal sursa As String) As String
    Select Case sursa
        Case "BF": NumeSursa = "Betfair"
        Case "B365": NumeSursa = "Bet365"
        Case "": NumeSursa = ""
        Case Else: NumeSursa = "medie case"
    End Select
End Function

Private Function CotaEfectiva(ByVal cv As Double, ByVal sursa As String, ByVal com As Double) As Double
    If cv <= 1 Then
        CotaEfectiva = 0
    ElseIf sursa = "BF" Then
        CotaEfectiva = 1 + (cv - 1) * (1 - com)
    Else
        CotaEfectiva = cv
    End If
End Function

Private Function ComisionBF() As Double
    Dim ws As Worksheet, v As Variant
    ComisionBF = 0.05
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Scanare")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    v = ws.Range("AP19").Value2
    If IsNumeric(v) And Not IsEmpty(v) Then
        If v > 1 Then v = v / 100
        If v >= 0 And v < 0.2 Then ComisionBF = CDbl(v)
    End If
End Function

' scorul exact cel mai probabil
Private Sub CelMaiProbabilScor(ByVal lH As Double, ByVal lA As Double, ByRef eticheta As String, ByRef pMax As Double)
    Dim a As Long, b As Long, q As Double
    pMax = -1
    For a = 0 To 6
        For b = 0 To 6
            q = PoisExact(a, lH) * PoisExact(b, lA)
            If q > pMax Then pMax = q: eticheta = a & "-" & b
        Next b
    Next a
End Sub

' Sansa ca o variabila Poisson cu media lam sa fie mai mare decat k (adica Over k+0.5)
Private Function PoisPeste(ByVal k As Long, ByVal lam As Double) As Double
    Dim j As Long, t As Double, s As Double
    If lam <= 0 Then Exit Function
    t = Exp(-lam): s = t
    For j = 1 To k
        t = t * lam / j
        s = s + t
    Next j
    PoisPeste = 1 - s
End Function


Private Function CotaMin(ByVal p As Double) As Double
    If p > 0 Then CotaMin = Round(1 / p, 2)
End Function

Private Function FoaieScanare() As Worksheet
    Dim ws As Worksheet, m As Long, nume As Variant, impl As Variant, vechi As Variant, vv As Variant
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Scanare")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets("Listing"))
        ws.Name = "Scanare"
    End If
    nume = PieteNume()
    impl = PraguriImplicite()

    ' tabelul de praguri (AO4:AP17) + comision Betfair (AP19).
    ' La prima rulare preia pragurile vechi din AF5:AG14 (versiunile anterioare) sau din B2.
    If CStr(ws.Range("AO5").Value) <> "O0.5 HT" Then
        vechi = ws.Range("B2").Value2
        ws.Range("AO4:AP4").Value = Array("Piata", "Prag")
        For m = 0 To NPM
            ws.Range("AO" & (5 + m)).Value = nume(m)
            ws.Range("AP" & (5 + m)).Value = impl(m)
            If m <= 9 And CStr(ws.Range("AF5").Value) = "O0.5 HT" Then
                vv = ws.Range("AG" & (5 + m)).Value2
                If IsNumeric(vv) And Not IsEmpty(vv) Then ws.Range("AP" & (5 + m)).Value = vv
            ElseIf impl(m) = 0.75 And IsNumeric(vechi) And Not IsEmpty(vechi) Then
                If vechi > 0 And vechi < 1 Then ws.Range("AP" & (5 + m)).Value = vechi
            End If
        Next m
        ws.Range("AO19").Value = "Comision Betfair"
        ws.Range("AP19").Value = 0.05
    End If
    With ws.Range("AO4:AP4")
        .Font.Bold = True: .Interior.Color = RGB(31, 56, 140): .Font.Color = RGB(255, 255, 255)
    End With
    If CStr(ws.Range("AO21").Value) <> "Greutate piata" Then
        ws.Range("AO21").Value = "Greutate piata"
        ws.Range("AP21").Value = 0.5
        ws.Range("AO22").Value = "0% = doar modelul, 100% = doar cotele"
    End If
    If CStr(ws.Range("AO23").Value) <> "Greutate Pinnacle" Then
        ws.Range("AO23").Value = "Greutate Pinnacle"
        ws.Range("AP23").Value = 0.5
        ws.Range("AO24").Value = "la VALUE: cat conteaza sansa Pinnacle fata de sansa botului"
    End If
    ws.Range("AP5:AP17,AP19,AP21,AP23").NumberFormat = "0%"
    ws.Range("AP5:AP17,AP19,AP21,AP23").Interior.Color = RGB(255, 242, 204)
    ws.Range("AP5:AP17,AP19,AP21,AP23").Font.Bold = True
    ws.Range("AO4:AP17,AO19:AP19,AO21:AP21,AO23:AP23").Borders.Color = RGB(200, 205, 215)
    ws.Columns("AO").ColumnWidth = 16
    ws.Columns("AP").ColumnWidth = 8

    ws.Range("A1:D2").Clear
    ws.Range("A1").Value = "Scanare - goluri, GG, gazde/oaspeti O1.5, HT/FT, scor corect, cornere, cartonase"
    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "Pragurile pe fiecare piata si comisionul Betfair sunt in tabelul galben din dreapta (coloanele AO-AP). " & _
        "Schimba-le acolo si apasa din nou Scanare value. Cota min = 1 / sansa."
    ws.Range("A4:AM4").Value = Array("Liga", "Ora", "Meci", "Istoric (acasa / depl.)", _
        "O0.5 HT", "Cota min", "O1.5 HT", "Cota min", "GG HT", "Cota min", _
        "O1.5", "Cota min", "O2.5", "Cota min", "Cota O2.5", "Avantaj O2.5", _
        "O3.5", "Cota min", "GG", "Cota min", _
        "Cornere asteptate", "Linie cornere", "Sansa linie", "Cota min", _
        "HT/FT probabil", "Sansa HT/FT", "Scor corect", "Sansa scor", _
        "Gazde O1.5", "Cota min", "Oaspeti O1.5", "Cota min", _
        "Galbene asteptate", "Linie galbene", "Sansa linie", "Cota min", _
        "Sursa cota", "Peste prag", "Scor")
    With ws.Range("A4:AM4")
        .Font.Bold = True
        .Interior.Color = RGB(31, 56, 140)
        .Font.Color = RGB(255, 255, 255)
    End With
    Set FoaieScanare = ws
End Function


' ============================ JURNAL ============================
' Foaia Jurnal: un rand pe meci, o coloana pe piata. In celula scrie ce a recomandat
' botul; dupa meci celula devine verde (a iesit) sau rosie (nu a iesit).
'  A Data meci | B Liga | C Meci | D Gazde | E Oaspeti
'  F O0.5 HT | G O1.5 HT | H O1.5 | I O2.5 | J GG | K HT/FT | L Cornere
'  M Scor HT | N Scor FT | O Cornere | P Verificat
'  Q..W (ascunse): rezultat DA/NU pe fiecare piata, pentru rezumat

Private Function Recomandare(ByVal p As Double, ByVal prag As Double) As String
    If p >= prag Then Recomandare = Format(p, "0%")
End Function

' rezultatul HT/FT cel mai probabil, din goluri asteptate la pauza si pe meci
Private Sub CelMaiProbabilHTFT(ByVal lHTh As Double, ByVal lHTa As Double, ByVal lFTh As Double, ByVal lFTa As Double, _
        ByRef eticheta As String, ByRef pMax As Double)
    Dim l2h As Double, l2a As Double, a As Long, b As Long, c As Long, d As Long
    Dim pa(0 To 6) As Double, pb(0 To 6) As Double, pc(0 To 6) As Double, pd(0 To 6) As Double
    Dim m(0 To 2, 0 To 2) As Double, r1 As Long, r2 As Long, et As Variant
    l2h = lFTh - lHTh: If l2h < 0.05 Then l2h = 0.05
    l2a = lFTa - lHTa: If l2a < 0.05 Then l2a = 0.05
    If lHTh < 0.05 Then lHTh = 0.05
    If lHTa < 0.05 Then lHTa = 0.05
    For a = 0 To 6
        pa(a) = PoisExact(a, lHTh): pb(a) = PoisExact(a, lHTa)
        pc(a) = PoisExact(a, l2h): pd(a) = PoisExact(a, l2a)
    Next a
    For a = 0 To 6
        For b = 0 To 6
            r1 = Rez(a, b)
            For c = 0 To 6
                For d = 0 To 6
                    r2 = Rez(a + c, b + d)
                    m(r1, r2) = m(r1, r2) + pa(a) * pb(b) * pc(c) * pd(d)
                Next d
            Next c
        Next b
    Next a
    et = Array("1", "X", "2")
    pMax = -1
    For a = 0 To 2
        For b = 0 To 2
            If m(a, b) > pMax Then pMax = m(a, b): eticheta = et(a) & "/" & et(b)
        Next b
    Next a
End Sub

Private Function Rez(ByVal g1 As Long, ByVal g2 As Long) As Long
    If g1 > g2 Then
        Rez = 0
    ElseIf g1 = g2 Then
        Rez = 1
    Else
        Rez = 2
    End If
End Function

Private Function PoisExact(ByVal k As Long, ByVal lam As Double) As Double
    Dim j As Long, t As Double
    t = Exp(-lam)
    For j = 1 To k
        t = t * lam / j
    Next j
    PoisExact = t
End Function

Private Function AdaugaInJurnal(jr() As Variant, ByVal nj As Long) As Long
    Dim ws As Worksheet, last As Long, vals As Variant, chei As Object, i As Long, k As String, r As Long, c As Long
    Set ws = FoaieJurnal()
    Set chei = CreateObject("Scripting.Dictionary")
    last = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    If last >= J_START Then
        vals = ws.Range("A" & J_START & ":E" & last).Value2
        For i = 1 To UBound(vals, 1)
            If IsNumeric(vals(i, 1)) Then
                k = CStr(vals(i, 4)) & "|" & CStr(vals(i, 5)) & "|" & CLng(Int(vals(i, 1)))
                If Not chei.Exists(k) Then chei.Add k, 1
            End If
        Next i
    Else
        last = J_START - 1
    End If
    r = last
    For i = 1 To nj
        k = jr(i, 3) & "|" & jr(i, 4) & "|" & CLng(Int(jr(i, 1)))
        If Not chei.Exists(k) Then
            chei.Add k, 1
            r = r + 1
            ws.Range("A" & r).Value = jr(i, 1)
            ws.Range("B" & r).Value = jr(i, 2)
            ws.Range("C" & r).Value = jr(i, 3) & " - " & jr(i, 4)
            ws.Range("D" & r).Value = jr(i, 3)
            ws.Range("E" & r).Value = jr(i, 4)
            For c = 5 To 17
                If CStr(jr(i, c)) <> "" Then ws.Cells(r, c + 1).Value = "'" & jr(i, c)
            Next c
            AdaugaInJurnal = AdaugaInJurnal + 1
        End If
    Next i
    On Error Resume Next
    Application.ErrorCheckingOptions.NumberAsText = False
    On Error GoTo 0
    If r >= J_START Then
        ws.Range("A" & J_START & ":A" & r).NumberFormat = "dd.mm.yyyy hh:mm"
        ws.Range("F" & J_START & ":W" & r).HorizontalAlignment = xlCenter
        ws.Range("A" & J_START & ":W" & r).Borders.Color = RGB(200, 205, 215)
    End If
End Function

' Verifica meciurile neverificate din Jurnal cu rezultatele din Istoric
Private Function VerificaJurnal() As Long
    Dim ws As Worksheet, hist As Worksheet, last As Long, lastH As Long
    Dim H As Variant, J As Variant, idx As Object, i As Long, k As String, d As Long, z As Long
    Dim hh As Double, ha As Double, fh As Double, fa As Double, corn As Double, galb As Double, rw As Long, m As Long, c As Long
    Dim hasHT As Boolean, hasC As Boolean, hasY As Boolean, txt As String, rezReal As String, ok As Boolean, okSet As Boolean
    Dim et3 As Variant
    et3 = Array("1", "X", "2")

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Jurnal")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    Set ws = FoaieJurnal()
    last = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    If last < J_START Then Exit Function
    Set hist = FoaieIstoric()
    lastH = hist.Cells(hist.Rows.Count, 1).End(xlUp).Row
    If lastH < 3 Then Exit Function

    H = hist.Range("A2:U" & lastH).Value2
    Set idx = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(H, 1)
        If IsNumeric(H(i, 3)) Then
            k = Trim$(CStr(H(i, 4))) & "|" & Trim$(CStr(H(i, 5))) & "|" & CLng(Int(H(i, 3)))
            If Not idx.Exists(k) Then idx.Add k, i
        End If
    Next i

    J = ws.Range("A" & J_START & ":W" & last).Value2
    For i = 1 To UBound(J, 1)
        If CStr(J(i, 23)) = "" And IsNumeric(J(i, 1)) And CStr(J(i, 4)) <> "" Then
            z = 0
            For d = 0 To 1
                k = CStr(J(i, 4)) & "|" & CStr(J(i, 5)) & "|" & (CLng(Int(J(i, 1))) + d)
                If idx.Exists(k) Then z = idx(k): Exit For
                k = CStr(J(i, 4)) & "|" & CStr(J(i, 5)) & "|" & (CLng(Int(J(i, 1))) - d)
                If idx.Exists(k) Then z = idx(k): Exit For
            Next d
            If z > 0 Then
                rw = i + J_START - 1
                fh = H(z, 8): fa = H(z, 9)
                hasHT = Not Gol(H(z, 6)) And Not Gol(H(z, 7))
                hasC = Not Gol(H(z, 10)) And Not Gol(H(z, 11))
                hasY = Not Gol(H(z, 17)) And Not Gol(H(z, 18))
                hh = 0: ha = 0: corn = 0: galb = 0
                If hasHT Then
                    hh = H(z, 6): ha = H(z, 7)
                    ws.Range("S" & rw).Value = "'" & hh & ":" & ha
                End If
                ws.Range("T" & rw).Value = "'" & fh & ":" & fa
                If hasC Then
                    corn = H(z, 10) + H(z, 11)
                    ws.Range("U" & rw).Value = corn
                End If
                If hasY Then
                    galb = H(z, 17) + H(z, 18)
                    ws.Range("V" & rw).Value = galb
                End If

                For m = 0 To NPM
                    c = 6 + m
                    txt = Trim$(CStr(J(i, c)))
                    If txt <> "" Then
                        okSet = True
                        Select Case m
                            Case 0: If hasHT Then ok = (hh + ha >= 1) Else okSet = False
                            Case 1: If hasHT Then ok = (hh + ha >= 2) Else okSet = False
                            Case 2: If hasHT Then ok = (hh >= 1 And ha >= 1) Else okSet = False
                            Case 3: ok = (fh + fa >= 2)
                            Case 4: ok = (fh + fa >= 3)
                            Case 5: ok = (fh + fa >= 4)
                            Case 6: ok = (fh >= 1 And fa >= 1)
                            Case 7
                                If hasHT Then
                                    rezReal = et3(Rez(CLng(hh), CLng(ha))) & "/" & et3(Rez(CLng(fh), CLng(fa)))
                                    ok = (Left$(txt, 3) = rezReal)
                                Else
                                    okSet = False
                                End If
                            Case 8
                                ok = (Left$(txt, InStr(txt & " ", " ") - 1) = (CLng(fh) & "-" & CLng(fa)))
                            Case 9: If hasC Then ok = (corn > Val(Mid$(txt, 2))) Else okSet = False
                            Case 10: ok = (fh >= 2)
                            Case 11: ok = (fa >= 2)
                            Case 12: If hasY Then ok = (galb > Val(Mid$(txt, 2))) Else okSet = False
                        End Select
                        If okSet Then
                            With ws.Cells(rw, c)
                                .Interior.Color = IIf(ok, RGB(146, 208, 80), RGB(255, 120, 120))
                                .Font.Bold = True
                            End With
                            ws.Cells(rw, 24 + m).Value = IIf(ok, "DA", "NU")
                        Else
                            ws.Cells(rw, c).Interior.Color = RGB(217, 217, 217)
                        End If
                    End If
                Next m
                ws.Range("W" & rw).Value = "DA"
                VerificaJurnal = VerificaJurnal + 1
            End If
        End If
    Next i
End Function

Private Function FoaieJurnal() As Worksheet
    Dim ws As Worksheet, nume As Variant, r As Long, colR As Variant, m As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Jurnal")
    On Error GoTo 0

    If Not ws Is Nothing Then
        If CStr(ws.Range("S" & (J_START - 1)).Value) = "Scor HT" Then
            ' format nou - nimic de facut
        ElseIf CStr(ws.Range("P16").Value) = "Scor HT" And CStr(ws.Range("N16").Value) = "Scor corect" Then
            ' jurnalul din versiunea anterioara: pastram tot si facem loc pietelor noi
            ws.Rows("16:18").Insert
            ws.Columns("P:R").Insert
            ws.Columns("V").Insert
            ws.Columns("A:W").Hidden = False
            ScrieAntetJurnal ws
        Else
            ' format foarte vechi: il pastram separat
            On Error Resume Next
            ws.Name = "Jurnal vechi " & Format(Now, "ddmmhhnn")
            On Error GoTo 0
            Set ws = Nothing
        End If
    End If

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets("Listing"))
        ws.Name = "Jurnal"
        ScrieAntetJurnal ws
    End If

    ' rezumatul se rescrie de fiecare data (in coloanele F-J, care sunt mereu vizibile)
    ws.Range("A1:J" & (J_START - 2)).Clear
    ws.Range("A1").Value = "Jurnal - ce a recomandat botul si ce a iesit"
    ws.Range("A1").Font.Size = 14: ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "Verde = a iesit, rosu = nu a iesit, gri = nu se poate verifica."
    ws.Range("A3").Value = "HT/FT si Scor corect se noteaza la fiecare meci (cel mai probabil rezultat)."
    ws.Range("A4").Value = "Se verifica automat la 'Actualizeaza istoric'."
    ws.Range("F3:J3").Value = Array("Piata", "Recomandate", "Iesite", "Pierdute", "Reusita")
    colR = Array("X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ")
    nume = PieteNume()
    r = 4
    For m = 0 To NPM
        ws.Range("F" & r).Value = nume(m)
        ws.Range("H" & r).Formula = "=COUNTIF($" & colR(m) & "$" & J_START & ":$" & colR(m) & "$20000,""DA"")"
        ws.Range("I" & r).Formula = "=COUNTIF($" & colR(m) & "$" & J_START & ":$" & colR(m) & "$20000,""NU"")"
        ws.Range("G" & r).Formula = "=H" & r & "+I" & r
        ws.Range("J" & r).Formula = "=IFERROR(H" & r & "/G" & r & ",""-"")"
        r = r + 1
    Next m
    ' r = 17: total, r = 18: fara bombe (fara HT/FT si Scor corect, randurile 11 si 12)
    ws.Range("F" & r).Value = "Total"
    ws.Range("G" & r).Formula = "=SUM(G4:G16)"
    ws.Range("H" & r).Formula = "=SUM(H4:H16)"
    ws.Range("I" & r).Formula = "=SUM(I4:I16)"
    ws.Range("J" & r).Formula = "=IFERROR(H" & r & "/G" & r & ",""-"")"
    r = r + 1
    ws.Range("F" & r).Value = "Fara bombe"
    ws.Range("G" & r).Formula = "=SUM(G4:G10)+SUM(G13:G16)"
    ws.Range("H" & r).Formula = "=SUM(H4:H10)+SUM(H13:H16)"
    ws.Range("I" & r).Formula = "=SUM(I4:I10)+SUM(I13:I16)"
    ws.Range("J" & r).Formula = "=IFERROR(H" & r & "/G" & r & ",""-"")"
    ws.Range("F" & (r - 1) & ":J" & (r - 1)).Font.Bold = True
    ws.Range("J4:J" & r).NumberFormat = "0%"
    With ws.Range("F3:J3")
        .Font.Bold = True: .Interior.Color = RGB(31, 56, 140): .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    ws.Range("H4:H" & r).Interior.Color = RGB(198, 239, 206)
    ws.Range("I4:I" & r).Interior.Color = RGB(255, 214, 214)
    ws.Range("F" & r & ":J" & r).Font.Bold = True
    ws.Range("F3:J" & r).Borders.Color = RGB(200, 205, 215)
    ws.Range("G4:J" & r).HorizontalAlignment = xlCenter
    Set FoaieJurnal = ws
End Function

Private Sub ScrieAntetJurnal(ws As Worksheet)
    ws.Range("A" & (J_START - 1) & ":AJ" & (J_START - 1)).Value = Array("Data meci", "Liga", "Meci", "Gazde", "Oaspeti", _
        "O0.5 HT", "O1.5 HT", "GG HT", "O1.5", "O2.5", "O3.5", "GG", "HT/FT", "Scor corect", "Cornere", _
        "Gazde O1.5", "Oaspeti O1.5", "Galbene", _
        "Scor HT", "Scor FT", "Cornere total", "Galbene total", "Verificat", _
        "r1", "r2", "r3", "r4", "r5", "r6", "r7", "r8", "r9", "r10", "r11", "r12", "r13")
    With ws.Range("A" & (J_START - 1) & ":W" & (J_START - 1))
        .Font.Bold = True: .Interior.Color = RGB(31, 56, 140): .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    ws.Columns("A").ColumnWidth = 16
    ws.Columns("B").ColumnWidth = 24
    ws.Columns("C").ColumnWidth = 32
    ws.Columns("F:W").ColumnWidth = 11
    ws.Columns("D:E").Hidden = True
    ws.Columns("X:AJ").Hidden = True
End Sub

Public Sub VerificaJurnalAcum()
    Mesaj "Meciuri verificate: " & VerificaJurnal() & ".", vbInformation, "BET PRO"
End Sub


' ============================ BACKTEST ============================
' Joaca pe hartie toate meciurile din Istoric, in ordinea datei. Pentru fiecare meci
' foloseste DOAR meciurile jucate inaintea lui (exact ca in realitate), apoi compara:
'  1) pariuri simulate pe Over/Under 2.5 cu cotele Bet365 de atunci (si cu cota de inchidere)
'  2) calibrarea: cand botul zice X%, cat iese de fapt, pe fiecare piata
Public Sub Backtest()
    Dim hist As Worksheet, ws As Worksheet, H As Variant, lastH As Long, nTot As Long
    Dim ord() As Long, cheie() As Double, t As Long, i As Long, m As Long, b As Long, s As Long
    Dim liste As Object, home As String, away As String, hasHT As Boolean, hasC As Boolean
    Dim nh As Long, na As Long, hF As Double, hA As Double, hHF As Double, hHA As Double
    Dim aF As Double, aA As Double, aHF As Double, aHA As Double
    Dim nch As Long, nca As Long, hCF As Double, hCA As Double, aCF As Double, aCA As Double, x1 As Double, x2 As Double
    Dim lH As Double, lA As Double, lFT As Double, lHTh As Double, lHTa As Double, lHT As Double, lC As Double
    Dim p(0 To NPM) As Double, lab(0 To NPM) As String, hit(0 To NPM) As Boolean, ok(0 To NPM) As Boolean
    Dim gHTh As Double, gHTa As Double, fh As Double, fa As Double, et3 As Variant
    Dim calN(0 To NPM, 0 To 9) As Long, calP(0 To NPM, 0 To 9) As Double, calH(0 To NPM, 0 To 9) As Long
    Dim sN(0 To 7) As Long, sW(0 To 7) As Long, sP(0 To 7) As Double, sClv(0 To 7) As Double, sClvN(0 To 7) As Long
    Dim oO As Double, oU As Double, cO As Double, cU As Double, evO As Double, evU As Double, peste As Boolean
    Dim ligaN As Object, ligaP As Object, liga As String, testate As Long, r As Long, k As Variant
    Dim nume As Variant, d1 As Double, d2 As Double, dif As Double
    Dim ligiStat As Object, okC As Boolean, lCh As Double, lCa As Double, okY As Boolean, lYh As Double, lYa As Double
    Dim hasY As Boolean, com As Double, eO As Double, eU As Double, wPiata As Double

    On Error GoTo Esec
    Set hist = FoaieIstoric()
    lastH = hist.Cells(hist.Rows.Count, 1).End(xlUp).Row
    If lastH < 100 Then
        Mesaj "Istoricul e prea mic. Apasa intai 'Actualizeaza istoric'.", vbInformation, "BET PRO"
        Exit Sub
    End If
    If CStr(hist.Range("L1").Value) <> "Cota O2.5" Then
        Mesaj "Istoricul nu are inca cotele salvate. Apasa o data 'Actualizeaza istoric', apoi Backtest.", vbInformation, "BET PRO"
        Exit Sub
    End If

    StartLucru
    com = ComisionBF()
    wPiata = GreutatePiata()
    H = hist.Range("A2:U" & lastH).Value2
    nTot = UBound(H, 1)
    ReDim ord(1 To nTot): ReDim cheie(1 To nTot)
    For i = 1 To nTot
        ord(i) = i
        If IsNumeric(H(i, 3)) Then cheie(i) = CDbl(H(i, 3))
    Next i
    SorteazaIndex cheie, ord, 1, nTot

    ' optimizare: incearca mai multe setari si o pastreaza pe cea care prezice cel mai bine
    Dim gridK As Variant, gridD As Variant, gk1 As Variant, gd1 As Variant, scor As Double, bestScor As Double
    Dim bestK As Double, bestD As Double, grid(1 To 12, 1 To 3) As Double, ng As Long
    gridK = Array(3, 6, 10, 15)
    gridD = Array(0.9, 0.95, 0.98)
    gNMax = 20
    bestScor = 1E+30
    For Each gk1 In gridK
        For Each gd1 In gridD
            gK = gk1: gDecay = gd1
            ng = ng + 1
            Application.StatusBar = "BET PRO: optimizare " & ng & " din 12..."
            DoEvents
            scor = EvalSetari(H, ord, cheie, nTot)
            grid(ng, 1) = gk1: grid(ng, 2) = gd1: grid(ng, 3) = scor
            If scor < bestScor Then bestScor = scor: bestK = gk1: bestD = gd1
        Next gd1
    Next gk1
    gK = bestK: gDecay = bestD
    SalveazaSetari

    Set liste = CreateObject("Scripting.Dictionary")
    Set ligiStat = CreateObject("Scripting.Dictionary")
    Set ligaN = CreateObject("Scripting.Dictionary")
    Set ligaP = CreateObject("Scripting.Dictionary")
    et3 = Array("1", "X", "2")
    nume = PieteNume()

    For t = 1 To nTot
        i = ord(t)
        If cheie(i) = 0 Then GoTo UrmT
        If t Mod 500 = 0 Then
            Application.StatusBar = "BET PRO: backtest " & Format(t / nTot, "0%") & "..."
            DoEvents
        End If
        home = Trim$(CStr(H(i, 4))): away = Trim$(CStr(H(i, 5)))
        hasHT = Not Gol(H(i, 6)) And Not Gol(H(i, 7))
        hasC = Not Gol(H(i, 10)) And Not Gol(H(i, 11))
        fh = H(i, 8): fa = H(i, 9)
        gHTh = 0: gHTa = 0
        If hasHT Then gHTh = H(i, 6): gHTa = H(i, 7)

        hasY = Not Gol(H(i, 17)) And Not Gol(H(i, 18))
        If Lambdas(H, liste, ligiStat, CStr(H(i, 1)), home, away, nh, na, lH, lA, lHTh, lHTa, okC, lCh, lCa, okY, lYh, lYa) Then
            testate = testate + 1
            AjusteazaLaPiata lH, lA, lHTh, lHTa, NumarCota(H(i, 19)), NumarCota(H(i, 20)), NumarCota(H(i, 21)), _
                NumarCota(H(i, 12)), NumarCota(H(i, 13)), wPiata
            lFT = lH + lA: lHT = lHTh + lHTa
            For m = 0 To NPM: p(m) = 0: ok(m) = False: hit(m) = False: Next m
            p(0) = PoisPeste(0, lHT): p(1) = PoisPeste(1, lHT)
            p(2) = (1 - Exp(-lHTh)) * (1 - Exp(-lHTa))
            p(3) = PoisPeste(1, lFT): p(4) = PoisPeste(2, lFT): p(5) = PoisPeste(3, lFT)
            p(6) = (1 - Exp(-lH)) * (1 - Exp(-lA))
            CelMaiProbabilHTFT lHTh, lHTa, lH, lA, lab(7), p(7)
            p(7) = LimitaHTFT(p(7), lab(7), lH, lA)
            CelMaiProbabilScor lH, lA, lab(8), p(8)

            If hasHT Then
                ok(0) = True: hit(0) = (gHTh + gHTa >= 1)
                ok(1) = True: hit(1) = (gHTh + gHTa >= 2)
                ok(2) = True: hit(2) = (gHTh >= 1 And gHTa >= 1)
                ok(7) = True: hit(7) = (lab(7) = et3(Rez(CLng(gHTh), CLng(gHTa))) & "/" & et3(Rez(CLng(fh), CLng(fa))))
            End If
            ok(3) = True: hit(3) = (fh + fa >= 2)
            ok(4) = True: hit(4) = (fh + fa >= 3)
            ok(5) = True: hit(5) = (fh + fa >= 4)
            ok(6) = True: hit(6) = (fh >= 1 And fa >= 1)
            ok(8) = True: hit(8) = (lab(8) = CLng(fh) & "-" & CLng(fa))

            If hasC And okC Then
                lC = lCh + lCa
                p(9) = PoisPeste(8, lC)
                ok(9) = True: hit(9) = ((H(i, 10) + H(i, 11)) > 8)
            End If
            p(10) = PoisPeste(1, lH): ok(10) = True: hit(10) = (fh >= 2)
            p(11) = PoisPeste(1, lA): ok(11) = True: hit(11) = (fa >= 2)
            If hasY And okY Then
                p(12) = PoisPeste(3, lYh + lYa)
                ok(12) = True: hit(12) = ((H(i, 17) + H(i, 18)) > 3)
            End If

            ' calibrare
            For m = 0 To NPM
                If ok(m) Then
                    b = Int(p(m) * 10): If b > 9 Then b = 9
                    If b < 0 Then b = 0
                    calN(m, b) = calN(m, b) + 1
                    calP(m, b) = calP(m, b) + p(m)
                    If hit(m) Then calH(m, b) = calH(m, b) + 1
                End If
            Next m

            ' pariuri simulate pe Over / Under 2.5
            oO = 0: oU = 0: cO = 0: cU = 0
            If IsNumeric(H(i, 12)) Then oO = CDbl(H(i, 12))
            If IsNumeric(H(i, 13)) Then oU = CDbl(H(i, 13))
            If IsNumeric(H(i, 14)) Then cO = CDbl(H(i, 14))
            If IsNumeric(H(i, 15)) Then cU = CDbl(H(i, 15))
            If oO > 1 And oU > 1 Then
                peste = (fh + fa >= 3)
                ' pe Betfair castigul se micsoreaza cu comisionul
                eO = CotaEfectiva(oO, CStr(H(i, 16)), com)
                eU = CotaEfectiva(oU, CStr(H(i, 16)), com)
                evO = p(4) * eO - 1
                evU = (1 - p(4)) * eU - 1
                liga = CStr(H(i, 2))
                If evO >= 0 Then Pariaza sN, sW, sP, sClv, sClvN, 0, peste, eO, cO
                If evO >= 0.05 Then Pariaza sN, sW, sP, sClv, sClvN, 1, peste, eO, cO
                If evO >= 0.1 Then Pariaza sN, sW, sP, sClv, sClvN, 2, peste, eO, cO
                If evU >= 0 Then Pariaza sN, sW, sP, sClv, sClvN, 3, Not peste, eU, cU
                If evU >= 0.05 Then Pariaza sN, sW, sP, sClv, sClvN, 4, Not peste, eU, cU
                If evU >= 0.1 Then Pariaza sN, sW, sP, sClv, sClvN, 5, Not peste, eU, cU
                If p(4) >= 0.75 Then Pariaza sN, sW, sP, sClv, sClvN, 6, peste, eO, cO
                If p(4) >= 0.75 And evO >= 0 Then Pariaza sN, sW, sP, sClv, sClvN, 7, peste, eO, cO
                If evO >= 0 Then AdaugaLiga ligaN, ligaP, liga, IIf(peste, eO - 1, -1)
            End If
        End If

        ' abia acum meciul intra in istoricul echipelor (ca sa nu "vada viitorul")
        AdaugaMeci H, i, liste, ligiStat
UrmT:
    Next t

    ' calibrarea se salveaza si va corecta procentele din Scanare
    ScrieCalibrare calN, calP, calH, ligaN, ligaP

    ' ---------------- raport ----------------
    Set ws = FoaieNoua("Backtest")
    ws.Range("A1").Value = "Backtest - botul jucat pe hartie pe " & testate & " meciuri din trecut"
    ws.Range("A1").Font.Size = 14: ws.Range("A1").Font.Bold = True
    d1 = cheie(ord(1)): d2 = cheie(ord(nTot))
    For t = 1 To nTot
        If cheie(ord(t)) > 0 Then d1 = cheie(ord(t)): Exit For
    Next t
    ws.Range("A3").Value = "Model ajustat spre cotele Betfair cu greutatea " & Format(wPiata, "0%") & " (Scanare, AP21)."
    ws.Range("A2").Value = "Perioada " & Format(d1, "dd.mm.yyyy") & " - " & Format(d2, "dd.mm.yyyy") & _
        ". Fiecare meci e calculat doar cu meciurile jucate inaintea lui. Rulat la " & Format(Now, "dd.mm.yyyy hh:mm") & "."

    ws.Range("I4").Value = "0. Setari incercate automat (eroare mai mica = prezice mai bine)"
    ws.Range("I4").Font.Bold = True
    ws.Range("I5:K5").Value = Array("Tragere spre medie", "Greutate meciuri vechi", "Eroare")
    AntetTabel ws.Range("I5:K5")
    For i = 1 To 12
        ws.Range("I" & (5 + i)).Value = grid(i, 1)
        ws.Range("J" & (5 + i)).Value = grid(i, 2)
        ws.Range("K" & (5 + i)).Value = grid(i, 3)
        If grid(i, 1) = gK And grid(i, 2) = gDecay Then ws.Range("I" & (5 + i) & ":K" & (5 + i)).Interior.Color = RGB(198, 239, 206)
    Next i
    ws.Range("K6:K17").NumberFormat = "0.0000"
    ws.Range("I5:K17").Borders.Color = RGB(200, 205, 215)
    ws.Range("I18").Value = "Verde = setarea aleasa si folosita de acum in Scanare."
    ws.Columns("I:K").ColumnWidth = 20

    r = 4
    ws.Range("A" & r).Value = "1. Pariuri simulate pe Over/Under 2.5 cu cotele Betfair Exchange dupa comision (Bet365 unde lipsesc), miza 1"
    ws.Range("A" & r).Font.Bold = True
    r = r + 1
    ws.Range("A" & r & ":G" & r).Value = Array("Strategie", "Pariuri", "Castigate", "Rata", "Profit (unitati)", "ROI", "Bate cota de inchidere")
    AntetTabel ws.Range("A" & r & ":G" & r)
    Dim numeS As Variant
    numeS = Array("Over 2.5 cand are value (>= 0%)", "Over 2.5 cand are value >= 5%", "Over 2.5 cand are value >= 10%", _
        "Under 2.5 cand are value (>= 0%)", "Under 2.5 cand are value >= 5%", "Under 2.5 cand are value >= 10%", _
        "Over 2.5 cand sansa >= 75% (orice cota)", "Over 2.5 cand sansa >= 75% si are value")
    For s = 0 To 7
        r = r + 1
        ws.Range("A" & r).Value = numeS(s)
        ws.Range("B" & r).Value = sN(s)
        ws.Range("C" & r).Value = sW(s)
        If sN(s) > 0 Then
            ws.Range("D" & r).Value = sW(s) / sN(s)
            ws.Range("E" & r).Value = sP(s)
            ws.Range("F" & r).Value = sP(s) / sN(s)
            CuloareSemn ws.Range("E" & r & ":F" & r), sP(s)
        End If
        If sClvN(s) > 0 Then
            ws.Range("G" & r).Value = sClv(s) / sClvN(s)
            CuloareSemn ws.Range("G" & r), sClv(s)
        End If
    Next s
    ws.Range("D6:D" & r & ",F6:G" & r).NumberFormat = "0.0%"
    ws.Range("E6:E" & r).NumberFormat = "+0.0;-0.0;0"
    ws.Range("A5:G" & r).Borders.Color = RGB(200, 205, 215)
    r = r + 1
    ws.Range("A" & r).Value = "ROI peste 0 = profit. 'Bate cota de inchidere' peste 0 = ai luat cote mai bune decat piata la final; " & _
        "e cel mai sigur semn ca un model chiar are avantaj."
    ws.Range("A" & r).Font.Italic = True

    r = r + 2
    ws.Range("A" & r).Value = "2. Pe ligi (Over 2.5 cu value) - Recomandarile folosesc doar ligile pe plus"
    ws.Range("A" & r).Font.Bold = True
    r = r + 1
    ws.Range("A" & r & ":D" & r).Value = Array("Liga", "Pariuri", "Profit (unitati)", "ROI")
    AntetTabel ws.Range("A" & r & ":D" & r)
    Dim r0 As Long
    r0 = r + 1
    For Each k In ligaN.Keys
        r = r + 1
        ws.Range("A" & r).Value = k
        ws.Range("B" & r).Value = ligaN(k)
        ws.Range("C" & r).Value = ligaP(k)
        ws.Range("D" & r).Value = ligaP(k) / ligaN(k)
    Next k
    If r >= r0 Then
        ws.Range("A" & (r0 - 1) & ":D" & r).Sort Key1:=ws.Range("D" & (r0 - 1)), Order1:=xlDescending, Header:=xlYes
        For i = r0 To r
            CuloareSemn ws.Range("C" & i & ":D" & i), ws.Range("C" & i).Value
        Next i
        ws.Range("C" & r0 & ":C" & r).NumberFormat = "+0.0;-0.0;0"
        ws.Range("D" & r0 & ":D" & r).NumberFormat = "0.0%"
        ws.Range("A" & (r0 - 1) & ":D" & r).Borders.Color = RGB(200, 205, 215)
    End If

    r = r + 2
    ws.Range("A" & r).Value = "3. Calibrare: cand botul zice X%, cat iese de fapt (doar intervalele cu cel putin 30 de meciuri)"
    ws.Range("A" & r).Font.Bold = True
    r = r + 1
    ws.Range("A" & r & ":F" & r).Value = Array("Piata", "Interval sansa", "Meciuri", "Sansa medie model", "Rata reala", "Diferenta")
    AntetTabel ws.Range("A" & r & ":F" & r)
    r0 = r + 1
    For m = 0 To NPM
        For b = 0 To 9
            If calN(m, b) >= 30 Then
                r = r + 1
                ws.Range("A" & r).Value = IIf(m = 9, "Cornere O8.5", IIf(m = 12, "Galbene O3.5", nume(m)))
                ws.Range("B" & r).Value = "'" & (b * 10) & "-" & (b * 10 + 10) & "%"
                ws.Range("C" & r).Value = calN(m, b)
                ws.Range("D" & r).Value = calP(m, b) / calN(m, b)
                ws.Range("E" & r).Value = calH(m, b) / calN(m, b)
                dif = calH(m, b) / calN(m, b) - calP(m, b) / calN(m, b)
                ws.Range("F" & r).Value = dif
                If Abs(dif) <= 0.03 Then
                    ws.Range("F" & r).Interior.Color = RGB(198, 239, 206)
                ElseIf dif < 0 Then
                    ws.Range("F" & r).Interior.Color = RGB(255, 199, 206)
                Else
                    ws.Range("F" & r).Interior.Color = RGB(255, 235, 156)
                End If
            End If
        Next b
    Next m
    If r >= r0 Then
        ws.Range("D" & r0 & ":E" & r).NumberFormat = "0%"
        ws.Range("F" & r0 & ":F" & r).NumberFormat = "+0%;-0%;0%"
        ws.Range("A" & (r0 - 1) & ":F" & r).Borders.Color = RGB(200, 205, 215)
    End If
    r = r + 1
    ws.Range("A" & r).Value = "Verde = botul e corect (diferenta sub 3%). Rosu = botul e prea optimist. Galben = botul e prea prudent. " & _
        "Tabelul asta se foloseste automat ca sa corecteze procentele din Scanare."
    ws.Range("A" & r).Font.Italic = True

    ws.Columns("A").ColumnWidth = 44
    ws.Columns("B:G").ColumnWidth = 16
    ws.Range("B:G").HorizontalAlignment = xlCenter
    ws.Activate
    ws.Range("A1").Select

Final:
    GataLucru
    Exit Sub
Esec:
    Mesaj "A aparut o eroare la backtest:" & vbLf & Err.Description, vbExclamation, "BET PRO"
    Resume Final
End Sub

Private Sub Pariaza(sN() As Long, sW() As Long, sP() As Double, sClv() As Double, sClvN() As Long, _
        ByVal s As Long, ByVal castigat As Boolean, ByVal cota As Double, ByVal inchidere As Double)
    sN(s) = sN(s) + 1
    If castigat Then
        sW(s) = sW(s) + 1
        sP(s) = sP(s) + cota - 1
    Else
        sP(s) = sP(s) - 1
    End If
    If inchidere > 1 Then
        sClv(s) = sClv(s) + cota / inchidere - 1
        sClvN(s) = sClvN(s) + 1
    End If
End Sub

Private Sub AdaugaLiga(ligaN As Object, ligaP As Object, ByVal liga As String, ByVal profit As Double)
    If Not ligaN.Exists(liga) Then ligaN.Add liga, 0: ligaP.Add liga, 0#
    ligaN(liga) = ligaN(liga) + 1
    ligaP(liga) = ligaP(liga) + profit
End Sub

Private Sub AdaugaInLista(liste As Object, ByVal cheieL As String, ByVal i As Long)
    Dim c As Collection
    If Not liste.Exists(cheieL) Then
        Set c = New Collection
        liste.Add cheieL, c
    End If
    liste(cheieL).Add i
End Sub


' sortare rapida a indicilor dupa cheie (data)
Private Sub SorteazaIndex(cheie() As Double, ord() As Long, ByVal lo As Long, ByVal hi As Long)
    Dim a As Long, b As Long, piv As Double, tmp As Long
    a = lo: b = hi
    piv = cheie(ord((lo + hi) \ 2))
    Do While a <= b
        Do While cheie(ord(a)) < piv: a = a + 1: Loop
        Do While cheie(ord(b)) > piv: b = b - 1: Loop
        If a <= b Then
            tmp = ord(a): ord(a) = ord(b): ord(b) = tmp
            a = a + 1: b = b - 1
        End If
    Loop
    If lo < b Then SorteazaIndex cheie, ord, lo, b
    If a < hi Then SorteazaIndex cheie, ord, a, hi
End Sub

Private Function FoaieNoua(ByVal nume As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nume)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = nume
    Else
        ws.Cells.Clear
    End If
    Set FoaieNoua = ws
End Function

Private Sub AntetTabel(rg As Range)
    rg.Font.Bold = True
    rg.Interior.Color = RGB(31, 56, 140)
    rg.Font.Color = RGB(255, 255, 255)
End Sub

Private Sub CuloareSemn(rg As Range, ByVal v As Double)
    If v > 0 Then
        rg.Interior.Color = RGB(198, 239, 206)
    ElseIf v < 0 Then
        rg.Interior.Color = RGB(255, 199, 206)
    End If
End Sub


' ============================ MODEL ============================
' Forta fiecarei echipe = ultimele gNMax meciuri pe acelasi teren, cu greutate mai mare
' pentru cele recente, comparata cu media ligii si "trasa" spre medie cand datele sunt putine.

Private Sub AdaugaMeci(H As Variant, ByVal i As Long, liste As Object, ligiStat As Object)
    Dim dv As String, a As Variant
    AdaugaInLista liste, Trim$(CStr(H(i, 4))) & "|H", i
    AdaugaInLista liste, Trim$(CStr(H(i, 5))) & "|A", i
    dv = CStr(H(i, 1))
    If Not ligiStat.Exists(dv) Then ligiStat.Add dv, Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
    a = ligiStat(dv)
    a(0) = a(0) + 1: a(1) = a(1) + H(i, 8): a(2) = a(2) + H(i, 9)
    If Not Gol(H(i, 6)) And Not Gol(H(i, 7)) Then
        a(3) = a(3) + 1: a(4) = a(4) + H(i, 6): a(5) = a(5) + H(i, 7)
    End If
    If Not Gol(H(i, 10)) And Not Gol(H(i, 11)) Then
        a(6) = a(6) + 1: a(7) = a(7) + H(i, 10): a(8) = a(8) + H(i, 11)
    End If
    If UBound(H, 2) >= 18 Then
        If Not Gol(H(i, 17)) And Not Gol(H(i, 18)) Then
            a(9) = a(9) + 1: a(10) = a(10) + H(i, 17): a(11) = a(11) + H(i, 18)
        End If
    End If
    ligiStat(dv) = a
End Sub

Private Sub MediiLiga(ligiStat As Object, ByVal dv As String, ByRef gLh As Double, ByRef gLa As Double, _
        ByRef gHh As Double, ByRef gHa As Double, ByRef gCh As Double, ByRef gCa As Double, _
        ByRef gYh As Double, ByRef gYa As Double)
    Dim a As Variant
    gLh = 1.5: gLa = 1.2: gHh = 0.67: gHa = 0.53: gCh = 5.4: gCa = 4.4: gYh = 1.9: gYa = 2.1
    If Not ligiStat.Exists(dv) Then Exit Sub
    a = ligiStat(dv)
    If a(0) >= 30 Then
        If a(1) > 0 Then gLh = a(1) / a(0)
        If a(2) > 0 Then gLa = a(2) / a(0)
    End If
    If a(3) >= 30 Then
        If a(4) > 0 Then gHh = a(4) / a(3)
        If a(5) > 0 Then gHa = a(5) / a(3)
    End If
    If a(6) >= 30 Then
        If a(7) > 0 Then gCh = a(7) / a(6)
        If a(8) > 0 Then gCa = a(8) / a(6)
    End If
    If a(9) >= 30 Then
        If a(10) > 0 Then gYh = a(10) / a(9)
        If a(11) > 0 Then gYa = a(11) / a(9)
    End If
End Sub

Private Sub StatPonderat(H As Variant, liste As Object, ByVal cheieL As String, ByVal acasa As Boolean, _
        ByRef n As Long, ByRef w As Double, ByRef sF As Double, ByRef sA As Double, _
        ByRef wHT As Double, ByRef sHF As Double, ByRef sHA As Double, _
        ByRef nC As Long, ByRef wC As Double, ByRef sCF As Double, ByRef sCA As Double, _
        ByRef nY As Long, ByRef wY As Double, ByRef sYF As Double, ByRef sYA As Double)
    Dim c As Collection, j As Long, i As Long, wt As Double, areY As Boolean
    Dim kF As Long, kA As Long, kHF As Long, kHA As Long, kCF As Long, kCA As Long, kYF As Long, kYA As Long
    n = 0: w = 0: sF = 0: sA = 0: wHT = 0: sHF = 0: sHA = 0: nC = 0: wC = 0: sCF = 0: sCA = 0
    nY = 0: wY = 0: sYF = 0: sYA = 0
    If Not liste.Exists(cheieL) Then Exit Sub
    areY = (UBound(H, 2) >= 18)
    If acasa Then
        kF = 8: kA = 9: kHF = 6: kHA = 7: kCF = 10: kCA = 11: kYF = 17: kYA = 18
    Else
        kF = 9: kA = 8: kHF = 7: kHA = 6: kCF = 11: kCA = 10: kYF = 18: kYA = 17
    End If
    Set c = liste(cheieL)
    For j = c.Count To 1 Step -1
        If n >= gNMax Then Exit For
        i = c(j)
        wt = gDecay ^ n
        n = n + 1
        w = w + wt: sF = sF + wt * H(i, kF): sA = sA + wt * H(i, kA)
        If Not Gol(H(i, kHF)) And Not Gol(H(i, kHA)) Then
            wHT = wHT + wt: sHF = sHF + wt * H(i, kHF): sHA = sHA + wt * H(i, kHA)
        End If
        If Not Gol(H(i, kCF)) And Not Gol(H(i, kCA)) Then
            nC = nC + 1: wC = wC + wt: sCF = sCF + wt * H(i, kCF): sCA = sCA + wt * H(i, kCA)
        End If
        If areY Then
            If Not Gol(H(i, kYF)) And Not Gol(H(i, kYA)) Then
                nY = nY + 1: wY = wY + wt: sYF = sYF + wt * H(i, kYF): sYA = sYA + wt * H(i, kYA)
            End If
        End If
    Next j
End Sub

' forta relativa la media ligii, trasa spre 1 cu gK meciuri imaginare
Private Function Rel(ByVal suma As Double, ByVal greutate As Double, ByVal medie As Double) As Double
    If medie <= 0 Then
        Rel = 1
        Exit Function
    End If
    Rel = ((suma + gK * medie) / (greutate + gK)) / medie
End Function

' golurile asteptate (meci, pauza) si cornerele asteptate pentru gazde si oaspeti
Private Function Lambdas(H As Variant, liste As Object, ligiStat As Object, ByVal dv As String, _
        ByVal home As String, ByVal away As String, ByRef nh As Long, ByRef na As Long, _
        ByRef lH As Double, ByRef lA As Double, ByRef lHTh As Double, ByRef lHTa As Double, _
        ByRef okC As Boolean, ByRef lCh As Double, ByRef lCa As Double, _
        ByRef okY As Boolean, ByRef lYh As Double, ByRef lYa As Double) As Boolean
    Dim hW As Double, hSF As Double, hSA As Double, hWHT As Double, hSHF As Double, hSHA As Double
    Dim hNC As Long, hWC As Double, hSCF As Double, hSCA As Double, hNYc As Long, hWY As Double, hSYF As Double, hSYA As Double
    Dim aW As Double, aSF As Double, aSA As Double, aWHT As Double, aSHF As Double, aSHA As Double
    Dim aNC As Long, aWC As Double, aSCF As Double, aSCA As Double, aNYc As Long, aWY As Double, aSYF As Double, aSYA As Double
    Dim gLh As Double, gLa As Double, gHh As Double, gHa As Double, gCh As Double, gCa As Double, gYh As Double, gYa As Double

    okC = False: lCh = 0: lCa = 0: okY = False: lYh = 0: lYa = 0
    StatPonderat H, liste, home & "|H", True, nh, hW, hSF, hSA, hWHT, hSHF, hSHA, hNC, hWC, hSCF, hSCA, hNYc, hWY, hSYF, hSYA
    StatPonderat H, liste, away & "|A", False, na, aW, aSF, aSA, aWHT, aSHF, aSHA, aNC, aWC, aSCF, aSCA, aNYc, aWY, aSYF, aSYA
    If nh < MIN_FORMA Or na < MIN_FORMA Then Exit Function

    MediiLiga ligiStat, dv, gLh, gLa, gHh, gHa, gCh, gCa, gYh, gYa
    lH = gLh * Rel(hSF, hW, gLh) * Rel(aSA, aW, gLh)
    lA = gLa * Rel(aSF, aW, gLa) * Rel(hSA, hW, gLa)
    lHTh = gHh * Rel(hSHF, hWHT, gHh) * Rel(aSHA, aWHT, gHh)
    lHTa = gHa * Rel(aSHF, aWHT, gHa) * Rel(hSHA, hWHT, gHa)
    If lHTh > lH * 0.9 Then lHTh = lH * 0.9
    If lHTa > lA * 0.9 Then lHTa = lA * 0.9
    If hNC >= MIN_FORMA And aNC >= MIN_FORMA Then
        okC = True
        lCh = gCh * Rel(hSCF, hWC, gCh) * Rel(aSCA, aWC, gCh)
        lCa = gCa * Rel(aSCF, aWC, gCa) * Rel(hSCA, hWC, gCa)
    End If
    If hNYc >= MIN_FORMA And aNYc >= MIN_FORMA Then
        okY = True
        lYh = gYh * Rel(hSYF, hWY, gYh) * Rel(aSYA, aWY, gYh)
        lYa = gYa * Rel(aSYF, aWY, gYa) * Rel(hSYA, hWY, gYa)
    End If
    Lambdas = True
End Function

' ---------- calibrare: corecteaza procentele dupa ce a iesit in backtest ----------
Private Sub ScrieCalibrare(calN() As Long, calP() As Double, calH() As Long, ligaN As Object, ligaP As Object)
    Dim ws As Worksheet, m As Long, b As Long, r As Long, k As Variant
    Set ws = FoaieNoua("Calibrare")
    ws.Range("A1:E1").Value = Array("Piata", "Interval", "Meciuri", "Sansa medie model", "Rata reala")
    r = 1
    For m = 0 To NPM
        For b = 0 To 9
            If calN(m, b) >= 30 Then
                r = r + 1
                ws.Range("A" & r).Value = m
                ws.Range("B" & r).Value = b
                ws.Range("C" & r).Value = calN(m, b)
                ws.Range("D" & r).Value = calP(m, b) / calN(m, b)
                ws.Range("E" & r).Value = calH(m, b) / calN(m, b)
            End If
        Next b
    Next m
    ws.Range("G1").Value = "Actualizat"
    ws.Range("G2").Value = Now
    ws.Range("I1:K1").Value = Array("Liga", "Pariuri", "ROI Over 2.5 value")
    r = 1
    For Each k In ligaN.Keys
        r = r + 1
        ws.Range("I" & r).Value = k
        ws.Range("J" & r).Value = ligaN(k)
        ws.Range("K" & r).Value = ligaP(k) / ligaN(k)
    Next k
    ws.Visible = xlSheetHidden
End Sub

Private Function IncarcaCalibrare(calX() As Double, calY() As Double, calOk() As Boolean) As Boolean
    Dim ws As Worksheet, last As Long, v As Variant, i As Long, m As Long, b As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Calibrare")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If last < 3 Then Exit Function
    v = ws.Range("A2:E" & last).Value2
    For i = 1 To UBound(v, 1)
        If IsNumeric(v(i, 1)) And IsNumeric(v(i, 2)) And Not IsEmpty(v(i, 1)) Then
            m = CLng(v(i, 1)): b = CLng(v(i, 2))
            If m >= 0 And m <= NPM And b >= 0 And b <= 9 Then
                calX(m, b) = v(i, 4): calY(m, b) = v(i, 5): calOk(m, b) = True
            End If
        End If
    Next i
    IncarcaCalibrare = True
End Function

Private Function Calibreaza(ByVal m As Long, ByVal p As Double, calX() As Double, calY() As Double, calOk() As Boolean) As Double
    Dim xs(1 To 10) As Double, ys(1 To 10) As Double, c As Long, b As Long, j As Long, r As Double
    For b = 0 To 9
        If calOk(m, b) Then
            c = c + 1: xs(c) = calX(m, b): ys(c) = calY(m, b)
        End If
    Next b
    If c < 2 Then
        Calibreaza = p
        Exit Function
    End If
    If p <= xs(1) Then
        r = p + (ys(1) - xs(1))
    ElseIf p >= xs(c) Then
        r = p + (ys(c) - xs(c))
    Else
        r = p
        For j = 1 To c - 1
            If p >= xs(j) And p <= xs(j + 1) Then
                If xs(j + 1) > xs(j) Then
                    r = ys(j) + (p - xs(j)) * (ys(j + 1) - ys(j)) / (xs(j + 1) - xs(j))
                Else
                    r = ys(j)
                End If
                Exit For
            End If
        Next j
    End If
    If r < 0.005 Then r = 0.005
    If r > 0.995 Then r = 0.995
    Calibreaza = r
End Function


' ============================ SETARI MODEL ============================
' Setarile modelului se aleg automat in Backtest si se pastreaza in foaia ascunsa "Setari".
Private Sub IncarcaSetari()
    Dim ws As Worksheet, v As Variant
    gNMax = 20: gDecay = 0.92: gK = 4
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Setari")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    v = ws.Range("B1").Value2
    If IsNumeric(v) And Not IsEmpty(v) Then
        If v >= 5 And v <= 60 Then gNMax = CLng(v)
    End If
    v = ws.Range("B2").Value2
    If IsNumeric(v) And Not IsEmpty(v) Then
        If v > 0.5 And v <= 1 Then gDecay = CDbl(v)
    End If
    v = ws.Range("B3").Value2
    If IsNumeric(v) And Not IsEmpty(v) Then
        If v >= 0 And v <= 50 Then gK = CDbl(v)
    End If
End Sub

Private Sub SalveazaSetari()
    Dim ws As Worksheet
    Set ws = FoaieNoua("Setari")
    ws.Range("A1:A3").Value = Application.Transpose(Array("Meciuri pe echipa", "Greutate meciuri vechi", "Tragere spre medie"))
    ws.Range("B1").Value = gNMax
    ws.Range("B2").Value = gDecay
    ws.Range("B3").Value = gK
    ws.Range("A5").Value = "Alese automat de Backtest la " & Format(Now, "dd.mm.yyyy hh:mm")
    ws.Visible = xlSheetHidden
End Sub

' cat de bine prezice modelul (eroare log - mai mic = mai bun), pe Over 1.5, Over 2.5, GG si Over 0.5 HT
Private Function EvalSetari(H As Variant, ord() As Long, cheie() As Double, ByVal nTot As Long) As Double
    Dim liste As Object, ligiStat As Object, t As Long, i As Long, cnt As Long, ll As Double
    Dim nh As Long, na As Long, lH As Double, lA As Double, lHTh As Double, lHTa As Double
    Dim okC As Boolean, lCh As Double, lCa As Double, okY As Boolean, lYh As Double, lYa As Double
    Dim fh As Double, fa As Double, start As Long
    Set liste = CreateObject("Scripting.Dictionary")
    Set ligiStat = CreateObject("Scripting.Dictionary")
    start = CLng(nTot * 0.3)
    For t = 1 To nTot
        i = ord(t)
        If cheie(i) > 0 Then
            If t > start Then
                If Lambdas(H, liste, ligiStat, CStr(H(i, 1)), Trim$(CStr(H(i, 4))), Trim$(CStr(H(i, 5))), _
                        nh, na, lH, lA, lHTh, lHTa, okC, lCh, lCa, okY, lYh, lYa) Then
                    fh = H(i, 8): fa = H(i, 9)
                    ll = ll + LogEr(PoisPeste(1, lH + lA), (fh + fa >= 2))
                    ll = ll + LogEr(PoisPeste(2, lH + lA), (fh + fa >= 3))
                    ll = ll + LogEr((1 - Exp(-lH)) * (1 - Exp(-lA)), (fh >= 1 And fa >= 1))
                    cnt = cnt + 3
                    If Not Gol(H(i, 6)) And Not Gol(H(i, 7)) Then
                        ll = ll + LogEr(PoisPeste(0, lHTh + lHTa), (H(i, 6) + H(i, 7) >= 1))
                        cnt = cnt + 1
                    End If
                End If
            End If
            AdaugaMeci H, i, liste, ligiStat
        End If
    Next t
    If cnt > 0 Then EvalSetari = ll / cnt Else EvalSetari = 99
End Function

Private Function LogEr(ByVal p As Double, ByVal aIesit As Boolean) As Double
    If p < 0.01 Then p = 0.01
    If p > 0.99 Then p = 0.99
    If aIesit Then LogEr = -Log(p) Else LogEr = -Log(1 - p)
End Function

' ============================ RECOMANDARI ============================
' Foaia "Recomandari": ce ar juca botul azi, pe baza Scanarii (procente corectate)
' si a ligilor care au iesit pe plus in Backtest.
Public Sub Recomandari()
    Dim ws As Worksheet, wsS As Worksheet, wc As Worksheet, wj As Worksheet
    Dim v As Variant, last As Long, i As Long, r As Long, j As Long, n As Long, pr As Variant
    Dim ligaRoi As Object, liga As String, lastC As Long, cv As Variant, areLigi As Boolean
    Dim idx() As Long, pv() As Double, txtP() As String, nume As String, bestP As Double, bestN As String, prag As Double
    Dim pComb As Double, linii As Collection, ln As Variant, k As Long
    Dim pePrg As Object, coteR As Object, areCote As Boolean
    Set pePrg = CreateObject("Scripting.Dictionary")

    ScaneazaValue
    On Error Resume Next
    IaCoteZi
    VerificaValue
    On Error GoTo Esec
    Set coteR = IncarcaCote()
    areCote = (coteR.Count > 0)
    Set wsS = FoaieScanare()
    last = wsS.Cells(wsS.Rows.Count, 3).End(xlUp).Row
    If last < 5 Then
        Set ws = FoaieNoua("Recomandari")
        ws.Range("A1").Value = "Recomandarile botului - " & Format(Date, "dd.mm.yyyy")
        ws.Range("A1").Font.Size = 16: ws.Range("A1").Font.Bold = True
        ws.Range("A3").Value = "Azi nu sunt meciuri de analizat in cele 22 de ligi (sau programul de pe site nu e inca actualizat)."
        ws.Range("A4").Value = "Nu e nimic de jucat azi."
        ws.Columns("A").ColumnWidth = 10
        Mesaj "Nu sunt meciuri de analizat. Apasa intai 'Meciurile zilei'.", vbInformation, "BET PRO"
        Exit Sub
    End If
    StartLucru
    v = wsS.Range("A5:AM" & last).Value2
    n = UBound(v, 1)
    pr = Praguri(wsS)
    Set linii = New Collection

    ' ROI pe ligi pentru Over 2.5 cu value, din ultimul Backtest
    Set ligaRoi = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Set wc = ThisWorkbook.Sheets("Calibrare")
    On Error GoTo Esec
    If Not wc Is Nothing Then
        lastC = wc.Cells(wc.Rows.Count, "I").End(xlUp).Row
        If lastC >= 2 Then
            cv = wc.Range("I1:K" & lastC).Value2
            For i = 2 To UBound(cv, 1)
                If CStr(cv(i, 1)) <> "" And IsNumeric(cv(i, 3)) Then
                    If CDbl(cv(i, 2)) >= 30 Then ligaRoi(CStr(cv(i, 1))) = CDbl(cv(i, 3))
                End If
            Next i
        End If
    End If
    areLigi = (ligaRoi.Count > 0)

    Set ws = FoaieNoua("Recomandari")
    ws.Range("A1").Value = "Recomandarile botului - " & Format(Date, "dd.mm.yyyy")
    ws.Range("A1").Font.Size = 16: ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "Procentele sunt corectate dupa Backtest (cat a iesit in realitate). Nicio selectie nu e sigura."
    ws.Range("A2").Font.Italic = True
    r = 4
    RaportBilete ws, r, linii

    ' ---------- 1. VALUE cu cote reale ----------
    Dim vI() As Long, vP() As Double, vE() As Double, vO() As Double, vN() As String, vC() As String, nv As Long
    Dim pz As Variant, pzN As Variant, pzX As Variant, q As Long, pp As Double, oo As Double, cs As String, eg As Double
    Dim exclV(0 To NPM) As Boolean, wjv As Worksheet, gv0 As Variant, hv0 As Variant, mm As Long
    On Error Resume Next
    Set wjv = ThisWorkbook.Sheets("Jurnal")
    On Error GoTo Esec
    If Not wjv Is Nothing Then
        For mm = 0 To NPM
            gv0 = wjv.Range("G" & (4 + mm)).Value2: hv0 = wjv.Range("H" & (4 + mm)).Value2
            If IsNumeric(gv0) And IsNumeric(hv0) And Not IsEmpty(gv0) Then
                If gv0 >= 10 Then
                    If hv0 / gv0 < pr(mm) - 0.1 Then exclV(mm) = True
                End If
            End If
        Next mm
    End If
If areCote Then
    Dim wPin As Double, pB As Double, pQ As Double, vB() As Double, vQ() As Double, idMeci As Object
    Dim nM As Long, nS As Long, ord() As Long, ordK() As String, tK As String, rezV As String
    wPin = GreutatePinnacle()
    Set idMeci = IncarcaIduri()
    ws.Range("A" & r).Value = "1. VALUE - cota casei e peste cota corecta (sansa botului amestecata " & Format(wPin, "0%") & " cu sansa Pinnacle)"
    ws.Range("A" & r).Font.Bold = True: ws.Range("A" & r).Font.Size = 12
    r = r + 1
    ws.Range("A" & r).Value = "Pinnacle e casa cea mai precisa. Cand botul e mult peste Pinnacle, de obicei greseste botul. De aceea aratam doar avantaje intre 3% si 15%."
    ws.Range("A" & r).Font.Italic = True
    r = r + 1
    ws.Range("A" & r & ":I" & r).Value = Array("Ora", "Meci", "Piata", "Sansa bot", "Sansa Pinnacle", "Sansa finala", "Cota", "Casa", "Avantaj")
    AntetTabel ws.Range("A" & r & ":I" & r)
    ReDim vI(1 To n * 12): ReDim vP(1 To n * 12): ReDim vE(1 To n * 12): ReDim vO(1 To n * 12): ReDim vN(1 To n * 12): ReDim vC(1 To n * 12)
    ReDim vB(1 To n * 12): ReDim vQ(1 To n * 12)
    For i = 1 To n
        pz = Array(5, 7, 11, 13, 17, 19, 29, 31, 26, 28, 23, 35)
        pzN = Array("Over 0.5 pauza", "Over 1.5 pauza", "Over 1.5", "Over 2.5", "Over 3.5", "GG", "Gazde peste 1.5", _
            "Oaspeti peste 1.5", "HT/FT " & v(i, 25), "Scor corect " & v(i, 27), "Cornere " & v(i, 22), "Galbene " & v(i, 34))
        pzX = Array(0, 1, 3, 4, 5, 6, 10, 11, 7, 8, 9, 12)
        For q = 0 To 11
            If IsNumeric(v(i, pz(q))) And Not IsEmpty(v(i, pz(q))) And Not exclV(pzX(q)) Then
                pB = CDbl(v(i, pz(q)))
                oo = CotaReala(coteR, CStr(v(i, 3)), CStr(pzN(q)), cs)
                If oo > 1 And pB >= 0.08 Then
                    pQ = SansaPinnacle(coteR, CStr(v(i, 3)), CStr(pzN(q)))
                    If pQ > 0 Then pp = (1 - wPin) * pB + wPin * pQ Else pp = pB
                    eg = pp * oo - 1
                    If eg >= 0.03 Then
                        nv = nv + 1
                        vI(nv) = i: vP(nv) = pp: vE(nv) = eg: vO(nv) = oo: vN(nv) = pzN(q): vC(nv) = cs
                        vB(nv) = pB: vQ(nv) = pQ
                    End If
                End If
            End If
        Next q
    Next i
    ' sortare dupa avantaj
    Dim t1 As Long, t2 As Long, tL As Long, tD As Double, tS As String
    For t1 = 1 To nv - 1
        For t2 = t1 + 1 To nv
            If vE(t2) > vE(t1) Then
                tL = vI(t1): vI(t1) = vI(t2): vI(t2) = tL
                tD = vP(t1): vP(t1) = vP(t2): vP(t2) = tD
                tD = vE(t1): vE(t1) = vE(t2): vE(t2) = tD
                tD = vO(t1): vO(t1) = vO(t2): vO(t2) = tD
                tD = vB(t1): vB(t1) = vB(t2): vB(t2) = tD
                tD = vQ(t1): vQ(t1) = vQ(t2): vQ(t2) = tD
                tS = vN(t1): vN(t1) = vN(t2): vN(t2) = tS
                tS = vC(t1): vC(t1) = vC(t2): vC(t2) = tS
            End If
        Next t2
    Next t1
    ' cele mai bune 15 realiste, afisate dupa ora
    ReDim ord(1 To 15): ReDim ordK(1 To 15)
    nM = 0
    For j = 1 To nv
        If vE(j) <= 0.15 And nM < 15 Then
            nM = nM + 1: ord(nM) = j: ordK(nM) = OraTxt(v(vI(j), 2)) & CStr(v(vI(j), 3))
        End If
    Next j
    For t1 = 1 To nM - 1
        For t2 = t1 + 1 To nM
            If ordK(t2) < ordK(t1) Then
                tL = ord(t1): ord(t1) = ord(t2): ord(t2) = tL
                tK = ordK(t1): ordK(t1) = ordK(t2): ordK(t2) = tK
            End If
        Next t2
    Next t1
    If nM = 0 Then
        r = r + 1
        ws.Range("A" & r).Value = "Azi nu e niciun value realist (3-15%). Nu e nimic de jucat la value - si asta e o recomandare."
    End If
    For t1 = 1 To nM
        j = ord(t1): i = vI(j): r = r + 1
        ws.Range("A" & r & ":I" & r).Value = Array("'" & OraTxt(v(i, 2)), v(i, 3), vN(j), vB(j), IIf(vQ(j) > 0, vQ(j), "-"), vP(j), vO(j), vC(j), vE(j))
        ws.Range("D" & r & ":F" & r & ",I" & r).NumberFormat = "0%"
        ws.Range("G" & r).NumberFormat = "0.00"
        ws.Range("I" & r).Interior.Color = RGB(198, 239, 206)
        If vQ(j) <= 0 Then ws.Range("E" & r).Interior.Color = RGB(255, 235, 156)
        AdaugaZi pePrg, OraTxt(v(i, 2)), CStr(v(i, 3)), "VALUE: " & vN(j) & " la " & Format(vO(j), "0.00") & " " & vC(j) & _
            " (" & Format(vP(j), "0%") & ", +" & Format(vE(j), "0%") & ")", ""
        SalveazaValue OraTxt(v(i, 2)), CStr(v(i, 3)), vN(j), vB(j), vQ(j), vP(j), vO(j), vC(j), vE(j), idMeci
    Next t1
    r = r + 1
    ws.Range("A" & r).Value = "Avantaj = sansa finala x cota - 1. Galben la Pinnacle = Pinnacle nu are piata, deci nu e verificat. Miza mica si egala pe fiecare (o unitate)."
    ws.Range("A" & r).Font.Italic = True
    ' suspecte
    nS = 0
    For j = 1 To nv
        If vE(j) > 0.15 Then nS = nS + 1
    Next j
    If nS > 0 Then
        r = r + 2
        ws.Range("A" & r).Value = "1b. Suspecte - avantaj prea mare (peste 15%). NU le juca fara sa verifici echipele (accidentari, rotatii, motivatie)."
        ws.Range("A" & r).Font.Bold = True
        r = r + 1
        ws.Range("A" & r & ":I" & r).Value = Array("Ora", "Meci", "Piata", "Sansa bot", "Sansa Pinnacle", "Sansa finala", "Cota", "Casa", "Avantaj")
        AntetTabel ws.Range("A" & r & ":I" & r)
        nS = 0
        For j = 1 To nv
            If vE(j) > 0.15 And nS < 8 Then
                nS = nS + 1: i = vI(j): r = r + 1
                ws.Range("A" & r & ":I" & r).Value = Array("'" & OraTxt(v(i, 2)), v(i, 3), vN(j), vB(j), IIf(vQ(j) > 0, vQ(j), "-"), vP(j), vO(j), vC(j), vE(j))
                ws.Range("D" & r & ":F" & r & ",I" & r).NumberFormat = "0%"
                ws.Range("G" & r).NumberFormat = "0.00"
                ws.Range("A" & r & ":I" & r).Interior.Color = RGB(255, 235, 156)
            End If
        Next j
    End If
    rezV = RezumatValue()
    If rezV <> "" Then
        r = r + 1
        ws.Range("A" & r).Value = rezV
        ws.Range("A" & r).Font.Bold = True
        linii.Add rezV
    End If
Else
    ' ---------- 1. value testat ----------
    ws.Range("A" & r).Value = "1. Value pe Over 2.5 (strategia care a iesit pe plus in backtest)"
    ws.Range("A" & r).Font.Bold = True: ws.Range("A" & r).Font.Size = 12
    r = r + 1
    ws.Range("A" & r & ":H" & r).Value = Array("Ora", "Meci", "Liga", "Sansa O2.5", "Cota", "Avantaj", "ROI liga in backtest", "Sursa")
    AntetTabel ws.Range("A" & r & ":H" & r)
    ReDim idx(1 To n): ReDim pv(1 To n)
    k = 0
    For i = 1 To n
        If IsNumeric(v(i, 15)) And IsNumeric(v(i, 16)) And IsNumeric(v(i, 13)) And Not IsEmpty(v(i, 15)) Then
            If v(i, 15) > 1 And v(i, 16) >= 0.05 And v(i, 13) >= 0.55 Then
                liga = CStr(v(i, 1))
                If (Not areLigi) Or (ligaRoi.Exists(liga)) Then
                    If (Not areLigi) Or ligaRoi(liga) >= 0 Then
                        k = k + 1: idx(k) = i: pv(k) = v(i, 16)
                    End If
                End If
            End If
        End If
    Next i
    SorteazaDesc idx, pv, k
    If k = 0 Then
        r = r + 1
        ws.Range("A" & r).Value = "Azi nu e niciun Over 2.5 cu value in ligile bune. Nu e nimic de jucat aici - si asta e o recomandare."
    End If
    For j = 1 To k
        i = idx(j): r = r + 1
        ws.Range("A" & r).Value = "'" & OraTxt(v(i, 2))
        ws.Range("B" & r).Value = v(i, 3)
        ws.Range("C" & r).Value = v(i, 1)
        ws.Range("D" & r).Value = v(i, 13)
        ws.Range("E" & r).Value = v(i, 15)
        ws.Range("F" & r).Value = v(i, 16)
        If areLigi Then ws.Range("G" & r).Value = ligaRoi(CStr(v(i, 1)))
        ws.Range("H" & r).Value = v(i, 37)
        ws.Range("D" & r & ",F" & r & ":G" & r).NumberFormat = "0%"
        ws.Range("E" & r).NumberFormat = "0.00"
        ws.Range("F" & r).Interior.Color = RGB(198, 239, 206)
        If areLigi Then AdaugaZi pePrg, OraTxt(v(i, 2)), CStr(v(i, 3)), "Over 2.5 VALUE (" & Format(v(i, 13), "0%") & ", cota " & Format(v(i, 15), "0.00") & " " & v(i, 37) & ")", ""
    Next j
    If Not areLigi Then
        r = r + 1
        ws.Range("A" & r).Value = "ATENTIE: ligile nu sunt inca verificate (ruleaza Backtest). Pana atunci selectiile de mai sus nu intra in textul pentru comunitate."
        ws.Range("A" & r).Font.Italic = True
    End If
End If

    ' ---------- 2. cele mai sigure ----------
    r = r + 2
    ws.Range("A" & r).Value = "2. Cele mai sigure de azi - Over 1.5, Over 2.5 sau GG (cea mai buna pentru fiecare meci, max 8)"
    ws.Range("A" & r).Font.Bold = True: ws.Range("A" & r).Font.Size = 12
    r = r + 1
    ws.Range("A" & r & ":E" & r).Value = Array("Ora", "Meci", "Piata", "Sansa", "Cota minima")
    AntetTabel ws.Range("A" & r & ":E" & r)
    ReDim idx(1 To n): ReDim pv(1 To n): ReDim txtP(1 To n)
    k = 0
    For i = 1 To n
        bestP = 0: bestN = ""
        AlegeMaxim v(i, 11), pr(3), "Over 1.5", bestP, bestN
        AlegeMaxim v(i, 13), pr(4), "Over 2.5", bestP, bestN
        AlegeMaxim v(i, 19), pr(6), "GG", bestP, bestN
        If bestN <> "" Then
            k = k + 1: idx(k) = i: pv(k) = bestP: txtP(k) = bestN
        End If
    Next i
    SorteazaDescTxt idx, pv, txtP, k
    If k = 0 Then
        r = r + 1
        ws.Range("A" & r).Value = "Niciun meci nu trece de pragurile tale azi."
    ElseIf k < 5 Then
        r = r + 1
        ws.Range("A" & r).Value = "Azi doar " & k & " meciuri trec de prag. Mai bine putine si bune decat 8 fortate."
        ws.Range("A" & r).Font.Italic = True
    End If
    pComb = 1
    For j = 1 To k
        If j > 8 Then Exit For
        i = idx(j): r = r + 1
        ws.Range("A" & r).Value = "'" & OraTxt(v(i, 2))
        ws.Range("B" & r).Value = v(i, 3)
        ws.Range("C" & r).Value = txtP(j)
        ws.Range("D" & r).Value = pv(j)
        ws.Range("E" & r).Value = CotaMin(pv(j))
        ws.Range("D" & r).NumberFormat = "0%"
        ws.Range("E" & r).NumberFormat = "0.00"
        ws.Range("D" & r).Interior.Color = RGB(198, 239, 206)
        pComb = pComb * pv(j)
        AdaugaZi pePrg, OraTxt(v(i, 2)), CStr(v(i, 3)), txtP(j) & " (" & Format(pv(j), "0%") & ", cota min " & Format(CotaMin(pv(j)), "0.00") & ")", ""
    Next j
    ' sansa reala a unui bilet combinat (toate trebuie sa iasa)
    If k >= 2 Then
        pComb = 1
        For j = 1 To k
            If j > 8 Then Exit For
            pComb = pComb * pv(j)
            If j = 2 Or j = 3 Or j = 5 Or j = 8 Or (j = k And j > 3) Then
                r = r + 1
                ws.Range("A" & r).Value = "Bilet din primele " & j & ": sansa ca toate sa iasa " & Format(pComb, "0%") & _
                    ", cota corecta " & Format(1 / pComb, "0.00")
                If pComb < 0.3 Then ws.Range("A" & r).Font.Color = RGB(192, 0, 0)
            End If
        Next j
        r = r + 1
        ws.Range("A" & r).Value = "Fiecare meci in plus pe bilet scade sansa. Un bilet merita doar daca cota de la casa e peste cota corecta."
        ws.Range("A" & r).Font.Italic = True
    End If

    ' ---------- 2b. bilete pe cote tinta ----------
    Dim tinte As Variant, tt As Variant, kL As Long, kMin As Long, rr As Double, nCand As Long, jj As Long
    Dim cP() As Double, cN() As String, mJ() As Long, mI() As Long, mD() As Double
    Dim prodO As Double, prodP As Double, gasit As Boolean, dd As Double, bestD As Double, txtBilet As String, bj As Long
    Dim legH() As String, legA() As String, legP() As String, mg As String, mo As String, cO() As Double, prodF As Double
    r = r + 2
    ws.Range("A" & r).Value = IIf(areCote, "2b. Bilete pe cote tinta din cote reale (doar selectii unde casa plateste peste cota corecta)", _
        "2b. Bilete pe cote tinta (un meci pe bilet o singura data, sanse corectate dupa backtest)")
    ws.Range("A" & r).Font.Bold = True: ws.Range("A" & r).Font.Size = 12
    r = r + 1
    ws.Range("A" & r).Value = "Cota corecta = cat ar trebui sa plateasca biletul ca sa fie corect. Sansa reala a unui bilet e cam 1 / cota corecta. " & _
        "Merita jucat doar daca cota totala de la casa (sau Betfair) e PESTE cota corecta."
    ws.Range("A" & r).Font.Italic = True
    ' pietele care in Jurnal ies clar mai rar decat promite botul (cel putin 10 verificate,
    ' reusita cu peste 10 puncte sub prag) nu intra in bilete pana nu isi revin
    Dim exclus(0 To NPM) As Boolean, txtEx As String, wjj As Worksheet, gv As Variant, hv As Variant, m2 As Long, numeP As Variant
    numeP = PieteNume()
    On Error Resume Next
    Set wjj = ThisWorkbook.Sheets("Jurnal")
    On Error GoTo Esec
    If Not wjj Is Nothing Then
        For m2 = 0 To NPM
            gv = wjj.Range("G" & (4 + m2)).Value2: hv = wjj.Range("H" & (4 + m2)).Value2
            If IsNumeric(gv) And IsNumeric(hv) And Not IsEmpty(gv) Then
                If gv >= 10 Then
                    If hv / gv < pr(m2) - 0.1 Then
                        exclus(m2) = True
                        txtEx = txtEx & IIf(txtEx = "", "", ", ") & numeP(m2) & " (" & Format(hv / gv, "0%") & " in Jurnal)"
                    End If
                End If
            End If
        Next m2
    End If
    If txtEx <> "" Then
        r = r + 1
        ws.Range("A" & r).Value = "Scoase din bilete pentru ca ies mai rar decat promite botul: " & txtEx & "."
        ws.Range("A" & r).Font.Color = RGB(192, 0, 0)
    End If
    ReDim cP(1 To n, 1 To 8): ReDim cN(1 To n, 1 To 8): ReDim cO(1 To n, 1 To 8)
    For i = 1 To n
        If Not exclus(3) Then AdaugaCand cP, cN, i, 1, v(i, 11), "Over 1.5"
        If Not exclus(4) Then AdaugaCand cP, cN, i, 2, v(i, 13), "Over 2.5"
        If Not exclus(6) Then AdaugaCand cP, cN, i, 3, v(i, 19), "GG"
        If Not exclus(10) Then AdaugaCand cP, cN, i, 4, v(i, 29), "Gazde peste 1.5"
        If Not exclus(11) Then AdaugaCand cP, cN, i, 5, v(i, 31), "Oaspeti peste 1.5"
        If Not exclus(0) Then AdaugaCand cP, cN, i, 6, v(i, 5), "Over 0.5 pauza"
        If CStr(v(i, 22)) <> "" And Not exclus(9) Then AdaugaCand cP, cN, i, 7, v(i, 23), "Cornere " & v(i, 22)
        If CStr(v(i, 34)) <> "" And Not exclus(12) Then AdaugaCand cP, cN, i, 8, v(i, 35), "Galbene " & v(i, 34)
        ' cu cote reale: intra doar selectiile unde casa plateste cel putin cat valoreaza
        For jj = 1 To 8
            If cP(i, jj) > 0 Then
                If areCote Then
                    cO(i, jj) = CotaReala(coteR, CStr(v(i, 3)), cN(i, jj), cs)
                    If cO(i, jj) <= 1 Or cP(i, jj) * cO(i, jj) < 1 Then cP(i, jj) = 0: cO(i, jj) = 0
                Else
                    cO(i, jj) = 1 / cP(i, jj)
                End If
            End If
        Next jj
    Next i
    tinte = Array(2, 3.5, 5, 7.5, 10, 15, 20)
    For Each tt In tinte
        gasit = False
        kMin = 1
        Do While CDbl(tt) ^ (1 / kMin) > 1.67 And kMin < 12
            kMin = kMin + 1
        Loop
        For kL = kMin To 12
            rr = CDbl(tt) ^ (1 / kL)
            ' pentru fiecare meci, candidatul cu cota corecta cea mai apropiata de cota dorita pe meci
            ReDim mI(1 To n): ReDim mJ(1 To n): ReDim mD(1 To n)
            nCand = 0
            For i = 1 To n
                bestD = 1E+30: bj = 0
                For jj = 1 To 8
                    If cP(i, jj) > 0 Then
                        dd = Abs(Log(cO(i, jj)) - Log(rr))
                        If dd < bestD Then bestD = dd: bj = jj
                    End If
                Next jj
                If bj > 0 Then
                    nCand = nCand + 1
                    mI(nCand) = i: mD(nCand) = -bestD: mJ(nCand) = bj
                End If
            Next i
            If nCand < kL Then Exit For
            SorteazaDescL mI, mD, mJ, nCand
            prodO = 1: prodP = 1
            For jj = 1 To kL
                prodP = prodP * cP(mI(jj), mJ(jj))
                prodO = prodO * cO(mI(jj), mJ(jj))
            Next jj
            If prodO >= CDbl(tt) * 0.85 And prodO <= CDbl(tt) * 1.5 Then
                gasit = True
                Exit For
            End If
        Next kL

        r = r + 2
        If Not gasit Then
            ws.Range("A" & r).Value = "Bilet cota ~" & tt & ": azi nu sunt destule meciuri potrivite."
        Else
            ws.Range("A" & r).Value = "Bilet cota ~" & tt & "  (" & kL & " meciuri)"
            ws.Range("A" & r).Font.Bold = True
            ws.Range("A" & r).Interior.Color = RGB(255, 242, 204)
            r = r + 1
            ws.Range("A" & r & ":E" & r).Value = Array("Ora", "Meci", "Piata", "Sansa", IIf(areCote, "Cota reala", "Cota corecta"))
            AntetTabel ws.Range("A" & r & ":E" & r)
            txtBilet = "Bilet ~" & tt & ": "
            ReDim legH(1 To kL): ReDim legA(1 To kL): ReDim legP(1 To kL)
            For jj = 1 To kL
                r = r + 1
                i = mI(jj)
                ws.Range("A" & r & ":E" & r).Value = Array("'" & OraTxt(v(i, 2)), v(i, 3), cN(i, mJ(jj)), cP(i, mJ(jj)), cO(i, mJ(jj)))
                ws.Range("D" & r).NumberFormat = "0%": ws.Range("E" & r).NumberFormat = "0.00"
                txtBilet = txtBilet & IIf(jj > 1, " + ", "") & v(i, 3) & " (" & cN(i, mJ(jj)) & ")"
                ImparteMeci CStr(v(i, 3)), mg, mo
                AdaugaZi pePrg, OraTxt(v(i, 2)), CStr(v(i, 3)), "", Replace(CStr(tt), ",", ".") & " (" & cN(i, mJ(jj)) & ")"
                legH(jj) = mg: legA(jj) = mo: legP(jj) = cN(i, mJ(jj))
            Next jj
            r = r + 1
            If areCote Then
                ws.Range("A" & r).Value = "Cota totala la case ~" & Format(prodO, "0.00") & "  -  sansa " & Format(prodP, "0%") & _
                    "  -  cota corecta " & Format(1 / prodP, "0.00") & "  -  avantaj " & Format(prodP * prodO - 1, "+0%;-0%")
            Else
                ws.Range("A" & r).Value = "Cota corecta totala " & Format(prodO, "0.00") & "  -  sansa reala " & Format(prodP, "0%") & _
                    "  -  joaca doar daca cota totala e cel putin " & Format(prodO, "0.00")
            End If
            ws.Range("A" & r).Font.Bold = True
            SalveazaBilet "Bilet", Replace(CStr(tt), ",", "."), prodO, prodP, Mid$(txtBilet, InStr(txtBilet, ": ") + 2), legH, legA, legP, kL
        End If
    Next tt

    ' ---------- 3. bombe ----------
    r = r + 2
    ws.Range("A" & r).Value = "3. Bombe (sanse mici, cote mari)"
    ws.Range("A" & r).Font.Bold = True: ws.Range("A" & r).Font.Size = 12
    r = r + 1
    ws.Range("A" & r & ":E" & r).Value = Array("Ora", "Meci", "Piata", "Sansa", "Cota minima")
    AntetTabel ws.Range("A" & r & ":E" & r)
    ReDim idx(1 To n): ReDim pv(1 To n)
    k = 0
    For i = 1 To n
        If IsNumeric(v(i, 26)) And Not IsEmpty(v(i, 26)) Then k = k + 1: idx(k) = i: pv(k) = v(i, 26)
    Next i
    SorteazaDesc idx, pv, k
    For j = 1 To k
        If j > 3 Then Exit For
        i = idx(j): r = r + 1
        ws.Range("A" & r & ":E" & r).Value = Array("'" & OraTxt(v(i, 2)), v(i, 3), "HT/FT " & v(i, 25), pv(j), CotaMin(pv(j)))
        ws.Range("D" & r).NumberFormat = "0%": ws.Range("E" & r).NumberFormat = "0.00"
        SalveazaBomba CStr(v(i, 3)), "HT/FT", "HT/FT " & v(i, 25), pv(j)
        AdaugaZi pePrg, OraTxt(v(i, 2)), CStr(v(i, 3)), "Bomba HT/FT " & v(i, 25) & " (" & Format(pv(j), "0%") & ")", ""
    Next j
    ReDim idx(1 To n): ReDim pv(1 To n)
    k = 0
    For i = 1 To n
        If IsNumeric(v(i, 28)) And Not IsEmpty(v(i, 28)) Then k = k + 1: idx(k) = i: pv(k) = v(i, 28)
    Next i
    SorteazaDesc idx, pv, k
    For j = 1 To k
        If j > 3 Then Exit For
        i = idx(j): r = r + 1
        ws.Range("A" & r & ":E" & r).Value = Array("'" & OraTxt(v(i, 2)), v(i, 3), "Scor corect " & v(i, 27), pv(j), CotaMin(pv(j)))
        ws.Range("D" & r).NumberFormat = "0%": ws.Range("E" & r).NumberFormat = "0.00"
        SalveazaBomba CStr(v(i, 3)), "Scor corect", "Scor corect " & v(i, 27), pv(j)
        AdaugaZi pePrg, OraTxt(v(i, 2)), CStr(v(i, 3)), "Bomba scor " & v(i, 27) & " (" & Format(pv(j), "0%") & ")", ""
    Next j

    ' ---------- 4. text pentru comunitate ----------
    r = r + 2
    ws.Range("A" & r).Value = "4. Programul zilei - tot ce recomanda botul, pe ore"
    ws.Range("A" & r).Font.Bold = True: ws.Range("A" & r).Font.Size = 12
    For Each ln In linii
        r = r + 1
        ws.Range("A" & r).Value = "'" & ln
        ws.Range("A" & r).Font.Bold = True
    Next ln
    r = r + 1
    ws.Range("A" & r & ":D" & r).Value = Array("Ora", "Meci", "Ce recomanda botul", "Pe biletele (cota)")
    AntetTabel ws.Range("A" & r & ":D" & r)
    ScrieProgram ws, r, pePrg
    r = r + 2
    ws.Range("A" & r).Value = "Procente calculate din istoric si verificate pe meciurile trecute. Nu sunt sigure. Joaca responsabil, doar cat iti permiti sa pierzi."
    On Error Resume Next
    Set wj = ThisWorkbook.Sheets("Jurnal")
    On Error GoTo Esec
    If Not wj Is Nothing Then Set wj = FoaieJurnal()
    If Not wj Is Nothing Then
        If IsNumeric(wj.Range("G18").Value) And IsNumeric(wj.Range("H18").Value) Then
            If wj.Range("G18").Value > 0 Then
                r = r + 1
                ws.Range("A" & r).Value = "Rezultatele botului pana acum (fara bombe): " & wj.Range("H18").Value & " din " & _
                    wj.Range("G18").Value & " recomandari iesite (" & Format(wj.Range("H18").Value / wj.Range("G18").Value, "0%") & ")."
            End If
        End If
    End If

    ws.Columns("A").ColumnWidth = 10
    ws.Columns("B").ColumnWidth = 34
    ws.Columns("C").ColumnWidth = 40
    ws.Columns("D").ColumnWidth = 30
    ws.Columns("E:G").ColumnWidth = 14
    ws.Range("E:G").HorizontalAlignment = xlCenter
    ws.Activate
    ws.Range("A1").Select

Final:
    GataLucru
    Exit Sub
Esec:
    Mesaj "A aparut o eroare la recomandari:" & vbLf & Err.Description, vbExclamation, "BET PRO"
    Resume Final
End Sub

' candidat pentru bilete: doar sanse de cel putin 55%
Private Sub AdaugaCand(cP() As Double, cN() As String, ByVal i As Long, ByVal j As Long, ByVal p As Variant, ByVal nume As String)
    If IsEmpty(p) Or Not IsNumeric(p) Then Exit Sub
    If CDbl(p) >= 0.55 And CDbl(p) < 0.99 Then
        cP(i, j) = CDbl(p)
        cN(i, j) = nume
    End If
End Sub

Private Sub SorteazaDescL(idx() As Long, pv() As Double, aux() As Long, ByVal n As Long)
    Dim a As Long, b As Long, ti As Long, tv As Double
    For a = 1 To n - 1
        For b = a + 1 To n
            If pv(b) > pv(a) Then
                tv = pv(a): pv(a) = pv(b): pv(b) = tv
                ti = idx(a): idx(a) = idx(b): idx(b) = ti
                ti = aux(a): aux(a) = aux(b): aux(b) = ti
            End If
        Next b
    Next a
End Sub

Private Sub AlegeMaxim(ByVal p As Variant, ByVal prag As Double, ByVal nume As String, ByRef bestP As Double, ByRef bestN As String)
    If IsEmpty(p) Or Not IsNumeric(p) Then Exit Sub
    If CDbl(p) >= prag And CDbl(p) > bestP Then
        bestP = CDbl(p)
        bestN = nume
    End If
End Sub

Private Sub SorteazaDesc(idx() As Long, pv() As Double, ByVal n As Long)
    Dim a As Long, b As Long, ti As Long, tv As Double
    For a = 1 To n - 1
        For b = a + 1 To n
            If pv(b) > pv(a) Then
                tv = pv(a): pv(a) = pv(b): pv(b) = tv
                ti = idx(a): idx(a) = idx(b): idx(b) = ti
            End If
        Next b
    Next a
End Sub

Private Sub SorteazaDescTxt(idx() As Long, pv() As Double, txt() As String, ByVal n As Long)
    Dim a As Long, b As Long, ti As Long, tv As Double, ts As String
    For a = 1 To n - 1
        For b = a + 1 To n
            If pv(b) > pv(a) Then
                tv = pv(a): pv(a) = pv(b): pv(b) = tv
                ti = idx(a): idx(a) = idx(b): idx(b) = ti
                ts = txt(a): txt(a) = txt(b): txt(b) = ts
            End If
        Next b
    Next a
End Sub

Private Function OraTxt(ByVal x As Variant) As String
    If IsNumeric(x) And Not IsEmpty(x) Then
        OraTxt = Format(CDbl(x), "hh:mm")
    Else
        OraTxt = CStr(x)
    End If
End Function


' ============================ AUTOMAT + EMAIL ============================
' RuleazaAutomat face singur tot: actualizeaza istoricul (verifica si Jurnalul),
' aduce meciurile zilei, face Recomandarile si le trimite pe email.
' Lunea ruleaza si Backtest (recalibrare). Se porneste zilnic din Windows Task Scheduler
' cu fisierul BetPro_Automat.vbs, sau manual cu butonul "Trimite email".

Private Sub Mesaj(ByVal t As String, Optional ByVal st As Long = 0, Optional ByVal ti As String = "BET PRO")
    If gTacut Then
        gLog = gLog & t & vbLf
    Else
        MsgBox t, st, ti
    End If
End Sub

Public Sub RuleazaAutomat()
    Dim ok As Boolean
    gTacut = True: gLog = ""
    On Error Resume Next
    Application.DisplayAlerts = False
    ActualizeazaIstoric
    If Weekday(Date, vbMonday) = 1 Then Backtest
    getMS
    Recomandari
    On Error GoTo 0
    ok = TrimiteEmailRecomandari()
    gTacut = False
    Application.DisplayAlerts = True
End Sub

Public Sub TrimiteEmailAcum()
    Dim ok As Boolean
    gTacut = True: gLog = ""
    On Error Resume Next
    Recomandari
    On Error GoTo 0
    ok = TrimiteEmailRecomandari()
    gTacut = False
    If ok Then
        MsgBox "Email trimis. Verifica si folderul Spam la primul email.", vbInformation, "BET PRO"
    Else
        MsgBox "Emailul nu a plecat:" & vbLf & vbLf & FoaieEmail().Range("B7").Value & vbLf & vbLf & _
            "Verifica datele din foaia Email.", vbExclamation, "BET PRO"
    End If
End Sub

Private Function FoaieEmail() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Email")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets("Listing"))
        ws.Name = "Email"
        ws.Range("A1").Value = "Setari email - recomandarile vin singure pe email"
        ws.Range("A1").Font.Size = 14: ws.Range("A1").Font.Bold = True
        ws.Range("A3").Value = "Trimite catre (adresa ta):"
        ws.Range("A4").Value = "Trimite de pe (contul Gmail):"
        ws.Range("A5").Value = "Parola de aplicatie Gmail (16 litere):"
        ws.Range("A7").Value = "Ultima trimitere:"
        ws.Range("B3:B5").Interior.Color = RGB(255, 242, 204)
        ws.Range("A9").Value = "Parola de aplicatie NU e parola ta de Gmail. O faci asa: myaccount.google.com > Securitate > " & _
            "activezi Verificarea in doi pasi > cauti 'Parole pentru aplicatii' > creezi una cu numele BET PRO > copiezi codul de 16 litere in B5."
        ws.Range("A10").Value = "Daca lasi B4 si B5 goale, botul incearca sa trimita prin Outlook (daca il ai instalat si configurat)."
        ws.Range("A11").Value = "Nu da fisierul altcuiva cu parola completata."
        ws.Columns("A").ColumnWidth = 36
        ws.Columns("B").ColumnWidth = 40
    End If
    Set FoaieEmail = ws
End Function

Private Function TrimiteEmailRecomandari() As Boolean
    Dim we As Worksheet, wr As Worksheet, wj As Worksheet, dest As String, user As String, pass As String
    Dim html As String, subiect As String, err1 As String

    Set we = FoaieEmail()
    dest = Trim$(CStr(we.Range("B3").Value))
    user = Trim$(CStr(we.Range("B4").Value))
    pass = Replace(Trim$(CStr(we.Range("B5").Value)), " ", "")
    If dest = "" Then
        we.Range("B7").Value = Format(Now, "dd.mm.yyyy hh:mm") & " - lipseste adresa din B3"
        Exit Function
    End If

    On Error Resume Next
    Set wr = ThisWorkbook.Sheets("Recomandari")
    Set wj = ThisWorkbook.Sheets("Jurnal")
    On Error GoTo 0
    subiect = "BET PRO - recomandari " & Format(Date, "dd.mm.yyyy")
    html = "<div style='font-family:Calibri,Arial,sans-serif;font-size:14px'>"
    If Not wr Is Nothing Then html = html & FoaieInHtml(wr, "A1:I" & wr.Cells(wr.Rows.Count, 1).End(xlUp).Row)
    If Not wj Is Nothing Then
        html = html & "<h3 style='margin-top:24px'>Cum a iesit botul pana acum (Jurnal)</h3>"
        html = html & FoaieInHtml(wj, "F3:J18")
    End If
    If gLog <> "" Then html = html & "<p style='color:#888;font-size:12px'>" & Replace(EscHtml(gLog), vbLf, "<br>") & "</p>"
    html = html & "<p style='color:#888;font-size:12px'>Trimis automat de BET PRO. Nicio selectie nu e sigura; joaca doar cat iti permiti sa pierzi.</p></div>"

    If user <> "" And pass <> "" Then
        err1 = TrimiteGmail(user, pass, dest, subiect, html)
    Else
        err1 = TrimiteOutlook(dest, subiect, html)
    End If

    If err1 = "" Then
        we.Range("B7").Value = Format(Now, "dd.mm.yyyy hh:mm") & " - trimis catre " & dest
        TrimiteEmailRecomandari = True
    Else
        we.Range("B7").Value = Format(Now, "dd.mm.yyyy hh:mm") & " - EROARE: " & err1
    End If
End Function

Private Function TrimiteGmail(ByVal user As String, ByVal pass As String, ByVal dest As String, _
        ByVal subiect As String, ByVal html As String) As String
    Dim m As Object, c As Object, sch As String
    On Error GoTo Esec
    sch = "http://schemas.microsoft.com/cdo/configuration/"
    Set m = CreateObject("CDO.Message")
    Set c = CreateObject("CDO.Configuration")
    c.Load -1
    With c.Fields
        .Item(sch & "sendusing") = 2
        .Item(sch & "smtpserver") = "smtp.gmail.com"
        .Item(sch & "smtpserverport") = 465
        .Item(sch & "smtpusessl") = True
        .Item(sch & "smtpauthenticate") = 1
        .Item(sch & "sendusername") = user
        .Item(sch & "sendpassword") = pass
        .Item(sch & "smtpconnectiontimeout") = 60
        .Update
    End With
    Set m.Configuration = c
    m.From = "BET PRO <" & user & ">"
    m.To = dest
    m.Subject = subiect
    m.HTMLBody = html
    m.Send
    Exit Function
Esec:
    TrimiteGmail = "Gmail: " & Err.Description & " (verifica adresa din B4 si parola de aplicatie din B5)"
End Function

Private Function TrimiteOutlook(ByVal dest As String, ByVal subiect As String, ByVal html As String) As String
    Dim o As Object, it As Object
    On Error GoTo Esec
    Set o = CreateObject("Outlook.Application")
    Set it = o.CreateItem(0)
    it.To = dest
    it.Subject = subiect
    it.HTMLBody = html
    it.Send
    Exit Function
Esec:
    TrimiteOutlook = "Outlook nu e disponibil (" & Err.Description & "). Completeaza Gmail in B4 si B5."
End Function

' transforma un tabel din foaie in HTML simplu, cu culorile celulelor
Private Function FoaieInHtml(ws As Worksheet, ByVal adresa As String) As String
    Dim rg As Range, rw As Range, cl As Range, t As String, linie As String, randGol As Boolean, bg As String
    Set rg = ws.Range(adresa)
    t = "<table style='border-collapse:collapse;margin:4px 0'>"
    For Each rw In rg.Rows
        linie = "": randGol = True
        For Each cl In rw.Cells
            If CStr(cl.Text) <> "" Then randGol = False
            bg = ""
            If cl.Interior.ColorIndex <> xlColorIndexNone Then bg = "background:" & CuloareHtml(cl.Interior.Color) & ";"
            linie = linie & "<td style='padding:3px 8px;vertical-align:top;" & bg & IIf(cl.Font.Bold, "font-weight:bold;", "") & _
                IIf(cl.Font.Color = RGB(255, 255, 255), "color:#fff;", "") & "'>" & EscHtml(cl.Text) & "</td>"
        Next cl
        If Not randGol Then t = t & "<tr>" & linie & "</tr>"
    Next rw
    FoaieInHtml = t & "</table>"
End Function

Private Function CuloareHtml(ByVal c As Long) As String
    CuloareHtml = "#" & Right$("0" & Hex(c Mod 256), 2) & Right$("0" & Hex((c \ 256) Mod 256), 2) & Right$("0" & Hex((c \ 65536) Mod 256), 2)
End Function

Private Function EscHtml(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    EscHtml = Replace(s, vbLf, "<br>")
End Function


' ============================ ISTORIC BILETE ============================
' Foaia "Bilete": fiecare bilet si fiecare bomba recomandate, verificate automat dupa meci.
'  A Data | B Tip (Bilet / Bomba) | C Eticheta (cota tinta sau HT/FT / Scor corect) | D Cota corecta
'  E Sansa | F Descriere | G Rezultat | H Verificat la | I.. pentru fiecare meci: Gazde, Oaspeti, Piata
Private Function FoaieBilete() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Bilete")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "Bilete"
        ws.Range("A1:H1").Value = Array("Data", "Tip", "Bilet", "Cota corecta", "Sansa", "Meciuri", "Rezultat", "Verificat la")
        AntetTabel ws.Range("A1:H1")
        ws.Columns("A").ColumnWidth = 11
        ws.Columns("B:C").ColumnWidth = 12
        ws.Columns("D:E").ColumnWidth = 12
        ws.Columns("F").ColumnWidth = 90
        ws.Columns("G").ColumnWidth = 14
        ws.Columns("H").ColumnWidth = 16
    End If
    Set FoaieBilete = ws
End Function

' salveaza un bilet o singura data pe zi (primul trimis pe email ramane)
Private Sub SalveazaBilet(ByVal tip As String, ByVal eticheta As String, ByVal cotaC As Double, ByVal sansa As Double, _
        ByVal descr As String, legH() As String, legA() As String, legP() As String, ByVal nL As Long)
    Dim ws As Worksheet, last As Long, r As Long, k As Long, v As Variant
    Set ws = FoaieBilete()
    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If last >= 2 Then
        v = ws.Range("A2:F" & last).Value2
        For r = 1 To UBound(v, 1)
            If IsNumeric(v(r, 1)) And Not IsEmpty(v(r, 1)) Then
                If CLng(v(r, 1)) = CLng(Date) And CStr(v(r, 2)) = tip And CStr(v(r, 3)) = eticheta Then
                    If tip = "Bilet" Or CStr(v(r, 6)) = descr Then Exit Sub
                End If
            End If
        Next r
    End If
    r = last + 1
    ws.Range("A" & r).Value = Date
    ws.Range("A" & r).NumberFormat = "dd.mm.yyyy"
    ws.Range("B" & r).Value = tip
    ws.Range("C" & r).Value = "'" & eticheta
    ws.Range("D" & r).Value = cotaC
    ws.Range("D" & r).NumberFormat = "0.00"
    ws.Range("E" & r).Value = sansa
    ws.Range("E" & r).NumberFormat = "0%"
    ws.Range("F" & r).Value = descr
    ws.Range("G" & r).Value = "In asteptare"
    For k = 1 To nL
        ws.Cells(r, 9 + (k - 1) * 3).Value = legH(k)
        ws.Cells(r, 10 + (k - 1) * 3).Value = legA(k)
        ws.Cells(r, 11 + (k - 1) * 3).Value = legP(k)
    Next k
End Sub

' rezultatul unei piete pentru un meci din Istoric: 1 = a iesit, 0 = nu, -1 = nu se poate verifica
Private Function EvalPiata(ByVal piata As String, H As Variant, ByVal z As Long) As Long
    Dim fh As Double, fa As Double, hh As Double, ha As Double, ok As Boolean, et3 As Variant
    Dim hasHT As Boolean, hasC As Boolean, hasY As Boolean
    et3 = Array("1", "X", "2")
    fh = H(z, 8): fa = H(z, 9)
    hasHT = Not Gol(H(z, 6)) And Not Gol(H(z, 7))
    hasC = Not Gol(H(z, 10)) And Not Gol(H(z, 11))
    hasY = Not Gol(H(z, 17)) And Not Gol(H(z, 18))
    If hasHT Then hh = H(z, 6): ha = H(z, 7)
    EvalPiata = -1
    If piata = "Over 1.5" Then
        ok = (fh + fa >= 2)
    ElseIf piata = "Over 2.5" Then
        ok = (fh + fa >= 3)
    ElseIf piata = "GG" Then
        ok = (fh >= 1 And fa >= 1)
    ElseIf piata = "Gazde peste 1.5" Then
        ok = (fh >= 2)
    ElseIf piata = "Oaspeti peste 1.5" Then
        ok = (fa >= 2)
    ElseIf piata = "Over 0.5 pauza" Then
        If Not hasHT Then Exit Function
        ok = (hh + ha >= 1)
    ElseIf Left$(piata, 9) = "Cornere O" Then
        If Not hasC Then Exit Function
        ok = (H(z, 10) + H(z, 11) > Val(Mid$(piata, 10)))
    ElseIf Left$(piata, 9) = "Galbene O" Then
        If Not hasY Then Exit Function
        ok = (H(z, 17) + H(z, 18) > Val(Mid$(piata, 10)))
    ElseIf Left$(piata, 6) = "HT/FT " Then
        If Not hasHT Then Exit Function
        ok = (Mid$(piata, 7, 3) = et3(Rez(CLng(hh), CLng(ha))) & "/" & et3(Rez(CLng(fh), CLng(fa))))
    ElseIf Left$(piata, 12) = "Scor corect " Then
        ok = (Trim$(Mid$(piata, 13)) = CLng(fh) & "-" & CLng(fa))
    Else
        Exit Function
    End If
    EvalPiata = IIf(ok, 1, 0)
End Function

' verifica biletele in asteptare cu rezultatele din Istoric
Private Function VerificaBilete() As Long
    Dim ws As Worksheet, hist As Worksheet, last As Long, lastH As Long, H As Variant, idx As Object
    Dim i As Long, r As Long, k As Long, z As Long, d As Long, key As String, gazde As String, oaspeti As String
    Dim rezL As Long, toate As Boolean, pierdut As Boolean, zi As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Bilete")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If last < 2 Then Exit Function
    Set hist = FoaieIstoric()
    lastH = hist.Cells(hist.Rows.Count, 1).End(xlUp).Row
    If lastH < 3 Then Exit Function
    H = hist.Range("A2:U" & lastH).Value2
    Set idx = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(H, 1)
        If IsNumeric(H(i, 3)) Then
            key = Trim$(CStr(H(i, 4))) & "|" & Trim$(CStr(H(i, 5))) & "|" & CLng(Int(H(i, 3)))
            If Not idx.Exists(key) Then idx.Add key, i
        End If
    Next i

    For r = 2 To last
        If CStr(ws.Range("G" & r).Value) = "In asteptare" And IsNumeric(ws.Range("A" & r).Value2) Then
            zi = CLng(ws.Range("A" & r).Value2)
            toate = True: pierdut = False
            k = 0
            Do While CStr(ws.Cells(r, 9 + k * 3).Value) <> ""
                gazde = CStr(ws.Cells(r, 9 + k * 3).Value)
                oaspeti = CStr(ws.Cells(r, 10 + k * 3).Value)
                z = 0
                For d = 0 To 2
                    key = gazde & "|" & oaspeti & "|" & (zi + d)
                    If idx.Exists(key) Then z = idx(key): Exit For
                Next d
                If z = 0 Then
                    toate = False
                Else
                    rezL = EvalPiata(CStr(ws.Cells(r, 11 + k * 3).Value), H, z)
                    If rezL = 0 Then
                        pierdut = True
                        ws.Cells(r, 11 + k * 3).Interior.Color = RGB(255, 120, 120)
                    ElseIf rezL = 1 Then
                        ws.Cells(r, 11 + k * 3).Interior.Color = RGB(146, 208, 80)
                    Else
                        toate = False
                    End If
                End If
                k = k + 1
            Loop
            If pierdut Then
                ws.Range("G" & r).Value = "PIERDUT"
                ws.Range("G" & r).Interior.Color = RGB(255, 120, 120)
                ws.Range("H" & r).Value = Now
                VerificaBilete = VerificaBilete + 1
            ElseIf toate And k > 0 Then
                ws.Range("G" & r).Value = "CASTIGAT"
                ws.Range("G" & r).Interior.Color = RGB(146, 208, 80)
                ws.Range("H" & r).Value = Now
                VerificaBilete = VerificaBilete + 1
            ElseIf zi < CLng(Date) - 10 Then
                ' meci care nu apare in istoric dupa 10 zile (amanat sau nume diferit)
                ws.Range("G" & r).Value = "Nu se poate verifica"
                ws.Range("G" & r).Interior.Color = RGB(217, 217, 217)
            End If
            ws.Range("H" & r).NumberFormat = "dd.mm.yyyy hh:mm"
        End If
    Next r
End Function

' scrie in foaia Recomandari rezultatele biletelor din ultimele 7 zile si totalurile
Private Sub RaportBilete(ws As Worksheet, ByRef r As Long, linii As Collection)
    Dim wb As Worksheet, last As Long, v As Variant, i As Long, nr As Long, tinte As Variant, t As Variant
    Dim nV As Long, nC As Long, st As String, ieri As String
    On Error Resume Next
    Set wb = ThisWorkbook.Sheets("Bilete")
    On Error GoTo 0
    If wb Is Nothing Then Exit Sub
    last = wb.Cells(wb.Rows.Count, 1).End(xlUp).Row
    If last < 2 Then Exit Sub
    v = wb.Range("A2:G" & last).Value2

    ws.Range("A" & r).Value = "0. Cum au iesit biletele si bombele (ultimele 7 zile)"
    ws.Range("A" & r).Font.Bold = True: ws.Range("A" & r).Font.Size = 12
    r = r + 1
    ws.Range("A" & r & ":E" & r).Value = Array("Data", "Bilet", "Cota corecta", "Meciuri", "Rezultat")
    AntetTabel ws.Range("A" & r & ":E" & r)
    For i = UBound(v, 1) To 1 Step -1
        If IsNumeric(v(i, 1)) And Not IsEmpty(v(i, 1)) Then
            If CLng(v(i, 1)) < CLng(Date) And CLng(v(i, 1)) >= CLng(Date) - 7 Then
                r = r + 1: nr = nr + 1
                st = CStr(v(i, 7))
                ws.Range("A" & r).Value = "'" & Format(v(i, 1), "dd.mm")
                ws.Range("B" & r).Value = IIf(CStr(v(i, 2)) = "Bilet", "Cota " & v(i, 3), "Bomba " & v(i, 3))
                ws.Range("C" & r).Value = v(i, 4)
                ws.Range("C" & r).NumberFormat = "0.00"
                ws.Range("D" & r).Value = v(i, 6)
                ws.Range("E" & r).Value = IIf(st = "In asteptare", "in asteptare", st)
                If st = "CASTIGAT" Then ws.Range("E" & r).Interior.Color = RGB(146, 208, 80)
                If st = "PIERDUT" Then ws.Range("E" & r).Interior.Color = RGB(255, 120, 120)
                If CLng(v(i, 1)) = CLng(Date) - 1 And (st = "CASTIGAT" Or st = "PIERDUT") Then
                    ieri = ieri & IIf(ieri = "", "", ", ") & ws.Range("B" & r).Value & " " & st
                End If
            End If
        End If
    Next i
    If nr = 0 Then
        r = r + 1
        ws.Range("A" & r).Value = "Inca nu sunt bilete din zilele trecute."
    End If
    r = r + 1
    ws.Range("A" & r).Value = "Rezultatele apar pe site de obicei la 1-3 zile dupa meci, pana atunci biletul ramane 'in asteptare'."
    ws.Range("A" & r).Font.Italic = True

    ' totaluri de la inceput
    r = r + 2
    ws.Range("A" & r & ":D" & r).Value = Array("Total de la inceput", "Verificate", "Castigate", "Reusita")
    AntetTabel ws.Range("A" & r & ":D" & r)
    tinte = Array("2", "3.5", "5", "7.5", "10", "15", "20", "HT/FT", "Scor corect")
    For Each t In tinte
        nV = 0: nC = 0
        For i = 1 To UBound(v, 1)
            If CStr(v(i, 3)) = CStr(t) And (CStr(v(i, 7)) = "CASTIGAT" Or CStr(v(i, 7)) = "PIERDUT") Then
                nV = nV + 1
                If CStr(v(i, 7)) = "CASTIGAT" Then nC = nC + 1
            End If
        Next i
        r = r + 1
        ws.Range("A" & r).Value = IIf(t = "HT/FT" Or t = "Scor corect", "Bombe " & t, "Bilet cota " & t)
        ws.Range("B" & r).Value = nV
        ws.Range("C" & r).Value = nC
        If nV > 0 Then
            ws.Range("D" & r).Value = nC / nV
            ws.Range("D" & r).NumberFormat = "0%"
        Else
            ws.Range("D" & r).Value = "-"
        End If
    Next t
    r = r + 2
    If ieri <> "" Then linii.Add "Ieri: " & ieri
End Sub

Private Sub SalveazaBomba(ByVal meci As String, ByVal eticheta As String, ByVal piata As String, ByVal p As Double)
    Dim legH(1 To 1) As String, legA(1 To 1) As String, legP(1 To 1) As String, mg As String, mo As String
    If p <= 0 Then Exit Sub
    ImparteMeci meci, mg, mo
    legH(1) = mg: legA(1) = mo: legP(1) = piata
    SalveazaBilet "Bomba", eticheta, 1 / p, p, meci & " - " & piata, legH, legA, legP, 1
End Sub

' "Gazde - Oaspeti" -> gazde, oaspeti
Private Sub ImparteMeci(ByVal meci As String, ByRef gazde As String, ByRef oaspeti As String)
    Dim k As Long
    k = InStr(meci, " - ")
    If k > 0 Then
        gazde = Trim$(Left$(meci, k - 1))
        oaspeti = Trim$(Mid$(meci, k + 3))
    Else
        gazde = meci: oaspeti = ""
    End If
End Sub


' adauga o recomandare la programul zilei (un rand pe meci)
Private Sub AdaugaZi(d As Object, ByVal ora As String, ByVal meci As String, ByVal pick As String, ByVal bilet As String)
    Dim a As Variant
    If Not d.Exists(meci) Then d.Add meci, Array(ora, "", "")
    a = d(meci)
    If pick <> "" And InStr(a(1), pick) = 0 Then a(1) = a(1) & IIf(a(1) = "", "", vbLf) & pick
    If bilet <> "" And InStr(a(2), bilet) = 0 Then a(2) = a(2) & IIf(a(2) = "", "", vbLf) & bilet
    d(meci) = a
End Sub

' scrie programul zilei sortat dupa ora
Private Sub ScrieProgram(ws As Worksheet, ByRef r As Long, d As Object)
    Dim chei As Variant, n As Long, i As Long, j As Long, t As Variant, a As Variant, r0 As Long
    Dim ai As Variant, aj As Variant
    If d.Count = 0 Then
        r = r + 1
        ws.Range("A" & r).Value = "Azi nu e nimic recomandat."
        Exit Sub
    End If
    chei = d.Keys
    n = d.Count
    For i = 0 To n - 2
        For j = i + 1 To n - 1
            ai = d(chei(i)): aj = d(chei(j))
            If aj(0) & chei(j) < ai(0) & chei(i) Then
                t = chei(i): chei(i) = chei(j): chei(j) = t
            End If
        Next j
    Next i
    r0 = r + 1
    For i = 0 To n - 1
        a = d(chei(i))
        r = r + 1
        ws.Range("A" & r).Value = "'" & a(0)
        ws.Range("B" & r).Value = chei(i)
        ws.Range("C" & r).Value = IIf(a(1) = "", "-", a(1))
        ws.Range("D" & r).Value = IIf(a(2) = "", "-", a(2))
        If i Mod 2 = 1 Then ws.Range("A" & r & ":D" & r).Interior.Color = RGB(242, 245, 250)
    Next i
    With ws.Range("A" & r0 & ":D" & r)
        .WrapText = True
        .VerticalAlignment = xlTop
        .Borders.Color = RGB(200, 205, 215)
    End With
    ws.Range("A" & r0 & ":A" & r).Font.Bold = True
    ws.Range("A" & r0 & ":D" & r).Rows.AutoFit
End Sub


' ============================ PIATA ============================
' Botul "asculta" si de piata: golurile asteptate se muta o parte spre ce spun cotele
' (Over/Under 2.5 pentru totalul de goluri, 1X2 pentru cine e favorit).
' Greutatea pietei e in Scanare, celula AP21 (implicit 50%).

Private Function GreutatePiata() As Double
    Dim ws As Worksheet, v As Variant
    GreutatePiata = 0.5
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Scanare")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    v = ws.Range("AP21").Value2
    If IsNumeric(v) And Not IsEmpty(v) Then
        If v > 1 Then v = v / 100
        If v >= 0 And v <= 1 Then GreutatePiata = CDbl(v)
    End If
End Function

' sansele 1X2 dintr-un model Poisson
Private Sub P1X2(ByVal lH As Double, ByVal lA As Double, ByRef ph As Double, ByRef pd As Double, ByRef pa As Double)
    Dim a As Long, b As Long, eh(0 To 9) As Double, ea(0 To 9) As Double
    ph = 0: pd = 0: pa = 0
    If lH < 0.01 Then lH = 0.01
    If lA < 0.01 Then lA = 0.01
    eh(0) = Exp(-lH): ea(0) = Exp(-lA)
    For a = 1 To 9
        eh(a) = eh(a - 1) * lH / a
        ea(a) = ea(a - 1) * lA / a
    Next a
    For a = 0 To 9
        For b = 0 To 9
            If a > b Then
                ph = ph + eh(a) * ea(b)
            ElseIf a = b Then
                pd = pd + eh(a) * ea(b)
            Else
                pa = pa + eh(a) * ea(b)
            End If
        Next b
    Next a
End Sub

Private Function NumarCota(ByVal x As Variant) As Double
    If IsEmpty(x) Or IsNull(x) Then Exit Function
    If IsNumeric(x) Then NumarCota = CDbl(x)
End Function

Private Sub AjusteazaLaPiata(ByRef lH As Double, ByRef lA As Double, ByRef lHTh As Double, ByRef lHTa As Double, _
        ByVal oH As Double, ByVal oD As Double, ByVal oA As Double, ByVal oO As Double, ByVal oU As Double, _
        ByVal w As Double)
    Dim T As Double, Tn As Double, pm As Double, tgt As Double, lo As Double, hi As Double, mj As Double
    Dim lH1 As Double, lA1 As Double, ph As Double, pd As Double, pa As Double, s As Double, it As Long
    If w <= 0 Or lH <= 0 Or lA <= 0 Then Exit Sub
    T = lH + lA
    Tn = T
    ' totalul de goluri spre piata Over/Under 2.5
    If oO > 1 And oU > 1 Then
        pm = (1 / oO) / ((1 / oO) + (1 / oU))
        tgt = (1 - w) * PoisPeste(2, T) + w * pm
        lo = 0.3: hi = 7
        For it = 1 To 30
            mj = (lo + hi) / 2
            If PoisPeste(2, mj) < tgt Then lo = mj Else hi = mj
        Next it
        Tn = (lo + hi) / 2
    End If
    lH1 = lH * Tn / T: lA1 = lA * Tn / T
    ' cine e favorit spre piata 1X2
    If oH > 1 And oD > 1 And oA > 1 Then
        s = 1 / oH + 1 / oD + 1 / oA
        pm = (1 / oH) / s
        P1X2 lH1, lA1, ph, pd, pa
        tgt = (1 - w) * ph + w * pm
        lo = -Tn * 0.95: hi = Tn * 0.95
        For it = 1 To 25
            mj = (lo + hi) / 2
            P1X2 (Tn + mj) / 2, (Tn - mj) / 2, ph, pd, pa
            If ph < tgt Then lo = mj Else hi = mj
        Next it
        mj = (lo + hi) / 2
        lH1 = (Tn + mj) / 2: lA1 = (Tn - mj) / 2
    End If
    lHTh = lHTh * lH1 / lH
    lHTa = lHTa * lA1 / lA
    lH = lH1: lA = lA1
End Sub

' HT/FT nu poate fi mai probabil decat ~70% din sansa rezultatului final
Private Function LimitaHTFT(ByVal p As Double, ByVal eticheta As String, ByVal lH As Double, ByVal lA As Double) As Double
    Dim ph As Double, pd As Double, pa As Double, fin As String, lim As Double
    P1X2 lH, lA, ph, pd, pa
    fin = Right$(eticheta, 1)
    Select Case fin
        Case "1": lim = 0.7 * ph
        Case "2": lim = 0.7 * pa
        Case Else: lim = 0.7 * pd
    End Select
    If p > lim Then LimitaHTFT = lim Else LimitaHTFT = p
End Function


' ============================ LIVE ============================
' Butonul LIVE: o singura cerere la API-Football (meciurile de azi cu scor si minut),
' apoi recalculeaza sansele live pentru meciurile scanate azi si arata cum stau biletele.
' "LIVE auto" repeta la fiecare 5 minute cat timp Excel e deschis (cam 12 cereri pe ora,
' planul gratuit are 100 pe zi). Cheia API e in foaia Email, celula B13.

Public Sub LiveStart()
    gLiveAuto = True
    LiveActualizeaza
End Sub

Public Sub LiveStop()
    gLiveAuto = False
    On Error Resume Next
    If gLiveNext > 0 Then Application.OnTime gLiveNext, "LiveAuto", , False
    On Error GoTo 0
    gLiveNext = 0
    FoaieLive().Range("A2").Value = "LIVE auto oprit la " & Format(Now, "hh:mm") & "."
End Sub

Public Sub LiveAuto()
    If gLiveAuto Then LiveActualizeaza
End Sub

Public Sub LiveActualizeaza()
    Dim ws As Worksheet, res As Object, fx As Variant, cheie As String, wl As Worksheet
    Dim lam As Object, lastL As Long, L As Variant, i As Long, r As Long, nLive As Long, ramase As String
    Dim home As String, away As String, f As Variant, st As String, mnt As Long, gh As Long, ga As Long
    Dim hth As Variant, hta As Variant, a As Variant, ps As Variant, pl As Long, deJucat As Long

    On Error GoTo Esec
    StartLucru
    Set ws = FoaieLive()
    cheie = CheieApi()
    If cheie = "" Then
        ws.Range("A2").Value = "Lipseste cheia API-Football in foaia Email, celula B13."
        GoTo Final
    End If
    Application.StatusBar = "BET PRO LIVE: descarc scorurile de azi..."
    DoEvents
    If Not MeciuriAziApi(cheie, res, ramase) Then
        ws.Range("A2").Value = "Nu am putut lua datele live: " & mEroare
        GoTo Final
    End If

    If mCornHT Is Nothing Then Set mCornHT = CreateObject("Scripting.Dictionary")
    Application.StatusBar = "BET PRO LIVE: pregatesc datele..."
    DoEvents
    PregatesteLive res
    ' lambdas salvate la scanarea de azi
    Set lam = IncarcaLambda()
    ws.Range("A4:P400").Clear

    ' ---------- prima repriza: toate meciurile scanate azi, in ordinea orei ----------
    Dim mOra() As String, mGz() As String, mOp() As String, mF() As Variant, mA() As Variant, nM As Long, j As Long
    Dim ids As String, nIds As Long, res2 As Object, corn As Object, idF As String, cr As Long, lin As Double
    Set wl = Sheets("Listing")
    lastL = wl.Cells(wl.Rows.Count, "E").End(xlUp).Row
    If lastL >= 2 Then
        L = wl.Range("A1:V" & lastL).Value2
        ReDim mOra(1 To lastL): ReDim mGz(1 To lastL): ReDim mOp(1 To lastL): ReDim mF(1 To lastL): ReDim mA(1 To lastL)
    End If
    For i = 2 To lastL
        home = Trim$(CStr(L(i, 3))): away = Trim$(CStr(L(i, 4)))
        If lam.Exists(home & "|" & away) Then
            nM = nM + 1
            mOra(nM) = Right$(CStr(L(i, 2)), 5)
            mGz(nM) = home: mOp(nM) = away
            mA(nM) = lam(home & "|" & away)
            Set mF(nM) = GasesteMeciLive(res, home, away)
            If Not mF(nM) Is Nothing Then
                st = JTxt(JGet(mF(nM), "fixture", "status", "short"))
                If (st = "1H" Or st = "HT") And nIds < 60 Then
                    ids = ids & IIf(ids = "", "", "-") & JTxt(JGet(mF(nM), "fixture", "id"))
                    nIds = nIds + 1
                End If
            End If
        End If
    Next i

    ' cornerele live (o cerere in plus, doar cand sunt meciuri in prima repriza)
    Set corn = CreateObject("Scripting.Dictionary")
    If nIds > 0 Then
        Application.StatusBar = "BET PRO LIVE: iau cornerele..."
        DoEvents
        Dim idL As Variant, grp As String, gi As Long
        idL = Split(ids, "-")
        For gi = 0 To UBound(idL)
            grp = grp & IIf(grp = "", "", "-") & idL(gi)
            If (gi + 1) Mod 20 = 0 Or gi = UBound(idL) Then
                If ApiFootball("/fixtures?ids=" & grp, cheie, res2, ramase) Then CitesteCornere res2, corn
                grp = ""
            End If
        Next gi
    End If

    ws.Range("A4").Value = "Prima repriza (HT) - toate meciurile, in ordinea orei"
    ws.Range("A4").Font.Bold = True: ws.Range("A4").Font.Size = 12
    ws.Range("A5:F5").Value = Array("Ora / minut", "Meci", "Over 0.5 HT", "Over 1.5 HT", "Over 2.5 HT", "Over cornere HT")
    AntetTabel ws.Range("A5:F5")
    r = 5
    For j = 1 To nM
        r = r + 1
        If mF(j) Is Nothing Then
            ws.Range("A" & r).Value = "'" & mOra(j)
            ws.Range("B" & r).Value = mGz(j) & " - " & mOp(j)
            ws.Range("C" & r).Value = "negasit live"
            ws.Range("C" & r & ":F" & r).Font.Color = RGB(150, 150, 150)
        Else
            StareMeci mF(j), st, mnt, gh, ga, hth, hta
            If st = "NS" Or st = "TBD" Then
                ws.Range("A" & r).Value = "'" & mOra(j)
                ws.Range("B" & r).Value = mGz(j) & " - " & mOp(j)
            Else
                ws.Range("A" & r).Value = "'" & IIf(st = "1H" Or st = "2H", mnt & "'", TextStare(st, mnt))
                ws.Range("B" & r).Value = mGz(j) & "  " & gh & "-" & ga & "  " & mOp(j)
            End If
            If st = "NS" Or st = "1H" Then deJucat = deJucat + 1
            If EsteLive(st) Then nLive = nLive + 1
            ' goluri in prima repriza: 0.5 / 1.5 / 2.5
            ScrieHT ws, "C", r, SansaHTLinie(0, mA(j), st, mnt, gh, ga, hth, hta)
            ScrieHT ws, "D", r, SansaHTLinie(1, mA(j), st, mnt, gh, ga, hth, hta)
            ScrieHT ws, "E", r, SansaHTLinie(2, mA(j), st, mnt, gh, ga, hth, hta)
            ' cornere in prima repriza
            idF = JTxt(JGet(mF(j), "fixture", "id"))
            cr = -1
            If corn.Exists(idF) Then cr = corn(idF)
            If st = "HT" And cr >= 0 Then mCornHT(idF) = cr
            If cr < 0 And mCornHT.Exists(idF) Then cr = mCornHT(idF)
            ScrieCornereHT ws, r, mA(j)(4), st, mnt, cr, mCornHT.Exists(idF)
            If st = "1H" Then ws.Range("A" & r & ":F" & r).Interior.Color = RGB(255, 247, 214)
            If st <> "NS" And st <> "1H" And st <> "HT" Then ws.Range("A" & r & ":B" & r).Font.Color = RGB(130, 130, 130)
        End If
    Next j
    If nM = 0 Then
        r = r + 1
        ws.Range("A" & r).Value = "Nu sunt meciuri scanate azi. Apasa intai Meciurile zilei si Recomandari."
    End If
    ws.Range("A5:F" & r).Borders.Color = RGB(200, 205, 215)
    ws.Range("C6:F" & r).HorizontalAlignment = xlCenter
    ws.Range("A6:A" & r).HorizontalAlignment = xlCenter

    ' ---------- biletele de azi ----------
    r = r + 2
    BileteLive ws, r, res, lam

    ws.Range("A2").Value = "Actualizat la " & Format(Now, "hh:mm") & "  -  meciuri in desfasurare: " & nLive & _
        "  -  cereri API ramase azi: " & IIf(ramase = "", "?", ramase) & _
        IIf(gLiveAuto, "  -  LIVE auto: urmatoarea actualizare peste 2 minute", "")
    ws.Columns("A").ColumnWidth = 8
    ws.Columns("B").ColumnWidth = 34
    ws.Columns("A").ColumnWidth = 11
    ws.Columns("C:F").ColumnWidth = 16
    ws.Activate
    ws.Range("A1").Select

    ' oprire automata cand nu mai e nimic de urmarit
    If gLiveAuto And deJucat = 0 And nLive = 0 Then
        gLiveAuto = False
        ws.Range("A2").Value = ws.Range("A2").Value & "  -  toate meciurile s-au terminat, LIVE auto oprit."
    End If

Final:
    If gLiveAuto Then
        gLiveNext = Now + TimeSerial(0, 2, 0)
        On Error Resume Next
        Application.OnTime gLiveNext, "LiveAuto"
        On Error GoTo 0
    End If
    GataLucru
    Exit Sub
Esec:
    Mesaj "A aparut o eroare la LIVE:" & vbLf & Err.Description, vbExclamation, "BET PRO"
    Resume Final
End Sub

' sansele live: (0) inca un gol, (1) inca 2 goluri, (2) GG, (3)(4)(5) 1/X/2 final, (6) gol pana la pauza
Private Function SanseLive(ByVal lH As Double, ByVal lA As Double, ByVal lHTh As Double, ByVal lHTa As Double, _
        ByVal st As String, ByVal mnt As Long, ByVal gh As Long, ByVal ga As Long) As Variant
    Dim rH As Double, rA As Double, rHTh As Double, rHTa As Double, out(0 To 6) As Double
    Dim a As Long, b As Long, eh(0 To 9) As Double, ea(0 To 9) As Double, q As Double
    RamaseLive lH, lA, lHTh, lHTa, st, mnt, rH, rA, rHTh, rHTa
    out(0) = PoisPeste(0, rH + rA)
    out(1) = PoisPeste(1, rH + rA)
    out(2) = IIf(gh > 0, 1, 1 - Exp(-rH)) * IIf(ga > 0, 1, 1 - Exp(-rA))
    eh(0) = Exp(-rH): ea(0) = Exp(-rA)
    For a = 1 To 9
        eh(a) = eh(a - 1) * rH / a: ea(a) = ea(a - 1) * rA / a
    Next a
    For a = 0 To 9
        For b = 0 To 9
            q = eh(a) * ea(b)
            If gh + a > ga + b Then
                out(3) = out(3) + q
            ElseIf gh + a = ga + b Then
                out(4) = out(4) + q
            Else
                out(5) = out(5) + q
            End If
        Next b
    Next a
    out(6) = PoisPeste(0, rHTh + rHTa)
    SanseLive = out
End Function

' golurile asteptate care mai raman (pe tot meciul si pana la pauza)
Private Sub RamaseLive(ByVal lH As Double, ByVal lA As Double, ByVal lHTh As Double, ByVal lHTa As Double, _
        ByVal st As String, ByVal mnt As Long, ByRef rH As Double, ByRef rA As Double, ByRef rHTh As Double, ByRef rHTa As Double)
    Dim h2 As Double, a2 As Double, t As Double
    h2 = lH - lHTh: a2 = lA - lHTa
    If h2 < 0.05 Then h2 = 0.05
    If a2 < 0.05 Then a2 = 0.05
    t = mnt
    If st = "NS" Then
        rHTh = lHTh: rHTa = lHTa: rH = lH: rA = lA
    ElseIf st = "1H" Then
        If t > 45 Then t = 45
        rHTh = lHTh * (45 - t) / 45 + 0.05 * lHTh: rHTa = lHTa * (45 - t) / 45 + 0.05 * lHTa
        rH = rHTh + h2: rA = rHTa + a2
    ElseIf st = "HT" Then
        rHTh = 0: rHTa = 0: rH = h2: rA = a2
    ElseIf st = "2H" Or st = "LIVE" Or st = "BT" Then
        If t < 45 Then t = 45
        If t > 90 Then t = 90
        rHTh = 0: rHTa = 0
        rH = h2 * (90 - t) / 45 + 0.04 * h2: rA = a2 * (90 - t) / 45 + 0.04 * a2
    Else
        rHTh = 0: rHTa = 0: rH = 0.0001: rA = 0.0001
    End If
End Sub

Private Function EsteLive(ByVal st As String) As Boolean
    EsteLive = (st = "1H" Or st = "HT" Or st = "2H" Or st = "LIVE" Or st = "BT" Or st = "ET" Or st = "P" Or st = "INT")
End Function

Private Function TextStare(ByVal st As String, ByVal mnt As Long) As String
    Select Case st
        Case "NS", "TBD": TextStare = "neinceput"
        Case "1H", "2H", "LIVE", "ET": TextStare = "min " & mnt
        Case "HT": TextStare = "pauza"
        Case "FT", "AET", "PEN": TextStare = "final"
        Case "PST": TextStare = "amanat"
        Case "CANC", "ABD", "AWD", "WO": TextStare = "anulat"
        Case Else: TextStare = st
    End Select
End Function

Private Sub StareMeci(f As Variant, ByRef st As String, ByRef mnt As Long, ByRef gh As Long, ByRef ga As Long, _
        ByRef hth As Variant, ByRef hta As Variant)
    st = JTxt(JGet(f, "fixture", "status", "short"))
    mnt = CLng(JNumar(JGet(f, "fixture", "status", "elapsed")))
    gh = CLng(JNumar(JGet(f, "goals", "home")))
    ga = CLng(JNumar(JGet(f, "goals", "away")))
    hth = JGet(f, "score", "halftime", "home")
    hta = JGet(f, "score", "halftime", "away")
    If IsNull(hth) Or IsNull(hta) Then hth = Empty: hta = Empty
    If st = "1H" Then hth = Empty: hta = Empty
End Sub

' scrie sansa si cota minima; 1 = IESIT, 0 = PIERDUT, -1 = fara date
Private Sub ScrieSansa(ws As Worksheet, ByVal cP As String, ByVal cC As String, ByVal r As Long, ByVal p As Double)
    If p < 0 Then
        ws.Range(cP & r).Value = "?"
    ElseIf p >= 0.999 Then
        ws.Range(cP & r).Value = "IESIT"
        ws.Range(cP & r).Interior.Color = RGB(146, 208, 80)
    ElseIf p <= 0.001 Then
        ws.Range(cP & r).Value = "PIERDUT"
        ws.Range(cP & r).Interior.Color = RGB(255, 120, 120)
    Else
        ws.Range(cP & r).Value = p
        ws.Range(cP & r).NumberFormat = "0%"
        ws.Range(cC & r).Value = CotaMin(p)
        ws.Range(cC & r).NumberFormat = "0.00"
        If p >= 0.75 Then ws.Range(cP & r).Font.Bold = True
    End If
    ws.Range(cP & r & ":" & cC & r).HorizontalAlignment = xlCenter
End Sub

' Over 1.5 in prima repriza, live
Private Function SansaHT15(a As Variant, ByVal st As String, ByVal mnt As Long, ByVal gh As Long, ByVal ga As Long) As Double
    Dim rH As Double, rA As Double, rHTh As Double, rHTa As Double
    If gh + ga >= 2 Then SansaHT15 = 1: Exit Function
    If st <> "NS" And st <> "1H" Then SansaHT15 = 0: Exit Function
    RamaseLive a(0), a(1), a(2), a(3), st, mnt, rH, rA, rHTh, rHTa
    SansaHT15 = PoisPeste(1 - (gh + ga), rHTh + rHTa)
End Function

' sansa live pentru Over k+0.5 goluri in prima repriza (k = 0, 1, 2)
Private Function SansaHTLinie(ByVal k As Long, a As Variant, ByVal st As String, ByVal mnt As Long, _
        ByVal gh As Long, ByVal ga As Long, ByVal hth As Variant, ByVal hta As Variant) As Double
    Dim rH As Double, rA As Double, rHTh As Double, rHTa As Double, g As Long
    If Not IsEmpty(hth) Then
        SansaHTLinie = IIf(hth + hta > k, 1, 0)
        Exit Function
    End If
    If st = "NS" Or st = "1H" Then
        g = IIf(st = "1H", gh + ga, 0)
        If g > k Then SansaHTLinie = 1: Exit Function
        RamaseLive a(0), a(1), a(2), a(3), st, mnt, rH, rA, rHTh, rHTa
        SansaHTLinie = PoisPeste(k - g, rHTh + rHTa)
        If SansaHTLinie > 0.999 Then SansaHTLinie = 0.999
    Else
        SansaHTLinie = -1
    End If
End Function

' o celula de piata: procent + cota minima, sau IESIT / PIERDUT colorat
Private Sub ScrieHT(ws As Worksheet, ByVal c As String, ByVal r As Long, ByVal p As Double)
    With ws.Range(c & r)
        If p < 0 Then
            .Value = "?"
        ElseIf p >= 0.999 Then
            .Value = "IESIT"
            .Interior.Color = RGB(146, 208, 80): .Font.Bold = True
        ElseIf p <= 0.001 Then
            .Value = "PIERDUT"
            .Interior.Color = RGB(255, 120, 120): .Font.Bold = True
        Else
            .Value = Format(p, "0%") & "  (min " & Format(CotaMin(p), "0.00") & ")"
            If p >= 0.75 Then .Font.Bold = True
        End If
    End With
End Sub

' cornere in prima repriza: alege linia (2.5 - 6.5) cea mai mare cu cel putin 70% inainte de meci
Private Sub ScrieCornereHT(ws As Worksheet, ByVal r As Long, ByVal lC As Variant, ByVal st As String, ByVal mnt As Long, _
        ByVal cr As Long, ByVal stiutLaPauza As Boolean)
    Dim lHTc As Double, k As Long, linie As Long, rest As Double, p As Double, t As Double
    If IsEmpty(lC) Or Not IsNumeric(lC) Then ws.Range("F" & r).Value = "?": Exit Sub
    If lC <= 0 Then ws.Range("F" & r).Value = "?": Exit Sub
    lHTc = 0.46 * lC
    linie = 2
    For k = 6 To 2 Step -1
        If PoisPeste(k, lHTc) >= 0.7 Then linie = k: Exit For
    Next k
    If st = "NS" Then
        p = PoisPeste(linie, lHTc)
    ElseIf st = "1H" Then
        If cr < 0 Then
            ws.Range("F" & r).Value = "O" & linie & ".5: ? (fara date live)"
            Exit Sub
        End If
        If cr > linie Then
            p = 1
        Else
            t = mnt: If t > 45 Then t = 45
            rest = lHTc * (45 - t) / 45 + 0.05 * lHTc
            p = PoisPeste(linie - cr, rest)
            If p > 0.999 Then p = 0.999
        End If
    ElseIf cr >= 0 And (st = "HT" Or stiutLaPauza) Then
        p = IIf(cr > linie, 1, 0)
    Else
        ws.Range("F" & r).Value = "O" & linie & ".5: ?"
        Exit Sub
    End If
    ScrieHT ws, "F", r, p
    If p > 0.001 And p < 0.999 Then
        ws.Range("F" & r).Value = "O" & linie & ".5: " & ws.Range("F" & r).Value & IIf(cr >= 0 And st = "1H", "  [" & cr & " pana acum]", "")
    Else
        ws.Range("F" & r).Value = "O" & linie & ".5: " & ws.Range("F" & r).Value & IIf(cr >= 0, " (" & cr & ")", "")
    End If
End Sub

' cornerele din statisticile live: id meci -> total cornere pana acum
Private Sub CitesteCornere(res2 As Object, corn As Object)
    Dim fx As Variant, tm As Variant, it As Variant, tot As Long, gasit As Boolean
    If res2 Is Nothing Then Exit Sub
    If Not IsObject(res2("response")) Then Exit Sub
    For Each fx In res2("response")
        tot = 0: gasit = False
        If IsObject(JGet(fx, "statistics")) Then
            For Each tm In JGet(fx, "statistics")
                If IsObject(JGet(tm, "statistics")) Then
                    For Each it In JGet(tm, "statistics")
                        If JTxt(JGet(it, "type")) = "Corner Kicks" Then
                            tot = tot + CLng(JNumar(JGet(it, "value")))
                            gasit = True
                        End If
                    Next it
                End If
            Next tm
        End If
        If gasit Then corn(JTxt(JGet(fx, "fixture", "id"))) = tot
    Next fx
End Sub

' ---------- biletele de azi, live ----------
Private Sub BileteLive(ws As Worksheet, ByRef r As Long, res As Object, lam As Object)
    Dim wb As Worksheet, last As Long, rr As Long, k As Long, gz As String, op As String, piata As String
    Dim f As Variant, st As String, mnt As Long, gh As Long, ga As Long, hth As Variant, hta As Variant
    Dim pLeg As Double, pTot As Double, pierdut As Boolean, toate As Boolean, txt As String, stare As String
    Dim r0 As Long, a As Variant, fara As Boolean
    ws.Range("A" & r).Value = "Biletele si bombele de azi - live"
    ws.Range("A" & r).Font.Bold = True: ws.Range("A" & r).Font.Size = 12
    r = r + 1
    ws.Range("A" & r & ":E" & r).Value = Array("Bilet", "Meciuri (stare acum)", "Sansa sa iasa acum", "Stare", "Cota corecta acum")
    AntetTabel ws.Range("A" & r & ":E" & r)
    r0 = r
    On Error Resume Next
    Set wb = ThisWorkbook.Sheets("Bilete")
    On Error GoTo 0
    If wb Is Nothing Then GoTo Gata
    last = wb.Cells(wb.Rows.Count, 1).End(xlUp).Row
    For rr = 2 To last
        If IsNumeric(wb.Range("A" & rr).Value2) And Not IsEmpty(wb.Range("A" & rr).Value2) Then
            If CLng(wb.Range("A" & rr).Value2) = CLng(Date) Then
                pTot = 1: pierdut = False: toate = True: txt = "": fara = False
                k = 0
                Do While CStr(wb.Cells(rr, 9 + k * 3).Value) <> ""
                    gz = CStr(wb.Cells(rr, 9 + k * 3).Value)
                    op = CStr(wb.Cells(rr, 10 + k * 3).Value)
                    piata = CStr(wb.Cells(rr, 11 + k * 3).Value)
                    Set f = GasesteMeciLive(res, gz, op)
                    If f Is Nothing Or Not lam.Exists(gz & "|" & op) Then
                        pLeg = -1
                        txt = txt & IIf(txt = "", "", vbLf) & gz & " - " & op & " (" & piata & "): fara date live"
                    Else
                        a = lam(gz & "|" & op)
                        StareMeci f, st, mnt, gh, ga, hth, hta
                        pLeg = SansaPiataLive(piata, a, st, mnt, gh, ga, hth, hta)
                        txt = txt & IIf(txt = "", "", vbLf) & gz & " - " & op & " (" & piata & "): " & _
                            IIf(st = "NS", "neinceput", "'" & gh & ":" & ga & " " & TextStare(st, mnt)) & " - " & _
                            IIf(pLeg >= 0.999, "IESIT", IIf(pLeg <= 0.001 And pLeg >= 0, "PIERDUT", IIf(pLeg < 0, "?", Format(pLeg, "0%"))))
                    End If
                    If pLeg < 0 Then
                        fara = True
                    Else
                        If pLeg <= 0.001 Then pierdut = True
                        If pLeg < 0.999 Then toate = False
                        pTot = pTot * pLeg
                    End If
                    k = k + 1
                Loop
                If pierdut Then
                    stare = "PIERDUT": pTot = 0
                ElseIf toate And Not fara Then
                    stare = "IESIT"
                Else
                    stare = "in joc"
                End If
                r = r + 1
                ws.Range("A" & r).Value = IIf(CStr(wb.Range("B" & rr).Value) = "Bilet", "Cota " & wb.Range("C" & rr).Value, _
                    "Bomba " & wb.Range("C" & rr).Value)
                ws.Range("B" & r).Value = Replace(txt, "'", "")
                ws.Range("C" & r).Value = pTot
                ws.Range("C" & r).NumberFormat = "0%"
                ws.Range("D" & r).Value = stare & IIf(fara And stare = "in joc", " (fara date pt unele piete)", "")
                If pTot > 0.001 And stare = "in joc" Then
                    ws.Range("E" & r).Value = 1 / pTot
                    ws.Range("E" & r).NumberFormat = "0.00"
                End If
                If stare = "IESIT" Then ws.Range("D" & r).Interior.Color = RGB(146, 208, 80)
                If stare = "PIERDUT" Then ws.Range("D" & r).Interior.Color = RGB(255, 120, 120)
            End If
        End If
    Next rr
Gata:
    If r = r0 Then
        r = r + 1
        ws.Range("A" & r).Value = "Azi nu sunt bilete salvate (se salveaza cand ruleaza Recomandari)."
    Else
        With ws.Range("A" & (r0 + 1) & ":E" & r)
            .WrapText = True
            .VerticalAlignment = xlTop
            .Borders.Color = RGB(200, 205, 215)
        End With
        ws.Range("B" & (r0 + 1) & ":B" & r).ColumnWidth = 60
        ws.Range("A" & (r0 + 1) & ":E" & r).Rows.AutoFit
    End If
End Sub

' sansa live ca o piata sa iasa: 1 = deja iesit, 0 = pierdut, -1 = fara date live
Private Function SansaPiataLive(ByVal piata As String, a As Variant, ByVal st As String, ByVal mnt As Long, _
        ByVal gh As Long, ByVal ga As Long, ByVal hth As Variant, ByVal hta As Variant) As Double
    Dim rH As Double, rA As Double, rHTh As Double, rHTa As Double, nevoie As Long, fin As Boolean
    Dim x As Long, y As Long, k As Long, eticheta As String, et3 As Variant, hh As Long, ha As Long
    Dim p As Double, a1 As Long, b1 As Long, a2 As Long, b2 As Long, eh1(0 To 6) As Double, ea1(0 To 6) As Double
    Dim eh2(0 To 6) As Double, ea2(0 To 6) As Double, h2 As Double, aw2 As Double, htKnown As Boolean
    et3 = Array("1", "X", "2")
    fin = (st = "FT" Or st = "AET" Or st = "PEN")
    If st = "PST" Or st = "CANC" Or st = "ABD" Or st = "AWD" Or st = "WO" Then SansaPiataLive = -1: Exit Function
    RamaseLive a(0), a(1), a(2), a(3), st, mnt, rH, rA, rHTh, rHTa
    htKnown = Not IsEmpty(hth)

    If piata = "Over 1.5" Or piata = "Over 2.5" Then
        nevoie = IIf(piata = "Over 1.5", 2, 3) - (gh + ga)
        If nevoie <= 0 Then SansaPiataLive = 1: Exit Function
        If fin Then SansaPiataLive = 0: Exit Function
        SansaPiataLive = PoisPeste(nevoie - 1, rH + rA)
    ElseIf piata = "GG" Then
        If gh > 0 And ga > 0 Then SansaPiataLive = 1: Exit Function
        If fin Then SansaPiataLive = 0: Exit Function
        SansaPiataLive = IIf(gh > 0, 1, 1 - Exp(-rH)) * IIf(ga > 0, 1, 1 - Exp(-rA))
    ElseIf piata = "Gazde peste 1.5" Then
        If gh >= 2 Then SansaPiataLive = 1: Exit Function
        If fin Then SansaPiataLive = 0: Exit Function
        SansaPiataLive = PoisPeste(1 - gh, rH)
    ElseIf piata = "Oaspeti peste 1.5" Then
        If ga >= 2 Then SansaPiataLive = 1: Exit Function
        If fin Then SansaPiataLive = 0: Exit Function
        SansaPiataLive = PoisPeste(1 - ga, rA)
    ElseIf piata = "Over 0.5 pauza" Then
        If htKnown Then
            SansaPiataLive = IIf(hth + hta >= 1, 1, 0)
        ElseIf st = "1H" And gh + ga >= 1 Then
            SansaPiataLive = 1
        ElseIf st = "NS" Or st = "1H" Then
            SansaPiataLive = PoisPeste(0, rHTh + rHTa)
        Else
            SansaPiataLive = -1
        End If
    ElseIf Left$(piata, 12) = "Scor corect " Then
        eticheta = Trim$(Mid$(piata, 13))
        k = InStr(eticheta, "-")
        If k = 0 Then SansaPiataLive = -1: Exit Function
        x = Val(Left$(eticheta, k - 1)): y = Val(Mid$(eticheta, k + 1))
        If gh > x Or ga > y Then SansaPiataLive = 0: Exit Function
        If fin Then
            SansaPiataLive = IIf(gh = x And ga = y, 1, 0)
        Else
            SansaPiataLive = PoisExact(x - gh, rH) * PoisExact(y - ga, rA)
        End If
    ElseIf Left$(piata, 6) = "HT/FT " Then
        eticheta = Mid$(piata, 7, 3)
        If fin And htKnown Then
            SansaPiataLive = IIf(eticheta = et3(Rez(CLng(hth), CLng(hta))) & "/" & et3(Rez(gh, ga)), 1, 0)
            Exit Function
        End If
        If htKnown Then
            ' pauza stiuta: trebuie sa fie aceeasi, apoi conteaza doar finalul
            If Left$(eticheta, 1) <> et3(Rez(CLng(hth), CLng(hta))) Then SansaPiataLive = 0: Exit Function
            p = 0
            For a1 = 0 To 6
                For b1 = 0 To 6
                    If et3(Rez(gh + a1, ga + b1)) = Right$(eticheta, 1) Then p = p + PoisExact(a1, rH) * PoisExact(b1, rA)
                Next b1
            Next a1
            SansaPiataLive = p
        Else
            ' inainte de pauza: goluri ramase in repriza 1 si apoi in repriza 2
            h2 = rH - rHTh: aw2 = rA - rHTa
            If h2 < 0.01 Then h2 = 0.01
            If aw2 < 0.01 Then aw2 = 0.01
            For k = 0 To 6
                eh1(k) = PoisExact(k, rHTh): ea1(k) = PoisExact(k, rHTa)
                eh2(k) = PoisExact(k, h2): ea2(k) = PoisExact(k, aw2)
            Next k
            p = 0
            For a1 = 0 To 6
                For b1 = 0 To 6
                    If et3(Rez(gh + a1, ga + b1)) = Left$(eticheta, 1) Then
                        For a2 = 0 To 6
                            For b2 = 0 To 6
                                If et3(Rez(gh + a1 + a2, ga + b1 + b2)) = Right$(eticheta, 1) Then
                                    p = p + eh1(a1) * ea1(b1) * eh2(a2) * ea2(b2)
                                End If
                            Next b2
                        Next a2
                    End If
                Next b1
            Next a1
            SansaPiataLive = p
        End If
    Else
        ' cornere / galbene: nu avem date live gratuite
        SansaPiataLive = -1
    End If
    If SansaPiataLive > 0.999 Then SansaPiataLive = 0.999
End Function

' ---------- lambdas salvate la scanare ----------
Private Sub SalveazaLambda(ByVal home As String, ByVal away As String, ByVal lH As Double, ByVal lA As Double, _
        ByVal lHTh As Double, ByVal lHTa As Double, ByVal prima As Boolean, Optional ByVal lC As Double = 0)
    Dim ws As Worksheet, r As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("LiveDate")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "LiveDate"
        ws.Visible = xlSheetHidden
    End If
    If prima Then
        ws.Cells.Clear
        ws.Range("A1:G1").Value = Array("Gazde", "Oaspeti", "lH", "lA", "lHTh", "lHTa", "Cornere")
        ws.Range("H1").Value = Date
    End If
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Range("A" & r & ":G" & r).Value = Array(home, away, lH, lA, lHTh, lHTa, lC)
End Sub

Private Function IncarcaLambda() As Object
    Dim ws As Worksheet, d As Object, last As Long, v As Variant, i As Long
    Set d = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("LiveDate")
    On Error GoTo 0
    If Not ws Is Nothing Then
        last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If last >= 2 Then
            v = ws.Range("A2:G" & last).Value2
            For i = 1 To UBound(v, 1)
                If Not d.Exists(v(i, 1) & "|" & v(i, 2)) Then d.Add v(i, 1) & "|" & v(i, 2), Array(v(i, 3), v(i, 4), v(i, 5), v(i, 6), NumarCota(v(i, 7)))
            Next i
        End If
    End If
    Set IncarcaLambda = d
End Function

' ---------- potrivirea numelor (football-data vs API-Football) ----------
' numele se normalizeaza o singura data pentru toate meciurile primite de la API
Private Sub PregatesteLive(res As Object)
    Dim fx As Variant, n As Long
    Set mLiveF = New Collection
    mLiveN = 0
    If res Is Nothing Then Exit Sub
    If Not IsObject(res("response")) Then Exit Sub
    n = res("response").Count
    If n = 0 Then Exit Sub
    ReDim mLiveH(1 To n): ReDim mLiveA(1 To n)
    For Each fx In res("response")
        mLiveN = mLiveN + 1
        mLiveF.Add fx
        mLiveH(mLiveN) = NormEchipa(JTxt(JGet(fx, "teams", "home", "name")))
        mLiveA(mLiveN) = NormEchipa(JTxt(JGet(fx, "teams", "away", "name")))
        If mLiveN Mod 200 = 0 Then
            Application.StatusBar = "BET PRO LIVE: pregatesc meciurile " & mLiveN & " din " & n & "..."
            DoEvents
        End If
    Next fx
End Sub

Private Function GasesteMeciLive(res As Object, ByVal home As String, ByVal away As String) As Object
    Dim i As Long, sH As Long, sA As Long, best As Long, nh As String, na As String
    Set GasesteMeciLive = Nothing
    If mLiveN = 0 Then Exit Function
    nh = NormEchipa(home): na = NormEchipa(away)
    For i = 1 To mLiveN
        sH = SimEchipa(nh, mLiveH(i))
        If sH >= 1 Then
            sA = SimEchipa(na, mLiveA(i))
            If sA >= 1 And sH + sA >= 3 And sH + sA > best Then
                best = sH + sA
                Set GasesteMeciLive = mLiveF(i)
            End If
        End If
    Next i
End Function

Private Function NormEchipa(ByVal s As String) As String
    Dim i As Long, c As String, o As String, cuv As Variant, w As Variant, r As String
    s = LCase$(s)
    s = Replace(s, ChrW(225), "a"): s = Replace(s, ChrW(224), "a"): s = Replace(s, ChrW(228), "a"): s = Replace(s, ChrW(227), "a")
    s = Replace(s, ChrW(226), "a"): s = Replace(s, ChrW(229), "a"): s = Replace(s, ChrW(233), "e"): s = Replace(s, ChrW(232), "e")
    s = Replace(s, ChrW(235), "e"): s = Replace(s, ChrW(234), "e"): s = Replace(s, ChrW(237), "i"): s = Replace(s, ChrW(239), "i")
    s = Replace(s, ChrW(243), "o"): s = Replace(s, ChrW(246), "o"): s = Replace(s, ChrW(244), "o"): s = Replace(s, ChrW(248), "o")
    s = Replace(s, ChrW(250), "u"): s = Replace(s, ChrW(252), "u"): s = Replace(s, ChrW(241), "n"): s = Replace(s, ChrW(231), "c")
    s = Replace(s, ChrW(351), "s"): s = Replace(s, ChrW(287), "g"): s = Replace(s, ChrW(305), "i"): s = Replace(s, ChrW(223), "ss")
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If (c >= "a" And c <= "z") Or (c >= "0" And c <= "9") Then o = o & c Else o = o & " "
    Next i
    ' prescurtari folosite de football-data
    o = " " & o & " "
    o = Replace(o, " man united ", " manchester united "): o = Replace(o, " man city ", " manchester city ")
    o = Replace(o, " nott m forest ", " nottingham forest "): o = Replace(o, " wolves ", " wolverhampton ")
    o = Replace(o, " sheffield weds ", " sheffield wednesday "): o = Replace(o, " sheffield utd ", " sheffield united ")
    o = Replace(o, " qpr ", " queens park rangers "): o = Replace(o, " ein frankfurt ", " eintracht frankfurt ")
    o = Replace(o, " m gladbach ", " monchengladbach "): o = Replace(o, " ath madrid ", " atletico madrid ")
    o = Replace(o, " ath bilbao ", " athletic club "): o = Replace(o, " paris sg ", " paris saint germain ")
    o = Replace(o, " st etienne ", " saint etienne "): o = Replace(o, " sp gijon ", " sporting gijon ")
    o = Replace(o, " sp lisbon ", " sporting cp "): o = Replace(o, " espanol ", " espanyol ")
    o = Replace(o, " munich ", " munchen "): o = Replace(o, " inverness c ", " inverness ")
    ' cuvinte fara importanta
    cuv = Array("fc", "cf", "afc", "sc", "ac", "as", "ss", "us", "sv", "vfb", "vfl", "fk", "sk", "bk", "cd", "ud", "rcd", _
        "sd", "ca", "rc", "the", "club", "calcio", "1", "04", "05", "07", "09", "1846", "1899", "1909")
    r = ""
    For Each w In Split(Trim$(o), " ")
        If w <> "" Then
            If Not EsteInLista(CStr(w), cuv) Then r = r & " " & w
        End If
    Next w
    NormEchipa = Trim$(r)
    If NormEchipa = "" Then NormEchipa = Trim$(o)
End Function

Private Function EsteInLista(ByVal w As String, lista As Variant) As Boolean
    Dim x As Variant
    For Each x In lista
        If x = w Then EsteInLista = True: Exit Function
    Next x
End Function

' 3 = identic, 2 = un cuvant important comun sau unul il contine pe celalalt, 1 = acelasi inceput
Private Function SimEchipa(ByVal a As String, ByVal b As String) As Long
    Dim wa As Variant, wb As Variant, x As Variant, y As Variant
    If a = "" Or b = "" Then Exit Function
    If a = b Then SimEchipa = 3: Exit Function
    If InStr(" " & b & " ", " " & a & " ") > 0 Or InStr(" " & a & " ", " " & b & " ") > 0 Then SimEchipa = 2: Exit Function
    wa = Split(a, " "): wb = Split(b, " ")
    For Each x In wa
        For Each y In wb
            If Len(x) >= 4 And x = y Then SimEchipa = 2: Exit Function
        Next y
    Next x
    If Len(a) >= 4 And Len(b) >= 4 Then
        If Left$(a, 4) = Left$(b, 4) Then SimEchipa = 1
    End If
End Function

' ---------- API-Football ----------
Private Function CheieApi() As String
    Dim we As Worksheet
    Set we = FoaieEmail()
    If CStr(we.Range("A13").Value) <> "Cheie API-Football (pentru LIVE):" Then
        we.Range("A13").Value = "Cheie API-Football (pentru LIVE):"
        we.Range("B13").Interior.Color = RGB(255, 242, 204)
    End If
    If Trim$(CStr(we.Range("B13").Value)) = "" Then we.Range("B13").Value = API_DEF
    CheieApi = Trim$(CStr(we.Range("B13").Value))
End Function

Private Function ApiFootball(ByVal cale As String, ByVal cheie As String, ByRef res As Object, ByRef ramase As String) As Boolean
    Dim http As Object, t As String, p As Long, e As Variant, k As Variant, msg As String
    On Error GoTo Esec
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    http.Open "GET", "https://v3.football.api-sports.io" & cale, False
    http.setRequestHeader "x-apisports-key", cheie
    http.send
    On Error Resume Next
    ramase = http.getResponseHeader("x-ratelimit-requests-remaining")
    On Error GoTo Esec
    If http.Status <> 200 Then mEroare = "HTTP " & http.Status: Exit Function
    t = http.responseText
    If Left$(LTrim$(t), 1) <> "{" Then mEroare = "raspuns neasteptat": Exit Function
    p = 1: mNextB = 0
    Set res = JParse(t, p)
    If res.Exists("errors") Then
        If IsObject(res("errors")) Then
            Set e = res("errors")
            If e.Count > 0 Then
                If TypeName(e) = "Dictionary" Then
                    For Each k In e.Keys
                        msg = msg & k & ": " & JTxt(e(k)) & " "
                    Next k
                Else
                    For Each k In e
                        msg = msg & JTxt(k) & " "
                    Next k
                End If
                mEroare = msg
                Exit Function
            End If
        End If
    End If
    ApiFootball = res.Exists("response")
    Exit Function
Esec:
    mEroare = Err.Description
End Function

Private Function FoaieLive() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Live")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets("Listing"))
        ws.Name = "Live"
    End If
    ws.Range("A1").Value = "BET PRO LIVE"
    ws.Range("A1").Font.Size = 16: ws.Range("A1").Font.Bold = True
    Set FoaieLive = ws
End Function

' ---------- cititor JSON ----------
Private Function JTxt(v As Variant) As String
    If IsObject(v) Then Exit Function
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    JTxt = CStr(v)
End Function

Private Function JNumar(v As Variant) As Double
    If IsObject(v) Then Exit Function
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    If IsNumeric(v) Then JNumar = CDbl(v)
End Function

' Citeste o valoare din JSON pe un drum de chei, ex: JGet(fx, "teams", "home", "name")
Private Function JGet(ByVal o As Variant, ParamArray keys() As Variant) As Variant
    Dim i As Long
    For i = LBound(keys) To UBound(keys)
        If Not IsObject(o) Then JGet = Empty: Exit Function
        If TypeName(o) = "Collection" Then
            If keys(i) < 1 Or keys(i) > o.Count Then JGet = Empty: Exit Function
        Else
            If Not o.Exists(keys(i)) Then JGet = Empty: Exit Function
        End If
        If IsObject(o(keys(i))) Then
            Set o = o(keys(i))
        Else
            o = o(keys(i))
        End If
    Next i
    If IsObject(o) Then Set JGet = o Else JGet = o
End Function

' ---------- cititor JSON (obiecte -> Dictionary, liste -> Collection) ----------
Private Function JParse(ByRef s As String, ByRef p As Long) As Variant
    Dim ch As String
    JSkip s, p
    ch = Mid$(s, p, 1)
    Select Case ch
        Case "{": Set JParse = JObj(s, p)
        Case "[": Set JParse = JArr(s, p)
        Case """": JParse = JStr(s, p)
        Case "t": JParse = True: p = p + 4
        Case "f": JParse = False: p = p + 5
        Case "n": JParse = Null: p = p + 4
        Case Else: JParse = JNum(s, p)
    End Select
End Function

Private Function JObj(ByRef s As String, ByRef p As Long) As Object
    Dim d As Object, k As String, ch As String
    Set d = CreateObject("Scripting.Dictionary")
    p = p + 1
    JSkip s, p
    If Mid$(s, p, 1) = "}" Then p = p + 1: Set JObj = d: Exit Function
    Do
        JSkip s, p
        k = JStr(s, p)
        JSkip s, p
        p = p + 1
        If d.Exists(k) Then d.Remove k
        d.Add k, JParse(s, p)
        JSkip s, p
        ch = Mid$(s, p, 1)
        p = p + 1
        If ch <> "," Then Exit Do
    Loop
    Set JObj = d
End Function

Private Function JArr(ByRef s As String, ByRef p As Long) As Object
    Dim c As Collection, ch As String
    Set c = New Collection
    p = p + 1
    JSkip s, p
    If Mid$(s, p, 1) = "]" Then p = p + 1: Set JArr = c: Exit Function
    Do
        c.Add JParse(s, p)
        JSkip s, p
        ch = Mid$(s, p, 1)
        p = p + 1
        If ch <> "," Then Exit Do
    Loop
    Set JArr = c
End Function

Private Function JStr(ByRef s As String, ByRef p As Long) As String
    Dim outS As String, q As Long, b As Long, e As String
    p = p + 1
    Do
        q = InStr(p, s, """")
        If q = 0 Then
            JStr = outS & Mid$(s, p)
            p = Len(s) + 1
            Exit Function
        End If
        If mNextB < p Then
            mNextB = InStr(p, s, "\")
            If mNextB = 0 Then mNextB = Len(s) + 1
        End If
        b = mNextB
        If b > q Then
            outS = outS & Mid$(s, p, q - p)
            p = q + 1
            Exit Do
        End If
        outS = outS & Mid$(s, p, b - p)
        e = Mid$(s, b + 1, 1)
        Select Case e
            Case "n": outS = outS & vbLf: p = b + 2
            Case "t": outS = outS & vbTab: p = b + 2
            Case "r": outS = outS & vbCr: p = b + 2
            Case "b", "f": p = b + 2
            Case "u": outS = outS & ChrW(CLng("&H" & Mid$(s, b + 2, 4))): p = b + 6
            Case Else: outS = outS & e: p = b + 2
        End Select
    Loop
    JStr = outS
End Function

Private Function JNum(ByRef s As String, ByRef p As Long) As Variant
    Dim st As Long
    st = p
    Do While p <= Len(s)
        If InStr("+-0123456789.eE", Mid$(s, p, 1)) = 0 Then Exit Do
        p = p + 1
    Loop
    JNum = Val(Mid$(s, st, p - st))
End Function

Private Sub JSkip(ByRef s As String, ByRef p As Long)
    Dim c As String
    Do While p <= Len(s)
        c = Mid$(s, p, 1)
        If c = " " Or c = vbTab Or c = vbCr Or c = vbLf Then
            p = p + 1
        Else
            Exit Do
        End If
    Loop
End Sub


' ============================ COTE REALE (API-Football) ============================
' Pentru fiecare meci scanat azi ia cotele de la toate casele si pastreaza cea mai buna cota
' pe fiecare piata, in foaia ascunsa "Cote". Recomandarile si biletele le folosesc pentru value.

Public Sub IaCoteZi()
    Dim cheie As String, res As Object, rc As Object, ramase As String, lam As Object, k As Variant
    Dim gz As String, op As String, f As Object, idF As String, ws As Worksheet, wd As Worksheet
    Dim cote As Object, mk As Variant, r As Long, nMeci As Long, nCote As Long, numeBet As Object, p As Long, tv As Variant

    cheie = CheieApi()
    If cheie = "" Then Exit Sub
    Set lam = IncarcaLambda()
    If lam.Count = 0 Then Exit Sub
    Application.StatusBar = "BET PRO: caut meciurile de azi in API..."
    DoEvents
    If Not MeciuriAziApi(cheie, res, ramase) Then
        Mesaj "Nu am putut lua meciurile din API: " & mEroare, vbExclamation, "BET PRO"
        Exit Sub
    End If
    PregatesteLive res

    Set ws = FoaieNoua("Cote")
    ws.Range("A1:F1").Value = Array("Gazde", "Oaspeti", "Piata", "Cota", "Casa", "Sansa Pinnacle")
    ws.Range("K1:M1").Value = Array("Gazde", "Oaspeti", "ID meci")
    Set numeBet = CreateObject("Scripting.Dictionary")
    Dim pinP As Object, rk As Long
    rk = 1
    r = 1
    For Each k In lam.Keys
        p = InStr(k, "|")
        gz = Left$(k, p - 1): op = Mid$(k, p + 1)
        Set f = GasesteMeciLive(res, gz, op)
        If Not f Is Nothing Then
            idF = JTxt(JGet(f, "fixture", "id"))
            nMeci = nMeci + 1
            Application.StatusBar = "BET PRO: cote pentru " & gz & " - " & op & " (" & nMeci & ")..."
            DoEvents
            rk = rk + 1
            ws.Range("K" & rk & ":M" & rk).Value = Array(gz, op, idF)
            Set cote = CreateObject("Scripting.Dictionary")
            Set pinP = CreateObject("Scripting.Dictionary")
            Set rc = Nothing
            If ApiFootball("/odds?fixture=" & idF & "&bookmaker=8", cheie, rc, ramase) Then CitesteCote rc, cote, numeBet
            Set rc = Nothing
            If ApiFootball("/odds?fixture=" & idF & "&bookmaker=4", cheie, rc, ramase) Then
                CitesteCote rc, cote, numeBet
                CitestePinnacle rc, pinP
            End If
            For Each mk In cote.Keys
                r = r + 1: nCote = nCote + 1
                tv = cote(mk)
                ws.Range("A" & r & ":E" & r).Value = Array(gz, op, mk, tv(0), tv(1))
                If pinP.Exists(mk) Then ws.Range("F" & r).Value = pinP(mk)
            Next mk
        End If
    Next k
    ws.Range("G1").Value = "Actualizat"
    ws.Range("G2").Value = Now
    ws.Range("I1").Value = "Piete primite de la API (pentru verificare)"
    r = 1
    For Each k In numeBet.Keys
        r = r + 1
        ws.Range("I" & r).Value = k
    Next k
    ws.Visible = xlSheetHidden
    gLog = gLog & "Cote reale: " & nCote & " cote pentru " & nMeci & " meciuri. Cereri API ramase azi: " & ramase & vbLf
End Sub

' cea mai buna cota pe fiecare piata, de la toate casele
Private Sub CitesteCote(rc As Object, cote As Object, numeBet As Object)
    Dim bk As Variant, bet As Variant, vv As Variant, nb As String, vs As String, od As Double, casa As String, piata As String
    Dim tmp As Variant
    If rc Is Nothing Then Exit Sub
    If Not IsObject(rc("response")) Then Exit Sub
    If rc("response").Count = 0 Then Exit Sub
    If Not IsObject(JGet(rc, "response", 1, "bookmakers")) Then Exit Sub
    For Each bk In JGet(rc, "response", 1, "bookmakers")
        casa = JTxt(JGet(bk, "name"))
        If IsObject(JGet(bk, "bets")) Then
            For Each bet In JGet(bk, "bets")
                nb = LCase$(JTxt(JGet(bet, "name")))
                If Not numeBet.Exists(nb) Then numeBet.Add nb, 1
                If IsObject(JGet(bet, "values")) Then
                    For Each vv In JGet(bet, "values")
                        vs = JTxt(JGet(vv, "value"))
                        od = Val(JTxt(JGet(vv, "odd")))
                        piata = PiataDinApi(nb, vs)
                        If piata <> "" And od > 1 Then
                            If Not cote.Exists(piata) Then
                                cote.Add piata, Array(od, casa)
                            Else
                                tmp = cote(piata)
                                If od > tmp(0) Then cote(piata) = Array(od, casa)
                            End If
                        End If
                    Next vv
                End If
            Next bet
        End If
    Next bk
End Sub

' traduce (piata API, valoare) in numele pietelor botului
Private Function PiataDinApi(ByVal nb As String, ByVal vl As String) As String
    Dim prima As Boolean, adoua As Boolean, gaz As Boolean, oas As Boolean, x As String, k As Long
    prima = (InStr(nb, "first half") > 0 Or InStr(nb, "1st half") > 0)
    adoua = (InStr(nb, "second half") > 0 Or InStr(nb, "2nd half") > 0)
    gaz = (InStr(nb, "home") > 0)
    oas = (InStr(nb, "away") > 0)
    If adoua Then Exit Function
    If nb = "goals over/under" Then
        If vl = "Over 1.5" Or vl = "Over 2.5" Or vl = "Over 3.5" Then PiataDinApi = vl
    ElseIf InStr(nb, "over/under") > 0 And InStr(nb, "goal") > 0 And prima And Not gaz And Not oas Then
        If vl = "Over 0.5" Then PiataDinApi = "Over 0.5 pauza"
        If vl = "Over 1.5" Then PiataDinApi = "Over 1.5 pauza"
        If vl = "Over 2.5" Then PiataDinApi = "Over 2.5 pauza"
    ElseIf nb = "both teams score" Then
        If vl = "Yes" Then PiataDinApi = "GG"
    ElseIf (nb = "total - home" Or nb = "home team total goals" Or nb = "home team goals over/under") And Not prima Then
        If vl = "Over 1.5" Then PiataDinApi = "Gazde peste 1.5"
    ElseIf (nb = "total - away" Or nb = "away team total goals" Or nb = "away team goals over/under") And Not prima Then
        If vl = "Over 1.5" Then PiataDinApi = "Oaspeti peste 1.5"
    ElseIf nb = "ht/ft double" Then
        x = Replace(Replace(Replace(vl, "Home", "1"), "Draw", "X"), "Away", "2")
        If Len(x) = 3 Then PiataDinApi = "HT/FT " & x
    ElseIf nb = "exact score" Then
        k = InStr(vl, ":")
        If k > 0 Then PiataDinApi = "Scor corect " & Trim$(Left$(vl, k - 1)) & "-" & Trim$(Mid$(vl, k + 1))
    ElseIf InStr(nb, "corner") > 0 And InStr(nb, "over") > 0 And Not prima And Not gaz And Not oas Then
        If Left$(vl, 5) = "Over " Then PiataDinApi = "Cornere O" & Trim$(Mid$(vl, 6))
    ElseIf InStr(nb, "card") > 0 And InStr(nb, "over") > 0 And Not prima And Not gaz And Not oas Then
        If Left$(vl, 5) = "Over " Then PiataDinApi = "Galbene O" & Trim$(Mid$(vl, 6))
    End If
End Function

' cotele din foaia "Cote": cheie "Gazde|Oaspeti|Piata" -> Array(cota, casa)
Private Function IncarcaCote() As Object
    Dim ws As Worksheet, d As Object, last As Long, v As Variant, i As Long, key As String
    Set d = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Cote")
    On Error GoTo 0
    If Not ws Is Nothing Then
        If IsDate(ws.Range("G2").Value) Then
            If Int(ws.Range("G2").Value) = Date Then
                last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
                If last >= 2 Then
                    v = ws.Range("A2:F" & last).Value2
                    For i = 1 To UBound(v, 1)
                        key = v(i, 1) & "|" & v(i, 2) & "|" & v(i, 3)
                        If Not d.Exists(key) Then
                            If IsNumeric(v(i, 6)) And Not IsEmpty(v(i, 6)) Then
                                d.Add key, Array(v(i, 4), v(i, 5), CDbl(v(i, 6)))
                            Else
                                d.Add key, Array(v(i, 4), v(i, 5), -1#)
                            End If
                        End If
                    Next i
                End If
            End If
        End If
    End If
    Set IncarcaCote = d
End Function

Private Function CotaReala(cote As Object, ByVal meci As String, ByVal piata As String, ByRef casa As String) As Double
    Dim gz As String, op As String, key As String, tv As Variant
    casa = ""
    ImparteMeci meci, gz, op
    key = gz & "|" & op & "|" & piata
    If cote.Exists(key) Then
        tv = cote(key)
        CotaReala = CDbl(tv(0))
        casa = CStr(tv(1))
    End If
End Function


' ============================ VALUE: Pinnacle + jurnal cu cota de start (CLV) ============================

' sansa corecta Pinnacle (fara marja) pentru o piata, -1 daca Pinnacle nu o are
Private Function SansaPinnacle(cote As Object, ByVal meci As String, ByVal piata As String) As Double
    Dim gz As String, op As String, key As String, tv As Variant
    SansaPinnacle = -1
    ImparteMeci meci, gz, op
    key = gz & "|" & op & "|" & piata
    If cote.Exists(key) Then
        tv = cote(key)
        If UBound(tv) >= 2 Then SansaPinnacle = CDbl(tv(2))
    End If
End Function

' scoate marja din cotele Pinnacle: pereche Over/Under, Da/Nu, sau toata piata (HT/FT, scor corect)
Private Sub CitestePinnacle(rc As Object, pinP As Object)
    Dim bk As Variant, bet As Variant, vv As Variant, nb As String, vs As String, od As Double, piata As String
    Dim vals As Object, k As Variant, sumAll As Double, comp As String, sm As Double
    If rc Is Nothing Then Exit Sub
    If Not IsObject(rc("response")) Then Exit Sub
    If rc("response").Count = 0 Then Exit Sub
    If Not IsObject(JGet(rc, "response", 1, "bookmakers")) Then Exit Sub
    For Each bk In JGet(rc, "response", 1, "bookmakers")
        If IsObject(JGet(bk, "bets")) Then
            For Each bet In JGet(bk, "bets")
                nb = LCase$(JTxt(JGet(bet, "name")))
                If IsObject(JGet(bet, "values")) Then
                    Set vals = CreateObject("Scripting.Dictionary")
                    sumAll = 0
                    For Each vv In JGet(bet, "values")
                        vs = JTxt(JGet(vv, "value"))
                        od = Val(JTxt(JGet(vv, "odd")))
                        If od > 1 And Not vals.Exists(vs) Then
                            vals.Add vs, od
                            sumAll = sumAll + 1 / od
                        End If
                    Next vv
                    For Each k In vals.Keys
                        piata = PiataDinApi(nb, CStr(k))
                        If piata <> "" And Not pinP.Exists(piata) Then
                            sm = 0
                            If Left$(CStr(k), 5) = "Over " Then
                                comp = "Under " & Mid$(CStr(k), 6)
                                If vals.Exists(comp) Then sm = 1 / vals(k) + 1 / vals(comp)
                            ElseIf CStr(k) = "Yes" Then
                                If vals.Exists("No") Then sm = 1 / vals(k) + 1 / vals("No")
                            Else
                                sm = sumAll
                                If sm < 1.02 Then sm = 1.05
                            End If
                            If sm >= 1 And sm < 1.4 Then pinP.Add piata, (1 / vals(k)) / sm
                        End If
                    Next k
                End If
            Next bet
        End If
    Next bk
End Sub

' "Gazde|Oaspeti" -> ID meci API, din foaia Cote
Private Function IncarcaIduri() As Object
    Dim ws As Worksheet, d As Object, last As Long, v As Variant, i As Long
    Set d = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Cote")
    On Error GoTo 0
    If Not ws Is Nothing Then
        last = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row
        If last >= 2 Then
            v = ws.Range("K2:M" & last).Value2
            For i = 1 To UBound(v, 1)
                If Not d.Exists(v(i, 1) & "|" & v(i, 2)) Then d.Add v(i, 1) & "|" & v(i, 2), CStr(v(i, 3))
            Next i
        End If
    End If
    Set IncarcaIduri = d
End Function

Private Function GreutatePinnacle() As Double
    Dim ws As Worksheet, gv As Variant
    GreutatePinnacle = 0.5
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Scanare")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    gv = ws.Range("AP23").Value2
    If IsNumeric(gv) And Not IsEmpty(gv) Then
        If gv > 1 Then gv = gv / 100
        If gv >= 0 And gv <= 1 Then GreutatePinnacle = CDbl(gv)
    End If
End Function

Private Function FoaieValue() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Value")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "Value"
        ws.Range("A1").Value = "Jurnal VALUE - fiecare pariu value, cu cota de la ora pontului si cota de la start"
        ws.Range("A1").Font.Size = 14: ws.Range("A1").Font.Bold = True
        ws.Range("A3").Value = "CLV = cota noastra / cota de la start - 1. Daca e pe plus des, botul gaseste valoare reala (cota a scazut dupa pont)."
        ws.Range("A3").Font.Italic = True
        ws.Range("A5:O5").Value = Array("Data", "Ora", "Meci", "Piata", "Sansa bot", "Sansa Pinnacle", "Sansa finala", _
            "Cota", "Casa", "Avantaj", "ID meci", "Cota la start", "CLV", "Rezultat", "Profit (1 unitate)")
        AntetTabel ws.Range("A5:O5")
        ws.Columns("A:B").ColumnWidth = 10
        ws.Columns("C").ColumnWidth = 34
        ws.Columns("D").ColumnWidth = 18
        ws.Columns("E:O").ColumnWidth = 12
    End If
    Set FoaieValue = ws
End Function

Private Sub SalveazaValue(ByVal ora As String, ByVal meci As String, ByVal piata As String, ByVal pB As Double, _
        ByVal pQ As Double, ByVal pF As Double, ByVal cotaV As Double, ByVal casa As String, ByVal eg As Double, idMeci As Object)
    Dim ws As Worksheet, last As Long, i As Long, v As Variant, gz As String, op As String, idF As String
    Set ws = FoaieValue()
    last = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    If last >= 6 Then
        v = ws.Range("A6:D" & last).Value2
        For i = 1 To UBound(v, 1)
            If IsNumeric(v(i, 1)) And Not IsEmpty(v(i, 1)) Then
                If CLng(v(i, 1)) = CLng(Date) And CStr(v(i, 3)) = meci And CStr(v(i, 4)) = piata Then Exit Sub
            End If
        Next i
    End If
    ImparteMeci meci, gz, op
    If idMeci.Exists(gz & "|" & op) Then idF = idMeci(gz & "|" & op)
    last = IIf(last < 5, 5, last) + 1
    ws.Range("A" & last).Value = Date
    ws.Range("A" & last).NumberFormat = "dd.mm.yyyy"
    ws.Range("B" & last & ":K" & last).Value = Array("'" & ora, meci, piata, pB, IIf(pQ > 0, pQ, "-"), pF, cotaV, casa, eg, "'" & idF)
    ws.Range("E" & last & ":G" & last & ",J" & last).NumberFormat = "0%"
    ws.Range("H" & last).NumberFormat = "0.00"
End Sub

Private Function IdCasa(ByVal casa As String) As Long
    If InStr(LCase$(casa), "pinnacle") > 0 Then IdCasa = 4 Else IdCasa = 8
End Function

' verifica pariurile value de ieri si de mai demult: rezultat, profit si cota de la start
Public Sub VerificaValue()
    Dim ws As Worksheet, last As Long, i As Long, v As Variant, cheie As String, ramase As String
    Dim ids As Object, k As Variant, grp As String, nG As Long, res As Object, fx As Variant, rc As Object
    Dim info As Object, a As Variant, st As String, rz As Long, cl As Object, tv As Variant, oc As Double
    Dim fh As Long, fa As Long, hh As Long, ha As Long, nc As Long, ny As Long, td As Object, tn As Object
    On Error GoTo Esec
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Value")
    On Error GoTo Esec
    If ws Is Nothing Then Exit Sub
    last = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    If last < 6 Then Exit Sub
    cheie = CheieApi()
    If cheie = "" Then Exit Sub
    v = ws.Range("A6:O" & last).Value2
    Set ids = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(v, 1)
        If CStr(v(i, 14)) = "" And CStr(v(i, 11)) <> "" And IsNumeric(v(i, 1)) And Not IsEmpty(v(i, 1)) Then
            If CLng(v(i, 1)) < CLng(Date) Then
                If Not ids.Exists(CStr(v(i, 11))) Then ids.Add CStr(v(i, 11)), 1
            End If
        End If
    Next i
    If ids.Count = 0 Then Exit Sub
    Application.StatusBar = "BET PRO: verific pariurile value..."
    DoEvents

    ' rezultatele, cate 20 de meciuri pe cerere
    Set info = CreateObject("Scripting.Dictionary")
    grp = "": nG = 0
    For Each k In ids.Keys
        grp = grp & IIf(grp = "", "", "-") & k: nG = nG + 1
        If nG = 20 Then
            Set res = Nothing
            If ApiFootball("/fixtures?ids=" & grp, cheie, res, ramase) Then CitesteRezultate res, info
            grp = "": nG = 0
        End If
    Next k
    If grp <> "" Then
        Set res = Nothing
        If ApiFootball("/fixtures?ids=" & grp, cheie, res, ramase) Then CitesteRezultate res, info
    End If

    ' cotele de la start, o cerere pe meci si casa
    Set cl = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(v, 1)
        If CStr(v(i, 14)) = "" And info.Exists(CStr(v(i, 11))) Then
            a = info(CStr(v(i, 11)))
            st = CStr(a(0))
            If st = "FT" Or st = "AET" Or st = "PEN" Then
                fh = a(1): fa = a(2): hh = a(3): ha = a(4): nc = a(5): ny = a(6)
                rz = EvalValue(CStr(v(i, 4)), fh, fa, hh, ha, nc, ny)
                If rz >= 0 Then
                    ws.Range("N" & (i + 5)).Value = IIf(rz = 1, "Iesit", "Pierdut")
                    ws.Range("N" & (i + 5)).Interior.Color = IIf(rz = 1, RGB(198, 239, 206), RGB(255, 199, 206))
                    ws.Range("O" & (i + 5)).Value = IIf(rz = 1, CDbl(v(i, 8)) - 1, -1)
                    ws.Range("O" & (i + 5)).NumberFormat = "0.00"
                End If
                k = CStr(v(i, 11)) & "|" & IdCasa(CStr(v(i, 9)))
                If Not cl.Exists(k) Then
                    Set td = CreateObject("Scripting.Dictionary")
                    Set tn = CreateObject("Scripting.Dictionary")
                    Set rc = Nothing
                    If ApiFootball("/odds?fixture=" & v(i, 11) & "&bookmaker=" & IdCasa(CStr(v(i, 9))), cheie, rc, ramase) Then
                        CitesteCote rc, td, tn
                    End If
                    cl.Add k, td
                End If
                Set td = cl(k)
                If td.Exists(CStr(v(i, 4))) Then
                    tv = td(CStr(v(i, 4)))
                    oc = CDbl(tv(0))
                    If oc > 1 Then
                        ws.Range("L" & (i + 5)).Value = oc
                        ws.Range("L" & (i + 5)).NumberFormat = "0.00"
                        ws.Range("M" & (i + 5)).Value = CDbl(v(i, 8)) / oc - 1
                        ws.Range("M" & (i + 5)).NumberFormat = "0%"
                    End If
                End If
            ElseIf st = "PST" Or st = "CANC" Or st = "ABD" Or st = "AWD" Or st = "WO" Then
                ws.Range("N" & (i + 5)).Value = "Anulat"
                ws.Range("O" & (i + 5)).Value = 0
            End If
        End If
    Next i
    Application.StatusBar = False
    Exit Sub
Esec:
    Application.StatusBar = False
    gLog = gLog & "Verificare value: " & Err.Description & vbLf
End Sub

' ID meci -> Array(stare, goluri gazde, goluri oaspeti, pauza gazde, pauza oaspeti, cornere, galbene)
Private Sub CitesteRezultate(res As Object, info As Object)
    Dim fx As Variant, tm As Variant, it As Variant, nc As Long, ny As Long, okC As Boolean, okY As Boolean, idF As String
    If res Is Nothing Then Exit Sub
    If Not IsObject(res("response")) Then Exit Sub
    For Each fx In res("response")
        idF = JTxt(JGet(fx, "fixture", "id"))
        nc = 0: ny = 0: okC = False: okY = False
        If IsObject(JGet(fx, "statistics")) Then
            For Each tm In JGet(fx, "statistics")
                If IsObject(JGet(tm, "statistics")) Then
                    For Each it In JGet(tm, "statistics")
                        If JTxt(JGet(it, "type")) = "Corner Kicks" Then nc = nc + CLng(JNumar(JGet(it, "value"))): okC = True
                        If JTxt(JGet(it, "type")) = "Yellow Cards" Then ny = ny + CLng(JNumar(JGet(it, "value"))): okY = True
                    Next it
                End If
            Next tm
        End If
        If Not okC Then nc = -1
        If Not okY Then ny = -1
        If idF <> "" Then
            info(idF) = Array(JTxt(JGet(fx, "fixture", "status", "short")), _
                CLng(JNumar(JGet(fx, "score", "fulltime", "home"))), CLng(JNumar(JGet(fx, "score", "fulltime", "away"))), _
                CLng(JNumar(JGet(fx, "score", "halftime", "home"))), CLng(JNumar(JGet(fx, "score", "halftime", "away"))), nc, ny)
        End If
    Next fx
End Sub

' 1 = iesit, 0 = pierdut, -1 = nu se poate verifica
Private Function EvalValue(ByVal piata As String, ByVal fh As Long, ByVal fa As Long, ByVal hh As Long, ByVal ha As Long, _
        ByVal nc As Long, ByVal ny As Long) As Long
    Dim x As String, k As Long, lin As Double, rh As String, rf As String
    EvalValue = -1
    Select Case piata
        Case "Over 0.5 pauza": EvalValue = IIf(hh + ha >= 1, 1, 0)
        Case "Over 1.5 pauza": EvalValue = IIf(hh + ha >= 2, 1, 0)
        Case "Over 2.5 pauza": EvalValue = IIf(hh + ha >= 3, 1, 0)
        Case "Over 1.5": EvalValue = IIf(fh + fa >= 2, 1, 0)
        Case "Over 2.5": EvalValue = IIf(fh + fa >= 3, 1, 0)
        Case "Over 3.5": EvalValue = IIf(fh + fa >= 4, 1, 0)
        Case "GG": EvalValue = IIf(fh > 0 And fa > 0, 1, 0)
        Case "Gazde peste 1.5": EvalValue = IIf(fh >= 2, 1, 0)
        Case "Oaspeti peste 1.5": EvalValue = IIf(fa >= 2, 1, 0)
    End Select
    If EvalValue >= 0 Then Exit Function
    If Left$(piata, 6) = "HT/FT " Then
        x = Mid$(piata, 7)
        rh = IIf(hh > ha, "1", IIf(hh = ha, "X", "2"))
        rf = IIf(fh > fa, "1", IIf(fh = fa, "X", "2"))
        EvalValue = IIf(x = rh & "/" & rf, 1, 0)
    ElseIf Left$(piata, 12) = "Scor corect " Then
        x = Mid$(piata, 13)
        k = InStr(x, "-")
        If k > 0 Then EvalValue = IIf(Val(Left$(x, k - 1)) = fh And Val(Mid$(x, k + 1)) = fa, 1, 0)
    ElseIf Left$(piata, 9) = "Cornere O" Then
        If nc >= 0 Then
            lin = Val(Mid$(piata, 10))
            EvalValue = IIf(nc > lin, 1, 0)
        End If
    ElseIf Left$(piata, 9) = "Galbene O" Then
        If ny >= 0 Then
            lin = Val(Mid$(piata, 10))
            EvalValue = IIf(ny > lin, 1, 0)
        End If
    End If
End Function

' o linie cu cum a mers value pana acum
Private Function RezumatValue() As String
    Dim ws As Worksheet, last As Long, v As Variant, i As Long, nP As Long, nW As Long, prof As Double
    Dim nCl As Long, sCl As Double, nClP As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Value")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    last = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    If last < 6 Then Exit Function
    v = ws.Range("A6:O" & last).Value2
    For i = 1 To UBound(v, 1)
        If CStr(v(i, 14)) = "Iesit" Or CStr(v(i, 14)) = "Pierdut" Then
            nP = nP + 1
            If CStr(v(i, 14)) = "Iesit" Then nW = nW + 1
            If IsNumeric(v(i, 15)) Then prof = prof + CDbl(v(i, 15))
        End If
        If IsNumeric(v(i, 13)) And Not IsEmpty(v(i, 13)) Then
            nCl = nCl + 1: sCl = sCl + CDbl(v(i, 13))
            If CDbl(v(i, 13)) > 0 Then nClP = nClP + 1
        End If
    Next i
    If nP = 0 And nCl = 0 Then Exit Function
    RezumatValue = "Value pana acum: " & nP & " pariuri, " & nW & " iesite, profit " & Format(prof, "+0.0;-0.0") & " unitati"
    If nP > 0 Then RezumatValue = RezumatValue & " (ROI " & Format(prof / nP, "+0%;-0%") & ")"
    If nCl > 0 Then RezumatValue = RezumatValue & ", CLV mediu " & Format(sCl / nCl, "+0.0%;-0.0%") & _
        " (" & nClP & " din " & nCl & " cu cota scazuta dupa pont)"
    ws.Range("A2").Value = RezumatValue
    ws.Range("A2").Font.Bold = True
End Function


' meciurile de azi doar din ligile noastre (cereri mici, una pe liga); daca nu merge, toate meciurile zilei
Private Function MeciuriAziApi(ByVal cheie As String, ByRef res As Object, ByRef ramase As String) As Boolean
    Dim wl As Worksheet, lastL As Long, v As Variant, i As Long, coduri As Object, cod As Variant, idL As Long
    Dim rr As Object, fx As Variant, colM As Collection, sez As Long, nOk As Long
    Set wl = Sheets("Listing")
    lastL = wl.Cells(wl.Rows.Count, "E").End(xlUp).Row
    Set coduri = CreateObject("Scripting.Dictionary")
    If lastL >= 2 Then
        v = wl.Range("F1:F" & lastL).Value2
        For i = 2 To lastL
            If CStr(v(i, 1)) <> "" And Not coduri.Exists(CStr(v(i, 1))) Then coduri.Add CStr(v(i, 1)), 1
        Next i
    End If
    sez = Year(Date): If Month(Date) < 7 Then sez = sez - 1
    Set colM = New Collection
    For Each cod In coduri.Keys
        idL = LigaApi(CStr(cod))
        If idL > 0 Then
            Application.StatusBar = "BET PRO: meciurile din " & NumeLiga(CStr(cod)) & "..."
            DoEvents
            If ApiFootball("/fixtures?date=" & Format(Date, "yyyy-mm-dd") & "&league=" & idL & "&season=" & sez & _
                    "&timezone=Europe/London", cheie, rr, ramase) Then
                nOk = nOk + 1
                For Each fx In rr("response")
                    colM.Add fx
                Next fx
            End If
        End If
    Next cod
    If nOk > 0 Then
        Set res = CreateObject("Scripting.Dictionary")
        res.Add "response", colM
        MeciuriAziApi = True
    Else
        MeciuriAziApi = ApiFootball("/fixtures?date=" & Format(Date, "yyyy-mm-dd") & "&timezone=Europe/London", cheie, res, ramase)
    End If
End Function

' codul ligii de pe football-data -> id-ul ligii in API-Football
Private Function LigaApi(ByVal cod As String) As Long
    Select Case cod
        Case "E0": LigaApi = 39
        Case "E1": LigaApi = 40
        Case "E2": LigaApi = 41
        Case "E3": LigaApi = 42
        Case "EC": LigaApi = 43
        Case "SC0": LigaApi = 179
        Case "SC1": LigaApi = 180
        Case "SC2": LigaApi = 183
        Case "SC3": LigaApi = 184
        Case "D1": LigaApi = 78
        Case "D2": LigaApi = 79
        Case "I1": LigaApi = 135
        Case "I2": LigaApi = 136
        Case "SP1": LigaApi = 140
        Case "SP2": LigaApi = 141
        Case "F1": LigaApi = 61
        Case "F2": LigaApi = 62
        Case "N1": LigaApi = 88
        Case "B1": LigaApi = 144
        Case "P1": LigaApi = 94
        Case "T1": LigaApi = 203
        Case "G1": LigaApi = 197
    End Select
End Function

' ============================ BUTOANE ============================
' Ruleaza o singura data (Alt+F8 > AdaugaButoane), apoi salveaza fisierul.
Public Sub AdaugaButoane()
    Dim ws As Worksheet, nume As Variant, i As Long
    Dim texte As Variant, macro As Variant, numeB As Variant, b As Object
    Set ws = Sheets("Listing")
    texte = Array("Meciurile zilei", "Actualizeaza istoric", "Test conexiune", "Scanare value", "Backtest", "Recomandari", "Trimite email", "LIVE", "Stop LIVE")
    macro = Array("getMS", "ActualizeazaIstoric", "TestCheie", "ScaneazaValue", "Backtest", "Recomandari", "TrimiteEmailAcum", "LiveStart", "LiveStop")
    numeB = Array("btnMeciuri", "btnRezultate", "btnTest", "btnScanare", "btnBacktest", "btnRecom", "btnEmail", "btnLive", "btnLiveStop")
    On Error Resume Next
    For Each nume In numeB
        ws.Buttons(nume).Delete
    Next nume
    On Error GoTo 0
    For i = 0 To 8
        Set b = ws.Buttons.Add(330 + i * 125, 16, 118, 25)
        b.Name = numeB(i)
        b.Characters.Text = texte(i)
        b.OnAction = macro(i)
    Next i
    FoaieEmail
    Mesaj "Butoanele au fost actualizate si foaia Email a fost creata. Completeaza foaia Email si salveaza (Ctrl+S).", vbInformation, "BET PRO"
End Sub

' ============================ AJUTOARE ============================
Private Function DescarcaCsv(ByVal url As String, ByRef lines As Variant, ByRef hdr As Object) As Boolean
    Dim txt As String, h As Variant, i As Long, incercari As Variant, k As Long
    Dim u As String, motor As String

    ' incearcam pe rand: https, http, apoi alt motor de descarcare din Windows
    incercari = Array("MSXML2.XMLHTTP.6.0|https", "MSXML2.XMLHTTP.6.0|http", _
                      "MSXML2.ServerXMLHTTP.6.0|https", "MSXML2.ServerXMLHTTP.6.0|http", _
                      "WinHttp.WinHttpRequest.5.1|https")
    mEroare = ""
    For k = 0 To UBound(incercari)
        motor = Split(incercari(k), "|")(0)
        u = Replace(url, "https://", Split(incercari(k), "|")(1) & "://")
        txt = DescarcaText(motor, u)
        If txt <> "" Then
            i = InStr(1, txt, "Div")
            If i > 0 And i <= 10 Then Exit For
            mEroare = motor & " " & u & ": raspuns neasteptat: " & Left$(Replace(Replace(txt, vbCr, " "), vbLf, " "), 120)
            txt = ""
        End If
    Next k
    If txt = "" Then Exit Function

    txt = Mid$(txt, InStr(1, txt, "Div"))
    txt = Replace(txt, vbCr, "")
    lines = Split(txt, vbLf)
    Set hdr = CreateObject("Scripting.Dictionary")
    h = Split(lines(0), ",")
    For i = 0 To UBound(h)
        If Not hdr.Exists(Trim$(h(i))) Then hdr.Add Trim$(h(i)), i
    Next i
    mEroare = ""
    DescarcaCsv = True
End Function

Private Function DescarcaText(ByVal motor As String, ByVal url As String) As String
    Dim http As Object
    On Error GoTo Esec
    Set http = CreateObject(motor)
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Excel"
    http.setRequestHeader "Cache-Control", "no-cache"
    http.send
    If http.Status <> 200 Then
        mEroare = motor & " " & url & ": HTTP " & http.Status
        Exit Function
    End If
    DescarcaText = http.responseText
    Exit Function
Esec:
    mEroare = motor & " " & url & ": " & Err.Description
End Function

Private Function Col(hdr As Object, ByVal nume As String) As Long
    If hdr.Exists(nume) Then Col = hdr(nume) Else Col = -1
End Function

Private Function Camp(f As Variant, ByVal i As Long) As String
    If i >= 0 And i <= UBound(f) Then Camp = Trim$(f(i))
End Function

Private Function CotaDin(f As Variant, hdr As Object, ParamArray nume() As Variant) As Double
    Dim k As Long, i As Long
    For k = LBound(nume) To UBound(nume)
        i = Col(hdr, CStr(nume(k)))
        If i >= 0 Then
            If Val(Camp(f, i)) > 1 Then CotaDin = Val(Camp(f, i)): Exit Function
        End If
    Next k
End Function

Private Function Cota(ByVal x As Double) As Variant
    If x > 1 Then Cota = x Else Cota = Empty
End Function

' "25/10/2026" sau "25/10/26" + "15:00" -> data cu ora
Private Function CitesteData(ByVal ds As String, ByVal ts As String) As Double
    Dim p As Variant, y As Long
    p = Split(Trim$(ds), "/")
    If UBound(p) < 2 Then Exit Function
    If Not (IsNumeric(p(0)) And IsNumeric(p(1)) And IsNumeric(p(2))) Then Exit Function
    y = CLng(p(2))
    If y < 100 Then y = y + 2000
    CitesteData = DateSerial(y, CLng(p(1)), CLng(p(0)))
    If Len(ts) >= 4 And InStr(ts, ":") > 0 Then
        CitesteData = CitesteData + TimeSerial(Val(Left$(ts, InStr(ts, ":") - 1)), Val(Mid$(ts, InStr(ts, ":") + 1, 2)), 0)
    End If
End Function

' sezonul curent (off=0) sau trecut (off=1), ex. "2627"
Private Function CodSezon(ByVal off As Long) As String
    Dim y As Long
    y = Year(Date)
    If Month(Date) < 7 Then y = y - 1
    y = y - off
    CodSezon = Right$(CStr(y), 2) & Right$(CStr(y + 1), 2)
End Function

Private Function CoduriLigi() As Variant
    Dim p As Variant, i As Long, r() As String
    p = Split(LIGI, ";")
    ReDim r(0 To UBound(p))
    For i = 0 To UBound(p)
        r(i) = Split(p(i), "=")(0)
    Next i
    CoduriLigi = r
End Function

Private Function NumeLiga(ByVal cod As String) As String
    Dim p As Variant, i As Long
    p = Split(LIGI, ";")
    For i = 0 To UBound(p)
        If Split(p(i), "=")(0) = cod Then NumeLiga = Split(p(i), "=")(1): Exit Function
    Next i
    NumeLiga = cod
End Function

Private Function FoaieIstoric() As Worksheet
    Dim ws As Worksheet, act As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SH_IST)
    On Error GoTo 0
    If ws Is Nothing Then
        Set act = ActiveSheet
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = SH_IST
        ws.Range("A1:I1").Value = Array("Div", "Liga", "Data", "Gazde", "Oaspeti", "HT gazde", "HT oaspeti", "FT gazde", "FT oaspeti")
        ws.Visible = xlSheetHidden
        act.Activate
    End If
    Set FoaieIstoric = ws
End Function

Private Function RanduriIstoric() As Long
    Dim ws As Worksheet
    Set ws = FoaieIstoric()
    RanduriIstoric = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
End Function

Private Function RandMeci(ByVal id As String) As Long
    Dim ws As Worksheet, vals As Variant, i As Long, last As Long
    Set ws = Sheets("Listing")
    last = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
    If last < 2 Then Exit Function
    vals = ws.Range("E1:E" & last).Value
    For i = 2 To last
        If CStr(vals(i, 1)) = CStr(id) Then RandMeci = i: Exit Function
    Next i
End Function

Private Function Array2D(ParamArray a() As Variant) As Variant
    Dim r() As Variant, i As Long
    ReDim r(1 To (UBound(a) + 1) \ 3, 1 To 3)
    For i = 0 To UBound(a)
        r(i \ 3 + 1, i Mod 3 + 1) = a(i)
    Next i
    Array2D = r
End Function

Private Function Gol(ByVal v As Variant) As Boolean
    Gol = IsEmpty(v) Or IsNull(v)
    If Not Gol Then Gol = (CStr(v) = "")
End Function

Private Sub StartLucru()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = "BET PRO: descarc datele..."
End Sub

Private Sub GataLucru()
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
End Sub
