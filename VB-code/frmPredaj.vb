Public Class frmPredaj
    'Upravit riadky 1 30 2x75 '131
    Dim rad, rad1, rad2 As String
    Dim Y, Y1, Y2 As Integer
    Dim X, x1, x2 As Integer
    Dim s, s1, s2 As String
    Dim CenaListka As Double

    Private PocetVsetkychMiest As Short
    Private ZlavaListka As Double
    Private ZlavMiest As Integer
    Dim HASH As String
    Dim SluzobMiesta As String
    Dim Rezervacie As Integer
    Dim ObsadeneMiesta As Integer
    Dim PredaneMiesta As Integer
    Dim ZobrazSedadla As Boolean 'Pre Amfiteatre s volnym sedenim a obrovskou kapacitou

    'Farby sedadiel
    Private FarbaZakl As String = CStr(&H8080) '&HC000&
    Private FarbaPredane As String = CStr(&H8080FF) '&H8080FF  '  &HC0&
    Private FarbaZlava As String = CStr(&HC0E0FF)
    Private FarbaVyber As String = CStr(&HFFC0C0) ' &H80FFFF
    Private FarbaRezerv As Integer = RGB(111, 49, 152)
    Private FarbaSluzob As Integer = RGB(230, 230, 230)

    Dim VizualUprava As Boolean = True
    Dim KontrolaRiadku As Boolean

    Dim Sedadlo As btnArray1

    Private Sub frmPredaj_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = e.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = e.CloseReason
        e.Cancel = Cancel
    End Sub

    Private Sub frmPredaj_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
1:      'Zafarbi
        Dim sala As Integer = frmSpracovavanie.PredstaveniaBindingSource.Current("Sala")

        If frmSpracovavanie.PoslOtvSala <> 0 And frmSpracovavanie.PoslOtvSala <> sala Then
            For i = 1 To 2000 'frmSpracovavanie.PocetMiest(frmSpracovavanie.PoslOtvSala)
                'Dim b As Button = TryCast(Me.Controls(37), Button)
                ' If b IsNot Nothing AndAlso b.Tag IsNot Nothing Then
                '  b.Dispose()       '' NOTE: disposing the button also removes it
                'End If
                Try
                    Sedadlo(i).Dispose()
                Catch ex As Exception
                    Exit For
                End Try
            Next
            frmSpracovavanie.VkladanieSedadiel1 = False
        End If

2:      Call mdlColors.Skining(Me)
3:
        PocetVsetkychMiest = frmSpracovavanie.PocetMiest(sala)
        M1MenoSaly = frmSpracovavanie.MenoSal(sala)
7:      If sala = 1 Then
9:          SluzobMiesta = frmSpracovavanie.SluzobMiesta1
11:         ZobrazSedadla = frmSpracovavanie.ZobrazMiest1
12:     ElseIf sala = 2 Then
14:         SluzobMiesta = frmSpracovavanie.SluzobMiesta2
16:         ZobrazSedadla = frmSpracovavanie.ZobrazMiest2
17:     ElseIf sala = 3 Then
19:         SluzobMiesta = frmSpracovavanie.SluzobMiesta3
21:         ZobrazSedadla = frmSpracovavanie.ZobrazMiest3
22:     ElseIf sala = 4 Then
24:         SluzobMiesta = frmSpracovavanie.SluzobMiesta4
26:         ZobrazSedadla = frmSpracovavanie.ZobrazMiest4
27:     ElseIf sala = 5 Then
29:         SluzobMiesta = frmSpracovavanie.SluzobMiesta5
31:         ZobrazSedadla = frmSpracovavanie.ZobrazMiest5
32:     End If
33:
        If frmSpracovavanie.VkladanieSedadiel1 = False And ZobrazSedadla Then Sedadlo = New btnArray1() 'Me)
        frmSpracovavanie.VkladanieSedadiel1 = True
        frmSpracovavanie.PoslOtvSala = sala

34:     Me.Text = "Pokladna: Predaj - " & M1MenoSaly
35:
36:     With frmSpracovavanie.PredstaveniaBindingSource
37:         lblDatum.Text = Format(CDate(.Current("Datum")), "dd.MM.yyyy")
38:         lblCas.Text = Format(CDate(.Current("Datum")), "HH:mm")
39:         lblFilm.Text = .Current("Nazovfilmu")
40:         CenaListka = CDbl(.Current("CenaListka"))
41:         ZlavaListka = CDbl(.Current("CenaZlavListka"))
43:     End With
45:     lblCenaListka.Text = Format(CenaListka, "###0.00") & " EUR"
        M1Datum = lblDatum.Text
        M1Cas = lblCas.Text
        M1Film = lblFilm.Text
        ' M1OSob = lblOsob.Text
        ' M1Rad = lblRad.Text
46:
47:     'Farby legendy
48:     btnVolVz.BackColor = System.Drawing.ColorTranslator.FromOle(CInt(FarbaZakl))
49:     btnPredVz.BackColor = System.Drawing.ColorTranslator.FromOle(CInt(FarbaPredane))
50:     btnRezVz.BackColor = System.Drawing.ColorTranslator.FromOle(CInt(FarbaRezerv))
51:     btnZlvVz.BackColor = System.Drawing.ColorTranslator.FromOle(CInt(FarbaZlava))
52:
53:     'Zisti zlavu a zablokuje vydavanie zlavnenych ak zlava neni
54:     If CenaListka = ZlavaListka Then
55:         lblZlava.Text = ""
56:         txtZlavnenych.Enabled = False
57:         txtZlavnenych.ReadOnly = True
58:     Else
59:         lblZlava.Text = "(" & ZlavaListka & " EUR" & ")"
60:         txtZlavnenych.Enabled = True
61:         txtZlavnenych.ReadOnly = False
62:     End If
63:
64:     'Ci bude vidno Rezervacie, Presadenie a Bezhotovost
        If ZobrazSedadla Then
65:         btnRezervovat.Enabled = frmSpracovavanie.Rezervacie
66:         btnRezervovat.Visible = frmSpracovavanie.Rezervacie
67:         btnPresadit.Enabled = frmSpracovavanie.Presadenie
68:         btnPresadit.Visible = frmSpracovavanie.Presadenie
        Else
            btnRezervovat.Enabled = False
            btnRezervovat.Visible = False
            btnPresadit.Enabled = False
            btnPresadit.Visible = False
            txtPocetOsobBezSed.Visible = True
            txtPocetOsobBezSed.Enabled = True
        End If
69:     cbxBezhotovost.Enabled = frmSpracovavanie.Bezhotovost
70:     cbxBezhotovost.Visible = frmSpracovavanie.Bezhotovost
        cbxRozdelLstk.Enabled = frmSpracovavanie.RozdelLstk
        cbxRozdelLstk.Visible = frmSpracovavanie.RozdelLstk
71:
72:     'Upravi velkost
73:     If VizualUprava Then
74:         Call mdlColors.Sizing(Me)
75:         Me.CenterToParent()
76:         VizualUprava = False
77:     End If
78:
79:     Call btnZrusit_Click(sender, e)
80:
        If frmSpracovavanie.DruheOknoPredaja Then
            With frmPredajObr
                .lblDatum.Text = lblDatum.Text
                .lblCas.Text = lblCas.Text
                .lblFilm.Text = lblFilm.Text
                .lblCenaListka.Text = lblCenaListka.Text
                .lblZlava.Text = lblZlava.Text
                .lblVolMiesta.Text = lblVolMiesta.Text
                .Show()
            End With
        End If

    End Sub

    Private Sub frmPredaj_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        If frmSpracovavanie.DruheOknoPredaja Then
            frmPredajObr.host = Me
            frmPredajObr.Timer1.Enabled = True
        End If
    End Sub

    Private Sub btnZrusit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnZrusit.Click
1:      If frmSpracovavanie.SietVerzia Then frmSpracovavanie.PredstaveniaTableAdapter.Fill(frmSpracovavanie.DataSet1.Predstavenia)
2:      lblRad.Text = ""
3:      Y1 = 0
4:      Y2 = 0
5:      rad = CStr(0)
6:      ZlavMiest = 0
7:      lblOD2.Text = ""
8:      lblDO2.Text = ""
9:      If CenaListka = ZlavaListka Then txtZlavnenych.Text = CStr(0) Else txtZlavnenych.Text = ""
10:     lblOsob.Text = ""
11:     lblCenaE.Text = ""
12:     cbxBezhotovost.Checked = False
        cbxRozdelLstk.Checked = False
13:     lblRad2.Text = ""
14:     lblIndex.Text = ""
15:     lblMst.Text = ""
16:
17:     frmSpracovavanie.PredstaveniaBindingSource.Position = frmVybPredst.PozPredstav
18:     frmSpracovavanie.FilmyBindingSource.Position = frmVybPredst.PozFilm
19:
20:     Call FarbenieSaly(sender, e)

        If frmSpracovavanie.DruheOknoPredaja Then
            frmPredajObr.host = Me
            frmPredajObr.Timer1.Enabled = True
        End If

    End Sub

    Private Sub lblCenaE_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblCenaE.TextChanged
1:      If lblCenaE.Text = "" Then
2:          lblCenaS.Text = ""
3:      Else
4:          lblCenaS.Text = Format(CDbl(lblCenaE.Text) * 30.126, "0.00")
5:      End If
    End Sub

    Private Sub Cenadokopy(ByVal sender As System.Object, ByVal e As System.EventArgs)
1:      '  On Error GoTo Chyba
2:
3:      'Pocet zliav z textboxu da do premennej (koli jednoduchsiemu ovladaniu pre usera)
4:      If txtZlavnenych.Text = "" Then
5:          ZlavMiest = 0
6:      ElseIf CDbl(txtZlavnenych.Text) < 0 Then
7:          ZlavMiest = 0
8:      Else
9:          ZlavMiest = CInt(txtZlavnenych.Text)
10:     End If
11:
12:     If Y1 = 0 And Y2 = 0 Then
13:         lblCenaE.Text = ""
14:         lblOsob.Text = ""
15:         Exit Sub
16:     ElseIf Y1 = 0 Or Y2 = 0 Then
17:         lblOsob.Text = CStr(1)
18:         If ZlavMiest = 0 Then lblCenaE.Text = CStr(CenaListka)
19:         If ZlavMiest = 1 Then lblCenaE.Text = CStr(ZlavaListka)
20:         If ZlavMiest > 1 Then GoTo VelaZlav
21:     ElseIf Y1 > 0 And Y2 > 0 Then
22:         If Y1 <= Y2 Then
23:             If (Y2 - Y1 + 1 - ZlavMiest) < 0 Then GoTo VelaZlav
24:             lblCenaE.Text = CStr((Y2 - Y1 + 1 - ZlavMiest) * CenaListka + (ZlavMiest * ZlavaListka))
25:             lblOsob.Text = CStr(Y2 - Y1 + 1)
26:         Else
27:             If (Y1 - Y2 + 1 - ZlavMiest) < 0 Then GoTo VelaZlav
28:             lblCenaE.Text = CStr((Y1 - Y2 + 1 - ZlavMiest) * CenaListka + (ZlavMiest * ZlavaListka))
29:             lblOsob.Text = CStr(Y1 - Y2 + 1)
30:         End If
31:     End If
32:     lblCenaE.Text = Format(CDbl(lblCenaE.Text), "###0.00")
33:
34:     Exit Sub
VelaZlav:
39:     MsgBox("Zlavnenych lístkov nemoze byt viac ako pocet osob!", MsgBoxStyle.Exclamation)
40:     txtZlavnenych.Text = ""
41:     Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub btnPredat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPredat.Click
1:      'On Error GoTo Chyba
2:
3:      If frmSpracovavanie.SietVerzia Then
4:          frmSpracovavanie.PredstaveniaTableAdapter.Fill(frmSpracovavanie.DataSet1.Predstavenia)
5:          frmSpracovavanie.FilmyTableAdapter.Fill(frmSpracovavanie.DataSet1.Filmy)
6:          frmSpracovavanie.ListkyTableAdapter.Fill(frmSpracovavanie.DataSet1.Listky)
7:      End If
8:
9:      If KontrolaPredaja(sender, e) = False Then Exit Sub
10:
11:     M1Bezhotovost = cbxBezhotovost.Checked
12:     Call Cenadokopy(sender, e)

        M1x1 = x1
        M1x2 = x2
        M1ZlvLstk = ZlavaListka
        M1ZlvMst = ZlavMiest
        M1Cena = CDbl(lblCenaE.Text)
        M1OSob = CInt(lblOsob.Text)
        M1Rad = lblRad.Text
13:
        Cursor = Cursors.AppStarting

18:     'Da 2ky na predane miesta a 3ky na zlavy
19:     s1 = LSet(s, x1 - 1)
20:     s2 = Mid(s, x2 + 1)
21:     s = s1 & StrDup(ZlavMiest, "3") & StrDup(x2 - x1 + 1 - ZlavMiest, "2") & s2
        M1s = s

23:     'Ulozi predaj
24:     With frmSpracovavanie.PredstaveniaBindingSource
25:         .Current("miesta") = M1s
26:         .Current("predaj") += M1OSob '(M1x2 - M1x1 + 1)
27:         .Current("trzbapredstavenia") += M1Cena
28:         'Vytvori HASH pre kontrolu upravy riadku
29:         HASH = Module1.RetazecRiadku("Predstavenia", frmVybPredst.PozPredstav)
30:         .Current("crc") = Module1.HASHString(HASH)
31:         Module1.UpdatePredstavenia(frmVybPredst.PozPredstav)
32:         .Position = frmVybPredst.PozPredstav
33:         If M1s <> .Current("miesta") Then MsgBox("Chyba pri ukladani obsadenych miest. Obratte sa na Programatora.", MsgBoxStyle.Critical) : Exit Sub
34:     End With
35:     With frmSpracovavanie.FilmyBindingSource
36:         .Current("predanelistky") += M1OSob '(M1x2 - M1x1 + 1)
37:         .Current("TrzbaFilmu") += M1Cena
38:         'Vytvori HASH pre kontrolu upravy riadku
39:         HASH = Module1.RetazecRiadku("Filmy", frmVybPredst.PozFilm)
40:         .Current("crc") = Module1.HASHString(HASH)
41:         UpdateFilmy(frmVybPredst.PozFilm)
42:         .Position = frmVybPredst.PozFilm
43:     End With
44:     'Call Module1.HASHdbKontrola()
45:     'Call Module1.HASHSubor("Data\data.pac")

        If cbxRozdelLstk.Checked Then
            Dim PredatZlav As Integer = ZlavMiest
            For i = 0 To M1OSob - 1
                M1x1 = x1 + i
                M1x2 = M1x1
                M1OSob = 1
                M1MstOD = Sedadlo(M1x1).Text & " "
                M1MstDO = M1MstOD
                If PredatZlav > 0 Then
                    M1ZlvMst = 1
                    M1Cena = ZlavaListka
                    PredatZlav -= 1
                Else
                    M1ZlvMst = 0
                    M1Cena = CenaListka
                End If
                PredajLiskta(" - Predaj listka (Pr:")
            Next i
        Else
            PredajLiskta(" - Predaj listka (Pr:")
        End If

        frmSpracovavanie.PredstaveniaBindingSource.Position = frmVybPredst.PozPredstav
47:     frmVydat.Zaplatit = CDbl(lblCenaE.Text)

        Cursor = Cursors.Default
107:    If cbxBezhotovost.Checked = False Then frmVydat.ShowDialog()
108:    Call btnZrusit_Click(sender, e)
        '109:    Module1.WriteLog("   OK")
110:    Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub FarbenieSaly(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Prezrie 1112121344443411121212 a urci ktore miesta su predane
        ' On Error GoTo Chyba
        On Error Resume Next
        Dim s As String
        Dim X As Short

        frmSpracovavanie.PredstaveniaBindingSource.Position = frmVybPredst.PozPredstav
        s = frmSpracovavanie.PredstaveniaBindingSource.Current("Miesta")
        Rezervacie = 0
        PredaneMiesta = 0

        If ZobrazSedadla = False Then
            s = StrDup(100, "1")
            'Informacia o volnych miestach
            With frmSpracovavanie.PredstaveniaBindingSource
                If CStr(.Current("Predaj")) <> "" Then
                    ObsadeneMiesta = CDbl(.Current("Predaj")) + Rezervacie
                    lblVolMiesta.Text = (PocetVsetkychMiest - ObsadeneMiesta) & " (" & .Current("Predaj") & "+" & Rezervacie & ")"
                Else : lblVolMiesta.Text = PocetVsetkychMiest & " (0+0)"
                End If
            End With
            Exit Sub
        End If

        For X = 1 To PocetVsetkychMiest
            If Mid(s, X, 1) = "2" Then
                Sedadlo(X).Enabled = frmSpracovavanie.MiestaEnabled
                Sedadlo(X).BackColor = System.Drawing.ColorTranslator.FromOle(CInt(FarbaPredane))
                PredaneMiesta += 1
            ElseIf Mid(s, X, 1) = "3" Then
                Sedadlo(X).Enabled = frmSpracovavanie.MiestaEnabled
                Sedadlo(X).BackColor = System.Drawing.ColorTranslator.FromOle(CInt(FarbaZlava))
                PredaneMiesta += 1
            ElseIf Mid(s, X, 1) = "4" Then
                Sedadlo(X).Enabled = frmSpracovavanie.MiestaEnabled
                Sedadlo(X).BackColor = System.Drawing.ColorTranslator.FromOle(CInt(FarbaRezerv))
                Rezervacie += 1
            Else
                Sedadlo(X).Enabled = True
                Sedadlo(X).BackColor = System.Drawing.ColorTranslator.FromOle(CInt(FarbaZakl))
            End If
        Next X

        Dim nElements As Short
        Dim aKey As Object
        aKey = Split(SluzobMiesta, " ")
        nElements = UBound(aKey) - LBound(aKey) + 1
        For X = 0 To nElements - 1
            If aKey(X) <> "" Then
                If Sedadlo(CInt(aKey(X))).BackColor = System.Drawing.ColorTranslator.FromOle(CInt(FarbaZakl)) Then
                    Sedadlo(CInt(aKey(X))).BackColor = System.Drawing.ColorTranslator.FromOle(CInt(FarbaSluzob))
                    Sedadlo(CInt(aKey(X))).Enabled = frmSpracovavanie.PredajSluzob
                End If
            End If
        Next X

        With frmSpracovavanie.PredstaveniaBindingSource
            If CStr(.Current("Predaj")) <> "" Then
                ObsadeneMiesta = CDbl(.Current("Predaj")) + Rezervacie
                If .Current("Predaj") <> PredaneMiesta Then MsgBox("Pocet obsadenych miest nesuhlasi s poctom predanych miest. Kontaktujte Programatora.", MsgBoxStyle.Critical)
                lblVolMiesta.Text = (PocetVsetkychMiest - ObsadeneMiesta) & " (" & .Current("Predaj") & "+" & Rezervacie & ")"
            Else : lblVolMiesta.Text = PocetVsetkychMiest & " (0+0)"
            End If
        End With

        Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub txtZlavnenych_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtZlavnenych.TextChanged
        Call Cenadokopy(sender, e)
    End Sub

    Private Sub lblIndex_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblIndex.TextChanged
        If lblIndex.Text = "" Or lblIndex.Text = "Index" Then Exit Sub

        If Y1 = 0 Then
            lblOD2.Text = lblMst.Text
            lblRad.Text = lblRad2.Text
            Y1 = CInt(lblIndex.Text)
            Call Cenadokopy(sender, e)
        Else
            lblDO2.Text = lblMst.Text
            Y2 = CInt(lblIndex.Text)
            If lblRad.Text <> lblRad2.Text Then
                MsgBox("Na jeden lístok len miesta z jedného radu!")
                Call btnZrusit_Click(sender, e)
                Exit Sub
            End If
            Call Cenadokopy(sender, e)
        End If

        'Uvolnovanie miest klikom ak je to povolene cez Command line 
        If frmSpracovavanie.MiestaEnabled Then
            Dim s, s1, s2, sp, HASH As String
            s = frmSpracovavanie.PredstaveniaBindingSource.Current("miesta")
            sp = Mid(s, CInt(lblIndex.Text), 1)
            s1 = LSet(s, CInt(lblIndex.Text) - 1)
            s2 = Mid(s, CInt(lblIndex.Text) + 1)
            If sp = "1" Then : s = s1 & "2" & s2
            ElseIf sp = "2" Then : s = s1 & "3" & s2
            ElseIf sp = "3" Then : s = s1 & "4" & s2
            ElseIf sp = "4" Then : s = s1 & "1" & s2
            End If

            With frmSpracovavanie.PredstaveniaBindingSource
                .Current("miesta") = s
                HASH = Module1.RetazecRiadku("Predstavenia", frmVybPredst.PozPredstav)
                .Current("crc") = Module1.HASHString(HASH)
                Module1.UpdatePredstavenia(frmVybPredst.PozPredstav)
                .Position = frmVybPredst.PozPredstav
            End With
            Call btnZrusit_Click(sender, e)
            Exit Sub
        End If

    End Sub

    Public Sub FocusSedadlo(ByVal sender As Object, ByVal e As System.EventArgs, ByVal index As Integer)
        'On Error GoTo Chyba
        On Error Resume Next

        If frmSpracovavanie.MiestaEnabled Then Exit Sub

        'Ked sa sipkou nastavi na miesto tak sa odklikne "samo"

        'Najprv zmeni farbu
        Sedadlo(index).BackColor = System.Drawing.ColorTranslator.FromOle(CInt(FarbaVyber))

        'Zisti ci ma vratit farbu miesta naspet na zakladnu
        'Riesi vratenie farby miest ak medzi dvoma vybranymi su ENABLED = FALSE miesta
        'Predpokladam ze ich neni viac ako 50

        'Riesi 1 miesto vybrane
        If Y1 = index Then
            For i = 1 To 50
                If Sedadlo(index + i).BackColor = System.Drawing.ColorTranslator.FromOle(CInt(FarbaZakl)) Then Exit For
                If Sedadlo(index + i).Enabled = True Then
                    Sedadlo(index + i).BackColor = System.Drawing.ColorTranslator.FromOle(CInt(FarbaZakl))
                End If
            Next i
            For i = 1 To 50
                If Sedadlo(index - i).BackColor = System.Drawing.ColorTranslator.FromOle(CInt(FarbaZakl)) Then Exit For
                If Sedadlo(index - i).Enabled = True Then
                    Sedadlo(index - i).BackColor = System.Drawing.ColorTranslator.FromOle(CInt(FarbaZakl))
                End If
            Next i
            'Riesi ak sa vyberie miesto medzi dvoma vybranymi miestami
        ElseIf Y1 < index And Y2 > index Then
            For i = 1 To 50
                If Sedadlo(index + i).BackColor = System.Drawing.ColorTranslator.FromOle(CInt(FarbaZakl)) Then Exit For
                If Sedadlo(index + i).Enabled = True Then
                    Sedadlo(index + i).BackColor = System.Drawing.ColorTranslator.FromOle(CInt(FarbaZakl))
                End If
            Next i
        ElseIf Y1 > index And Y2 < index Then
            For i = 1 To 50
                If Sedadlo(index - i).BackColor = System.Drawing.ColorTranslator.FromOle(CInt(FarbaZakl)) Then Exit For
                If Sedadlo(index - i).Enabled = True Then
                    Sedadlo(index - i).BackColor = System.Drawing.ColorTranslator.FromOle(CInt(FarbaZakl))
                End If
            Next i
        End If

        'Odklikne sa
        Call Sedadlo.ClickHandler(sender, e)

        'riesi prekliknutie pri vybere miest
        If Y2 = 0 Then Exit Sub
        If Y1 > Y2 Then
            For index = Y2 To Y1
                If Sedadlo(index).Enabled = True Then Sedadlo(index).BackColor = System.Drawing.ColorTranslator.FromOle(CInt(FarbaVyber))
            Next index
            For i = 1 To 50
                If Sedadlo(Y1 + i).BackColor = System.Drawing.ColorTranslator.FromOle(CInt(FarbaZakl)) Then Exit For
                If Sedadlo(Y1 + i).Enabled = True Then
                    Sedadlo(Y1 + i).BackColor = System.Drawing.ColorTranslator.FromOle(CInt(FarbaZakl))
                End If
            Next i
        ElseIf Y1 < Y2 Then
            For index = Y1 To Y2
                If Sedadlo(index).Enabled = True Then Sedadlo(index).BackColor = System.Drawing.ColorTranslator.FromOle(CInt(FarbaVyber))
            Next index
            For i = 1 To 50
                If Sedadlo(Y1 - i).BackColor = System.Drawing.ColorTranslator.FromOle(CInt(FarbaZakl)) Then Exit For
                If Sedadlo(Y1 - i).Enabled = True Then
                    Sedadlo(Y1 - i).BackColor = System.Drawing.ColorTranslator.FromOle(CInt(FarbaZakl))
                End If
            Next i
        End If

        Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub btnRezervovat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRezervovat.Click
1:      On Error GoTo Chyba
2:      Dim i As Short
3:
4:      If frmSpracovavanie.SietVerzia Then
5:          frmSpracovavanie.PredstaveniaTableAdapter.Fill(frmSpracovavanie.DataSet1.Predstavenia)
6:          frmSpracovavanie.RezervacieTableAdapter.Fill(frmSpracovavanie.DataSet1.Rezervacie)
7:      End If
8:
9:      If KontrolaPredaja(sender, e) = False Then Exit Sub
10:
11:     frmRezerv.MenaMiest = lblRad.Text & " "
12:     For i = x1 To x2
13:         frmRezerv.MenaMiest += Sedadlo(i).Text & " "
14:     Next i
15:     frmRezerv.MiestoOdRez = x1
16:     frmRezerv.MiestoDoRez = x2
        frmRezerv.Datum = lblDatum.Text
        frmRezerv.Cas = lblCas.Text
17:     frmRezerv.ShowDialog()
18:     Call btnZrusit_Click(sender, e)
19:
        Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub btnPresadit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPresadit.Click
        ' On Error GoTo Chyba

        If frmSpracovavanie.SietVerzia Then frmSpracovavanie.PredstaveniaTableAdapter.Fill(frmSpracovavanie.DataSet1.Predstavenia)

        If KontrolaPredaja(sender, e) = False Then Exit Sub

        ' Data pre frmPresad
        frmPresad.DatumPresad = lblDatum.Text
        frmPresad.CasPresad = lblCas.Text
        frmPresad.SalaPresad = frmSpracovavanie.PredstaveniaBindingSource.Current("Sala")
        frmPresad.MstOdPresad = x1
        frmPresad.MstDoPresad = x2
        frmPresad.NazMstPresad = lblRad.Text & " " & lblOD2.Text & " - "
        If x1 = x2 Then frmPresad.NazMstPresad = frmPresad.NazMstPresad & lblOD2.Text Else frmPresad.NazMstPresad = frmPresad.NazMstPresad & lblDO2.Text
        If txtZlavnenych.Text = "" Then frmPresad.ZlavPresad = 0 Else frmPresad.ZlavPresad = txtZlavnenych.Text
        frmPresad.SumaPresad = lblCenaE.Text

        'Co sa ma vytalcit na listok
        frmVybPredst.TlacitDatum = lblDatum.Text
        frmVybPredst.TlacitCas = lblCas.Text
        frmVybPredst.TlacitFilm = lblFilm.Text
        frmVybPredst.TlacitOsob = lblOsob.Text
        frmVybPredst.TlacitRad = lblRad.Text
        frmVybPredst.TlacitSala = M1MenoSaly
        frmVybPredst.TlacitZlavMiest = ZlavMiest
        frmVybPredst.TlacitZlavCena = ZlavaListka
        frmVybPredst.TlacitCenaCelkom = CDbl(lblCenaE.Text)
        frmVybPredst.TlacitMiestoOd = lblOD2.Text
        If x1 = x2 Then frmVybPredst.TlacitMiestoDo = lblOD2.Text Else frmVybPredst.TlacitMiestoDo = lblDO2.Text
        frmVybPredst.TlacitBezhotov = cbxBezhotovost.Checked

        frmPresad.ShowDialog()
        Call btnZrusit_Click(sender, e)

        Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Function KontrolaPredaja(ByVal sender As System.Object, ByVal e As System.EventArgs) As Boolean
        Dim x As Integer
        KontrolaPredaja = True

        If (lblDatum.Text & " " & lblCas.Text) <> frmSpracovavanie.PredstaveniaBindingSource.Current("Datum") Then
            MsgBox("Chyba pri urceni pozicii v databaze. Zavrite a otvorte okno predaja.", MsgBoxStyle.Critical)
            GoTo NepresloTestom
        End If

        If frmSpracovavanie.DatumPrihlasenia <> Today And frmLogin.Rights <> "A" Then
            MsgBox("S datumom bolo manipulovane. Zmente datum Vasho pocitaca.", MsgBoxStyle.Exclamation)
            GoTo NepresloTestom
        End If

        If CenaListka <> CDbl(frmSpracovavanie.PredstaveniaBindingSource.Current("CenaListka")) Or ZlavaListka <> CDbl(frmSpracovavanie.PredstaveniaBindingSource.Current("CenaZlavListka")) Then
            MsgBox("Chyba pri urcovani cien, skuste otvorit a zatvorit okno.", MsgBoxStyle.Exclamation)
            GoTo NepresloTestom
        End If

        If Y1 = 0 Or lblOD2.Text = "" Then
            MsgBox("Neboli vybrate miesta")
            GoTo NepresloTestom
        End If

        If Y2 <> 0 Then
            If Y1 <= Y2 Then
                x1 = Y1
                x2 = Y2
                M1MstOD = lblOD2.Text
                M1MstDO = lblDO2.Text
            End If
            If Y1 > Y2 Then
                x1 = Y2
                x2 = Y1
                M1MstOD = lblDO2.Text
                M1MstDO = lblOD2.Text
            End If
            'frmVybPredst.TlacitMiestoDo = lblDO2.Text
        Else
            x1 = Y1
            x2 = Y1
            'frmVybPredst.TlacitMiestoDo = lblOD2.Text
            M1MstOD = lblOD2.Text
            M1MstDO = lblOD2.Text
        End If

        If ZobrazSedadla Then
            For x = x1 To x2
                If Sedadlo(x).Enabled = False Then
                    MsgBox("Nemozno predat tieto miesta.")
                    GoTo NepresloTestom
                End If
            Next x
        End If

        If ZobrazSedadla Then s = frmSpracovavanie.PredstaveniaBindingSource.Current("miesta") Else s = StrDup(100, "1")
        s1 = Mid(s, x1, (x2 - x1 + 1))
        If InStr(s1, "2") > 0 Or InStr(s1, "3") > 0 Or InStr(s1, "4") > 0 Then
            MsgBox("Na miesto uz boli predane listky.")
            GoTo NepresloTestom
        End If

        If PocetVsetkychMiest < (ObsadeneMiesta + x2 - x1 + 1) Then
            MsgBox("Predaj blokovany. Sala je uz zaplnena. Ak este stale vidite volne miesta obratte sa na Programatora.", MsgBoxStyle.Exclamation)
            GoTo NepresloTestom
        End If

        If frmVybPredst.NarusFilm Then
            MsgBox("Data Filmu narusene! Nie je mozny predaj na toto predstavenie.", MsgBoxStyle.Exclamation)
            GoTo NepresloTestom
        End If

        frmSpracovavanie.Timer2_Tick()
        If frmSpracovavanie.NasloDatabazu = False Then
            MsgBox("Databaza sa nenasla.", MsgBoxStyle.Exclamation)
            GoTo NepresloTestom
        End If


        Exit Function
NepresloTestom:
        KontrolaPredaja = False
        Call btnZrusit_Click(sender, e)
    End Function

    Private Sub txtZlavnenych_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtZlavnenych.KeyDown
        Dim KeyCode As Short = e.KeyCode
        Dim Shift As Short = e.KeyData \ &H10000
        Call btnPredat_KeyDown(btnPredat, New System.Windows.Forms.KeyEventArgs(KeyCode Or Shift * &H10000))
    End Sub
    Private Sub txtZlavnenych_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtZlavnenych.KeyPress
        frmPredstavenia.txtCena_KeyPress(sender, e, Me)
    End Sub

    Private Sub txtPocetOsobBezSed_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPocetOsobBezSed.KeyPress
        frmPredstavenia.txtCena_KeyPress(sender, e, Me)
    End Sub
    Private Sub txtPocetOsobBezSed_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPocetOsobBezSed.TextChanged
        On Error GoTo Chyba
        If txtPocetOsobBezSed.Text = "" Or txtPocetOsobBezSed.Text = "0" Then
            Y1 = 0
            Y2 = 0
            Call btnZrusit_Click(sender, e)
        Else
            Y1 = 1
            Y2 = Y1 + CInt(txtPocetOsobBezSed.Text) - 1
            lblOD2.Text = " "
            If Y2 > Y1 Then lblDO2.Text = " "
            Call Cenadokopy(sender, e)
        End If

        Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub btnPredat_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles btnPredat.KeyDown
        Dim KeyCode As Short = e.KeyCode
        Dim Shift As Short = e.KeyData \ &H10000
        If frmLogin.Rights = "M" Or frmLogin.Rights = "A" Then
            If KeyCode = System.Windows.Forms.Keys.F5 Then
                lblIndex.Visible = True
                MsgBox("Zobrazovanie cisla miest aktivne.")
            End If
        End If
        If KeyCode = System.Windows.Forms.Keys.F3 Then frmStatPrehladTrzP.ShowDialog()
        If KeyCode = System.Windows.Forms.Keys.Escape Then Me.Close()
    End Sub
    Private Sub btnZrusit_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles btnZrusit.KeyDown
        Dim KeyCode As Short = e.KeyCode
        Dim Shift As Short = e.KeyData \ &H10000
        Call btnPredat_KeyDown(btnPredat, New System.Windows.Forms.KeyEventArgs(KeyCode Or Shift * &H10000))
    End Sub
    Private Sub btnRezervovat_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles btnRezervovat.KeyDown
        Dim KeyCode As Short = e.KeyCode
        Dim Shift As Short = e.KeyData \ &H10000
        Call btnPredat_KeyDown(btnPredat, New System.Windows.Forms.KeyEventArgs(KeyCode Or Shift * &H10000))
    End Sub
End Class