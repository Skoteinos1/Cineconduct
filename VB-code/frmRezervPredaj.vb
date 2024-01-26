Public Class frmRezervPredaj

    Dim MiestaNaSedenie As String
    Dim CenaN, CenaZ, Suma As Double
    Dim Datum, CasPr As Date
    Dim POsob, PZlav As Integer
    Dim PozRez As Integer
    Dim Film As String
    Dim VizualUprava As Boolean = True

    Private Sub frmRezervPredaj_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Zafarbi
        Call mdlColors.Skining(Me)
        If VizualUprava Then
            Call mdlColors.Sizing(Me)
            Me.CenterToParent()
            VizualUprava = False
        End If

        If frmSpracovavanie.SietVerzia Then
            frmSpracovavanie.PredstaveniaTableAdapter.Fill(frmSpracovavanie.DataSet1.Predstavenia)
            frmSpracovavanie.RezervacieTableAdapter.Fill(frmSpracovavanie.DataSet1.Rezervacie)
        End If

        '************ SPRAVA DATA GRIDU*************************
        With DataGridView1
            .SuspendLayout()
            .DataSource = Nothing
            .AllowUserToAddRows = False
            .AllowUserToDeleteRows = False
            .AllowUserToResizeRows = False
            .AllowUserToResizeColumns = False

            .AllowUserToOrderColumns = True
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .ReadOnly = True
            .MultiSelect = False
            .RowHeadersVisible = False
            .Columns.Clear()
            .Rows.Clear()

            .Columns.Add("dtgrdMeno", "Meno")
            .Columns(0).Width = (.Width - 20) * 0.2
            .Columns.Add("dtgrdUdaje", "Udaje")
            .Columns(1).Width = (.Width - 20) * 0.2
            .Columns.Add("dtgrdDatum", "Datum")
            .Columns(2).Width = (.Width - 20) * 0.11
            .Columns.Add("dtgrdCas", "Cas")
            .Columns(3).Width = (.Width - 20) * 0.07
            .Columns.Add("dtgrdFilm", "Názov filmu")
            .Columns(4).Width = (.Width - 20) * 0.2
            .Columns.Add("dtgrdOsob", "Osob")
            .Columns(5).Width = (.Width - 20) * 0.06
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            .Columns.Add("dtgrdRad", "Rad")
            .Columns(6).Width = (.Width - 20) * 0.06
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            .Columns.Add("dtgrdMiesta", "Miesta")
            .Columns(7).Width = (.Width - 20) * 0.1
            .ResumeLayout(True)
        End With

        'Viditelnost moznosti bezhotovostnej platby
        cbxBezhotovost.Enabled = frmSpracovavanie.Bezhotovost : cbxBezhotovost.Visible = frmSpracovavanie.Bezhotovost : cbxBezhotovost.Checked = False
        cbxRozdelLstk.Enabled = frmSpracovavanie.RozdelLstk : cbxRozdelLstk.Visible = frmSpracovavanie.RozdelLstk : cbxRozdelLstk.Checked = False

        'Potom sa postara o obsah comboboxu z predstaveniami
        Dim i As Integer
        cmbPredstavenie.Text = ""
        cmbPredstavenie.Items.Clear()
        With frmSpracovavanie.PredstaveniaBindingSource
            cmbPredstavenie.Items.Add("")
            For i = 0 To .Count - 1
                .Position = i
                If Format(CDate(.Current("datum")), "dd.MM.yyyy") >= Today Then
                    cmbPredstavenie.Items.Add(.Current("Datum") & ";" & .Current("NazovFilmu"))
                End If
            Next i
        End With

        btnPredat.Enabled = False
        btnZrusit.Enabled = False
        txtZlavnen.Enabled = False
        txtZlavnen.Text = 0

        'Call ZaplnitTabulku(sender, e)

    End Sub

    Private Sub ZaplnitTabulku(ByVal sender As System.Object, ByVal e As System.EventArgs)
1:      On Error GoTo Chyba
2:      Dim i, i2, x As Integer
3:      Dim CasPredst As Date
4:      Dim aKey As Object
5:      Dim nElements As Short
6:      'Dim MiestaTab(50) As String
7:
8:      If cmbPredstavenie.Text <> "" Then
9:          aKey = Split(cmbPredstavenie.Text, ";")
10:         CasPredst = CDate(aKey(0))
11:     End If
12:
13:     With DataGridView1
14:         .Rows.Clear()
15:         For i = 0 To frmSpracovavanie.RezervacieBindingSource.Count - 1
16:             frmSpracovavanie.RezervacieBindingSource.Position = i
17:             If cmbPredstavenie.Text = "" Or (Format(CDate(frmSpracovavanie.RezervacieBindingSource.Current("Datum")), "dd.MM.yyyy") = Format(CasPredst, "dd.MM.yyyy") And Format(CDate(frmSpracovavanie.RezervacieBindingSource.Current("cas")), "HH:mm") = Format(CasPredst, "HH:mm")) Then
18:                 .Rows.Add()
19:                 .Rows(.Rows.Count - 1).Cells(0).Value = frmSpracovavanie.RezervacieBindingSource.Current("Meno")
20:                 .Rows(.Rows.Count - 1).Cells(1).Value = frmSpracovavanie.RezervacieBindingSource.Current("Udaje") & ""
21:                 .Rows(.Rows.Count - 1).Cells(2).Value = Format(CDate(frmSpracovavanie.RezervacieBindingSource.Current("Datum")), "dd.MM.yyyy")
22:                 .Rows(.Rows.Count - 1).Cells(3).Value = Format(CDate(frmSpracovavanie.RezervacieBindingSource.Current("cas")), "HH:mm")
23:                 .Rows(.Rows.Count - 1).Cells(4).Value = frmSpracovavanie.RezervacieBindingSource.Current("Film") & ""
24:                 .Rows(.Rows.Count - 1).Cells(5).Value = CInt(frmSpracovavanie.RezervacieBindingSource.Current("MiestoDo")) - CInt(frmSpracovavanie.RezervacieBindingSource.Current("MiestoOd")) + 1
25:
26:                 'Zobrazi rad a cisla rezervovanych miest v ludskej reci
27:                 aKey = Split(frmSpracovavanie.RezervacieBindingSource.Current("MiestaMena"), " ")
28:                 nElements = UBound(aKey) - LBound(aKey) + 1
29:                 .Rows(.Rows.Count - 1).Cells(6).Value = aKey(0)
30:                 .Rows(.Rows.Count - 1).Cells(7).Value = aKey(1) & "-" & aKey(nElements - 2)
31:             End If
32:         Next i
33:         If .Rows.Count = 0 Then btnPredat.Enabled = False Else btnPredat.Enabled = True
            btnZrusit.Enabled = btnPredat.Enabled
34:     End With
35:
        cbxBezhotovost.Checked = False

        'If DataGridView1.Rows.Count < 5 Then
        Call DataGridView1_Click(sender, e)
        'Else
36:
        'End If
42:
        Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub DataGridView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.Click
1:      On Error GoTo Chyba
2:      If DataGridView1.Rows.Count = 0 Then
            lblMeno.Text = ""
            lblUdaje.Text = ""
            txtOsob.Text = 0
            txtZlavnen.Text = 0
            lblSuma.Text = 0
            lblCena.Text = 0
            Exit Sub
        End If
6:
7:      Dim i, i2 As Integer
9:      i = DataGridView1.CurrentRow.Index
10:
11:     Datum = CDate(DataGridView1.Rows(i).Cells(2).Value)
12:     CasPr = CDate(DataGridView1.Rows(i).Cells(3).Value)
13:     lblMeno.Text = DataGridView1.Rows(i).Cells(0).Value
14:     lblUdaje.Text = DataGridView1.Rows(i).Cells(1).Value
15:
16:     With frmSpracovavanie.PredstaveniaBindingSource
17:         i2 = .Find("Datum", Datum & " " & CasPr)
18:         If i2 <> -1 Then
19:             .Position = i2
20:             frmVybPredst.PozPredstav = .Position
21:             Film = .Current("NazovFilmu")
22:             CenaN = .Current("CenaListka")
23:             CenaZ = .Current("CenaZlavListka")
24:             btnPredat.Enabled = True
25:             btnZrusit.Enabled = True
26:         Else
27:             btnPredat.Enabled = False
28:             btnZrusit.Enabled = False
29:             MsgBox("Nenaslo sa predstavenie.", MsgBoxStyle.Exclamation)
30:             Exit Sub
31:         End If
32:     End With
33:
34:     lblCena.Text = CenaN & " EUR"
35:     If CenaZ <> CenaN Then
36:         lblCena.Text += " (" & CenaZ & ")"
37:         txtZlavnen.Enabled = True
38:     Else
39:         txtZlavnen.Enabled = False
40:     End If
41:
42:     txtOsob.Text = DataGridView1.Rows(i).Cells(5).Value
43:     txtZlavnen.Text = 0
44:
45:     Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub btnZrusit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnZrusit.Click
1:      If MsgBox("Chcete zmazat rezervaciu na meno " & lblMeno.Text, MsgBoxStyle.YesNo + MsgBoxStyle.Question) = MsgBoxResult.No Then Exit Sub
2:      On Error GoTo Chyba
3:      Module1.WriteLog(Chr(13) & Chr(10) & Now & " - Zrusenie rezervacie: " & lblMeno.Text & "; ")
4:      Dim X, x1, x2 As Integer
5:      Dim s, s1, s2 As String
6:      Dim HASH As String
7:
8:      With frmSpracovavanie.RezervacieBindingSource
9:          .Position = .Find("Meno", lblMeno.Text)
10:         x1 = .Current("MiestoOd")
11:         x2 = .Current("MiestoDo")
12:         Module1.WriteLog(.Current("Datum") & " " & Format(CDate(.Current("Cas")), "HH:mm") & " " & .Current("MiestaMena"))
13:     End With
14:
15:     With frmSpracovavanie.PredstaveniaBindingSource
16:         .Position = frmVybPredst.PozPredstav
17:         s = .Current("miesta")
18:
19:         ' **** 4 spat na 1 ****
20:         ' Najprv kontrola ci nejde o predane miesta
21:         s1 = Mid(s, x1, x2 - x1 + 1)
22:         If InStr(s1, "2") <> 0 Or InStr(s1, "3") <> 0 Then
23:             MsgBox("Vyskytla sa chyba pri evidencii predanych miest v sale. Rezervacia nebude zrusena. Obratte sa na Programatora.", MsgBoxStyle.Exclamation)
24:             Exit Sub
25:         End If
26:         s1 = LSet(s, x1 - 1)
27:         s2 = Mid(s, x2 + 1)
28:         s = s1 & StrDup(x2 - x1 + 1, "1") & s2
29:
30:         .Current("miesta") = s
31:         HASH = Module1.RetazecRiadku("Predstavenia", frmVybPredst.PozPredstav)
32:         .Current("crc") = Module1.HASHString(HASH)
33:         Module1.UpdatePredstavenia(frmVybPredst.PozPredstav)
34:         .Position = frmVybPredst.PozPredstav
35:         If s <> .Current("miesta") Then MsgBox("Chyba pri ukladani uvolnenych miest. Obratte sa na Programatora.", MsgBoxStyle.Critical) : Exit Sub
36:     End With
37:
38:     'Zmaze rezervaciu
39:     With frmSpracovavanie.RezervacieBindingSource
40:         frmSpracovavanie.RezervacieTableAdapter.Delete(.Current(0), .Current(1), .Current(2), .Current(3), .Current(4), .Current(5), .Current(6))
41:         frmSpracovavanie.RezervacieTableAdapter.Fill(frmSpracovavanie.DataSet1.Rezervacie)
42:     End With
43:
44:     'Call Module1.HASHdbKontrola()
45:     Call Module1.HASHSubor("Data\data.pac")
46:     Module1.WriteLog("   OK")
47:     Call ZaplnitTabulku(sender, e)
48:
49:     Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub btnPredat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPredat.Click
1:      If frmSpracovavanie.SietVerzia Then
2:          frmSpracovavanie.PredstaveniaTableAdapter.Fill(frmSpracovavanie.DataSet1.Predstavenia)
3:          frmSpracovavanie.RezervacieTableAdapter.Fill(frmSpracovavanie.DataSet1.Rezervacie)
4:          frmSpracovavanie.ListkyTableAdapter.Fill(frmSpracovavanie.DataSet1.Listky)
5:          frmSpracovavanie.FilmyTableAdapter.Fill(frmSpracovavanie.DataSet1.Filmy)
6:      End If
7:
        frmSpracovavanie.Timer2_Tick()
        If frmSpracovavanie.NasloDatabazu = False Then
            MsgBox("Databaza sa nenasla.", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

8:      If frmSpracovavanie.DatumPrihlasenia <> Today Then
9:          MsgBox("S datumom bolo manipulovane. Zmente datum Vasho pocitaca.", MsgBoxStyle.Exclamation)
10:         Exit Sub
11:     End If
12:
13:     If POsob = 0 Then MsgBox("Nemozno predat vstupenku pre 0 osob.", MsgBoxStyle.Exclamation) : Exit Sub
14:     If MsgBox("Vstupenka je pre " & POsob & " osob. Z toho je " & PZlav & " zvliav. Pokracovat?", MsgBoxStyle.YesNo + MsgBoxStyle.Question) = MsgBoxResult.No Then Exit Sub
15:
        M1Rad = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(6).Value
        M1MstOD = DataGridView1.Rows(DataGridView1.CurrentRow.Index).Cells(7).Value
        M1MstDO = "" 'M1MstOD v sebe skryva aj M1mstdo
        M1Bezhotovost = cbxBezhotovost.Checked
        M1OSob = POsob
        txtZlavnen.Text = PZlav
        M1Film = Film
        M1ZlvLstk = CenaZ
        M1ZlvMst = PZlav
20:
21:     Dim X, x1, x2, x3 As Short
22:     Dim s, s1, s2 As String
23:     Dim HASH As String
24:     On Error GoTo Chyba
25:

29:     With frmSpracovavanie.RezervacieBindingSource
30:         PozRez = .Find("Meno", lblMeno.Text)
31:         .Position = PozRez
32:         x1 = .Current("MiestoOd")
33:         x2 = .Current("MiestoOd") + POsob - 1
34:         x3 = .Current("MiestoDo")
35:     End With
36:
37:     s = frmSpracovavanie.PredstaveniaBindingSource.Current("miesta")
38:     s1 = Mid(s, x1, (x2 - x1 + 1))
39:     If InStr(s1, "2") > 0 Or InStr(s1, "3") > 0 Then
40:         MsgBox("Na miesto uz boli predane listky")
41:         Exit Sub
42:     End If
43:
44:     'Da 2ky na predane miesta a 3ky na zlavy a uvolni nepredane miesta
45:     s1 = LSet(s, x1 - 1)
46:     s2 = Mid(s, x3 + 1)
47:     s = s1 & StrDup(PZlav, "3") & StrDup(POsob - PZlav, "2") & StrDup(x3 - x2, "1") & s2
48:
        M1s = s
        M1Cena = Suma
        M1x1 = x1
        M1x2 = x1 + POsob - 1 'x3
        M1Datum = Datum
        M1Cas = Format(CasPr, "HH:mm")

49:     'Ulozi predaj
50:     With frmSpracovavanie.PredstaveniaBindingSource
51:         .Current("miesta") = M1s
52:         .Current("predaj") += M1OSob '(M1x2 - M1x1 + 1)
53:         .Current("trzbapredstavenia") += M1Cena
54:         'Vytvori HASH pre kontrolu upravy riadku
55:         HASH = Module1.RetazecRiadku("Predstavenia", frmVybPredst.PozPredstav)
56:         .Current("crc") = Module1.HASHString(HASH)
57:         Module1.UpdatePredstavenia(frmVybPredst.PozPredstav)
58:         .Position = frmVybPredst.PozPredstav
59:         If M1s <> .Current("miesta") Then MsgBox("Chyba pri ukladani obsadenych miest. Obratte sa na Programatora.", MsgBoxStyle.Critical) : Exit Sub
60:     End With
61:
62:     With frmSpracovavanie.FilmyBindingSource
63:         frmVybPredst.PozFilm = .Find("Film", frmSpracovavanie.PredstaveniaBindingSource.Current("NazovFilmu"))
64:         If frmVybPredst.PozFilm = -1 Then
65:             MsgBox("Nenasiel som film!", MsgBoxStyle.Critical) : Exit Sub
66:         Else
67:             .Position = frmVybPredst.PozFilm
68:         End If
69:         .Current("predanelistky") += M1OSob '(M1x2 - M1x1 + 1)
70:         .Current("TrzbaFilmu") += M1Cena
71:         'Vytvori HASH pre kontrolu upravy riadku
72:         HASH = Module1.RetazecRiadku("Filmy", frmVybPredst.PozFilm)
73:         .Current("crc") = Module1.HASHString(HASH)
74:         Module1.UpdateFilmy(frmVybPredst.PozFilm)
75:         .Position = frmVybPredst.PozFilm
76:     End With
77:
78:     frmVydat.Zaplatit = Suma
79:
        '80:     '****** VKLADANIE LISTKA DO DATABAZY ******
        '81:     Dim cis2, cis1, VKod As String
        '82:     Dim rnd1, rnd2 As Single
        '83:     Dim MstOd, MstDo As Short

        '85:     Dim Stav As String
        '86:     frmSpracovavanie.PredstaveniaBindingSource.Position = frmVybPredst.PozPredstav
        '87:     frmSpracovavanie.RezervacieBindingSource.Position = PozRez

        M1MenoSaly = frmSpracovavanie.MenoSal(CInt(frmSpracovavanie.PredstaveniaBindingSource.Current("Sala")))
95:
96:     Dim i, nElements As Integer
97:     Dim aKey As Object
98:     Dim Miesta(50) As String
99:     aKey = Split(frmSpracovavanie.RezervacieBindingSource.Current("MiestaMena"), " ")
100:    nElements = UBound(aKey) - LBound(aKey) + 1
101:    i = 0
102:    For X = 0 To nElements - 1
103:        If aKey(X) <> "" Then
104:            i = i + 1
105:            Miesta(i) = aKey(X)
106:        End If
107:    Next X
108:
        If cbxRozdelLstk.Checked Then
            Dim PredatZlav As Integer = PZlav
            For i = 0 To M1OSob - 1
                M1x1 = x1 + i
                M1x2 = M1x1
                M1OSob = 1
                M1MstOD = Miesta(i + 2)
                M1MstDO = M1MstOD
                If PredatZlav > 0 Then
                    M1ZlvMst = 1
                    M1Cena = CenaZ
                    PredatZlav -= 1
                Else
                    M1ZlvMst = 0
                    M1Cena = CenaN
                End If
                PredajLiskta(" - Predaj rezervacie (Meno: " & lblMeno.Text & " Pr:")
            Next i
        Else
            M1x1 = x1
            M1x2 = M1x1 + POsob - 1
            'M1OSob = POsob
            M1MstOD = Miesta(2)
            M1MstDO = Miesta(1 + POsob)
            PredajLiskta(" - Predaj rezervacie (Meno: " & lblMeno.Text & " Pr:")
        End If

        '109:    With frmSpracovavanie.ListkyTableAdapter
        'IneCislo:
        '111:        Randomize(System.DateTime.Now.Millisecond)
        '112:        rnd1 = System.Math.Round(89999 * Rnd() + 10000)
        '113:        rnd2 = System.Math.Round(89999 * Rnd() + 10000)
        '114:
        '115:        cis1 = CStr(rnd1)
        '116:        cis2 = CStr(rnd2)
        '117:        VKod = cis1 & frmSpracovavanie.KontrKodFilmu & cis2
        '118:
        '119:        'Skontroluje ci uz taky listok je v pameti
        '120:        If frmSpracovavanie.ListkyBindingSource.Find("kod", VKod) <> -1 Then GoTo IneCislo
        '121:        frmVybPredst.TlacitKodListka = VKod
        '122:
        '123:    'Vlozi listok do databazy
        '124:    frmSpracovavanie.RezervacieBindingSource.Position = PozRez
        '125:    MstOd = frmSpracovavanie.RezervacieBindingSource.Current("MiestoOd")
        '126:    MstDo = MstOd + POsob - 1
        '127:
        '128:        frmSpracovavanie.ListkyBindingSource.MoveFirst()
        '129:        If cbxBezhotovost.Checked Then Stav = "Bezhotov" Else Stav = "Predane"
        '130:        .Insert(CStr(VKod), Datum, CasPr, CInt(frmSpracovavanie.PredstaveniaBindingSource.Current("Sala")), CInt(MstOd), CInt(MstDo), (Miesta(1) & " " & Miesta(2) & " - " & Miesta(1 + POsob)), CInt(PZlav), CDbl(Suma), Format(Today, "dd.MM.yyyy"), Stav & ";" & frmLogin.UserName, "")
        '131:        .Fill(frmSpracovavanie.DataSet1.Listky)
        '132:        HASH = Module1.RetazecRiadku("Listky", frmSpracovavanie.ListkyBindingSource.Find("crc", ""))
        '133:        frmSpracovavanie.ListkyBindingSource.Current("CRC") = Module1.HASHString(HASH)
        '134:        frmSpracovavanie.ListkyBindingSource.EndEdit()
        '135:        .Update(frmSpracovavanie.DataSet1.Listky)
        '136:        .Fill(frmSpracovavanie.DataSet1.Listky)
        '137:    End With
        '138:
        '139:    frmVybPredst.TlacitDatum = Datum
        '140:    frmVybPredst.TlacitCas = Format(CasPr, "HH:mm")
        '141:    frmVybPredst.TlacitFilm = Film
        '142:    frmVybPredst.TlacitOsob = POsob
        '143:    frmVybPredst.TlacitRad = Miesta(1)
        '144:    frmVybPredst.TlacitSala = M1MenoSaly
        '145:    frmVybPredst.TlacitZlavMiest = PZlav
        '146:    frmVybPredst.TlacitZlavCena = CenaZ
        '147:    frmVybPredst.TlacitCenaCelkom = Suma
        '148:    frmVybPredst.TlacitMiestoOd = Miesta(2)
        '149:    frmVybPredst.TlacitMiestoDo = Miesta(1 + POsob)
        '150:    frmVybPredst.TlacitBezhotov = cbxBezhotovost.Checked
        '151:    Call frmVybPredst.TlacitListok()
152:
153:    With frmSpracovavanie.RezervacieBindingSource
154:        .Position = PozRez
155:        frmSpracovavanie.RezervacieTableAdapter.Delete(.Current(0), .Current(1), .Current(2), .Current(3), .Current(4), .Current(5), .Current(6))
156:        frmSpracovavanie.RezervacieTableAdapter.Fill(frmSpracovavanie.DataSet1.Rezervacie)
157:    End With
158:
        '159:    Call Module1.HASHdbKontrola("Listky")
        '160:    Call Module1.HASHSubor("Data\data.pac")
        '161:    Module1.WriteLog("   OK")
162:    If cbxBezhotovost.Checked = False Then frmVydat.ShowDialog() Else cbxBezhotovost.Checked = False
        cbxRozdelLstk.Checked = False
163:    Call ZaplnitTabulku(sender, e)
164:
165:    Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub txtOsob_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtOsob.Leave
        'On Error Resume Next
        If DataGridView1.Rows.Count = 0 Then Exit Sub
        Dim i As Integer
        i = DataGridView1.CurrentRow.Index
        If CInt(txtOsob.Text) > CInt(DataGridView1.Rows(i).Cells(5).Value) Then txtOsob.Text = DataGridView1.Rows(i).Cells(5).Value
        If txtZlavnen.Text = "" Then txtZlavnen.Text = "0"
        If CInt(txtZlavnen.Text) > CInt(txtOsob.Text) Then txtZlavnen.Text = txtOsob.Text
    End Sub

    Private Sub txtZlavnen_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtZlavnen.Leave
        On Error Resume Next
        If txtZlavnen.Text = "" Then txtZlavnen.Text = "0"
        If CInt(txtZlavnen.Text) > CInt(txtOsob.Text) Then txtZlavnen.Text = txtOsob.Text
    End Sub

    Private Sub txtOsob_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOsob.TextChanged
1:      On Error Resume Next
2:      If lblMeno.Text = "" Then Exit Sub
        If DataGridView1.Rows.Count = 0 Then Exit Sub
3:      Dim i As Integer
4:
5:      i = DataGridView1.CurrentRow.Index
6:      cbxBezhotovost.Checked = False
7:      If txtOsob.Text = "" Then POsob = 0 'txtOsob.Text = DataGridView1.Rows(i).Cells(5).Value
8:
9:      If CInt(txtOsob.Text) < 1 Then
10:         POsob = 0 'txtOsob.Text = DataGridView1.Rows(i).Cells(5).Value
11:     ElseIf CInt(txtOsob.Text) > CInt(DataGridView1.Rows(i).Cells(5).Value) Then
12:         POsob = CInt(DataGridView1.Rows(i).Cells(5).Value) 'txtOsob.Text = DataGridView1.Rows(i).Cells(5).Value
13:     Else
14:         POsob = CInt(txtOsob.Text)
15:     End If
16:
17:     If txtZlavnen.Text = "" Then
18:         PZlav = 0 'txtZlavnen.Text = 0
19:     ElseIf CInt(txtZlavnen.Text) > POsob Then
20:         PZlav = POsob
21:     ElseIf CInt(txtZlavnen.Text) < 0 Then
22:         PZlav = 0
23:     Else
24:         PZlav = CInt(txtZlavnen.Text)
25:     End If
26:
27:     Suma = PZlav * CenaZ + (POsob - PZlav) * CenaN
28:     lblSuma.Text = Suma & " EUR"
    End Sub

    Private Sub cmbPredstavenie_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmbPredstavenie.KeyPress
        frmPredstavenia.cmbFilmy_KeyPress(sender, e, Me)
    End Sub
    Private Sub btnPredat_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles btnPredat.KeyPress
        frmPredstavenia.btnPridaj_KeyPress(sender, e, Me)
    End Sub
    Private Sub btnZobrazit_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles btnZobrazit.KeyPress
        frmPredstavenia.btnPridaj_KeyPress(sender, e, Me)
    End Sub
    Private Sub btnZrusit_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles btnZrusit.KeyPress
        frmPredstavenia.btnPridaj_KeyPress(sender, e, Me)
    End Sub
    Private Sub btnZobrazit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnZobrazit.Click
        Call ZaplnitTabulku(sender, e)
    End Sub
    Private Sub txtZlavnen_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtZlavnen.TextChanged
        txtOsob_TextChanged(sender, e)
    End Sub
    Private Sub txtZlavnen_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtZlavnen.KeyPress
        frmPredstavenia.txtCena_KeyPress(sender, e, Me)
    End Sub
    Private Sub txtOsob_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtOsob.KeyPress
        frmPredstavenia.txtCena_KeyPress(sender, e, Me)
    End Sub

    Private Sub btnTlacit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTlacit.Click
1:      On Error GoTo Chyba
2:      frmMenu.Vytlacit = ""
3:      frmMenu.Velkost = 9
4:      Dim ODDELOVAC As String = StrDup(90, "─")
5:      Dim X, i As Integer
6:      Dim PomocRetazec As String
7:      Dim Hlavicka(10) As String
8:      Dim Tab(10) As Integer
9:
10:     Tab(0) = 0
11:     Tab(1) = 20
12:     Tab(2) = 40
13:     Tab(3) = 46
14:     Tab(4) = 52
15:     Tab(5) = 72
16:     Tab(6) = 77
17:
18:     Hlavicka(0) = "Meno"
19:     Hlavicka(1) = "Udaje"
20:     Hlavicka(2) = "Datum"
21:     Hlavicka(3) = "Cas"
22:     Hlavicka(4) = "Film"
23:     Hlavicka(5) = "Osob"
24:     Hlavicka(6) = "Miesta"
25:
26:     'Printer.FontSize = 13
27:     frmMenu.Vytlacit += "Evidovane rezervacie" & Chr(13) & Chr(10)
28:     frmMenu.Vytlacit += Chr(13) & Chr(10)
29:     If cmbPredstavenie.Text <> "" Then frmMenu.Vytlacit += "Predstavenie: " & cmbPredstavenie.Text Else frmMenu.Vytlacit += "Predstavenie: Vsetky predstavenia"
30:     frmMenu.Vytlacit += Chr(13) & Chr(10)
31:     frmMenu.Vytlacit += Chr(13) & Chr(10)
32:     frmMenu.Vytlacit += Chr(13) & Chr(10)
33:
34:     '-- ── ══ ==
35:     'Printer.FontSize = 10
36:     'frmMenu.Vytlacit += "---------1---------2---------3---------4---------5---------6---------7--" & Chr(13) & Chr(10)
37:     'frmMenu.Vytlacit += "123456789012345678901234567890123456789012345678901234567890123456789012" & Chr(13) & Chr(10)
38:     frmMenu.Vytlacit += ODDELOVAC & Chr(13) & Chr(10)
39:
40:     PomocRetazec = ""
41:     For i = 0 To 6
42:         PomocRetazec += StrDup((Tab(i) - Len(PomocRetazec)), " ")
43:         PomocRetazec += Hlavicka(i)
44:     Next i
45:     frmMenu.Vytlacit += PomocRetazec & Chr(13) & Chr(10)
46:
47:     frmMenu.Vytlacit += ODDELOVAC & Chr(13) & Chr(10)
48:     frmMenu.Vytlacit += Chr(13) & Chr(10)
49:
50:     With DataGridView1
51:         For X = 1 To .Rows.Count
52:             PomocRetazec = ""
53:             PomocRetazec += LSet(.Rows(X - 1).Cells(0).Value, 18)
54:             For i = 1 To 6
55:                 PomocRetazec += StrDup((Tab(i) - Len(PomocRetazec)), " ")
56:                 If i = 2 Then
57:                     PomocRetazec += Format(CDate(.Rows(X - 1).Cells(i).Value), "dd.MM")
58:                 ElseIf i = 3 Then
59:                     PomocRetazec += Format(CDate(.Rows(X - 1).Cells(i).Value), "HH:mm")
60:                 ElseIf i = 5 Then
61:                     PomocRetazec += CStr(.Rows(X - 1).Cells(i).Value)
62:                 ElseIf i = 6 Then
63:                     PomocRetazec += .Rows(X - 1).Cells(i).Value & "  " & .Rows(X - 1).Cells(7).Value
64:                 Else
65:                     PomocRetazec += LSet(.Rows(X - 1).Cells(i).Value, 18)
66:                 End If
67:             Next i
68:             frmMenu.Vytlacit += PomocRetazec & Chr(13) & Chr(10)
69:         Next X
70:     End With
71:
72:     frmMenu.Vytlacit += Chr(13) & Chr(10)
73:     frmMenu.Vytlacit += ODDELOVAC & Chr(13) & Chr(10)
74:     frmMenu.Vytlacit += Chr(13) & Chr(10)
75:
76:     frmMenu.Tlacit()
77:
78:     Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub
    Private Sub DataGridView1_CellEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellEnter
        If DataGridView1.Rows.Count > 1 Then Call DataGridView1_Click(sender, e)
    End Sub
End Class