Public Class frmPredstavenia

    Dim VizualUprava As Boolean = True

    Private Sub frmPredstavenia_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
1:      On Error GoTo Chyba
2:      Dim i As Short
3:
4:      'Zafarbi
5:      Call mdlColors.Skining(Me)
6:      If VizualUprava Then
7:          Call mdlColors.Sizing(Me)
8:          Me.CenterToParent()
9:          VizualUprava = False
10:     End If
11:
12:     '************ SPRAVA DATA GRIDU*************************
13:     With DataGridView1
14:         .SuspendLayout()
15:         .DataSource = Nothing
16:         .AllowUserToAddRows = False
17:         .AllowUserToDeleteRows = False
18:         .AllowUserToResizeRows = False
19:         .AllowUserToResizeColumns = False
20:
21:         .AllowUserToOrderColumns = True
22:         .SelectionMode = DataGridViewSelectionMode.FullRowSelect
23:         .ReadOnly = True
24:         .MultiSelect = False
25:         .RowHeadersVisible = False
26:         .Columns.Clear()
27:         .Rows.Clear()
28:         ' setup columns
29:
30:         .Columns.Add("dtgrdDatum", "Datum")
31:         .Columns(0).Width = (.Width - 20) * 0.13
32:         .Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
33:         .Columns.Add("dtgrdDen", "Den")
34:         .Columns(1).Width = (.Width - 20) * 0.05
35:         .Columns(1).SortMode = DataGridViewColumnSortMode.NotSortable
36:         .Columns.Add("dtgrdCas", "Cas")
37:         .Columns(2).Width = (.Width - 20) * 0.08
38:         .Columns(2).SortMode = DataGridViewColumnSortMode.NotSortable
39:         .Columns.Add("dtgrdFilm", "Názov filmu")
40:         .Columns(3).Width = (.Width - 20) * 0.28
41:         .Columns(3).SortMode = DataGridViewColumnSortMode.NotSortable
42:         .Columns.Add("dtgrdCena", "Cena (EUR)")
43:         .Columns(4).Width = (.Width - 20) * 0.08
44:         .Columns(4).SortMode = DataGridViewColumnSortMode.NotSortable
45:         .Columns.Add("dtgrdCenaZlav", "Zlavneny (EUR)")
46:         .Columns(5).Width = (.Width - 20) * 0.095
47:         .Columns(5).SortMode = DataGridViewColumnSortMode.NotSortable
48:         .Columns.Add("dtgrdSala2", "Sála")
49:         .Columns(6).Width = (.Width - 20) * 0.12
50:         .Columns(6).SortMode = DataGridViewColumnSortMode.NotSortable
51:         .Columns.Add("dtgrdSala3", "Predaj")
52:         .Columns(7).Width = (.Width - 20) * 0.075
53:         .Columns(7).SortMode = DataGridViewColumnSortMode.NotSortable
54:         .Columns.Add("dtgrdSala4", "Trzba (EUR)")
55:         .Columns(8).Width = (.Width - 20) * 0.09
56:         .Columns(8).SortMode = DataGridViewColumnSortMode.NotSortable
57:
58:         .ResumeLayout(True)
59:     End With
60:
61:     'Naplni combobox z Filmamy
62:     With frmSpracovavanie.FilmyBindingSource
63:         cmbFilmy.Items.Clear()
65:         For i = 0 To .Count - 1
66:             .Position = i
67:             cmbFilmy.Items.Add(.Current("Film"))
68:         Next i
69:     End With
70:
71:     cmbSala.Items.Clear()
72:     For i = 1 To frmSpracovavanie.PocetSal
73:         cmbSala.Items.Add(frmSpracovavanie.MenoSal(i))
            'If i = 1 Then cmbSala.Items.Add(frmSpracovavanie.MenoSal1)
74:         'If i = 2 Then cmbSala.Items.Add(frmSpracovavanie.MenoSal2)
75:         'If i = 3 Then cmbSala.Items.Add(frmSpracovavanie.MenoSal3)
76:         'If i = 4 Then cmbSala.Items.Add(frmSpracovavanie.MenoSal4)
77:         'If i = 5 Then cmbSala.Items.Add(frmSpracovavanie.MenoSal5)
78:     Next i
79:     cmbSala.SelectedIndex = 0
80:
81:     Call ZaplnitTabulku(sender, e)
82:
        Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub btnPridaj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPridaj.Click
1:      On Error GoTo Chyba
2:      Dim DatumDO As Date
3:
4:      If CDate(dtpDatum.Text) < Today Then If MsgBox("Den predstavenia uz bol. Vlozit predstavenia aj tak?", MsgBoxStyle.YesNo + MsgBoxStyle.Question) = MsgBoxResult.No Then Exit Sub
5:      DatumDO = CDate(dtpDatumDo.Text)
6:      If DatumDO < CDate(dtpDatum.Text) Then MsgBox("Datum do je mensi ako datum od.", MsgBoxStyle.Exclamation) : Exit Sub
7:      If (System.DateTime.FromOADate(DatumDO.ToOADate - CDate(dtpDatum.Text).ToOADate)) > System.DateTime.FromOADate(32) Then
8:          If MsgBox("Chcete vlozit predstavenia na " & DatumDO.ToOADate - CDate(dtpDatum.Text).ToOADate & " dni naraz?", MsgBoxStyle.YesNo + MsgBoxStyle.Question) = MsgBoxResult.No Then Exit Sub
9:      End If
10:
11:     If txtCas.Text = "" Then MsgBox("Zadaj cas predstavenia.", MsgBoxStyle.Exclamation) : Exit Sub
12:     If cmbFilmy.Text = "" Then MsgBox("Vyber predstavenie.", MsgBoxStyle.Exclamation) : Exit Sub
13:     If cmbSala.Text = "" Then MsgBox("Vyber salu predstavenia.", MsgBoxStyle.Exclamation) : Exit Sub
14:     If txtCena.Text = "" Then txtCena.Text = 0
15:     If txtCenaZlavnena.Text = "" Then txtCenaZlavnena.Text = txtCena.Text
16:     If CDbl(txtCena.Text) < CDbl(txtCenaZlavnena.Text) Then MsgBox("Zlavnena cena je vyssia ako bezne vstupne.", MsgBoxStyle.Exclamation) : Exit Sub
17:     If CDbl(txtCena.Text) <= 0 Then If MsgBox("Vstup na predstavenie zdarma?", MsgBoxStyle.YesNo + MsgBoxStyle.Question) = MsgBoxResult.No Then Exit Sub
18:     Module1.WriteLog(Chr(13) & Chr(10) & Now & " - Vlozenie predstaveni: " & cmbFilmy.Text)
19:
20:     Dim SalaMiest As Short
        Dim ZobrazSedadla As Boolean
21:     Dim x, i As Integer
22:     Dim nElements As Short
23:     Dim aKey As Object
24:     Dim Casy(20) As Date
25:     Dim DatumRobit As Double
26:     Dim HASH As String
27:
28:     Cursor = Cursors.AppStarting
29:
30:     'Pole casov premietania filmu, aby netrebalo davat na 3x
31:     aKey = Split(txtCas.Text, " ")
32:     nElements = UBound(aKey) - LBound(aKey) + 1
33:     i = 0
34:     For x = 0 To nElements - 1
35:         If aKey(x) <> "" Then
36:             i = i + 1
37:             Casy(i) = aKey(x)
38:         End If
39:     Next x
40:     i = 0
41:
42:     'DatumRobit je datum ktory prave spracuva
43:     'Casy(x) su casy v ktorych je premietany film
44:     With frmSpracovavanie.PredstaveniaBindingSource
45:         For DatumRobit = CDate(dtpDatum.Text).ToOADate To DatumDO.ToOADate
46:             For x = 1 To nElements
47:                 If .Find("Datum", DateTime.FromOADate(DatumRobit) & " " & Format(Casy(x), "HH:mm")) <> -1 Then
48:                     MsgBox("Dna " & DateTime.FromOADate(DatumRobit) & " o " & Format(Casy(x), "HH:mm") & " bude premietané iné predstavenie." & Chr(13) & "Predstavenia nepridane", MsgBoxStyle.Exclamation)
49:                     Cursor = Cursors.Default
50:                     Exit Sub
51:                 End If
52:             Next x
53:         Next DatumRobit
54:
55:         'Prida vsetky predstavenia
56:         If cmbSala.SelectedIndex = 0 Then SalaMiest = frmSpracovavanie.PocetMiest(1) : ZobrazSedadla = frmSpracovavanie.ZobrazMiest1
57:         If cmbSala.SelectedIndex = 1 Then SalaMiest = frmSpracovavanie.PocetMiest(2) : ZobrazSedadla = frmSpracovavanie.ZobrazMiest2
58:         If cmbSala.SelectedIndex = 2 Then SalaMiest = frmSpracovavanie.PocetMiest(3) : ZobrazSedadla = frmSpracovavanie.ZobrazMiest3
59:         If cmbSala.SelectedIndex = 3 Then SalaMiest = frmSpracovavanie.PocetMiest(4) : ZobrazSedadla = frmSpracovavanie.ZobrazMiest4
60:         If cmbSala.SelectedIndex = 4 Then SalaMiest = frmSpracovavanie.PocetMiest(5) : ZobrazSedadla = frmSpracovavanie.ZobrazMiest5
            If ZobrazSedadla = False Then SalaMiest = 100
61:         For DatumRobit = CDate(dtpDatum.Text).ToOADate To DatumDO.ToOADate
62:             For x = 1 To nElements
                    HASH = Module1.HASHString(Format(CDate(DateTime.FromOADate(DatumRobit) & " " & Casy(x)), "dd.MM.yyyy HH:mm") & cmbFilmy.Text & Format(CDbl(txtCena.Text), "##0.##") & Format(CDbl(txtCenaZlavnena.Text), "##0.##") & 0 & Format(0, "###0.##") & StrDup(SalaMiest, "1") & cmbSala.SelectedIndex + 1)
63:                 frmSpracovavanie.PredstaveniaTableAdapter.Insert((DateTime.FromOADate(DatumRobit) & " " & Casy(x)), cmbFilmy.Text, txtCena.Text, txtCenaZlavnena.Text, 0, 0, StrDup(SalaMiest, "1"), cmbSala.SelectedIndex + 1, HASH) ' "")
64:                 'frmSpracovavanie.PredstaveniaTableAdapter.Insert(Format((DateTime.FromOADate(DatumRobit) & " " & Casy(x)), "dd.MM.yyyy HH:mm"), cmbFilmy.Text, txtCena.Text, txtCenaZlavnena.Text, 0, 0, StrDup(SalaMiest, "1"), cmbSala.SelectedIndex + 1, "")
65:
66:             Next x
67:         Next DatumRobit
68:         frmSpracovavanie.PredstaveniaTableAdapter.Fill(frmSpracovavanie.DataSet1.Predstavenia)
69:
70:         'Prida CRC (po novom uz to robi priebezne)
72:         ' For i = 1 To .Count
73:         'x = .Find("crc", "")
74:         'If x = -1 Then Exit For
75:         '.Position = x
76:         'HASH = Module1.RetazecRiadku("Predstavenia", .Position)
77:         '.Current("CRC") = Module1.HASHString(HASH)
78:         'Next i
79:         '.EndEdit()
80:         'frmSpracovavanie.PredstaveniaTableAdapter.Update(frmSpracovavanie.DataSet1.Predstavenia)
81:         'frmSpracovavanie.PredstaveniaTableAdapter.Fill(frmSpracovavanie.DataSet1.Predstavenia)
82:     End With
83:
84:     Call Module1.HASHdbKontrola("Predstavenia")
85:
86:     'Ulozi datum posledneho predstavenia
87:     With frmSpracovavanie.FilmyBindingSource
88:         frmVybPredst.PozFilm = .Find("Film", cmbFilmy.Text)
89:         .Position = frmVybPredst.PozFilm
90:         If DatumDO > CDate(.Current("PoslednePredstavenie")) Then
91:             .Current("PoslednePredstavenie") = Format(DatumDO, "dd.MM.yyyy")
92:             HASH = Module1.RetazecRiadku("Filmy", .Position)
93:             .Current("CRC") = Module1.HASHString(HASH)
94:             Module1.UpdateFilmy(frmVybPredst.PozFilm)
95:         End If
96:     End With
97:
98:     Cursor = Cursors.Default
99:     Call Module1.HASHSubor("Data\data.pac")
100:    Call ZaplnitTabulku(sender, e)
101:    MsgBox("Predstavenia pridané.", MsgBoxStyle.Information)
102:    Module1.WriteLog("   OK")
103:    frmSpracovavanie.Zaloha = True
104:    Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub btnOdobrat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOdobrat.Click
1:      On Error GoTo Chyba
2:      Dim i As Integer
3:      i = DataGridView1.CurrentRow.Index
4:      If CDbl(DataGridView1.Rows(i).Cells(8).Value) <> 0 Then MsgBox("Na predstavenie uz boli predane listky!", MsgBoxStyle.Exclamation) : Exit Sub
5:      Dim DatumPr, Datum2 As Date
6:      Dim Flm, HASH As String
7:
8:      DatumPr = CDate(DataGridView1.Rows(i).Cells(0).Value & " " & DataGridView1.Rows(i).Cells(2).Value)
9:      Flm = DataGridView1.Rows(i).Cells(3).Value
10:
11:     If MsgBox("Chcete vymazat predstavenie " & DatumPr & ", " & Flm & " ?", MsgBoxStyle.YesNo + MsgBoxStyle.Question) = MsgBoxResult.No Then Exit Sub
12:     Module1.WriteLog(Chr(13) & Chr(10) & Now & " - Zmazanie predstaveni: " & Flm)
13:     'Vyhlada spravne zaznamy filmov
14:     If frmSpracovavanie.FilmyBindingSource.Find("Film", Flm) = -1 Then
15:         MsgBox("Nenasiel sa film! Predstavenie bude vymazane, ale v tabulke Filmov nebudu robene zmeny.", MsgBoxStyle.Critical)
16:         With frmSpracovavanie.PredstaveniaBindingSource
17:             frmVybPredst.PozPredstav = .Find("Datum", DatumPr)
18:             .Position = frmVybPredst.PozPredstav
19:             frmSpracovavanie.PredstaveniaTableAdapter.Delete(.Current(0), .Current(1), .Current(2), .Current(3), .Current(4), .Current(5), .Current(7), .Current(8))
20:             frmSpracovavanie.PredstaveniaTableAdapter.Fill(frmSpracovavanie.DataSet1.Predstavenia)
21:         End With
22:
23:     Else
24:
25:         frmVybPredst.PozFilm = frmSpracovavanie.FilmyBindingSource.Find("Film", Flm)
26:         frmSpracovavanie.FilmyBindingSource.Position = frmVybPredst.PozFilm
27:
28:         With frmSpracovavanie.PredstaveniaBindingSource
29:             frmVybPredst.PozPredstav = .Find("Datum", DatumPr)
30:             .Position = frmVybPredst.PozPredstav
31:             frmSpracovavanie.PredstaveniaTableAdapter.Delete(.Current(0), .Current(1), .Current(2), .Current(3), .Current(4), .Current(5), .Current(7), .Current(8))
32:             frmSpracovavanie.PredstaveniaTableAdapter.Fill(frmSpracovavanie.DataSet1.Predstavenia)
33:
34:             'Zisti najvacsi datum a ulozi
35:             'Ak datum predstavenia = datumu posledneho vysielania 
36:             If Format(DatumPr, "dd.MM.yyyy") = CDate(frmSpracovavanie.FilmyBindingSource.Current("PoslednePredstavenie")) Then
37:                 Datum2 = CDate("01.01.2000")
38:                 For i = 0 To .Count - 1
39:                     .Position = i
40:                     If .Current("nazovfilmu") = Flm Then
41:                         If CDate(.Current("Datum")) > Datum2 Then Datum2 = CDate(.Current("datum"))
42:                     End If
43:                 Next i
44:                 frmSpracovavanie.FilmyBindingSource.Current("PoslednePredstavenie") = Format(Datum2, "dd.MM.yyyy")
45:                 HASH = Module1.RetazecRiadku("Filmy", frmVybPredst.PozFilm)
46:                 frmSpracovavanie.FilmyBindingSource.Current("CRC") = Module1.HASHString(HASH)
47:                 Module1.UpdateFilmy(frmVybPredst.PozFilm)
48:             End If
49:         End With
50:     End If
51:
52:     Call Module1.HASHdbKontrola("Predstavenia")
53:     Call Module1.HASHSubor("Data\data.pac")
54:     Module1.WriteLog("   OK")
55:     Call ZaplnitTabulku(sender, e)
56:
57:     Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub btnStorno_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStorno.Click
1:      On Error GoTo Chyba
2:      Dim i As Integer = DataGridView1.CurrentRow.Index
3:      Dim datum As Date
4:      datum = CDate(DataGridView1.Rows(i).Cells(0).Value & " " & DataGridView1.Rows(i).Cells(2).Value)
5:
6:      'Vyhlada spravne zaznamy
7:      frmVybPredst.PozPredstav = frmSpracovavanie.PredstaveniaBindingSource.Find("Datum", datum)
8:
9:      frmPredstStorn.ShowDialog()
10:
11:     Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
1:      On Error GoTo Chyba
2:      Dim i As Integer = DataGridView1.CurrentRow.Index
3:      'If CDbl(DataGridView1.Rows(i).Cells(8).Value) <> 0 Then MsgBox("Na predstavenie uz boli predane listky!", MsgBoxStyle.Exclamation) : Exit Sub
4:      Dim datum As Date
5:      datum = CDate(DataGridView1.Rows(i).Cells(0).Value & " " & DataGridView1.Rows(i).Cells(2).Value)
6:
7:      'Vyhlada spravne zaznamy
8:      frmVybPredst.PozPredstav = frmSpracovavanie.PredstaveniaBindingSource.Find("Datum", datum)
9:
10:     frmPredstaveniaUprav.ShowDialog()
11:
12:     Call ZaplnitTabulku(sender, e)
13:
        Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub ZaplnitTabulku(ByVal sender As System.Object, ByVal e As System.EventArgs)
1:      On Error GoTo Chyba
2:      Dim i As Integer
3:
4:      With DataGridView1
5:          .Rows.Clear()
6:          frmSpracovavanie.PredstaveniaBindingSource.MoveFirst()
7:          For i = 1 To frmSpracovavanie.PredstaveniaBindingSource.Count
8:              .Rows.Add()
9:              .Rows(.Rows.Count - 1).Cells(0).Value = Format(CDate(frmSpracovavanie.PredstaveniaBindingSource.Current("Datum")), "dd.MM.yyyy")
10:             .Rows(.Rows.Count - 1).Cells(1).Value = Format(CDate(frmSpracovavanie.PredstaveniaBindingSource.Current("Datum")), "ddd")
11:             .Rows(.Rows.Count - 1).Cells(2).Value = Format(CDate(frmSpracovavanie.PredstaveniaBindingSource.Current("Datum")), "HH:mm")
12:             .Rows(.Rows.Count - 1).Cells(3).Value = frmSpracovavanie.PredstaveniaBindingSource.Current("NazovFilmu")
13:             .Rows(.Rows.Count - 1).Cells(4).Value = Format(frmSpracovavanie.PredstaveniaBindingSource.Current("cenalistka"), "###0.00")
14:             .Rows(.Rows.Count - 1).Cells(5).Value = Format(frmSpracovavanie.PredstaveniaBindingSource.Current("cenazlavlistka"), "###0.00")
15:
16:             .Rows(.Rows.Count - 1).Cells(6).Value = frmSpracovavanie.MenoSal(frmSpracovavanie.PredstaveniaBindingSource.Current("sala"))
                'If frmSpracovavanie.PredstaveniaBindingSource.Current("sala") = 1 Then .Rows(.Rows.Count - 1).Cells(6).Value = frmSpracovavanie.MenoSal1
17:             'If frmSpracovavanie.PredstaveniaBindingSource.Current("sala") = 2 Then .Rows(.Rows.Count - 1).Cells(6).Value = frmSpracovavanie.MenoSal2
18:             'If frmSpracovavanie.PredstaveniaBindingSource.Current("sala") = 3 Then .Rows(.Rows.Count - 1).Cells(6).Value = frmSpracovavanie.MenoSal3
19:             'If frmSpracovavanie.PredstaveniaBindingSource.Current("sala") = 4 Then .Rows(.Rows.Count - 1).Cells(6).Value = frmSpracovavanie.MenoSal4
20:             'If frmSpracovavanie.PredstaveniaBindingSource.Current("sala") = 5 Then .Rows(.Rows.Count - 1).Cells(6).Value = frmSpracovavanie.MenoSal5
21:
22:             .Rows(.Rows.Count - 1).Cells(7).Value = frmSpracovavanie.PredstaveniaBindingSource.Current("Predaj") & ""
23:             .Rows(.Rows.Count - 1).Cells(8).Value = Format(CDbl(frmSpracovavanie.PredstaveniaBindingSource.Current("TrzbaPredstavenia") & ""), "###0.00")
24:             frmSpracovavanie.PredstaveniaBindingSource.MoveNext()
25:         Next i
26:     End With
27:
        Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub txtCas_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCas.Leave
1:      Dim aKey As Object
2:      Dim Casy(20) As Date
3:      Dim x, i As Integer
4:      Dim nElements As Integer
5:      Dim KontrCas As Date
6:
7:      txtCas.Text = Trim(txtCas.Text)
8:      aKey = Split(txtCas.Text, " ")
9:      nElements = UBound(aKey) - LBound(aKey) + 1
10:     i = 0
11:     txtCas.Text = ""
12:
13:     On Error GoTo ZlyCas
14:     For x = 0 To nElements - 1
15:         If aKey(x) <> "" Then
16:             KontrCas = aKey(x)
17:             KontrCas = Format(CDate(KontrCas), "HH:mm")
18:             If KontrCas <> CDate("00:00") Then
19:                 i = i + 1
20:                 Casy(i) = aKey(x)
21:                 txtCas.Text = txtCas.Text & Format(Casy(i), "HH:mm") & " "
22:             End If
23:         End If
24:     Next x
25:
26:     txtCas.Text = Trim(txtCas.Text)
27:
        Exit Sub

ZlyCas:
        MsgBox("Nespravne zadany cas!", MsgBoxStyle.Exclamation)
        txtCas.Focus()
    End Sub

    Private Sub txtCas_KeyPress1(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCas.KeyPress
        txtCas_KeyPress(sender, e, Me)
    End Sub
    Private Sub cmbSala_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmbSala.KeyPress
        cmbFilmy_KeyPress(sender, e, Me)
    End Sub
    Private Sub btnOdobrat_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles btnOdobrat.KeyPress
        btnPridaj_KeyPress(sender, e, Me)
    End Sub
    Private Sub btnStorno_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles btnStorno.KeyPress
        btnPridaj_KeyPress(sender, e, Me)
    End Sub
    Private Sub txtCena_KeyPress1(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCena.KeyPress
        txtCena_KeyPress(sender, e, Me)
    End Sub
    Private Sub txtCenaZlavnena_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCenaZlavnena.KeyPress
        txtCena_KeyPress(sender, e, Me)
    End Sub
    Private Sub btnPridaj_KeyPress1(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles btnPridaj.KeyPress
        btnPridaj_KeyPress(sender, e, Me)
    End Sub
    Private Sub cmbFilmy_KeyPress1(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmbFilmy.KeyPress
        cmbFilmy_KeyPress(sender, e, Me)
    End Sub

    Public Sub txtCena_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs, ByVal host As System.Windows.Forms.Form) 'Handles txtCena.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        ' Zisti, ktora klavesa bola stalcena v ASCII hodnotach
        Dim TrackKey As String
        TrackKey = Chr(KeyAscii)
        'Ak Enter
        If KeyAscii = System.Windows.Forms.Keys.Enter Then
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{tab}")
        ElseIf KeyAscii = System.Windows.Forms.Keys.Escape Then
            host.Close()
            'ElseIf KeyAscii = 126 And frmLogin.Rights = "A" Then
            ' 126 - ~
            'Prikazovy riadok
            'KeyAscii = 0
            'frmCommand.ShowDialog()
        ElseIf KeyAscii = 44 Or KeyAscii = 46 Then
            'ak ciarka abo bodka tak to co je dobre
            KeyAscii = frmSpracovavanie.DesOdd
        ElseIf (Not IsNumeric(TrackKey) And Not (KeyAscii = System.Windows.Forms.Keys.Back) And Not (KeyAscii = 46)) And Not (KeyAscii = 44) Then
            ' Ak klavesa nebola a) cislo b) backspace c) desatinna bodka, akoby nebolo nic stlacene
            KeyAscii = 0
            Beep()
        End If
        e.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            e.Handled = True
        End If
    End Sub

    Public Sub txtDatum_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs, ByVal host As System.Windows.Forms.Form) ' Handles txtDatum.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        ' Zisti, ktora klavesa bola stalcena v ASCII hodnotach
        Dim TrackKey As String
        TrackKey = Chr(KeyAscii)
        'Ak Enter
        If KeyAscii = System.Windows.Forms.Keys.Enter Then
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{tab}")
        ElseIf KeyAscii = System.Windows.Forms.Keys.Escape Then
            host.Close()
        ElseIf KeyAscii = 44 Then
            'ak ciarka tak bodka
            KeyAscii = 46
        ElseIf (Not IsNumeric(TrackKey) And Not (KeyAscii = System.Windows.Forms.Keys.Back) And Not (KeyAscii = 46)) And Not (KeyAscii = 44) Then
            ' Ak klavesa nebola a) cislo b) backspace c) desatinna bodka, akoby nebolo nic stlacene
            KeyAscii = 0
            Beep()
        End If
        e.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtCas_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs, ByVal host As System.Windows.Forms.Form) 'Handles txtCas.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        Dim TrackKey As String
        If KeyAscii = System.Windows.Forms.Keys.Return Then
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{tab}")
        ElseIf KeyAscii = System.Windows.Forms.Keys.Escape Then
            host.Close()
        Else
            TrackKey = Chr(KeyAscii)
            If (Not IsNumeric(TrackKey) And Not (KeyAscii = System.Windows.Forms.Keys.Back) And Not (KeyAscii = 58) And Not (KeyAscii = 46)) And Not (KeyAscii = System.Windows.Forms.Keys.Space) Then
                KeyAscii = 0
                Beep()
            End If
        End If
        e.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            e.Handled = True
        End If
    End Sub

    Public Sub cmbFilmy_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs, ByVal host As System.Windows.Forms.Form) ' Handles cmbFilmy.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        'Dim TrackKey As String
        If KeyAscii = System.Windows.Forms.Keys.Return Then
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{tab}")
        ElseIf KeyAscii = System.Windows.Forms.Keys.Escape Then
            host.Close()
        Else
            'TrackKey = Chr(KeyAscii)
            If (Not KeyAscii = System.Windows.Forms.Keys.Back) Then
                KeyAscii = 0
                Beep()
            End If
        End If
        e.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            e.Handled = True
        End If
    End Sub

    Public Sub txtDatum_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
        If sender.Text = "" Then Exit Sub
        On Error GoTo ZlyDatum
        sender.Text = Format(CDate(sender.Text), "dd.MM.yyyy")
        Exit Sub
ZlyDatum:
        sender.Text = ""
        MsgBox("Neplatný dátum", MsgBoxStyle.Exclamation)
    End Sub

    Public Sub btnPridaj_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs, ByVal host As System.Windows.Forms.Form) ' Handles btnPridaj.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If KeyAscii = System.Windows.Forms.Keys.Escape Then
            host.Close()
        End If
        e.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            e.Handled = True
        End If
    End Sub

    Private Sub dtpDatum_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpDatum.ValueChanged
        If dtpDatum.Value > dtpDatumDo.Value Then dtpDatumDo.Value = dtpDatum.Value
    End Sub
    Private Sub dtpDatumDo_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpDatumDo.ValueChanged
        If dtpDatum.Value > dtpDatumDo.Value Then dtpDatum.Value = dtpDatumDo.Value
    End Sub

   
End Class