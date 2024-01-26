Public Class frmFilmy

    Dim VizualUprava As Boolean = True

    Private Sub frmFilmy_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
30:         .Columns.Add("dtgrdDatum", "Film")
31:         .Columns(0).Width = (.Width - 20) * 0.26
32:         .Columns.Add("dtgrdDen", "Distributor")
33:         .Columns(1).Width = (.Width - 20) * 0.17
34:         .Columns(1).SortMode = DataGridViewColumnSortMode.NotSortable
35:         .Columns.Add("dtgrdCas", "Predane listky (ks)")
36:         .Columns(2).Width = (.Width - 20) * 0.1
37:         .Columns(2).SortMode = DataGridViewColumnSortMode.NotSortable
38:         .Columns.Add("dtgrdFilm", "Trzba filmu (EUR)")
39:         .Columns(3).Width = (.Width - 20) * 0.11
40:         .Columns(3).SortMode = DataGridViewColumnSortMode.NotSortable
41:         .Columns.Add("dtgrdCena", "Posledne predstavenie")
42:         .Columns(4).Width = (.Width - 20) * 0.14
43:         .Columns(4).SortMode = DataGridViewColumnSortMode.NotSortable
44:         .Columns.Add("dtgrdCenaZlav", "Pevne pozicovne (EUR)")
45:         .Columns(5).Width = (.Width - 20) * 0.11
46:         .Columns(5).SortMode = DataGridViewColumnSortMode.NotSortable
47:         .Columns.Add("dtgrdSala", "Variabilne pozicovne (%)")
48:         .Columns(6).Width = (.Width - 20) * 0.11
49:         .Columns(6).SortMode = DataGridViewColumnSortMode.NotSortable
50:
51:         .ResumeLayout(True)
52:     End With
53:
54:     'Naplni combobox z distributormi
55:     With frmSpracovavanie.DistributoriBindingSource
56:         cmbDistrib.Items.Clear()
57:         .MoveFirst()
58:         For i = 1 To .Count
59:             cmbDistrib.Items.Add(.Current("Distributor"))
60:             .MoveNext()
61:         Next i
62:     End With
63:
64:     Call ZaplnitTabulku(sender, e)

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
6:          frmSpracovavanie.FilmyBindingSource.MoveFirst()
7:          For i = 1 To frmSpracovavanie.FilmyBindingSource.Count
8:              .Rows.Add()
9:              .Rows(.Rows.Count - 1).Cells(0).Value = frmSpracovavanie.FilmyBindingSource.Current("Film") & ""
10:             .Rows(.Rows.Count - 1).Cells(1).Value = frmSpracovavanie.FilmyBindingSource.Current("Distributor") & ""
11:             .Rows(.Rows.Count - 1).Cells(2).Value = frmSpracovavanie.FilmyBindingSource.Current("PredaneListky") & ""
12:             .Rows(.Rows.Count - 1).Cells(3).Value = frmSpracovavanie.FilmyBindingSource.Current("TrzbaFilmu") & ""
13:             .Rows(.Rows.Count - 1).Cells(4).Value = frmSpracovavanie.FilmyBindingSource.Current("PoslednePredstavenie") & ""
14:             .Rows(.Rows.Count - 1).Cells(5).Value = frmSpracovavanie.FilmyBindingSource.Current("PevnePozicovne") & ""
15:             .Rows(.Rows.Count - 1).Cells(6).Value = frmSpracovavanie.FilmyBindingSource.Current("PercentPozicovne") & ""
16:             frmSpracovavanie.FilmyBindingSource.MoveNext()
17:         Next i
18:     End With

        Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub btnZmazat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnZmazat.Click
1:      On Error GoTo Chyba
2:      Dim VybranyFilm As String
3:      Dim i As Integer
4:      i = DataGridView1.CurrentRow.Index
5:      VybranyFilm = DataGridView1.Rows(i).Cells(0).Value
6:
7:      If frmSpracovavanie.PredstaveniaBindingSource.Find("NazovFilmu", VybranyFilm) <> -1 Then MsgBox("Nemozno zmazat film." & Chr(13) & "K filmu " & VybranyFilm & " su vlozene predstavenia.", MsgBoxStyle.Exclamation) : Exit Sub
8:      If MsgBox("Naozaj zmazat film" & Chr(13) & VybranyFilm & "?", MsgBoxStyle.YesNo + MsgBoxStyle.Question) = MsgBoxResult.No Then Exit Sub
9:      Module1.WriteLog(Chr(13) & Chr(10) & Now & " - Zmazanie filmu: " & VybranyFilm)
10:     With frmSpracovavanie.FilmyBindingSource
11:         If .Find("Film", VybranyFilm) = -1 Then
12:             MsgBox("Chyba pri mazani filmu!", MsgBoxStyle.Critical)
13:             Exit Sub
14:         Else
15:             .Position = .Find("Film", VybranyFilm)
16:             If CDbl(.Current("PredaneListky")) <> 0 Then
17:                 MsgBox("K filmu uz boli predane lístky!", MsgBoxStyle.Critical)
18:             Else
19:                 frmSpracovavanie.FilmyTableAdapter.Delete(.Current(0), .Current(1), .Current(2), .Current(3), .Current(4), .Current(5), .Current(6), .Current(7))
20:                 frmSpracovavanie.FilmyTableAdapter.Fill(frmSpracovavanie.DataSet1.Filmy)
21:             End If
22:         End If
23:     End With
24:
25:     Call Module1.HASHdbKontrola("Filmy")
26:     Call Module1.HASHSubor("Data\data.pac")
27:     Call ZaplnitTabulku(sender, e)
28:     Module1.WriteLog("   OK")
        frmSpracovavanie.Zaloha = True
        Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub btnPridaj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPridaj.Click
1:      On Error GoTo Chyba
2:      If txtNFilmu.Text = "" Then MsgBox("Zadaj nazov filmu.") : Exit Sub
3:      If txtVarPozic.Text = "" Then txtVarPozic.Text = CStr(0)
4:      If txtPevPozic.Text = "" Then txtPevPozic.Text = CStr(0)
5:      If CDbl(txtVarPozic.Text) = 0 And CDbl(txtPevPozic.Text) = 0 Then MsgBox("Zadajte pozicovné filmu!", MsgBoxStyle.Exclamation) : Exit Sub
6:      If CDbl(txtVarPozic.Text) <> 0 And CDbl(txtPevPozic.Text) <> 0 Then MsgBox("Zadajte iba jeden druh pozicovného!", MsgBoxStyle.Exclamation) : Exit Sub
7:      If cmbDistrib.Text = "" Then MsgBox("Zvolte Distributora filmu!", MsgBoxStyle.Exclamation) : Exit Sub
8:      If frmSpracovavanie.FilmyBindingSource.Find("Film", txtNFilmu.Text) <> -1 Then MsgBox("Film s takým menom uz je ulozený.", MsgBoxStyle.Exclamation) : Exit Sub
9:      If CDbl(txtVarPozic.Text) > 100 Then MsgBox("Viac ako 100% pozicovne?" & Chr(13) & "Nemate radi svoju penazenku?", MsgBoxStyle.Exclamation) : Exit Sub
10:     If InStr(LCase(txtNFilmu.Text), "organizovane") <> 0 Or InStr(LCase(txtNFilmu.Text), "organizované") <> 0 Then
11:         If MsgBox("Nazov filmu obsahuje slovo: organizovane. Film s tymto nazvom bude mat vo vyuctovani nulovy odvod." & Chr(13) & Chr(10) & "Naozaj vlozit film s tymto nazovm?", MsgBoxStyle.YesNo + MsgBoxStyle.Question) = MsgBoxResult.No Then Exit Sub
12:     End If
13:
14:     Dim HASH As String
15:     Dim x As Integer
16:     Module1.WriteLog(Chr(13) & Chr(10) & Now & " - Vlozenie filmu: " & txtNFilmu.Text)
17:     With frmSpracovavanie.FilmyTableAdapter
18:         x = frmSpracovavanie.FilmyBindingSource.Find("Film", txtNFilmu.Text)
19:         If x <> -1 Then MsgBox("Film je uz zadany.", MsgBoxStyle.Exclamation) : Exit Sub
20:
21:         .Insert(txtNFilmu.Text, cmbDistrib.Text, 0, 0, Format(Today, "dd.MM.yyyy"), txtPevPozic.Text, txtVarPozic.Text, "")
22:         .Fill(frmSpracovavanie.DataSet1.Filmy)
23:         HASH = Module1.RetazecRiadku("Filmy", frmSpracovavanie.FilmyBindingSource.Find("crc", ""))
24:         frmSpracovavanie.FilmyBindingSource.Current("CRC") = Module1.HASHString(HASH)
25:         frmSpracovavanie.FilmyBindingSource.EndEdit()
26:         .Update(frmSpracovavanie.DataSet1.Filmy)
27:         .Fill(frmSpracovavanie.DataSet1.Filmy)
28:     End With
29:
30:     Call Module1.HASHdbKontrola("Filmy")
31:     Call Module1.HASHSubor("Data\data.pac")
32:
33:     txtNFilmu.Text = ""
34:     Call ZaplnitTabulku(sender, e)
35:     Module1.WriteLog("   OK")
        frmSpracovavanie.Zaloha = True
        Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub btnUprav_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUprav.Click
1:      On Error GoTo Chyba
2:
3:      Dim VybranyFilm As String
4:      Dim i As Integer
5:      i = DataGridView1.CurrentRow.Index
6:      VybranyFilm = DataGridView1.Rows(i).Cells(0).Value
7:
8:      frmVybPredst.PozFilm = frmSpracovavanie.FilmyBindingSource.Find("Film", VybranyFilm)
9:      frmFilmyNaklady.ShowDialog()

        Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub cmbDistrib_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmbDistrib.KeyPress
        frmPredstavenia.cmbFilmy_KeyPress(sender, e, Me)
    End Sub

    Private Sub btnPridaj_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles btnPridaj.KeyPress
        frmPredstavenia.btnPridaj_KeyPress(sender, e, Me)
    End Sub

    Private Sub btnUprav_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles btnUprav.KeyPress
        frmPredstavenia.btnPridaj_KeyPress(sender, e, Me)
    End Sub

    Private Sub btnZmazat_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles btnZmazat.KeyPress
        frmPredstavenia.btnPridaj_KeyPress(sender, e, Me)
    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
1:      On Error GoTo Chyba
2:      Dim VybranyFilm As String
3:      Dim i As Integer
4:      i = DataGridView1.CurrentRow.Index
5:      VybranyFilm = DataGridView1.Rows(i).Cells(0).Value
6:
7:      frmVybPredst.PozFilm = frmSpracovavanie.FilmyBindingSource.Find("Film", VybranyFilm)
8:
9:      frmFilmyUprav.ShowDialog()
10:
11:     Call ZaplnitTabulku(sender, e)
12:
        Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub DataGridView1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DataGridView1.KeyPress
        frmPredstavenia.btnPridaj_KeyPress(sender, e, Me)
    End Sub

    Private Sub txtNFilmu_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtNFilmu.KeyPress
        frmDistribut.txtDistributor_KeyPress(sender, e, Me)
    End Sub

    Private Sub txtPevPozic_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPevPozic.KeyPress
        frmPredstavenia.txtCena_KeyPress(sender, e, Me)
    End Sub

    Private Sub txtVarPozic_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtVarPozic.KeyPress
        frmPredstavenia.txtCena_KeyPress(sender, e, Me)
    End Sub

End Class