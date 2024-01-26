Public Class frmFilmyNaklady

    Dim VizualUprava As Boolean = True

    Private Sub frmFilmyNaklady_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
1:      On Error GoTo Chyba
2:      frmSpracovavanie.FilmyBindingSource.Position = frmVybPredst.PozFilm
3:      lblFilm.Text = frmSpracovavanie.FilmyBindingSource.Current("Film")
4:
5:      'Zafarbi
6:      Call mdlColors.Skining(Me)
7:      If VizualUprava Then
8:          Call mdlColors.Sizing(Me)
9:          Me.CenterToParent()
10:         VizualUprava = False
11:     End If
12:
13:     '************ SPRAVA DATA GRIDU*************************
14:     With DataGridView1
15:         .SuspendLayout()
16:         .DataSource = Nothing
17:         .AllowUserToAddRows = False
18:         .AllowUserToDeleteRows = False
19:         .AllowUserToResizeRows = False
20:         .AllowUserToResizeColumns = False
21:
22:         .AllowUserToOrderColumns = True
23:         .SelectionMode = DataGridViewSelectionMode.FullRowSelect
24:         .ReadOnly = True
25:         .MultiSelect = False
26:         .RowHeadersVisible = False
27:         .Columns.Clear()
28:         .Rows.Clear()
29:         ' setup columns
30:
31:         .Columns.Add("dtgrdDatum", "Názov filmu")
32:         .Columns(0).Width = (.Width - 20) * 0.3
33:         .Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
34:         .Columns.Add("dtgrdDen", "Názov nákladu")
35:         .Columns(1).Width = (.Width - 20) * 0.3
36:         .Columns(1).SortMode = DataGridViewColumnSortMode.NotSortable
37:         .Columns.Add("dtgrdCas", "Náklad návstevníka (EUR)")
38:         .Columns(2).Width = (.Width - 20) * 0.2
39:         .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
40:         .Columns(2).SortMode = DataGridViewColumnSortMode.NotSortable
41:         .Columns.Add("dtgrdFilm", "Náklad predstavenia (EUR)")
42:         .Columns(3).Width = (.Width - 20) * 0.2
43:         .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
44:         .Columns(3).SortMode = DataGridViewColumnSortMode.NotSortable
45:
46:         .ResumeLayout(True)
47:     End With
48:
49:     Call ZaplnitTabulku(sender, e)

        Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub btnPridaj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPridaj.Click
1:      On Error GoTo Chyba
2:      If txtNazNaklad.Text = "" Then MsgBox("Nevyplnený názov nákladu!", MsgBoxStyle.Exclamation) : Exit Sub
3:      If CDbl(txtSumaNakl.Text) <= 0 Then MsgBox("Zadajte výsku nákladu!", MsgBoxStyle.Exclamation) : Exit Sub
4:      Module1.WriteLog(Chr(13) & Chr(10) & Now & " - Vlozenie nakladu k filmu: " & lblFilm.Text)
5:      Dim i As Integer
6:
7:      With DataGridView1
8:          For i = 0 To .RowCount - 1
9:              If DataGridView1.Rows(i).Cells(1).Value = txtNazNaklad.Text Then MsgBox("Náklad k filmu uz bol zadaný!", MsgBoxStyle.Exclamation) : Exit Sub
10:         Next i
11:     End With
12:
13:     Dim HASH As String
14:
15:     With frmSpracovavanie.NakladFilmTableAdapter
16:         If radNaklNavstev.Checked = True Then .Insert(lblFilm.Text, txtNazNaklad.Text, txtSumaNakl.Text, 0, "")
17:         If radNaklPredst.Checked = True Then .Insert(lblFilm.Text, txtNazNaklad.Text, 0, txtSumaNakl.Text, "")
18:         .Fill(frmSpracovavanie.DataSet1.NakladFilm)
19:     End With
20:
21:     With frmSpracovavanie.NakladFilmBindingSource
22:         HASH = Module1.RetazecRiadku("NakladFilm", .Find("CRC", ""))
23:         .Current("CRC") = Module1.HASHString(HASH)
24:         .EndEdit()
25:         frmSpracovavanie.NakladFilmTableAdapter.Update(frmSpracovavanie.DataSet1.NakladFilm)
26:         frmSpracovavanie.NakladFilmTableAdapter.Fill(frmSpracovavanie.DataSet1.NakladFilm)
27:     End With
28:
29:     Call Module1.HASHdbKontrola("NakladFilm")
30:     Call Module1.HASHSubor("Data\data.pac")
31:     Call ZaplnitTabulku(sender, e)
32:     Module1.WriteLog("   OK")
        frmSpracovavanie.Zaloha = True
        Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub btnZmazat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnZmazat.Click
1:      On Error GoTo Chyba
2:      Dim MenFilmu, MenNakl As String
3:      Dim i As Integer
4:      Module1.WriteLog(Chr(13) & Chr(10) & Now & " - Odstranenie nakladu k filmu: " & lblFilm.Text)
5:      i = DataGridView1.CurrentRow.Index
6:      MenFilmu = DataGridView1.Rows(i).Cells(0).Value & ""
7:      MenNakl = DataGridView1.Rows(i).Cells(1).Value & ""
8:
9:      'zmazat z nakladov
10:     With frmSpracovavanie.NakladFilmBindingSource
11:         .MoveFirst()
12:         For i = 1 To .Count
13:             If .Current("NazovFilmu") = MenFilmu And .Current("MenoNakl") = MenNakl Then
14:                 frmSpracovavanie.NakladFilmTableAdapter.Delete(.Current(0), .Current(1), .Current(2), .Current(3), .Current(4))
15:                 frmSpracovavanie.NakladFilmTableAdapter.Fill(frmSpracovavanie.DataSet1.NakladFilm)
16:             End If
17:             .MoveNext()
18:         Next i
19:     End With
20:
21:     Call Module1.HASHdbKontrola("NakladFilm")
22:     Call Module1.HASHSubor("Data\data.pac")
23:     Call ZaplnitTabulku(sender, e)
24:     Module1.WriteLog("   OK")
        frmSpracovavanie.Zaloha = True
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
6:          frmSpracovavanie.NakladFilmBindingSource.MoveFirst()
7:          For i = 1 To frmSpracovavanie.NakladFilmBindingSource.Count
8:              If frmSpracovavanie.NakladFilmBindingSource.Current("NazovFilmu") = lblFilm.Text Then
9:                  .Rows.Add()
10:                 .Rows(.Rows.Count - 1).Cells(0).Value = frmSpracovavanie.NakladFilmBindingSource.Current("NazovFilmu") & ""
11:                 .Rows(.Rows.Count - 1).Cells(1).Value = frmSpracovavanie.NakladFilmBindingSource.Current("MenoNakl") & ""
12:                 .Rows(.Rows.Count - 1).Cells(2).Value = frmSpracovavanie.NakladFilmBindingSource.Current("NaklNavstev") & ""
13:                 .Rows(.Rows.Count - 1).Cells(3).Value = frmSpracovavanie.NakladFilmBindingSource.Current("NaklPredst") & ""
14:             End If
15:             frmSpracovavanie.NakladFilmBindingSource.MoveNext()
16:         Next i
17:     End With
18:
        Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub btnPridaj_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles btnPridaj.KeyPress
        frmPredstavenia.btnPridaj_KeyPress(sender, e, Me)
    End Sub

    Private Sub btnZmazat_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles btnZmazat.KeyPress
        frmPredstavenia.btnPridaj_KeyPress(sender, e, Me)
    End Sub

    Private Sub radNaklPredst_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles radNaklPredst.KeyPress
        frmDistribut.txtDistributor_KeyPress(sender, e, Me)
    End Sub

    Private Sub radNaklNavstev_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles radNaklNavstev.KeyPress
        frmDistribut.txtDistributor_KeyPress(sender, e, Me)
    End Sub

    Private Sub txtNazNaklad_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtNazNaklad.KeyPress
        frmDistribut.txtDistributor_KeyPress(sender, e, Me)
    End Sub

    Private Sub txtSumaNakl_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSumaNakl.KeyPress
        frmPredstavenia.txtCena_KeyPress(sender, e, Me)
    End Sub
End Class