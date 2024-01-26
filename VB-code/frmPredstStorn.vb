Public Class frmPredstStorn

    Dim VizualUprava As Boolean = True

    Private Sub frmPredstStorn_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
1:      On Error GoTo Chyba
2:
3:      'Zafarbi
4:      Call mdlColors.Skining(Me)
5:      If VizualUprava Then
6:          Call mdlColors.Sizing(Me)
7:          Me.CenterToParent()
8:          VizualUprava = False
9:      End If
10:
11:     '************ SPRAVA DATA GRIDU*************************
12:     With DataGridView1
13:         .SuspendLayout()
14:         .DataSource = Nothing
15:         .AllowUserToAddRows = False
16:         .AllowUserToDeleteRows = False
17:         .AllowUserToResizeRows = False
18:         .AllowUserToResizeColumns = False
19:
20:         .AllowUserToOrderColumns = True
21:         .SelectionMode = DataGridViewSelectionMode.FullRowSelect
22:         .ReadOnly = True
23:         .MultiSelect = False
24:         .RowHeadersVisible = False
25:         .Columns.Clear()
26:         .Rows.Clear()
27:         ' setup columns
28:
29:         .Columns.Add("dtgrdDatum", "Suma (EUR)")
30:         .Columns(0).Width = (.Width - 20) * 0.15
31:         .Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
32:         .Columns.Add("dtgrdDen", "Osob")
33:         .Columns(1).Width = (.Width - 20) * 0.15
34:         .Columns(1).SortMode = DataGridViewColumnSortMode.NotSortable
35:         .Columns.Add("dtgrdCas", "Dôvod storna")
36:         .Columns(2).Width = (.Width - 20) * 0.7
37:         .Columns(2).SortMode = DataGridViewColumnSortMode.NotSortable
38:
39:         .ResumeLayout(True)
40:     End With
41:
42:     With frmSpracovavanie.PredstaveniaBindingSource
43:         .Position = frmVybPredst.PozPredstav
44:         lblCas.Text = Format(CDate(.Current("Datum")), "HH:mm")
45:         lblDatum.Text = Format(CDate(.Current("datum")), "dd.MM.yyyy")
46:         lblFilm.Text = .Current("NazovFilmu")
47:     End With
48:
49:     Call ZaplnitTabulku(sender, e)
50:
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
6:          frmSpracovavanie.StornoPredstaveniBindingSource.MoveFirst()
7:          For i = 1 To frmSpracovavanie.StornoPredstaveniBindingSource.Count
8:              If CDate(lblDatum.Text) = CDate(frmSpracovavanie.StornoPredstaveniBindingSource.Current("datum")) _
                And CDate(lblCas.Text) = Format(CDate(frmSpracovavanie.StornoPredstaveniBindingSource.Current("cas")), "HH:mm") Then
10:                 .Rows.Add()
11:                 .Rows(.Rows.Count - 1).Cells(0).Value = frmSpracovavanie.StornoPredstaveniBindingSource.Current("suma") & ""
12:                 .Rows(.Rows.Count - 1).Cells(1).Value = frmSpracovavanie.StornoPredstaveniBindingSource.Current("osob") & ""
13:                 .Rows(.Rows.Count - 1).Cells(2).Value = frmSpracovavanie.StornoPredstaveniBindingSource.Current("dovod") & ""
14:             End If
15:             frmSpracovavanie.StornoPredstaveniBindingSource.MoveNext()
16:         Next i
17:     End With
18:
19:     txtDovod.Text = ""
20:     txtOsob.Text = ""
21:     txtSuma.Text = ""
22:
23:     Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub DataGridView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.Click
        On Error GoTo Chyba
        Dim i As Integer = DataGridView1.CurrentRow.Index

        txtSuma.Text = DataGridView1.Rows(i).Cells(0).Value & ""
        txtOsob.Text = DataGridView1.Rows(i).Cells(1).Value & ""
        txtDovod.Text = DataGridView1.Rows(i).Cells(2).Value & ""
        Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub btnPridaj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPridaj.Click
1:      On Error GoTo Chyba
2:      If txtSuma.Text = "" Or CDbl(txtSuma.Text) = 0 Then MsgBox("Nezadaná suma storna!", MsgBoxStyle.Exclamation) : Exit Sub
3:      If txtDovod.Text = "" Then MsgBox("Nevyplnený dôvod storna!", MsgBoxStyle.Exclamation) : Exit Sub
4:      If txtOsob.Text = "" Then txtOsob.Text = 0
5:      Module1.WriteLog(Chr(13) & Chr(10) & Now & " - Velke storno k predstaveniu: " & lblFilm.Text)
6:      Dim i As Object
7:
8:      With DataGridView1
9:          For i = 0 To .RowCount - 1
10:             If DataGridView1.Rows(i).Cells(2).Value = txtDovod.Text Then MsgBox("Náklad k filmu uz bol zadaný!", MsgBoxStyle.Exclamation) : Exit Sub
11:         Next i
12:     End With
13:
14:     Dim HASH As String
15:
16:     With frmSpracovavanie.StornoPredstaveniTableAdapter
17:         .Insert(lblDatum.Text, lblCas.Text, lblFilm.Text, txtSuma.Text, txtOsob.Text, txtDovod.Text, "")
18:         .Fill(frmSpracovavanie.DataSet1.StornoPredstaveni)
19:         HASH = Module1.RetazecRiadku("StornoPredstaveni", frmSpracovavanie.StornoPredstaveniBindingSource.Find("CRC", ""))
20:         frmSpracovavanie.StornoPredstaveniBindingSource.Current("CRC") = Module1.HASHString(HASH)
21:         frmSpracovavanie.StornoPredstaveniBindingSource.EndEdit()
22:         .Update(frmSpracovavanie.DataSet1.StornoPredstaveni)
23:         .Fill(frmSpracovavanie.DataSet1.StornoPredstaveni)
24:     End With
25:
26:     Call Module1.HASHdbKontrola("StornoPredstaveni")
27:     Call Module1.HASHSubor("Data\data.pac")
        Module1.WriteLog("   OK")
        frmSpracovavanie.Zaloha = True
28:     Call ZaplnitTabulku(sender, e)
29:
30:     Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub btnZmazat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnZmazat.Click
1:      On Error GoTo Chyba
2:      Dim i As Integer
3:      Module1.WriteLog(Chr(13) & Chr(10) & Now & " - Zrusenie velkeho storna k predstaveniu: " & lblFilm.Text)
4:      'zmazat z nakladov
5:      With frmSpracovavanie.StornoPredstaveniBindingSource
6:          .MoveFirst()
7:          For i = 1 To .Count
8:              If .Current("film") = lblFilm.Text Then
9:                  If CDate(.Current("datum")) = CDate(lblDatum.Text) And Format(CDate(.Current("cas")), "HH:mm") = CDate(lblCas.Text) Then
10:                     frmSpracovavanie.StornoPredstaveniTableAdapter.Delete(.Current(0), .Current(1), .Current(2), .Current(3), .Current(4), .Current(5), .Current(6))
11:                     frmSpracovavanie.StornoPredstaveniTableAdapter.Fill(frmSpracovavanie.DataSet1.StornoPredstaveni)
12:                 End If
13:             End If
14:             .MoveNext()
15:         Next i
16:     End With
17:
18:     Call Module1.HASHdbKontrola("StornoPredstaveni")
19:     Call Module1.HASHSubor("Data\data.pac")
        Module1.WriteLog("   OK")
        frmSpracovavanie.Zaloha = True
20:     Call ZaplnitTabulku(sender, e)
21:
        Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub btnPridaj_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles btnPridaj.KeyPress
        frmPredstavenia.btnPridaj_KeyPress(sender, e, Me)
    End Sub

    Private Sub DataGridView1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DataGridView1.KeyPress
        frmPredstavenia.btnPridaj_KeyPress(sender, e, Me)
    End Sub

    Private Sub btnZmazat_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles btnZmazat.KeyPress
        frmPredstavenia.btnPridaj_KeyPress(sender, e, Me)
    End Sub

    Private Sub txtDovod_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDovod.KeyPress
        frmDistribut.txtDistributor_KeyPress(sender, e, Me)
    End Sub

    Private Sub txtSuma_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSuma.KeyPress
        frmPredstavenia.txtCena_KeyPress(sender, e, Me)
    End Sub

    Private Sub txtOsob_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtOsob.KeyPress
        frmPredstavenia.txtCena_KeyPress(sender, e, Me)
    End Sub

End Class