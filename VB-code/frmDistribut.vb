Public Class frmDistribut

    Dim VizualUprava As Boolean = True

    Private Sub frmDistribut_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
1:      On Error GoTo Chyba
2:      'Zafarbi
3:      Call mdlColors.Skining(Me)
4:      If VizualUprava Then
5:          Call mdlColors.Sizing(Me)
6:          Me.CenterToParent()
7:          VizualUprava = False
8:      End If
9:
10:     '************ SPRAVA DATA GRIDU*************************
11:     With DataGridView1
12:         .SuspendLayout()
13:         .DataSource = Nothing
14:         .AllowUserToAddRows = False
15:         .AllowUserToDeleteRows = False
16:         .AllowUserToResizeRows = False
17:         .AllowUserToResizeColumns = False
18:
19:         .AllowUserToOrderColumns = True
20:         .SelectionMode = DataGridViewSelectionMode.FullRowSelect
21:         .ReadOnly = True
22:         .MultiSelect = False
23:         .RowHeadersVisible = False
24:         .Columns.Clear()
25:         .Rows.Clear()
26:         ' setup columns
27:
28:         .Columns.Add("dtgrdDatum", "Distributor")
29:         .Columns(0).Width = (.Width - 20) * 0.17
30:         .Columns.Add("dtgrdDen", "Mesto")
31:         .Columns(1).Width = (.Width - 20) * 0.15
32:         .Columns(1).SortMode = DataGridViewColumnSortMode.NotSortable
33:         .Columns.Add("dtgrdCas", "Ulica")
34:         .Columns(2).Width = (.Width - 20) * 0.15
35:         .Columns(2).SortMode = DataGridViewColumnSortMode.NotSortable
36:         .Columns.Add("dtgrdFilm", "PSC")
37:         .Columns(3).Width = (.Width - 20) * 0.08
38:         .Columns(3).SortMode = DataGridViewColumnSortMode.NotSortable
39:         .Columns.Add("dtgrdCena", "Stat")
40:         .Columns(4).Width = (.Width - 20) * 0.11
41:         .Columns(4).SortMode = DataGridViewColumnSortMode.NotSortable
42:         .Columns.Add("dtgrdCenaZlav", "ICO")
43:         .Columns(5).Width = (.Width - 20) * 0.11
44:         .Columns(5).SortMode = DataGridViewColumnSortMode.NotSortable
45:         .Columns.Add("dtgrdSala", "DIC")
46:         .Columns(6).Width = (.Width - 20) * 0.13
47:         .Columns(6).SortMode = DataGridViewColumnSortMode.NotSortable
48:         .Columns.Add("dtgrdTelefon", "Telefon")
49:         .Columns(7).Width = (.Width - 20) * 0.14
50:         .Columns(7).SortMode = DataGridViewColumnSortMode.NotSortable
51:         .Columns.Add("dtgrdUcet", "Cislo uctu")
52:         .Columns(8).Width = (.Width - 20) * 0.18
53:         .Columns(8).SortMode = DataGridViewColumnSortMode.NotSortable
54:         .Columns.Add("dtgrdVarSymb", "Variabilny symbol")
55:         .Columns(9).Width = (.Width - 20) * 0.1
56:         .Columns(9).SortMode = DataGridViewColumnSortMode.NotSortable
57:
58:         .ResumeLayout(True)
59:     End With
60:
61:     Call ZaplnitTabulku(sender, e)
62:
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
6:          frmSpracovavanie.DistributoriBindingSource.MoveFirst()
7:          For i = 1 To frmSpracovavanie.DistributoriBindingSource.Count
8:              .Rows.Add()
9:              .Rows(.Rows.Count - 1).Cells(0).Value = frmSpracovavanie.DistributoriBindingSource.Current("Distributor") & ""
10:             .Rows(.Rows.Count - 1).Cells(1).Value = frmSpracovavanie.DistributoriBindingSource.Current("Mesto") & ""
11:             .Rows(.Rows.Count - 1).Cells(2).Value = frmSpracovavanie.DistributoriBindingSource.Current("Ulica") & ""
12:             .Rows(.Rows.Count - 1).Cells(3).Value = frmSpracovavanie.DistributoriBindingSource.Current("PSC") & ""
13:             .Rows(.Rows.Count - 1).Cells(4).Value = frmSpracovavanie.DistributoriBindingSource.Current("Stat") & ""
14:             .Rows(.Rows.Count - 1).Cells(5).Value = frmSpracovavanie.DistributoriBindingSource.Current("ICO") & ""
15:             .Rows(.Rows.Count - 1).Cells(6).Value = frmSpracovavanie.DistributoriBindingSource.Current("DIC") & ""
16:             .Rows(.Rows.Count - 1).Cells(7).Value = frmSpracovavanie.DistributoriBindingSource.Current("Telefon") & ""
17:             .Rows(.Rows.Count - 1).Cells(8).Value = frmSpracovavanie.DistributoriBindingSource.Current("CisloUctu") & ""
18:             .Rows(.Rows.Count - 1).Cells(9).Value = frmSpracovavanie.DistributoriBindingSource.Current("VariabilnySymbol") & ""
19:             frmSpracovavanie.DistributoriBindingSource.MoveNext()
20:         Next i
21:     End With
22:
23:     txtDistributor.Text = ""
24:     txtMesto.Text = ""
25:     txtUlica.Text = ""
26:     txtPSC.Text = ""
27:     txtStat.Text = ""
28:     txtICO.Text = ""
29:     txtDIC.Text = ""
30:     txtTelefon.Text = ""
31:     txtCisloUctu.Text = ""
32:     txtVarSymb.Text = ""

        Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub DataGridView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.Click
1:      On Error GoTo Chyba
2:      Dim i As Integer
3:      i = DataGridView1.CurrentRow.Index
4:
5:      txtDistributor.Text = DataGridView1.Rows(i).Cells(0).Value & ""
6:      txtMesto.Text = DataGridView1.Rows(i).Cells(1).Value & ""
7:      txtUlica.Text = DataGridView1.Rows(i).Cells(2).Value & ""
8:      txtPSC.Text = DataGridView1.Rows(i).Cells(3).Value & ""
9:      txtStat.Text = DataGridView1.Rows(i).Cells(4).Value & ""
10:     txtICO.Text = DataGridView1.Rows(i).Cells(5).Value & ""
11:     txtDIC.Text = DataGridView1.Rows(i).Cells(6).Value & ""
12:     txtTelefon.Text = DataGridView1.Rows(i).Cells(7).Value & ""
13:     txtCisloUctu.Text = DataGridView1.Rows(i).Cells(8).Value & ""
14:     txtVarSymb.Text = DataGridView1.Rows(i).Cells(9).Value & ""

        Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
1:      On Error GoTo Chyba
2:      'Otvori menu na upravu riadku. Najprv sa nastavi na prislusny riadok databazy, Potom otvori okno
3:      frmMenu.PozDistrib = frmSpracovavanie.DistributoriBindingSource.Find("Distributor", txtDistributor.Text)
4:
5:      frmDistributUprav.ShowDialog()
6:
7:      Call ZaplnitTabulku(sender, e)
        Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub btnZmazat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnZmazat.Click
1:      On Error GoTo Chyba
2:      If MsgBox("Naozaj zmazat distributora" & Chr(13) & txtDistributor.Text & "?", MsgBoxStyle.YesNo + MsgBoxStyle.Question) = MsgBoxResult.No Then Exit Sub
3:      If frmSpracovavanie.FilmyBindingSource.Find("Distributor", txtDistributor.Text) <> -1 Then MsgBox("Nemozno zmazat distributora." & Chr(13) & "K distributorovi " & txtDistributor.Text & " su vlozene filmy.", MsgBoxStyle.Exclamation) : Exit Sub
4:
5:      Module1.WriteLog(Chr(13) & Chr(10) & Now & " - Zmazanie distributora: " & txtDistributor.Text)
6:
7:      With frmSpracovavanie.DistributoriBindingSource
8:          If .Find("Distributor", txtDistributor.Text) = -1 Then
9:              MsgBox("Chyba pri mazani distributora!", MsgBoxStyle.Critical)
10:             Exit Sub
11:         Else
12:             .Position = .Find("Distributor", txtDistributor.Text)
13:             frmSpracovavanie.DistributoriTableAdapter.Delete(.Current(0), .Current(1), .Current(2), .Current(3), .Current(4), .Current(5), .Current(6), .Current(7), .Current(8), .Current(9), .Current(10))
14:             frmSpracovavanie.DistributoriTableAdapter.Fill(frmSpracovavanie.DataSet1.Distributori)
15:         End If
16:     End With
17:
18:     Call Module1.HASHdbKontrola("Distributori")
19:     Call Module1.HASHSubor("Data\data.pac")
20:     Call ZaplnitTabulku(sender, e)
21:     Module1.WriteLog("   OK")
22:     frmSpracovavanie.Zaloha = True
23:     Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub btnPridaj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPridaj.Click
1:      On Error GoTo Chyba
2:      If txtDistributor.Text = "" Then MsgBox("Zadaj meno Distributora") : Exit Sub
3:      If frmSpracovavanie.DistributoriBindingSource.Find("Distributor", txtDistributor.Text) <> -1 Then MsgBox("Distributor s takým menom uz je ulozený.", MsgBoxStyle.Exclamation) : Exit Sub
4:      If MsgBox("Je meno distributora zadané správne?" & Chr(13) & "Po ulození ho uz nebude mozné menit.", MsgBoxStyle.YesNo + MsgBoxStyle.Question) = MsgBoxResult.No Then Exit Sub
5:
6:      Dim HASH As String
7:      Dim x As Integer
8:
9:      Module1.WriteLog(Chr(13) & Chr(10) & Now & " - Vlozenie distributora: " & txtDistributor.Text)
10:
11:     With frmSpracovavanie.DistributoriTableAdapter
12:         x = frmSpracovavanie.DistributoriBindingSource.Find("Distributor", txtDistributor.Text)
13:         If x <> -1 Then MsgBox("Distributor je uz zadany.", MsgBoxStyle.Exclamation) : Exit Sub
14:
15:         .Insert(txtDistributor.Text, txtMesto.Text, txtUlica.Text, txtPSC.Text, txtStat.Text, txtICO.Text, txtDIC.Text, txtTelefon.Text, txtCisloUctu.Text, txtVarSymb.Text, "")
16:         .Fill(frmSpracovavanie.DataSet1.Distributori)
17:         HASH = Module1.RetazecRiadku("Distributori", frmSpracovavanie.DistributoriBindingSource.Find("crc", ""))
18:         frmSpracovavanie.DistributoriBindingSource.Current("CRC") = Module1.HASHString(HASH)
19:         frmSpracovavanie.DistributoriBindingSource.EndEdit()
20:         .Update(frmSpracovavanie.DataSet1.Distributori)
21:         .Fill(frmSpracovavanie.DataSet1.Distributori)
22:     End With
23:
24:     Call Module1.HASHdbKontrola("Distributori")
25:     Call Module1.HASHSubor("Data\data.pac")
26:     txtDistributor.Text = ""
27:     Call ZaplnitTabulku(sender, e)
28:     Module1.WriteLog("   OK")
29:     frmSpracovavanie.Zaloha = True
30:     Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub frmDistribut_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        frmPredstavenia.btnPridaj_KeyPress(sender, e, Me)
    End Sub

    Private Sub btnPridaj_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles btnPridaj.KeyPress
        frmPredstavenia.btnPridaj_KeyPress(sender, e, Me)
    End Sub

    Private Sub txtCisloUctu_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCisloUctu.KeyPress
        txtDistributor_KeyPress(sender, e, Me)
    End Sub

    Private Sub txtDIC_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDIC.KeyPress
        txtDistributor_KeyPress(sender, e, Me)
    End Sub

    Private Sub txtICO_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtICO.KeyPress
        txtDistributor_KeyPress(sender, e, Me)
    End Sub

    Private Sub txtMesto_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMesto.KeyPress
        txtDistributor_KeyPress(sender, e, Me)
    End Sub

    Private Sub txtUlica_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtUlica.KeyPress
        txtDistributor_KeyPress(sender, e, Me)
    End Sub

    Private Sub txtPSC_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPSC.KeyPress
        txtDistributor_KeyPress(sender, e, Me)
    End Sub

    Private Sub txtStat_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtStat.KeyPress
        txtDistributor_KeyPress(sender, e, Me)
    End Sub

    Private Sub txtTelefon_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtTelefon.KeyPress
        txtDistributor_KeyPress(sender, e, Me)
    End Sub

    Private Sub txtVarSymb_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtVarSymb.KeyPress
        txtDistributor_KeyPress(sender, e, Me)
    End Sub

    Private Sub DataGridView1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DataGridView1.KeyPress
        frmPredstavenia.btnPridaj_KeyPress(sender, e, Me)
    End Sub

    Private Sub btnZmazat_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles btnZmazat.KeyPress
        frmPredstavenia.btnPridaj_KeyPress(sender, e, Me)
    End Sub

    Private Sub txtDistributor_KeyPress1(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDistributor.KeyPress
        txtDistributor_KeyPress(sender, e, Me)
    End Sub

    Public Sub txtDistributor_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs, ByVal host As System.Windows.Forms.Form)
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If KeyAscii = System.Windows.Forms.Keys.Return Then
            KeyAscii = 0
            System.Windows.Forms.SendKeys.Send("{tab}")
        ElseIf KeyAscii = System.Windows.Forms.Keys.Escape Then
            host.Close()
        End If
        e.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            e.Handled = True
        End If
    End Sub
End Class