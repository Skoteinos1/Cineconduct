Public Class frmFilmyUprav

    Dim VizualUprava As Boolean = True

    Private Sub btnZrusit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnZrusit.Click
        Me.Close()
    End Sub

    Private Sub btnUloz_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUloz.Click
1:      On Error GoTo Chyba
2:      If txtPevPozic.Text = "" Then txtPevPozic.Text = 0
3:      If txtVarPozic.Text = "" Then txtVarPozic.Text = 0
4:      If CDbl(txtVarPozic.Text) = 0 And CDbl(txtPevPozic.Text) = 0 Then MsgBox("Zadajte pozicovné filmu!", MsgBoxStyle.Exclamation) : Exit Sub
5:      If CDbl(txtVarPozic.Text) <> 0 And CDbl(txtPevPozic.Text) <> 0 Then MsgBox("Zadajte iba jeden druh pozicovného!", MsgBoxStyle.Exclamation) : Exit Sub
6:      If CDbl(txtVarPozic.Text) > 100 Then MsgBox("Nemozno vlozit viac ako 100% pozicovne!", MsgBoxStyle.Exclamation) : Exit Sub
7:      Module1.WriteLog(Chr(13) & Chr(10) & Now & " - Uprava filmu: " & lblFilm.Text)
8:      Dim HASH As String
9:
10:     With frmSpracovavanie.FilmyBindingSource
11:         .Position = frmVybPredst.PozFilm
12:
13:         .Current("Film") = lblFilm.Text
14:         .Current("distributor") = cmbDistrib.Text
15:         .Current("Pevnepozicovne") = txtPevPozic.Text
16:         .Current("Percentpozicovne") = txtVarPozic.Text
17:
18:         HASH = Module1.RetazecRiadku("Filmy", frmVybPredst.PozFilm)
19:         .Current("CRC") = Module1.HASHString(HASH)
20:
21:         Module1.UpdateFilmy(frmVybPredst.PozFilm)
22:     End With
23:
24:     Call Module1.HASHdbKontrola("Filmy")
25:     Call Module1.HASHSubor("Data\data.pac")
26:
27:     MsgBox("Udaje zmenene", MsgBoxStyle.Information)
28:     Module1.WriteLog("   OK")
29:     Me.Close()
30:
        Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub frmFilmyUprav_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
1:      On Error GoTo Chyba
2:      'Zafarbi
3:      Call mdlColors.Skining(Me)
4:      If VizualUprava Then
5:          Call mdlColors.Sizing(Me)
6:          Me.CenterToParent()
7:          VizualUprava = False
8:      End If
9:
10:     'Naplni combobox z distributormi
11:     With frmSpracovavanie.DistributoriBindingSource
12:         cmbDistrib.Items.Clear()
13:         .MoveFirst()
14:         For i = 1 To .Count
15:             cmbDistrib.Items.Add(.Current("Distributor"))
16:             .MoveNext()
17:         Next i
18:     End With
19:
20:     'Vyplni okna
21:     With frmSpracovavanie.FilmyBindingSource
22:         .Position = frmVybPredst.PozFilm
23:         lblFilm.Text = .Current("Film")
24:         cmbDistrib.SelectedIndex = cmbDistrib.Items.IndexOf(.Current("distributor"))
25:         txtPevPozic.Text = .Current("Pevnepozicovne")
26:         txtVarPozic.Text = .Current("Percentpozicovne")
27:     End With

        Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub btnUloz_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles btnUloz.KeyPress
        frmPredstavenia.btnPridaj_KeyPress(sender, e, Me)
    End Sub

    Private Sub btnZrusit_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles btnZrusit.KeyPress
        frmPredstavenia.btnPridaj_KeyPress(sender, e, Me)
    End Sub

    Private Sub txtPevPozic_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPevPozic.KeyPress
        frmPredstavenia.txtCena_KeyPress(sender, e, Me)
    End Sub

    Private Sub txtVarPozic_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtVarPozic.KeyPress
        frmPredstavenia.txtCena_KeyPress(sender, e, Me)
    End Sub

    Private Sub cmbDistrib_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmbDistrib.KeyPress
        frmPredstavenia.cmbFilmy_KeyPress(sender, e, Me)
    End Sub

End Class