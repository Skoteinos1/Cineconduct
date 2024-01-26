Public Class frmPredstaveniaUprav

    Dim VizualUprava As Boolean = True

    Private Sub frmPredstaveniaUprav_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
1:      On Error GoTo Chyba
2:      'Zafarbi
3:      Call mdlColors.Skining(Me)
4:      If VizualUprava Then
5:          Call mdlColors.Sizing(Me)
6:          Me.CenterToParent()
7:          VizualUprava = False
8:      End If
9:
10:     'Naplni combobox s salami
11:     cmbSala.Items.Clear()
12:     For i = 1 To frmSpracovavanie.PocetSal
13:         cmbSala.Items.Add(frmSpracovavanie.MenoSal(i))
            'If i = 1 Then cmbSala.Items.Add(frmSpracovavanie.MenoSal1)
14:         'If i = 2 Then cmbSala.Items.Add(frmSpracovavanie.MenoSal2)
15:         'If i = 3 Then cmbSala.Items.Add(frmSpracovavanie.MenoSal3)
16:         'If i = 4 Then cmbSala.Items.Add(frmSpracovavanie.MenoSal4)
17:         'If i = 5 Then cmbSala.Items.Add(frmSpracovavanie.MenoSal5)
18:     Next i
19:     cmbSala.SelectedIndex = 0
20:
21:     'Vyplni okna
22:     With frmSpracovavanie.PredstaveniaBindingSource
23:         .Position = frmVybPredst.PozPredstav
24:         lblFilm.Text = .Current("NazovFilmu") & ""
25:         lblDatum.Text = Format(CDate(.Current("Datum") & ""), "dd.MM.yyyy")
26:         lblCas.Text = Format(CDate(.Current("Datum") & ""), "HH:mm")
27:         txtCena.Text = Format(.Current("cenalistka"), "###0.00")
28:         txtCenaZlavnena.Text = Format(.Current("cenazlavlistka"), "###0.00")
29:         cmbSala.SelectedIndex = .Current("sala") - 1
            If CDbl(.Current("TrzbaPredstavenia") & "") <> 0 Then cmbSala.Enabled = False Else cmbSala.Enabled = True
30:     End With
31:
        Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub btnZrusit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnZrusit.Click
        Me.Close()
    End Sub

    Private Sub btnUlozit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUlozit.Click
1:      On Error GoTo Chyba
2:      If txtCena.Text = "" Then txtCena.Text = 0
        If txtCenaZlavnena.Text = "" Then txtCenaZlavnena.Text = txtCena.Text
        If CDbl(txtCena.Text) < CDbl(txtCenaZlavnena.Text) Then MsgBox("Zlavnena cena je vyssia ako bezne vstupne.", MsgBoxStyle.Exclamation) : Exit Sub
3:      If CDbl(txtCena.Text) = 0 Then
4:          If MsgBox("Vstup na predstavenie zdarma?", MsgBoxStyle.YesNo + MsgBoxStyle.Question) = MsgBoxResult.No Then Exit Sub
5:      End If
7:      Module1.WriteLog(Chr(13) & Chr(10) & Now & " - Uprava predstavena: " & lblFilm.Text)
8:      Dim HASH As String
9:
10:     With frmSpracovavanie.PredstaveniaBindingSource
11:         .Position = frmVybPredst.PozPredstav
12:
13:         .Current("CenaListka") = txtCena.Text
14:         .Current("CenaZlavListka") = txtCenaZlavnena.Text
15:         .Current("Sala") = cmbSala.SelectedIndex + 1
16:
17:         HASH = Module1.RetazecRiadku("Predstavenia", frmVybPredst.PozPredstav)
18:         .Current("CRC") = Module1.HASHString(HASH)
19:         Module1.UpdatePredstavenia(frmVybPredst.PozPredstav)
20:     End With
21:
22:     'Call Module1.HASHdbKontrola()
23:     Call Module1.HASHSubor("Data\data.pac")
24:
25:     MsgBox("Udaje zmenene", MsgBoxStyle.Information)
26:     Module1.WriteLog("   OK")
27:     Me.Close()
28:
29:     Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub txtCena_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCena.KeyPress
        frmPredstavenia.txtCena_KeyPress(sender, e, Me)
    End Sub

    Private Sub txtCenaZlavnena_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCenaZlavnena.KeyPress
        frmPredstavenia.txtCena_KeyPress(sender, e, Me)
    End Sub

    Private Sub cmbSala_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmbSala.KeyPress
        frmPredstavenia.cmbFilmy_KeyPress(sender, e, Me)
    End Sub

    Private Sub btnUlozit_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles btnUlozit.KeyPress
        frmPredstavenia.btnPridaj_KeyPress(sender, e, Me)
    End Sub

    Private Sub btnZrusit_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles btnZrusit.KeyPress
        frmPredstavenia.btnPridaj_KeyPress(sender, e, Me)
    End Sub
End Class