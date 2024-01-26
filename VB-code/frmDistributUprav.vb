Public Class frmDistributUprav

Dim VizualUprava As Boolean = True

Private Sub btnZrusit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnZrusit.Click
    Me.Close()
End Sub

Private Sub frmDistributUprav_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
1:  'Zafarbi
2:  Call mdlColors.Skining(Me)
3:  If VizualUprava Then
4:      Call mdlColors.Sizing(Me)
5:      Me.CenterToParent()
6:      VizualUprava = False
7:  End If
8:
9:  'Vyplni okna
10: With frmSpracovavanie.DistributoriBindingSource
11:     .Position = frmMenu.PozDistrib
12:     lblDistributor.Text = .Current("Distributor") & ""
13:     txtTelefon.Text = .Current("telefon") & ""
14:     txtUlica.Text = .Current("ulica") & ""
15:     txtPSC.Text = .Current("psc") & ""
16:     txtMesto.Text = .Current("Mesto") & ""
17:     txtStat.Text = .Current("stat") & ""
18:     txtICO.Text = .Current("ico") & ""
19:     txtDIC.Text = .Current("dic") & ""
20:     txtCisloUctu.Text = .Current("cislouctu") & ""
21:     txtVarSymb.Text = .Current("variabilnysymbol") & ""
    End With
End Sub

Private Sub btnUlozit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUlozit.Click
1:  On Error GoTo Chyba
2:  Dim HASH As String
3:
4:  Module1.WriteLog(Chr(13) & Chr(10) & Now & " - Uprava distributora: " & lblDistributor.Text)
5:
6:  With frmSpracovavanie.DistributoriBindingSource
7:      .Current("telefon") = txtTelefon.Text
8:      .Current("ulica") = txtUlica.Text
9:      .Current("psc") = txtPSC.Text
10:     .Current("Mesto") = txtMesto.Text
11:     .Current("stat") = txtStat.Text
12:     .Current("ico") = txtICO.Text
13:     .Current("dic") = txtDIC.Text
14:     .Current("cislouctu") = txtCisloUctu.Text
15:     .Current("variabilnysymbol") = txtVarSymb.Text
16:
17:     HASH = Module1.RetazecRiadku("Distributori", frmMenu.PozDistrib)
18:     .Current("CRC") = Module1.HASHString(HASH)
19:     .EndEdit()
20:     frmSpracovavanie.DistributoriTableAdapter.Update(frmSpracovavanie.DataSet1.Distributori)
21:     frmSpracovavanie.DistributoriTableAdapter.Fill(frmSpracovavanie.DataSet1.Distributori)
22: End With
23:
24: Call Module1.HASHdbKontrola("Distributori")
25: Call Module1.HASHSubor("Data\data.pac")
26: Module1.WriteLog("   OK")
27: MsgBox("Udaje zmenene", MsgBoxStyle.Information)
28:
29: Me.Close()
30:
31: Exit Sub
Chyba:
    Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
End Sub

Private Sub txtCisloUctu_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCisloUctu.KeyPress
    frmDistribut.txtDistributor_KeyPress(sender, e, Me)
End Sub

Private Sub txtDIC_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDIC.KeyPress
    frmDistribut.txtDistributor_KeyPress(sender, e, Me)
End Sub

Private Sub txtICO_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtICO.KeyPress
    frmDistribut.txtDistributor_KeyPress(sender, e, Me)
End Sub

Private Sub txtMesto_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMesto.KeyPress
    frmDistribut.txtDistributor_KeyPress(sender, e, Me)
End Sub

Private Sub txtUlica_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtUlica.KeyPress
    frmDistribut.txtDistributor_KeyPress(sender, e, Me)
End Sub

Private Sub txtPSC_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPSC.KeyPress
    frmDistribut.txtDistributor_KeyPress(sender, e, Me)
End Sub

Private Sub txtStat_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtStat.KeyPress
    frmDistribut.txtDistributor_KeyPress(sender, e, Me)
End Sub

Private Sub txtTelefon_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtTelefon.KeyPress
    frmDistribut.txtDistributor_KeyPress(sender, e, Me)
End Sub

Private Sub txtVarSymb_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtVarSymb.KeyPress
    frmDistribut.txtDistributor_KeyPress(sender, e, Me)
End Sub

Private Sub btnUlozit_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles btnUlozit.KeyPress
    frmPredstavenia.btnPridaj_KeyPress(sender, e, Me)
End Sub

Private Sub btnZrusit_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles btnZrusit.KeyPress
    frmPredstavenia.btnPridaj_KeyPress(sender, e, Me)
End Sub
End Class