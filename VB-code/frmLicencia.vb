Public Class frmLicencia
    Dim teraz As Date
    Dim pokusy As Integer
    Dim VizualUprava As Boolean = True

    Private Sub frmLicencia_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
1:      'Zafarbi
2:      'Call mdlColors.Skining(Me)
3:      If VizualUprava Then
4:          Call mdlColors.Sizing(Me)
5:          Me.CenterToParent()
6:          VizualUprava = False
7:      End If
8:
9:      pokusy = 0
10:     teraz = Now
11:
12:     txtDni.Text = 400
13:     lblOrg.Text = frmSpracovavanie.MenoKina & ", " & frmSpracovavanie.Mesto
14:     lblKod.Text = LSet(HASHString(lblOrg.Text & teraz), 10)
15:
16:     If 60 < frmSpracovavanie.Licenc.ToOADate - Today.ToOADate And frmLogin.Rights <> "A" Then
17:         MsgBox("Vasu licenciu zatial nie je potrebne predlzovat. Platnost Vasej licencie vyprsi " & Format(frmSpracovavanie.Licenc, "dd.MM.yyyy"))
18:         btnZavriet_Click(sender, e)
19:     End If
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
1:      Dim HASH As String
2:
3:      If frmSpracovavanie.DatumPrihlasenia <> Today Then
4:          MsgBox("S datumom bolo manipulovane. Zmente datum Vasho pocitaca.", MsgBoxStyle.Exclamation)
5:          Exit Sub
6:      End If
7:      Module1.WriteLog(Chr(13) & Chr(10) & Now & " - Pokus o predlzenie licencie")
8:      If txtKluc.Text = LSet(HASHString(lblKod.Text & txtDni.Text), 10) Then
9:          frmSpracovavanie.Licenc = Date.FromOADate(teraz.ToOADate + CInt(txtDni.Text))
10:         With frmSpracovavanie.NastaveniaBindingSource
11:             .Position = .Find("Option", "Demo")
12:             If .Current("setting") <> "01.01.2000" Then
13:                 .Current("setting") = "01.01.2000"
14:                 HASH = Module1.RetazecRiadku("Nastavenia", .Position)
15:                 .Current("CRC") = Module1.HASHString(HASH)
16:             End If
17:
18:             .Position = .Find("Option", "Licencia")
19:             .Current("setting") = Format(frmSpracovavanie.Licenc, "dd.MM.yyyy")
20:             HASH = Module1.RetazecRiadku("Nastavenia", .Position)
21:             .Current("CRC") = Module1.HASHString(HASH)
22:             .EndEdit()
23:             frmSpracovavanie.NastaveniaTableAdapter.Update(frmSpracovavanie.DataSet1.Nastavenia)
24:             frmSpracovavanie.NastaveniaTableAdapter.Fill(frmSpracovavanie.DataSet1.Nastavenia)
25:         End With
26:
27:         MsgBox("Vasa licencia bola predlzena.", MsgBoxStyle.Information)
28:         txtKluc.Text = ""
            Module1.WriteLog("   OK")
            frmSpracovavanie.Zaloha = True
29:         Me.Close()
30:     Else
31:         pokusy += 1
32:         MsgBox("Nespravny kluc.", MsgBoxStyle.Exclamation)
33:     End If
34:
35:     If pokusy > 9 Then
36:         MsgBox("Pocet moznych kombinacii: 1 099 511 627 776", MsgBoxStyle.Information)
37:         btnZavriet_Click(sender, e)
38:     End If
    End Sub

    Private Sub btnZavriet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnZavriet.Click
        If frmSpracovavanie.Visible Then End
        Me.Close()
    End Sub

    Private Sub txtKluc_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtKluc.KeyPress
        On Error Resume Next
        Dim KeyAscii As Short = Asc(e.KeyChar)
        ' Zisti, ktora klavesa bola stalcena v ASCII hodnotach
        Dim TrackKey As String
        TrackKey = Chr(KeyAscii)
        'Ak Enter
        If KeyAscii = System.Windows.Forms.Keys.Enter Then
            btnOK_Click(sender, e)
        ElseIf KeyAscii = System.Windows.Forms.Keys.Escape Then
            btnZavriet_Click(sender, e)
        End If
        e.KeyChar = UCase(Chr(KeyAscii))
        If KeyAscii = 0 Then
            e.Handled = True
        End If
    End Sub

    Private Sub btnOK_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles btnOK.KeyPress
        frmPredstavenia.btnPridaj_KeyPress(sender, e, Me)
    End Sub

    Private Sub btnZavriet_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles btnZavriet.KeyPress
        frmPredstavenia.btnPridaj_KeyPress(sender, e, Me)
    End Sub

    Private Sub txtDni_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDni.KeyPress
        frmPredstavenia.txtCena_KeyPress(sender, e, Me)
    End Sub
End Class