Public Class frmPocStav
    Public PocStav As Double

    Private Sub frmPocStav_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
1:      'Zafarbi
2:      Call mdlColors.Skining(Me)
3:      Call mdlColors.Sizing(Me)
        'Kontroluje log
        Call Chyby.OpravaLogu()
4:      Me.CenterToScreen()
5:      System.Windows.Forms.SendKeys.Send("{Home}+{End}")
    End Sub

    Private Sub txtPocStav_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPocStav.KeyPress
1:      Dim KeyAscii As Short = Asc(e.KeyChar)
2:      ' Zisti, ktora klavesa bola stalcena v ASCII hodnotach
3:      Dim TrackKey As String
4:      TrackKey = Chr(KeyAscii)
5:      'Ak Enter
6:      If KeyAscii = System.Windows.Forms.Keys.Enter Then
7:          Button1_Click(sender, e)
8:          Exit Sub
9:      ElseIf KeyAscii = 44 Or KeyAscii = 46 Then
10:         'ak ciarka abo bodka tak to co je dobre
11:         KeyAscii = frmSpracovavanie.DesOdd
12:     ElseIf (Not IsNumeric(TrackKey) And Not (KeyAscii = System.Windows.Forms.Keys.Back) And Not (KeyAscii = 46)) And Not (KeyAscii = 44) Then
13:         ' Ak klavesa nebola a) cislo b) backspace c) desatinna bodka, akoby nebolo nic stlacene
14:         KeyAscii = 0
15:         Beep()
16:     End If
17:     e.KeyChar = Chr(KeyAscii)
18:     If KeyAscii = 0 Then
19:         e.Handled = True
20:     End If
    End Sub

    Private Sub txtPocStav_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPocStav.TextChanged
        ' Prerata na koruny
        Dim cislo As Double
        On Error GoTo Chyba
        If txtPocStav.Text = "" Then cislo = 0 Else cislo = txtPocStav.Text
        cislo = cislo * 30.126
        lblPocStav2.Text = Format(CDbl(cislo), "###0.00")
        Exit Sub
Chyba:
        txtPocStav.Text = ""
        MsgBox("Nesprávny formát císla", MsgBoxStyle.Exclamation, "Chyba")
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If txtPocStav.Text = "" Then txtPocStav.Text = 0
        PocStav = txtPocStav.Text
        Me.Close()
        frmVybPredst.ShowDialog()
    End Sub
End Class