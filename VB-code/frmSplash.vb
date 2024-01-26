Public Class frmSplash

    Dim odpocet As Integer

    Private Sub frmSplash_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        lblVerzia.Text = "Verzia " & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & My.Application.Info.Version.Build
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text = "qw" Then
            Label5.Text = "Upozornenie: Táto aplikácia je chránená autorskými právami."
            frmLogin.Rights = "LenProgramator"
        End If
    End Sub

    Private Sub Timer1_Tick_1(sender As Object, e As EventArgs) Handles Timer1.Tick
        odpocet = odpocet + 1
        If odpocet >= 3 Then
1:          Timer1.Enabled = False
2:          'Zakaze opatovne spustenie
3:          If UBound(Diagnostics.Process.GetProcessesByName(Diagnostics.Process.GetCurrentProcess.ProcessName)) > 0 Then
4:              'Cineconduct.vshost a Cineconduct
                MsgBox("Aplikácia uz je spustená...")
5:              End
6:          Else
                'MsgBox(Diagnostics.Process.GetCurrentProcess.ProcessName)
7:              frmSpracovavanie.ShowDialog()

            End If
        End If
    End Sub

    Private Sub Label4_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label4.MouseMove
        TextBox1.Enabled = True
        TextBox1.Focus()
        Label5.Text = "Upozornenie: Táto aplikácia je chránená autorskými právami"
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub
End Class