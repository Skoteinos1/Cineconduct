Public Class frmVydat

    Dim VizualUprava As Boolean = True
    Public Zaplatit As Double

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Me.Dispose()
    End Sub

    Private Sub frmVydat_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Zafarbi
        Call mdlColors.Skining(Me)
        If VizualUprava Then
            Call mdlColors.Sizing(Me)
            Me.CenterToParent()
            VizualUprava = False
        End If

        If frmSpracovavanie.OpatovnaTlac Then
            btnTlacit.Visible = True
            Me.Height = 203 * frmSpracovavanie.Zvacsenie
        Else
            btnTlacit.Visible = False
            Me.Height = 167 * frmSpracovavanie.Zvacsenie
        End If

        lblCena.Text = Zaplatit

        txtPlatene.Text = "0" & Chr(frmSpracovavanie.DesOdd) & "00"
        System.Windows.Forms.SendKeys.Send("{Home}+{End}")

    End Sub

    Private Sub txtPlatene_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPlatene.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        ' Zisti, ktora klavesa bola stalcena v ASCII hodnotach
        Dim TrackKey As String
        TrackKey = Chr(KeyAscii)
        If KeyAscii = System.Windows.Forms.Keys.Return Then
            KeyAscii = 0
            Me.Close()
        ElseIf KeyAscii = System.Windows.Forms.Keys.Escape Then
            Me.Close()
        ElseIf KeyAscii = 44 Or KeyAscii = 46 Then
            'ak ciarka abo bodka tak to co je dobre
            KeyAscii = frmSpracovavanie.DesOdd
        Else
            If (Not IsNumeric(TrackKey) And Not (KeyAscii = System.Windows.Forms.Keys.Back) And Not (KeyAscii = 46)) Then
                ' Ak klavesa nebola a) cislo b) backspace c) desatinna bodka, akoby nebolo nic stlacene
                KeyAscii = 0
                Beep()
            End If
        End If
        e.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtPlatene_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPlatene.TextChanged
        Dim Peniaze As Double
        On Error Resume Next

        If txtPlatene.Text = "" Then Peniaze = 0 Else Peniaze = CDbl(txtPlatene.Text)

        If Zaplatit > Peniaze Then
            lblVydat.Text = "Doplatit"
        Else
            lblVydat.Text = Format((Peniaze - Zaplatit), "####0.00")
        End If

    End Sub

    Private Sub btnOK_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles btnOK.KeyPress
        frmPredstavenia.btnPridaj_KeyPress(sender, e, Me)
    End Sub

    Private Sub btnTlacit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTlacit.Click
        Module1.WriteLog(Chr(13) & Chr(10) & Now & " - Listok tlaceny este raz")
        Call frmVybPredst.TlacitListok()
    End Sub
End Class