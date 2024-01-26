Public Class frmLogin

    Public UserName As String
    Public Rights As String
    Public LoginSucceeded As Boolean
    Dim HASH As String
    Dim VizualUprava As Boolean = True

    Dim Manager As Boolean = False

    ' TODO: Insert code to perform custom authentication using the provided username and password 
    ' (See http://go.microsoft.com/fwlink/?LinkId=35339).  
    ' The custom principal can then be attached to the current thread's principal as follows: 
    '     My.User.CurrentPrincipal = CustomPrincipal
    ' where CustomPrincipal is the IPrincipal implementation used to perform authentication. 
    ' Subsequently, My.User will return identity information encapsulated in the CustomPrincipal object
    ' such as the username, display name, etc.

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
1:      Dim loginP, hesloP As String
2:      Dim i As Integer
3:      LoginSucceeded = False
4:
        'Ak je povoleny vstup len Programatorovi
        If Rights = "LenProgramator" And UsernameTextBox.Text <> "Programator" Then
            MsgBox("Pristup zamietnuty." & Chr(13) & "Skuste vypnut a zapnut aplikaciu.")
            Exit Sub
        End If

5:      'Skontroluje ci Programator
6:      If UsernameTextBox.Text = "Programator" And PasswordTextBox.Text = "phantomlancer" Then
7:          LoginSucceeded = True
8:          Rights = "A"
        ElseIf UsernameTextBox.Text = "Programator" And PasswordTextBox.Text = "pokus" Then
            Call frmCommand.Pokus()
            Exit Sub
9:      End If
10:     'Aby v Demo rezime sa nemohli prihlasit ako Programator a ak daju uzivatela "" zmeni to na demo
11:     If frmSpracovavanie.DemoRezim And LoginSucceeded = False Then
12:         If Rights <> "A" Then
13:             If UsernameTextBox.Text = "Programator" Then UsernameTextBox.Text = "demo"
14:         Else
15:             Rights = "M"
16:         End If
17:         If UsernameTextBox.Text = "" Then UsernameTextBox.Text = "demo"
18:         LoginSucceeded = True
19:     End If
20:
21:     Module1.WriteLog(Chr(13) & Chr(10) & Now & " - Pokus o prihlasenie; User: " & UsernameTextBox.Text)
22:
23:     'Prihlasenie
24:     'Podla M a P dava prava
25:     With frmSpracovavanie.PristupBindingSource
26:         i = .Find("login", Module1.Koduj(UsernameTextBox.Text))
27:         If i <> -1 Then
28:             .Position = i
29:             loginP = Module1.Dekoduj(.Current("login"))
30:             If Module1.Dekoduj(.Current("heslo")) = "Null" Then hesloP = "" Else hesloP = Module1.Dekoduj(.Current("heslo"))
31:             If UsernameTextBox.Text = loginP And PasswordTextBox.Text = hesloP Then
32:                 LoginSucceeded = True
33:                 Rights = Module1.Dekoduj(.Current("Program"))
34:                 If Rights <> "M" Then Manager = False
35:             End If
36:         End If
37:     End With
38:
39:     If LoginSucceeded = False Then
40:         'opdoved na platnost hesla
41:         MsgBox("Nesprávne Meno alebo Heslo!", MsgBoxStyle.Critical)
42:         PasswordTextBox.Focus()
43:         System.Windows.Forms.SendKeys.Send("{Home}+{End}")
44:         Exit Sub
45:     End If
46:
47:     UserName = UsernameTextBox.Text
48:     If UserName <> "Programator" Then
49:         'Zapise log
50:         With frmSpracovavanie.LogTableAdapter
51:             If Manager Then
52:                 .Insert(UserName, "Manager", Format(Today, "dd.MM.yyyy"), Format(TimeOfDay, "HH:mm:ss"), Format(TimeOfDay, "HH:mm:ss"), "")
53:             Else
54:                 .Insert(UserName, "Pokladna", Format(Today, "dd.MM.yyyy"), Format(TimeOfDay, "HH:mm:ss"), Format(TimeOfDay, "HH:mm:ss"), "")
55:             End If
56:             .Fill(frmSpracovavanie.DataSet1.Log)
57:             HASH = Module1.RetazecRiadku("Log", frmSpracovavanie.LogBindingSource.Find("crc", ""))
58:             frmSpracovavanie.LogBindingSource.Current("CRC") = Module1.HASHString(HASH)
59:             frmSpracovavanie.PozLogu = frmSpracovavanie.LogBindingSource.Position
60:             Module1.UpdateLog(frmSpracovavanie.PozLogu)
61:         End With
62:     End If
63:
64:     Call Module1.HASHdbKontrola("Log")
65:     Call Module1.HASHSubor("Data\data.pac")
66:
67:     frmSplash.Visible = False
68:     Me.Visible = False
69:
70:     If Manager = True Then Module1.WriteLog("; Manager   OK") Else Module1.WriteLog("; Pokladna   OK")
71:     If Manager = True Then frmMenu.ShowDialog() Else frmPocStav.ShowDialog()
72:


    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
1:      LoginSucceeded = False
2:      frmSplash.Dispose()
3:      frmSpracovavanie.Dispose()
4:      End
    End Sub

    Private Sub LogoPictureBox_Click(sender As Object, e As EventArgs) Handles LogoPictureBox.Click
        If Manager = False Then
            Manager = True
            Label1.Text = "Manager"
            LogoPictureBox.Image = System.Drawing.Image.FromFile(frmSpracovavanie.Adresa & "obr\Manager.ico")
        Else
            Manager = False
            Label1.Text = "Pokladna"
            LogoPictureBox.Image = System.Drawing.Image.FromFile(frmSpracovavanie.Adresa & "obr\Pokladna3.ico")
        End If
    End Sub
    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
        LogoPictureBox_Click(sender, e)
    End Sub

    Private Sub UsernameTextBox_KeyPress(sender As Object, e As KeyPressEventArgs) Handles UsernameTextBox.KeyPress
        ' If e.KeyChar = Chr(13) Then
        'e.KeyChar = Chr(0)
        'System.Windows.Forms.SendKeys.Send("{tab}")
        'End If
        frmDistribut.txtDistributor_KeyPress(sender, e, Me)
    End Sub
    Private Sub PasswordTextBox_KeyPress(sender As Object, e As KeyPressEventArgs) Handles PasswordTextBox.KeyPress
        frmDistribut.txtDistributor_KeyPress(sender, e, Me)
    End Sub
    Private Sub OK_KeyPress(sender As Object, e As KeyPressEventArgs) Handles OK.KeyPress
        frmPredstavenia.btnPridaj_KeyPress(sender, e, Me)
    End Sub
    Private Sub Cancel_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Cancel.KeyPress
        frmPredstavenia.btnPridaj_KeyPress(sender, e, Me)
    End Sub
End Class
