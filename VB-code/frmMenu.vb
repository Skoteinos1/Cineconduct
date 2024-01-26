Imports System.IO
Imports System.Drawing.Printing

Public Class frmMenu
    Inherits System.Windows.Forms.Form

    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Integer, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Integer) As Integer

    Private NasielSubor As Boolean
    Public PozDistrib As Integer
    Public Vytlacit As String
    Public Velkost As Integer

    Private Sub frmMenu_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        frmZavierka.frmZavierka_Disposed(sender, e)
    End Sub

    Private Sub frmMenu_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
1:      On Error GoTo Chyba
2:      Module1.WriteLog(Chr(13) & Chr(10) & Now & " - Odhlasenie uzivatela z Managera: " & frmLogin.UserName)
3:      Dim Cancel As Boolean = e.Cancel
4:      Dim UnloadMode As System.Windows.Forms.CloseReason = e.CloseReason
5:      Call Odhlasit()
6:      'Call Module1.HASHdbKontrola()
7:      Module1.WriteLog("   OK")
8:      e.Cancel = Cancel
9:      Me.Dispose()
10:
        Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub frmMenu_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        If frmLogin.UserName = "Programator" Then
            Dim KeyAscii As Short = Asc(e.KeyChar)
            If KeyAscii = 126 Then
                frmCommand.ShowDialog()
            End If
        End If
    End Sub

    Private Sub frmMenu_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        On Error Resume Next

        'Kontroluje log
        Call Chyby.OpravaLogu()

1:      'frmSplash.Close()
2:      'frmLogin.Close()
3:      frmSpracovavanie.Visible = False
4:      lblVerzia.Text = "V " & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & My.Application.Info.Version.Build
5:      If frmSpracovavanie.SietVerzia Then lblCopyright.Text = "Nezadavajte filmy a predstavenia ak na inych pocitacoch prebieh predaj.       " & My.Application.Info.Copyright Else lblCopyright.Text = My.Application.Info.Copyright
6:
7:      'Zafarbi
8:      Call mdlColors.Skining(Me)
9:      Call mdlColors.Sizing(Me)
10:     Me.CenterToScreen()

        Dim cesta As String = "\obr\"
        If Today >= CDate("1.1." & Year(Today)) And Today <= CDate("7.1." & Year(Today)) Then
            cesta += "NewYear.gif"
        ElseIf Today >= CDate("13.2." & Year(Today)) And Today <= CDate("15.2." & Year(Today)) Then
            cesta += "valentine.gif"
        ElseIf Today >= CDate("15.4.2019") And Today <= CDate("26.4.2019") Then
            cesta += "easter.gif"
        ElseIf Today >= CDate("8.4.2020") And Today <= CDate("19.4.2020") Then 'od pondelka po dalsi piatok
            cesta += "easter.gif"
        ElseIf Today >= CDate("28.10." & Year(Today)) And Today <= CDate("3.11." & Year(Today)) Then
            cesta += "pumpkin.gif"
        ElseIf Today >= CDate("5.12." & Year(Today)) And Today <= CDate("7.12." & Year(Today)) Then
            cesta += "santa2.gif"
        ElseIf Today >= CDate("8.12." & Year(Today)) And Today <= CDate("31.12." & Year(Today)) Then
            cesta += "ChristmasTree.gif"
        End If
        If InStr(cesta, ".gif") <> 0 Then PictureBox2.Image = System.Drawing.Image.FromFile(frmSpracovavanie.Adresa & cesta)
    End Sub

    Public Sub Odhlasit()
1:      On Error GoTo Chyba
2:      'Logout - Zapise cas odhlasenia
3:      If frmLogin.UserName = "Programator" Then Exit Sub
4:      Dim HASH As String
5:
6:      With frmSpracovavanie.LogBindingSource
7:          .Position = frmSpracovavanie.PozLogu
8:          .Current("odhlasenie") = CStr(Format(TimeOfDay, "HH:mm:ss"))
9:          HASH = RetazecRiadku("Log", .Position)
10:         .Current("crc") = Module1.HASHString(HASH)
11:         Module1.UpdateLog(frmSpracovavanie.PozLogu)
12:     End With
13:
        Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, "Odhlasit", ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub ZálohaDátToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ZálohaDátToolStripMenuItem.Click
1:      On Error GoTo Chyba
2:      Module1.WriteLog(Chr(13) & Chr(10) & Now & " - Zaloha databazy")
3:      FileCopy(frmSpracovavanie.Adresa & "Data\data.pac", frmSpracovavanie.Adresa2 & "Zalohy\" & Today & ".pac")
4:      If frmSpracovavanie.Adresa2 <> frmSpracovavanie.Adresa Then FileCopy(frmSpracovavanie.Adresa & "Data\data.pac", frmSpracovavanie.Adresa & "Zalohy\" & Today & ".pac")
5:      Call Module1.HASHSubor("Zalohy\data2.pac")
6:      Call Module1.HASHSubor("Zalohy\" & Today & ".pac")
7:      Module1.WriteLog("   OK")
8:      MsgBox("Zalohované")
9:
        Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub ČistenieDátToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ČistenieDátToolStripMenuItem.Click
1:      On Error GoTo Chyba
2:      Dim Subor As String
3:      Dim Datum As Date
4:
5:      Subor = frmSpracovavanie.Adresa2 & "\Zalohy\" & Today & ".pac"
6:      If File.Exists(Subor) = False Then
7:          MsgBox("Nenašiel sa súbor: " & Today & ".pac" & Chr(13) & "Zalohujte databazu!", vbCritical)
8:          Exit Sub
9:      End If
10:     Module1.WriteLog(Chr(13) & Chr(10) & Now & " - Cistenie databazy od starych dat")
11:     Datum = CDate("01.12." & Year(Today) - 1)
12:     lblUpozornenie.Text = "Cistenie moze trvat aj niekolko minut. Neprerusujte cinnost programu ani keby nebolo viditet ziadnu zmenu."
13:     If MsgBox("Naozaj vymazat stare udaje?" & Chr(13) & "Zmazu sa udaje vytvorene pred " & Datum, MsgBoxStyle.YesNo + MsgBoxStyle.Question) = MsgBoxResult.No Then
14:         lblUpozornenie.Text = ""
15:         Exit Sub
16:     End If
17:
18:     'Ak bola robena, prejde nasledujuce tabulky a zmaze "stare" veci (do 30.11. minuleho roku)
19:
20:     Dim x, PocetMaz As Integer
21:     ProgressBar1.Visible = True
22:     ProgressBar1.Maximum = frmSpracovavanie.FilmyBindingSource.Count + frmSpracovavanie.PredstaveniaBindingSource.Count + frmSpracovavanie.LogBindingSource.Count + frmSpracovavanie.ListkyBindingSource.Count + frmSpracovavanie.NakladFilmBindingSource.Count + frmSpracovavanie.StornoPredstaveniBindingSource.Count + 10
23:     ProgressBar1.Value = 0
24:
25:     PocetMaz = 0
26:
27:     With frmSpracovavanie.FilmyBindingSource
28:         For x = 1 To .Count
29:             .Position = x
30:             If CDate(.Current("PoslednePredstavenie")) < Datum Then
31:                 frmSpracovavanie.FilmyTableAdapter.Delete(.Current(0), .Current(1), .Current(2), .Current(3), .Current(4), .Current(5), .Current(6), .Current(7))
32:                 PocetMaz += 1
33:             End If
34:             ProgressBar1.Value += 1
35:         Next x
36:     End With
37:     frmSpracovavanie.FilmyTableAdapter.Fill(frmSpracovavanie.DataSet1.Filmy)
38:
39:     With frmSpracovavanie.PredstaveniaBindingSource
40:         For x = 1 To .Count
41:             .Position = x
42:             If CDate(.Current("Datum")) < Datum Then
43:                 frmSpracovavanie.PredstaveniaTableAdapter.Delete(.Current(0), .Current(1), .Current(2), .Current(3), .Current(4), .Current(5), .Current(7), .Current(8))
44:                 PocetMaz += 1
45:             End If
46:             ProgressBar1.Value += 1
47:         Next x
48:     End With
49:     frmSpracovavanie.PredstaveniaTableAdapter.Fill(frmSpracovavanie.DataSet1.Predstavenia)
50:
51:     With frmSpracovavanie.LogBindingSource
52:         For x = 1 To .Count
53:             .Position = x
54:             If CDate(.Current("Den")) < Datum Then
55:                 frmSpracovavanie.LogTableAdapter.Delete(.Current(0), .Current(1), .Current(2), .Current(3), .Current(4), .Current(5))
56:                 PocetMaz += 1
57:             End If
58:             ProgressBar1.Value += 1
59:         Next x
60:     End With
61:     frmSpracovavanie.LogTableAdapter.Fill(frmSpracovavanie.DataSet1.Log)
62:
63:     With frmSpracovavanie.ListkyBindingSource
64:         For x = 1 To .Count
65:             .Position = x
66:             If CDate(.Current("DatumPredst")) < Datum Then
67:                 frmSpracovavanie.ListkyTableAdapter.Delete(.Current(0), .Current(1), .Current(2), .Current(3), .Current(4), .Current(5), .Current(6) & "", .Current(7), .Current(8), .Current(9), .Current(10), .Current(11))
68:                 PocetMaz += 1
69:             End If
70:             ProgressBar1.Value += 1
71:         Next x
72:     End With
73:     frmSpracovavanie.ListkyTableAdapter.Fill(frmSpracovavanie.DataSet1.Listky)
74:
75:     With frmSpracovavanie.NakladFilmBindingSource
76:         For x = 1 To .Count
77:             .Position = x
78:             If frmSpracovavanie.FilmyBindingSource.Find("Film", .Current("NazovFilmu")) = -1 Then
79:                 frmSpracovavanie.NakladFilmTableAdapter.Delete(.Current(0), .Current(1), .Current(2), .Current(3), .Current(4))
80:                 PocetMaz += 1
81:             End If
82:             ProgressBar1.Value += 1
83:         Next x
84:     End With
85:     frmSpracovavanie.NakladFilmTableAdapter.Fill(frmSpracovavanie.DataSet1.NakladFilm)
86:
87:     With frmSpracovavanie.StornoPredstaveniBindingSource
88:         For x = 1 To .Count
89:             .Position = x
90:             If CDate(.Current("Datum")) < Datum Then
91:                 frmSpracovavanie.StornoPredstaveniTableAdapter.Delete(.Current(0), .Current(1), .Current(2), .Current(3), .Current(4), .Current(5), .Current(6))
92:                 PocetMaz += 1
93:             End If
94:             ProgressBar1.Value += 1
95:         Next x
96:     End With
97:     frmSpracovavanie.StornoPredstaveniTableAdapter.Fill(frmSpracovavanie.DataSet1.StornoPredstaveni)
98:
99:     ProgressBar1.Visible = False
100:    lblUpozornenie.Text = ""
101:
102:    Call Module1.HASHdbKontrola("all")
103:    Call Module1.HASHSubor("Data\data.pac")
104:
105:    If PocetMaz = 0 Then
106:        MsgBox("Bolo vymazanych " & PocetMaz & " udajov." & Chr(13) & "Cistenie dokoncene.")
107:    Else
108:        MsgBox("Bolo vymazanych " & PocetMaz & " udajov." & Chr(13) & "Pre istotu vykonajte cistenie este raz.")
109:    End If
        Module1.WriteLog("   OK")
        Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)

    End Sub

    Public Sub Tlacit()
        On Error GoTo NeniTlaciaren
        If frmSpracovavanie.DemoRezim Then MsgBox("Tato funkcia nie je v Demo rezime podporovana.", MsgBoxStyle.Information) : Exit Sub
        If frmSpracovavanie.VyberTlaciarneMan = "Tlac do textoveho suboru" Then
            FileOpen(1, frmSpracovavanie.Adresa2 & "tlac.txt", OpenMode.Append)
            WriteLine(1, Vytlacit)
            WriteLine(1)
            WriteLine(1)
            FileClose(1)
        Else
            FileOpen(1, frmSpracovavanie.Adresa2 & "tlac.txt", OpenMode.Output)
            WriteLine(1, Vytlacit)
            WriteLine(1)
            WriteLine(1)
            FileClose(1)
            If frmSpracovavanie.PrntQstMan Then
                'showDialog method makes the dialog box visible at run time
                If PrintDialog1.ShowDialog = DialogResult.OK Then PrintDocument1.Print()
            Else
                PrintDocument1.PrinterSettings.PrinterName = frmSpracovavanie.VyberTlaciarneMan ' TReba na vybratie tlaciarne ak sa nevyberie rucne
                PrintDocument1.Print()
            End If
        End If
        Exit Sub
NeniTlaciaren:
        MsgBox("Vyber tlaciaren pre Manager.", MsgBoxStyle.Exclamation)
    End Sub

    Private Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
1:      On Error GoTo Chyba
2:      Static intCurrentChar As Int32
3:      Module1.WriteLog(Chr(13) & Chr(10) & Now & " - Tlac dokumentu v Manageri")
4:      Dim font As New Font("Lucida Console", Velkost)
5:
6:      Dim intPrintAreaHeight, intPrintAreaWidth, marginLeft, marginTop As Int32
7:      With PrintDocument1.DefaultPageSettings
8:          .Margins.Top = 50
9:          .Margins.Bottom = 50
10:         .Margins.Left = 40
11:         .Margins.Right = 40
12:
13:         intPrintAreaHeight = .PaperSize.Height - .Margins.Top - .Margins.Bottom
14:         intPrintAreaWidth = .PaperSize.Width - .Margins.Left - .Margins.Right
15:
16:         marginLeft = .Margins.Left ' X coordinate
17:         marginTop = .Margins.Top ' Y coordinate
18:     End With
19:
20:     If PrintDocument1.DefaultPageSettings.Landscape Then
21:         Dim intTemp As Int32
22:         intTemp = intPrintAreaHeight
23:         intPrintAreaHeight = intPrintAreaWidth
24:         intPrintAreaWidth = intTemp
25:     End If
26:
27:     Dim intLineCount As Int32 = CInt(intPrintAreaHeight / font.Height)
28:     Dim rectPrintingArea As New RectangleF(marginLeft, marginTop, intPrintAreaWidth, intPrintAreaHeight)
29:
30:     Dim fmt As New StringFormat(StringFormatFlags.LineLimit)
31:
32:     Dim intLinesFilled, intCharsFitted As Int32
33:     e.Graphics.MeasureString(Mid(Vytlacit, intCurrentChar + 1), font, New SizeF(intPrintAreaWidth, intPrintAreaHeight), fmt, intCharsFitted, intLinesFilled)
34:
35:     e.Graphics.DrawString(Mid(Vytlacit, intCurrentChar + 1), font, Brushes.Black, rectPrintingArea, fmt)
36:
37:     intCurrentChar += intCharsFitted
38:
39:     If intCurrentChar < Vytlacit.Length Then
40:         e.HasMorePages = True
41:     Else
42:         e.HasMorePages = False
43:         intCurrentChar = 0
44:     End If
45:     Module1.WriteLog("   OK")
        Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub PokladnaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PokladnaToolStripMenuItem.Click
        If frmSpracovavanie.PrepnutiePM = True Then MsgBox("Raz uz ste aplikaciu prepli. Pre dalsie prepnutie je potrebne ju spustit znova.", MsgBoxStyle.Information) : Exit Sub

        If MsgBox("Prepnut do Pokladne?", MsgBoxStyle.YesNo + MsgBoxStyle.Question) = MsgBoxResult.No Then Exit Sub

1:      On Error GoTo Chyba
2:      Module1.WriteLog(Chr(13) & Chr(10) & Now & " - Prepnutie do Pokladne ( User: " & frmLogin.UserName & ")")
3:      Call Odhlasit()
4:
5:      'Log prihlasenia do pokladne
6:      If frmLogin.UserName <> "Programator" Then
7:          Dim HASH As String
8:          'Zapise log
9:          With frmSpracovavanie.LogTableAdapter
10:             .Insert(frmLogin.UserName, "Pokladna", Format(Today, "dd.MM.yyyy"), Format(TimeOfDay, "HH:mm:ss"), Format(TimeOfDay, "HH:mm:ss"), "")
11:             .Fill(frmSpracovavanie.DataSet1.Log)
12:             HASH = Module1.RetazecRiadku("Log", frmSpracovavanie.LogBindingSource.Find("crc", ""))
13:             frmSpracovavanie.LogBindingSource.Current("CRC") = Module1.HASHString(HASH)
14:             frmSpracovavanie.PozLogu = frmSpracovavanie.LogBindingSource.Position
15:             frmSpracovavanie.LogBindingSource.EndEdit()
16:             .Update(frmSpracovavanie.DataSet1.Log)
17:             'Refresh pre tabulku
18:             .Fill(frmSpracovavanie.DataSet1.Log)
19:         End With
20:     End If
21:
22:     Call Module1.HASHdbKontrola("Log")
23:
24:     Me.Visible = False
        'frmPocStav.Dispose()

        frmSpracovavanie.PrepnutiePM = True
        Module1.WriteLog("   OK")
25:     frmPocStav.ShowDialog()
26:
        Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub
    Private Sub btnKontrola_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnKontrola.Click
        frmSpracovavanie.Timer2_Tick()
    End Sub
    Private Sub PrehľadPredstaveníToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrehľadPredstaveníToolStripMenuItem.Click
        'frmPredstPrehl.ShowDialog()
    End Sub
    Private Sub ZaDistribútoraBezDPHToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ZaDistribútoraBezDPHToolStripMenuItem.Click
        frmUctDBD.ShowDialog()
    End Sub
    Private Sub ZaDistribútoraDPHToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ZaDistribútoraDPHToolStripMenuItem.Click
        frmUctDsD.ShowDialog()
    End Sub
    Private Sub VyťaženosťSályToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles VyťaženosťSályToolStripMenuItem.Click
        ' frmStatVytazSal.ShowDialog()
    End Sub
    Private Sub PrístupyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrístupyToolStripMenuItem.Click
        ' frmStatLogy.ShowDialog()
    End Sub
    Private Sub PoradieFilmovToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PoradieFilmovToolStripMenuItem.Click
        ' frmStatPoradF.ShowDialog()
    End Sub
    Private Sub ZaKinoBezDPHAVFToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ZaKinoBezDPHAVFToolStripMenuItem1.Click
        frmUctKBDavf.ShowDialog()
    End Sub
    Private Sub PrehladTrzbyToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrehladTrzbyToolStripMenuItem.Click
        'frmStatPrehladTrz.ShowDialog()
    End Sub
    Private Sub JednoduchýPrehľadToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles JednoduchýPrehľadToolStripMenuItem.Click
        frmJednPrehlad.ShowDialog()
    End Sub
    Private Sub PredpisTrzbyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PredpisTrzbyToolStripMenuItem.Click
        ' frmPredpisTrzby.ShowDialog()
    End Sub
    Private Sub VýkazVýsledkovPoDňochToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles VýkazVýsledkovPoDňochToolStripMenuItem.Click
        frmUctVykVysl.ShowDialog()
    End Sub
    Private Sub PredstaveniaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PredstaveniaToolStripMenuItem.Click
        frmPredstavenia.ShowDialog()
    End Sub
    Private Sub RočnýVýsledokKinaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RočnýVýsledokKinaToolStripMenuItem.Click
        frmStatRocVysl.ShowDialog()
    End Sub
    Private Sub FilmyToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles FilmyToolStripMenuItem.Click
        frmFilmy.ShowDialog()
    End Sub
    Private Sub DistribútoriToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DistribútoriToolStripMenuItem1.Click
        frmDistribut.ShowDialog()
    End Sub
    Private Sub ZaKinoBezDPHToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ZaKinoBezDPHToolStripMenuItem.Click
        frmUctKBD.ShowDialog()
    End Sub
    Private Sub ZaKinoDPHToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ZaKinoDPHToolStripMenuItem.Click
        frmUctKsD.ShowDialog()
    End Sub
    Private Sub NastaveniaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NastaveniaToolStripMenuItem.Click
        frmOptions.ShowDialog()
    End Sub
    Private Sub ManuálToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ManuálToolStripMenuItem.Click
        Call frmVybPredst.ManuálToolStripMenuItem_Click(sender, e)
    End Sub
    Private Sub OProgrameManagerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OProgrameManagerToolStripMenuItem.Click
        AboutBox1.ShowDialog()
    End Sub
    Private Sub LicenciaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LicenciaToolStripMenuItem.Click
        frmLicencia.ShowDialog()
    End Sub
    Private Sub NavštíviťStránkuToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NavštíviťStránkuToolStripMenuItem.Click
        Process.Start("http://nagypeter.webnode.sk/")
    End Sub
    Private Sub LogSúborToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LogSúborToolStripMenuItem.Click
        frmLogy.ShowDialog()
    End Sub
    Private Sub UkončiťManagerToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UkončiťManagerToolStripMenuItem.Click
        If MsgBox("Naozaj ukoncit aplikáciu?", MsgBoxStyle.YesNo + MsgBoxStyle.Question) = MsgBoxResult.No Then Exit Sub
        frmMenu_FormClosing(Me, New System.Windows.Forms.FormClosingEventArgs(System.Windows.Forms.CloseReason.UserClosing, 0))
    End Sub
    Private Sub VyúčtovaniePokladníkovToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles VyúčtovaniePokladníkovToolStripMenuItem.Click
        frmStatPokladn.ShowDialog()
    End Sub
    Private Sub KontrolneSuctyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles KontrolneSuctyToolStripMenuItem.Click
        frmStatSucty.ShowDialog()
    End Sub
End Class
