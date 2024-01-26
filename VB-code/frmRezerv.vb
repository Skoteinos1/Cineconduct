Public Class frmRezerv

    Public MiestoOdRez, MiestoDoRez As Integer
    Public MenaMiest As String
    Public VizualUprava As Boolean = True
    Public Datum As String
    Public Cas As String

    Private Sub btnRezervovat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRezervovat.Click
1:      On Error GoTo Chyba
2:      If btnOK.BackColor = Color.Red Then
3:          If txtMeno.Text.Length = 0 Then
4:              MsgBox("Vyplnte meno rezervacie.", MsgBoxStyle.Exclamation)
5:          ElseIf txtMeno.Text.Length = 1 Or txtMeno.Text.Length = 2 Then
6:              MsgBox("Meno rezervacie musi obsahovat aspon 3 znaky.", MsgBoxStyle.Exclamation)
7:          Else
8:              MsgBox("Na taketo meno su uz rezervovane listky.", MsgBoxStyle.Exclamation)
9:          End If
10:         Exit Sub
11:     End If
12:     Module1.WriteLog(Chr(13) & Chr(10) & Now & " - Vytvorenie rezervacie: " & Format(CDate(lblDatumRez.Text), "dd.MM.yyyy") & " " & Format(CDate(lblCasRez.Text), "HH:mm") & "; " & txtMeno.Text & "; " & txtUdaje.Text & "; " & MenaMiest)
13:     Dim X As Short
14:     Dim s, s1, s2 As String
15:     Dim HASH As String
16:     Dim DatumPlatnosti As Date
17:
18:     'Kontrola ci zapisuje do spravneho riadku
19:     With frmSpracovavanie.PredstaveniaBindingSource
20:         If Datum <> Format(CDate(.Current("Datum")), "dd.MM.yyyy") Then
21:             MsgBox("Vyskytla sa chyba pri ukladani rezervacie. Otvorte okno predaja este raz.", MsgBoxStyle.Critical)
22:             Me.Dispose()
23:             Exit Sub
24:         End If
25:         If Cas <> Format(CDate(.Current("Datum")), "HH:mm") Then
26:             MsgBox("Vyskytla sa chyba pri ukladani rezervacie. Otvorte okno predaja este raz.", MsgBoxStyle.Critical)
27:             Me.Dispose()
28:             Exit Sub
29:         End If
30:         s = .Current("miesta")
31:     End With
32:
33:     'Da 4ky na predane miesta
34:     s1 = LSet(s, MiestoOdRez - 1)
35:     s2 = Mid(s, MiestoDoRez + 1)
36:     s = s1 & StrDup(MiestoDoRez - MiestoOdRez + 1, "4") & s2
37:
38:     If cbxRezDni.Checked = True Then
39:         DatumPlatnosti = DateTime.FromOADate(Today.ToOADate + frmSpracovavanie.RezervPlatn)
40:         txtUdaje.Text = "/" & Format(DatumPlatnosti, "dd.MM.yyyy") & " " & txtUdaje.Text
41:     End If
42:
43:     With frmSpracovavanie.PredstaveniaBindingSource
44:         .Current("miesta") = s
45:         HASH = Module1.RetazecRiadku("Predstavenia", frmVybPredst.PozPredstav)
46:         .Current("crc") = Module1.HASHString(HASH)
47:         Module1.UpdatePredstavenia(frmVybPredst.PozPredstav)
48:         .Position = frmVybPredst.PozPredstav
49:         If s <> .Current("miesta") Then
50:             MsgBox("Chyba pri ukladani obsadenych miest. Obratte sa na Programatora.", MsgBoxStyle.Critical)
51:             Me.Dispose()
52:             Exit Sub
53:         End If
54:     End With
55:
56:     With frmSpracovavanie.RezervacieTableAdapter
57:         .Insert(Format(CDate(lblDatumRez.Text), "dd.MM.yyyy"), Format(CDate(lblCasRez.Text), "HH:mm"), lblFilm.Text, MiestoOdRez, MiestoDoRez, txtMeno.Text, txtUdaje.Text, MenaMiest)
58:         .Fill(frmSpracovavanie.DataSet1.Rezervacie)
59:     End With
60:
61:     'Call Module1.HASHdbKontrola()
62:     Call Module1.HASHSubor("Data\data.pac")
63:     Module1.WriteLog("   OK")
64:     frmSpracovavanie.Zaloha = True
65:     Me.Dispose()
66:
67:     Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub frmRezerv_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
1:      'Zafarbi
2:      Call mdlColors.Skining(Me)
3:      If VizualUprava Then
4:          Call mdlColors.Sizing(Me)
5:          Me.CenterToParent()
6:          VizualUprava = False
7:      End If
8:
9:      With frmSpracovavanie.PredstaveniaBindingSource
10:         .Position = frmVybPredst.PozPredstav
11:         lblDatumRez.Text = Format(CDate(.Current("Datum")), "dd.MM.yyyy")
12:         lblCasRez.Text = Format(CDate(.Current("Datum")), "HH:mm")
13:         lblFilm.Text = .Current("NazovFilmu")
14:     End With
15:
16:     lblOsobRez.Text = MiestoDoRez - MiestoOdRez + 1
17:     txtUdaje.Text = ""
18:     txtMeno.Text = ""
19:     btnOK.BackColor = Color.Red
20:     cbxRezDni.Checked = False
21:     cbxRezDni.Text = "Platnost rezervacie je " & frmSpracovavanie.RezervPlatn & " dni"

    End Sub

    Private Sub txtMeno_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMeno.TextChanged
1:      If txtMeno.Text.Length < 3 Then btnOK.BackColor = Color.Red : Exit Sub
2:
3:      If frmSpracovavanie.RezervacieBindingSource.Find("Meno", txtMeno.Text) = -1 Then
4:          btnOK.BackColor = Color.Green
5:      Else
6:          btnOK.BackColor = Color.Red
7:      End If
    End Sub
    Private Sub txtMeno_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMeno.KeyPress
        frmDistribut.txtDistributor_KeyPress(sender, e, Me)
    End Sub
    Private Sub txtUdaje_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtUdaje.KeyPress
        frmDistribut.txtDistributor_KeyPress(sender, e, Me)
    End Sub
    Private Sub cbxRezDni_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cbxRezDni.KeyPress
        frmPredstavenia.btnPridaj_KeyPress(sender, e, Me)
    End Sub
    Private Sub btnRezervovat_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles btnRezervovat.KeyPress
        frmPredstavenia.btnPridaj_KeyPress(sender, e, Me)
    End Sub
End Class