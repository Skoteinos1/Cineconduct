Public Class frmStorno

    Dim VizualUprava As Boolean = True
    Dim PolohaListka As Integer

    Private Sub frmStorno_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        gbxStorno.Visible = False
        btnStorno.Enabled = False
        btnOK.Enabled = False

        'Zafarbi
        Call mdlColors.Skining(Me)
        If VizualUprava Then
            Call mdlColors.Sizing(Me)
            Me.CenterToParent()
            VizualUprava = False
        End If
    End Sub

    Private Sub txtCisListk_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCisListk.KeyPress
        frmPredstavenia.txtCena_KeyPress(sender, e, Me)
    End Sub

    Private Sub txtCisListk_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCisListk.TextChanged
        btnStorno.Enabled = False
        gbxStorno.Visible = False

        If txtCisListk.Text = "" Then
            btnOK.Enabled = False
        Else
            btnOK.Enabled = True
        End If
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        If frmSpracovavanie.SietVerzia Then frmSpracovavanie.ListkyTableAdapter.Fill(frmSpracovavanie.DataSet1.Listky)
        On Error GoTo Chyba
        Dim Pokladnik, Hotov As String
        Dim x As Integer

        With frmSpracovavanie.ListkyBindingSource
            x = .Find("Kod", txtCisListk.Text)
            If x <> -1 Then
                .Filter = "Kod = '" & txtCisListk.Text & "'"
                x = .Count
                If x = 0 Then
                    gbxStorno.Visible = False
                    btnStorno.Enabled = False
                    MsgBox("Lístok sa nenachádza v pamäti." & Chr(13) & "Nie je mozné Stornovat.", MsgBoxStyle.Exclamation)
                ElseIf x = 1 Then
                    .MoveFirst()
                    If InStr(.Current("Stav"), "Predane") <> 0 Or InStr(.Current("Stav"), "Bezhotov") <> 0 Then
                        ZobrazListok()
                        gbxStorno.Visible = True
                        btnStorno.Enabled = True
                    ElseIf InStr(.Current("Stav"), "Storn") <> 0 Then 'Or InStr(.Current("Stav"), "BezhotStorn") <> 0 Then
                        gbxStorno.Visible = False
                        btnStorno.Enabled = False
                        Module1.WriteLog(Chr(13) & Chr(10) & Now & " - Chyba v databaze listkov - " & txtCisListk.Text)
                        MsgBox("Chyba v databaze listkov. Na adresu pnagy11@gmail.com zaslite subor log.txt", MsgBoxStyle.Critical)
                    End If
                ElseIf x = 2 Then
                    .MoveFirst()
                    If InStr(.Current("Stav"), "Predane") <> 0 Or InStr(.Current("Stav"), "Bezhotov") <> 0 Then .MoveNext()
                    If InStr(.Current("Stav"), "Storn") <> 0 Then 'Or InStr(.Current("Stav"), "BezhotStorn") <> 0 Then                     
                        ZobrazListok()
                        gbxStorno.Visible = True
                        btnStorno.Enabled = False
                    End If
                Else
                    gbxStorno.Visible = False
                    btnStorno.Enabled = False
                    Module1.WriteLog(Chr(13) & Chr(10) & Now & " - Chyba v databaze listkov - " & txtCisListk.Text)
                    MsgBox("Chyba v databaze listkov. Na adresu pnagy11@gmail.com zaslite subor log.txt", MsgBoxStyle.Critical)
                End If
                .RemoveFilter()
            Else
                gbxStorno.Visible = False
                btnStorno.Enabled = False
                MsgBox("Lístok sa nenachádza v pamäti." & Chr(13) & "Nie je mozné Stornovat.", MsgBoxStyle.Exclamation)
            End If
        End With

        Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub btnStorno_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStorno.Click
1:      If frmSpracovavanie.SietVerzia Then
2:          frmSpracovavanie.PredstaveniaTableAdapter.Fill(frmSpracovavanie.DataSet1.Predstavenia)
3:          frmSpracovavanie.FilmyTableAdapter.Fill(frmSpracovavanie.DataSet1.Filmy)
4:          frmSpracovavanie.ListkyTableAdapter.Fill(frmSpracovavanie.DataSet1.Listky)
5:      End If

        frmSpracovavanie.Timer2_Tick()
        If frmSpracovavanie.NasloDatabazu = False Then
            MsgBox("Databaza sa nenasla.", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        If MsgBox("Naozaj Stornovat lístok?", MsgBoxStyle.YesNo + MsgBoxStyle.Question) = MsgBoxResult.No Then
            txtCisListk.Text = ""
            Exit Sub
        End If

6:      On Error GoTo Chyba
7:      Module1.WriteLog(Chr(13) & Chr(10) & Now & " - Stornovanie listka: " & txtCisListk.Text)

8:      'Vyhlada spravne zaznamy
        frmVybPredst.PozPredstav = frmSpracovavanie.PredstaveniaBindingSource.Find("Datum", lblDatumPredst.Text & " " & lblCasPredst.Text)
        If frmVybPredst.PozPredstav = -1 Then
            MsgBox("Nenaslo sa predstavenie!", MsgBoxStyle.Critical) : Exit Sub
        Else
            frmSpracovavanie.PredstaveniaBindingSource.Position = frmVybPredst.PozPredstav
        End If
11:     frmVybPredst.PozFilm = frmSpracovavanie.FilmyBindingSource.Find("Film", frmSpracovavanie.PredstaveniaBindingSource.Current("NazovFilmu"))
12:     If frmVybPredst.PozFilm = -1 Then MsgBox("Nenasiel sa film!", MsgBoxStyle.Critical) : Exit Sub

        frmSpracovavanie.ListkyBindingSource.Position = frmSpracovavanie.ListkyBindingSource.Find("Kod", txtCisListk.Text)
22:
23:     Dim p2, p3 As Short
24:     Dim s, s1, s2, HASH As String
25:     Dim vratit As Double
26:
27:     vratit = CDbl(lblSuma.Text)

53:     'Vytvori rovnaky stornovany listok len den predaja bude dnom stornovania
54:     With frmSpracovavanie.ListkyBindingSource
55:
56:         If lblHotov.Text = "Hotovost" Then
57:             frmSpracovavanie.ListkyTableAdapter.Insert(.Current("kod"), Format(CDate(.Current("DatumPredst")), "dd.MM.yyyy"), Format(CDate(.Current("CasPredst")), "HH:mm"), .Current("sala"), .Current("miestoOd"), .Current("miestoDo"), .Current("NazovMiest"), .Current("Zlavnenych"), .Current("Suma"), Format(Today, "dd.MM.yyyy"), "Storno;" & frmLogin.UserName, "")
58:         ElseIf lblHotov.Text = "Bezhotovost" Then
59:             frmSpracovavanie.ListkyTableAdapter.Insert(.Current("kod"), Format(CDate(.Current("DatumPredst")), "dd.MM.yyyy"), Format(CDate(.Current("CasPredst")), "HH:mm"), .Current("sala"), .Current("miestoOd"), .Current("miestoDo"), .Current("NazovMiest"), .Current("Zlavnenych"), .Current("Suma"), Format(Today, "dd.MM.yyyy"), "BezhotStorn;" & frmLogin.UserName, "")
60:         End If
61:
62:         frmSpracovavanie.ListkyTableAdapter.Fill(frmSpracovavanie.DataSet1.Listky)
63:         HASH = Module1.RetazecRiadku("Listky", .Find("crc", ""))
64:         .Current("CRC") = Module1.HASHString(HASH)
65:         Module1.UpdateListky(.Position)
66:     End With
67:

        'Prehodene najprv zapise stornovanie listka do databazy a potom uvolni miesta
        'Ak to spravi naopak tak listok sa pri chybe nevystornuje ale miesta sa uvolnia
28:     With frmSpracovavanie.PredstaveniaBindingSource
            .Position = frmVybPredst.PozPredstav
            frmSpracovavanie.ListkyBindingSource.Position = frmSpracovavanie.ListkyBindingSource.Find("Kod", txtCisListk.Text)

29:         'Zmenit 2 na 1, odpocita trzbu a predaj
30:         s = .Current("miesta")
31:         p2 = CInt(frmSpracovavanie.ListkyBindingSource.Current("miestood"))
32:         p3 = CInt(frmSpracovavanie.ListkyBindingSource.Current("miestodo"))
33:         s1 = LSet(s, p2 - 1)
34:         s2 = Mid(s, p3 + 1)
35:         s = s1 & StrDup(p3 - p2 + 1, "1") & s2
36:         .Current("miesta") = s
37:
38:         .Current("Predaj") = CDbl(.Current("Predaj")) - CDbl(lblOsob.Text)
39:         .Current("TrzbaPredstavenia") = CDbl(.Current("TrzbaPredstavenia")) - CDbl(lblSuma.Text)
40:         HASH = Module1.RetazecRiadku("Predstavenia", .Position)
41:         .Current("crc") = Module1.HASHString(HASH)
42:         Module1.UpdatePredstavenia(.Position)
43:     End With
44:
45:     With frmSpracovavanie.FilmyBindingSource
            .Position = frmVybPredst.PozFilm
46:         .Current("Predanelistky") = CDbl(.Current("Predanelistky")) - CDbl(lblOsob.Text)
47:         .Current("TrzbaFilmu") = CDbl(.Current("TrzbaFilmu")) - CDbl(lblSuma.Text)
48:         HASH = Module1.RetazecRiadku("Filmy", .Position)
49:         .Current("crc") = Module1.HASHString(HASH)
50:         Module1.UpdateFilmy(frmVybPredst.PozFilm)
51:     End With



68:     Call Module1.HASHdbKontrola("Listky")
69:     Call Module1.HASHSubor("Data\data.pac")
70:     Module1.WriteLog(";   OK")
71:     frmSpracovavanie.Zaloha = True
72:     If lblHotov.Text = "Hotovost" Then
73:         MsgBox("Lístok stornovaný." & Chr(13) & "Vrátit " & vratit & " EUR v hotovosti")
74:     ElseIf lblHotov.Text = "Bezhotovost" Then
75:         MsgBox("Lístok stornovaný." & Chr(13) & "Vrátit " & vratit & " EUR bezhotovostne")
76:     End If
77:
78:     btnStorno.Enabled = False  ' txtCisListk.Text = ""
79:
80:     Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Sub ZobrazListok()
        Dim Pokladnik, Hotov As String
        Dim DatPredst, DatPredaja As Date
        Dim nElements As Short
        Dim aKey As Object

        With frmSpracovavanie.ListkyBindingSource
            '  .Position = .Find("Kod", txtCisListk.Text)
42:         aKey = Split(.Current("Stav"), ";")
43:         nElements = UBound(aKey) - LBound(aKey) + 1
44:         Hotov = aKey(0)
45:         If nElements = 1 Then Pokladnik = "" Else Pokladnik = aKey(1)
46:
47:         'Zobrazi info v textboxoch
48:         lblKod.Text = .Current("kod")
49:         lblDatumPredst.Text = Format(CDate(.Current("Datumpredst")), "dd.MM.yyyy")
50:         lblCasPredst.Text = Format(CDate(.Current("Caspredst")), "HH:mm")
51:         lblOsob.Text = CInt(.Current("miestodo")) - CInt(.Current("miestood")) + 1
52:         lblZlavnenych.Text = .Current("Zlavnenych")
53:         lblSuma.Text = .Current("Suma")
54:         lblMiesta.Text = .Current("NazovMiest")
55:         lblMeno.Text = Pokladnik

            If Hotov = "Predane" Then : lblHotov.Text = "Hotovost"
            ElseIf Hotov = "Bezhotov" Then : lblHotov.Text = "Bezhotovost"
            Else
                lblHotov.Text = Hotov
            End If

32:         'Zisti ci predstavenie bolo 
33:         DatPredst = Format(CDate(.Current("DatumPredst")), "dd.MM.yyyy")
34:         DatPredaja = Format(CDate(.Current("DenPredaja")), "dd.MM.yyyy")
35:
36:         ' If DatPredaja >= System.DateTime.FromOADate(Today.ToOADate - 8) And DatPredst >= Today Then
37:         'If DatPredst >= Today Or frmCommand.txtPrikazovy.Text = "frm storno" Then
            If DatPredst >= System.DateTime.FromOADate(Today.ToOADate - 8) Or frmCommand.txtPrikazovy.Text = "frm storno" Then
            Else
                btnStorno.Enabled = False
60:             MsgBox("Predstavenie uz bolo." & Chr(13) & "Nie je mozné stornovat.", MsgBoxStyle.Exclamation)
61:         End If

        End With
    End Sub

End Class