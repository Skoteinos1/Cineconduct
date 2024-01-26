Public Class frmOptBlocek

    Dim HASH As String
    Dim VizualUprava As Boolean = True

    Private Sub btnZrusit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnZrusit.Click
        Me.Close()
    End Sub

    Private Sub btnUlozit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUlozit.Click
1:      On Error GoTo Chyba
2:      If cmbFontNazPredst.Text = "" Or txtFontNazPredst.Text = "" Or cmbFontNazPredst2.Text = "" Or cmbFontOst.Text = "" Or txtFontOst.Text = "" Or _
        cmbFontOst2.Text = "" Or cmbFontKodList.Text = "" Or txtFontKodList.Text = "" Or cmbFontKodList2.Text = "" Then MsgBox("Nie su vyplnene vsetky polia") : Exit Sub
3:
4:      Dim Zmena As Boolean = False
5:
6:      With frmSpracovavanie.NastaveniaBindingSource
7:          .Position = .Find("Option", "LstkBlocekFont1")
8:          .Current("setting") = cmbFontKodList.Text & ";" & txtFontKodList.Text & ";" & cmbFontKodList2.Text
9:          HASH = Module1.RetazecRiadku("Nastavenia", .Position)
10:         .Current("CRC") = Module1.HASHString(HASH)
11:         .EndEdit()
12:         frmSpracovavanie.NastaveniaTableAdapter.Update(frmSpracovavanie.DataSet1.Nastavenia)
13:         frmSpracovavanie.NastaveniaTableAdapter.Fill(frmSpracovavanie.DataSet1.Nastavenia)
14:         frmSpracovavanie.LstkBlocekFont1 = cmbFontKodList.Text & ";" & txtFontKodList.Text & ";" & cmbFontKodList2.Text
15:
16:         .Position = .Find("Option", "LstkBlocekFont2")
17:         .Current("setting") = cmbFontOst.Text & ";" & txtFontOst.Text & ";" & cmbFontOst2.Text
18:         HASH = Module1.RetazecRiadku("Nastavenia", .Position)
19:         .Current("CRC") = Module1.HASHString(HASH)
20:         .EndEdit()
21:         frmSpracovavanie.NastaveniaTableAdapter.Update(frmSpracovavanie.DataSet1.Nastavenia)
22:         frmSpracovavanie.NastaveniaTableAdapter.Fill(frmSpracovavanie.DataSet1.Nastavenia)
23:         frmSpracovavanie.LstkBlocekFont2 = cmbFontOst.Text & ";" & txtFontOst.Text & ";" & cmbFontOst2.Text & ";" & txtFontKodList.Text & ";" & cmbFontKodList2.Text
24:
25:         .Position = .Find("Option", "LstkBlocekFont3")
26:         .Current("setting") = cmbFontNazPredst.Text & ";" & txtFontNazPredst.Text & ";" & cmbFontNazPredst2.Text
27:         HASH = Module1.RetazecRiadku("Nastavenia", .Position)
28:         .Current("CRC") = Module1.HASHString(HASH)
29:         .EndEdit()
30:         frmSpracovavanie.NastaveniaTableAdapter.Update(frmSpracovavanie.DataSet1.Nastavenia)
31:         frmSpracovavanie.NastaveniaTableAdapter.Fill(frmSpracovavanie.DataSet1.Nastavenia)
32:         frmSpracovavanie.LstkBlocekFont3 = cmbFontNazPredst.Text & ";" & txtFontNazPredst.Text & ";" & cmbFontNazPredst2.Text
33:     End With
34:
35:     If txtHlavickaListka1.Text <> frmSpracovavanie.HlavickaVstupenky Then
36:         With frmSpracovavanie.NastaveniaBindingSource
37:             .Position = .Find("Option", "HlavickaListka")
38:             .Current("setting") = txtHlavickaListka1.Text
39:             HASH = Module1.RetazecRiadku("Nastavenia", .Position)
40:             .Current("crc") = Module1.HASHString(HASH)
41:         End With
42:         frmSpracovavanie.HlavickaVstupenky = txtHlavickaListka1.Text
43:         Zmena = True
44:     End If
45:
46:     If txtHlavickaListka2.Text <> frmSpracovavanie.HlavickaVstupenky2 Then
47:         With frmSpracovavanie.NastaveniaBindingSource
48:             .Position = .Find("Option", "HlavickaListka2")
49:             .Current("setting") = txtHlavickaListka2.Text
50:             HASH = Module1.RetazecRiadku("Nastavenia", .Position)
51:             .Current("crc") = Module1.HASHString(HASH)
52:         End With
53:         frmSpracovavanie.HlavickaVstupenky2 = txtHlavickaListka2.Text
54:         Zmena = True
55:     End If
56:
57:     If txtHlavickaListka3.Text <> frmSpracovavanie.HlavickaVstupenky3 Then
58:         With frmSpracovavanie.NastaveniaBindingSource
59:             .Position = .Find("Option", "HlavickaListka3")
60:             .Current("setting") = txtHlavickaListka3.Text
61:             HASH = Module1.RetazecRiadku("Nastavenia", .Position)
62:             .Current("crc") = Module1.HASHString(HASH)
63:         End With
64:         frmSpracovavanie.HlavickaVstupenky3 = txtHlavickaListka3.Text
65:         Zmena = True
66:     End If
67:
68:     If txtHlavickaListka4.Text <> frmSpracovavanie.HlavickaVstupenky4 Then
69:         With frmSpracovavanie.NastaveniaBindingSource
70:             .Position = .Find("Option", "HlavickaListka4")
71:             .Current("setting") = txtHlavickaListka4.Text
72:             HASH = Module1.RetazecRiadku("Nastavenia", .Position)
73:             .Current("crc") = Module1.HASHString(HASH)
74:         End With
75:         frmSpracovavanie.HlavickaVstupenky4 = txtHlavickaListka4.Text
76:         Zmena = True
77:     End If
78:
79:     If txtHlavickaListka5.Text <> frmSpracovavanie.HlavickaVstupenky5 Then
80:         With frmSpracovavanie.NastaveniaBindingSource
81:             .Position = .Find("Option", "HlavickaListka5")
82:             .Current("setting") = txtHlavickaListka5.Text
83:             HASH = Module1.RetazecRiadku("Nastavenia", .Position)
84:             .Current("crc") = Module1.HASHString(HASH)
85:         End With
86:         frmSpracovavanie.HlavickaVstupenky5 = txtHlavickaListka5.Text
87:         Zmena = True
88:     End If

        If txtHlavickaListka6.Text <> frmSpracovavanie.HlavickaVstupenky6 Then
            With frmSpracovavanie.NastaveniaBindingSource
                .Position = .Find("Option", "HlavickaListka6")
                .Current("setting") = txtHlavickaListka6.Text
                HASH = Module1.RetazecRiadku("Nastavenia", .Position)
                .Current("crc") = Module1.HASHString(HASH)
            End With
            frmSpracovavanie.HlavickaVstupenky6 = txtHlavickaListka6.Text
            Zmena = True
        End If

        If txtHlavickaListka7.Text <> frmSpracovavanie.HlavickaVstupenky7 Then
            With frmSpracovavanie.NastaveniaBindingSource
                .Position = .Find("Option", "HlavickaListka7")
                .Current("setting") = txtHlavickaListka7.Text
                HASH = Module1.RetazecRiadku("Nastavenia", .Position)
                .Current("crc") = Module1.HASHString(HASH)
            End With
            frmSpracovavanie.HlavickaVstupenky7 = txtHlavickaListka7.Text
            Zmena = True
        End If

        If txtHlavickaListka8.Text <> frmSpracovavanie.HlavickaVstupenky8 Then
            With frmSpracovavanie.NastaveniaBindingSource
                .Position = .Find("Option", "HlavickaListka8")
                .Current("setting") = txtHlavickaListka8.Text
                HASH = Module1.RetazecRiadku("Nastavenia", .Position)
                .Current("crc") = Module1.HASHString(HASH)
            End With
            frmSpracovavanie.HlavickaVstupenky8 = txtHlavickaListka8.Text
            Zmena = True
        End If

        If txtHlavickaListka9.Text <> frmSpracovavanie.HlavickaVstupenky9 Then
            With frmSpracovavanie.NastaveniaBindingSource
                .Position = .Find("Option", "HlavickaListka9")
                .Current("setting") = txtHlavickaListka9.Text
                HASH = Module1.RetazecRiadku("Nastavenia", .Position)
                .Current("crc") = Module1.HASHString(HASH)
            End With
            frmSpracovavanie.HlavickaVstupenky9 = txtHlavickaListka9.Text
            Zmena = True
        End If

        If txtHlavickaListka10.Text <> frmSpracovavanie.HlavickaVstupenky10 Then
            With frmSpracovavanie.NastaveniaBindingSource
                .Position = .Find("Option", "HlavickaListka10")
                .Current("setting") = txtHlavickaListka10.Text
                HASH = Module1.RetazecRiadku("Nastavenia", .Position)
                .Current("crc") = Module1.HASHString(HASH)
            End With
            frmSpracovavanie.HlavickaVstupenky10 = txtHlavickaListka10.Text
            Zmena = True
        End If

89:
90:     If cbxRezatUstrizok.Checked <> frmSpracovavanie.RezatUstrizok Then
91:         With frmSpracovavanie.NastaveniaBindingSource
92:             .Position = .Find("Option", "RezatUstrizok")
93:             .Current("setting") = CStr(cbxRezatUstrizok.Checked)
94:             HASH = Module1.RetazecRiadku("Nastavenia", .Position)
95:             .Current("crc") = Module1.HASHString(HASH)
96:             frmSpracovavanie.RezatUstrizok = cbxRezatUstrizok.Checked
                Zmena = True
97:         End With
98:     End If
99:
100:    If Zmena = False Then Exit Sub
101:
102:    frmSpracovavanie.NastaveniaBindingSource.EndEdit()
103:    frmSpracovavanie.NastaveniaTableAdapter.Update(frmSpracovavanie.DataSet1.Nastavenia)
104:    frmSpracovavanie.NastaveniaTableAdapter.Fill(frmSpracovavanie.DataSet1.Nastavenia)
105:
106:    'Call Module1.HASHdbKontrola()
107:    Call Module1.HASHSubor("Data\data.pac")
108:
109:    Call ZobrazListok()
110:
111:    MsgBox("Detaily listka Blocek zmenene.")
112:
113:    Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, "Ulozit", ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub frmOptBlocek_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
1:      On Error GoTo Chyba
2:      'Zafarbi
3:      Call mdlColors.Skining(Me)
4:      If VizualUprava Then
5:          Call mdlColors.Sizing(Me)
6:          Me.CenterToParent()
7:          VizualUprava = False
8:      End If
9:
10:     cmbFontNazPredst.Items.Clear()
11:     cmbFontOst.Items.Clear()
12:     cmbFontKodList.Items.Clear()
13:     For i = 1 To FontFamily.Families.Count
14:         cmbFontNazPredst.Items.Add(FontFamily.Families.ElementAt(i - 1).Name)
15:         cmbFontOst.Items.Add(FontFamily.Families.ElementAt(i - 1).Name)
16:         cmbFontKodList.Items.Add(FontFamily.Families.ElementAt(i - 1).Name)
17:     Next i
18:
19:     'Nacita nastavenia listka
20:     Dim nElements As Short
21:     Dim aKey As Object
22:     Dim fnt(5) As String
23:     aKey = Split(frmSpracovavanie.LstkBlocekFont1, ";")
24:     nElements = UBound(aKey) - LBound(aKey) + 1
25:     For i = 0 To nElements - 1
26:         If aKey(i) <> "" Then
27:             fnt(i) = aKey(i)
28:         End If
29:     Next i
30:     cmbFontKodList.Text = fnt(0)
31:     txtFontKodList.Text = fnt(1)
32:     cmbFontKodList2.Text = fnt(2)
33:
34:     aKey = Split(frmSpracovavanie.LstkBlocekFont2, ";")
35:     nElements = UBound(aKey) - LBound(aKey) + 1
36:     For i = 0 To nElements - 1
37:         If aKey(i) <> "" Then
38:             fnt(i) = aKey(i)
39:         End If
40:     Next i
41:     cmbFontOst.Text = fnt(0)
42:     txtFontOst.Text = fnt(1)
43:     cmbFontOst2.Text = fnt(2)
44:
45:     aKey = Split(frmSpracovavanie.LstkBlocekFont3, ";")
46:     nElements = UBound(aKey) - LBound(aKey) + 1
47:     For i = 0 To nElements - 1
48:         If aKey(i) <> "" Then
49:             fnt(i) = aKey(i)
50:         End If
51:     Next i
52:     cmbFontNazPredst.Text = fnt(0)
53:     txtFontNazPredst.Text = fnt(1)
54:     cmbFontNazPredst2.Text = fnt(2)
55:
56:     txtHlavickaListka1.Text = frmSpracovavanie.HlavickaVstupenky
57:     txtHlavickaListka2.Text = frmSpracovavanie.HlavickaVstupenky2
58:     txtHlavickaListka3.Text = frmSpracovavanie.HlavickaVstupenky3
59:     txtHlavickaListka4.Text = frmSpracovavanie.HlavickaVstupenky4
60:     txtHlavickaListka5.Text = frmSpracovavanie.HlavickaVstupenky5
        txtHlavickaListka6.Text = frmSpracovavanie.HlavickaVstupenky6
        txtHlavickaListka7.Text = frmSpracovavanie.HlavickaVstupenky7
        txtHlavickaListka8.Text = frmSpracovavanie.HlavickaVstupenky8
        txtHlavickaListka9.Text = frmSpracovavanie.HlavickaVstupenky9
        txtHlavickaListka10.Text = frmSpracovavanie.HlavickaVstupenky10

61:     cbxRezatUstrizok.Checked = frmSpracovavanie.RezatUstrizok
62:
63:     Call ZobrazListok()
64:
        Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub ZobrazListok()
1:      On Error GoTo NeniFont
2:      Dim styl, stylKod, stylUd As FontStyle
3:
4:      If cmbFontNazPredst2.Text = "Tucne" Then styl = FontStyle.Bold
5:      If cmbFontNazPredst2.Text = "Kurziva" Then styl = FontStyle.Italic
6:      If cmbFontNazPredst2.Text = "Normalne" Then styl = FontStyle.Regular
7:      If cmbFontOst2.Text = "Tucne" Then stylUd = FontStyle.Bold
8:      If cmbFontOst2.Text = "Kurziva" Then stylUd = FontStyle.Italic
9:      If cmbFontOst2.Text = "Normalne" Then stylUd = FontStyle.Regular
10:     If cmbFontKodList2.Text = "Tucne" Then stylKod = FontStyle.Bold
11:     If cmbFontKodList2.Text = "Kurziva" Then stylKod = FontStyle.Italic
12:     If cmbFontKodList2.Text = "Normalne" Then stylKod = FontStyle.Regular
13:
14:     lblHlavicka.Font = New Font(cmbFontKodList.Text, CInt(txtFontKodList.Text), stylKod)
15:     lblNazovPredst.Font = New Font(cmbFontNazPredst.Text, CInt(txtFontNazPredst.Text), styl)
16:     lblZakladUd.Font = New Font(cmbFontOst.Text, CInt(txtFontOst.Text), stylUd)
17:     lblKodZlav.Font = New Font(cmbFontKodList.Text, CInt(txtFontKodList.Text), stylKod)
18:     lblCena.Font = New Font(cmbFontOst.Text, CInt(txtFontOst.Text), stylUd)
19:     lblPrajeme.Font = New Font(cmbFontKodList.Text, CInt(txtFontKodList.Text), stylKod)
20:     lblKontrol.Font = New Font(cmbFontKodList.Text, CInt(txtFontKodList.Text), stylKod)
21:
22:     If frmSpracovavanie.HlavickaVstupenky <> "" Then lblHlavicka.Text = frmSpracovavanie.HlavickaVstupenky
23:     If frmSpracovavanie.HlavickaVstupenky2 <> "" Then lblHlavicka.Text += Chr(13) & Chr(10) & frmSpracovavanie.HlavickaVstupenky2
24:     If frmSpracovavanie.HlavickaVstupenky3 <> "" Then lblHlavicka.Text += Chr(13) & Chr(10) & frmSpracovavanie.HlavickaVstupenky3
25:     If frmSpracovavanie.HlavickaVstupenky4 <> "" Then lblHlavicka.Text += Chr(13) & Chr(10) & frmSpracovavanie.HlavickaVstupenky4
26:     If frmSpracovavanie.HlavickaVstupenky5 <> "" Then lblHlavicka.Text += Chr(13) & Chr(10) & frmSpracovavanie.HlavickaVstupenky5
27:     lblHlavicka.Text += Chr(13) & Chr(10) & "Pokladnik: " & frmLogin.UserName
28:     lblHlavicka.Text += Chr(13) & Chr(10) & CStr(Now)
29:     lblHlavicka.Text += Chr(13) & Chr(10) & "-------------------------------"
30:
31:     lblNazovPredst.Top = lblHlavicka.Top + lblHlavicka.Height + 15
32:     lblNazovPredst.Text = "Nazov Predstavenia"
33:
34:     lblZakladUd.Top = lblNazovPredst.Top + lblNazovPredst.Height + 15
35:     lblZakladUd.Text = Today & "   " & "20:00" & Chr(13) & Chr(10)
36:     lblZakladUd.Text += "Osob: 10" & Chr(13) & Chr(10)
37:     lblZakladUd.Text += "Rad: 15." & Chr(13) & Chr(10)
38:     lblZakladUd.Text += "Miesta: 3L - 2L" & Chr(13) & Chr(10)
39:
40:     lblKodZlav.Top = lblZakladUd.Top + lblZakladUd.Height + 15
41:     lblKodZlav.Text = "        123456789012345" & Chr(13) & Chr(10)
42:     lblKodZlav.Text += "-------------------------------" & Chr(13) & Chr(10)
43:     lblKodZlav.Text += Chr(13) & Chr(10)
44:     lblKodZlav.Text += "Zlavnených miest:   3" & Chr(13) & Chr(10)
45:     lblKodZlav.Text += "Zlavnená cena:      1.10 €" & Chr(13) & Chr(10)
46:
47:     lblCena.Top = lblKodZlav.Top + lblKodZlav.Height
48:     lblCena.Text = "Cena celkom:    100.00 €"
49:
50:     lblPrajeme.Top = lblCena.Top + lblCena.Height + 15
51:     If frmSpracovavanie.HlavickaVstupenky6 <> "" Then lblPrajeme.Text = frmSpracovavanie.HlavickaVstupenky6
        If frmSpracovavanie.HlavickaVstupenky7 <> "" Then lblPrajeme.Text += Chr(13) & Chr(10) & frmSpracovavanie.HlavickaVstupenky7
        If frmSpracovavanie.HlavickaVstupenky8 <> "" Then lblPrajeme.Text += Chr(13) & Chr(10) & frmSpracovavanie.HlavickaVstupenky8
        If frmSpracovavanie.HlavickaVstupenky9 <> "" Then lblPrajeme.Text += Chr(13) & Chr(10) & frmSpracovavanie.HlavickaVstupenky9
        If frmSpracovavanie.HlavickaVstupenky10 <> "" Then lblPrajeme.Text += Chr(13) & Chr(10) & frmSpracovavanie.HlavickaVstupenky10
52:
53:     lblKontrol.Top = lblPrajeme.Top + lblPrajeme.Height
54:     If cbxRezatUstrizok.Checked Then lblKontrol.Top += 30 Else lblKontrol.Top += 15
55:     lblKontrol.Text = "-------Kontrolny ustrizok------" & Chr(13) & Chr(10) & Chr(13) & Chr(10)
56:     lblKontrol.Text += "Predstavenie " & Today & " 20:00" & Chr(13) & Chr(10)
57:     lblKontrol.Text += "            " & frmSpracovavanie.MenoSal(1) & Chr(13) & Chr(10)
58:     lblKontrol.Text += "        123456789012345" & Chr(13) & Chr(10)
59:     lblKontrol.Text += "Os.10 rad.15. Mst.3L - 2L"
60:
61:     Exit Sub
NeniFont:
        lblHlavicka.Text = ""
        lblNazovPredst.Text = ""
        lblZakladUd.Text = ""
        lblKodZlav.Text = ""
        lblCena.Text = ""
        lblPrajeme.Text = ""
        lblKontrol.Text = ""
    End Sub

    Private Sub btnDefault_Click(sender As Object, e As EventArgs) Handles btnDefault.Click
        ' If MsgBox("Obnovit povodne pismo?", MsgBoxStyle.YesNo + MsgBoxStyle.Question) = MsgBoxResult.No Then Exit Sub
        cmbFontKodList.Text = "Lucida Console"
        txtFontKodList.Text = 10
        cmbFontKodList2.Text = "Tucne"
        cmbFontOst.Text = "Lucida Console"
        txtFontOst.Text = 10
        cmbFontOst2.Text = "Tucne"
        cmbFontNazPredst.Text = "Lucida Console"
        txtFontNazPredst.Text = 12
        cmbFontNazPredst2.Text = "Tucne"
        Call ZobrazListok()
    End Sub

    Private Sub cbxRezatUstrizok_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxRezatUstrizok.CheckedChanged
        lblKontrol.Top = lblPrajeme.Top + lblPrajeme.Height
        If cbxRezatUstrizok.Checked Then lblKontrol.Top += 30 Else lblKontrol.Top += 15
    End Sub

    Private Sub txtHlavickaListka1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        frmDistribut.txtDistributor_KeyPress(sender, e, Me)
    End Sub
    Private Sub txtHlavickaListka2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        frmDistribut.txtDistributor_KeyPress(sender, e, Me)
    End Sub
    Private Sub txtHlavickaListka3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        frmDistribut.txtDistributor_KeyPress(sender, e, Me)
    End Sub
    Private Sub txtHlavickaListka4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        frmDistribut.txtDistributor_KeyPress(sender, e, Me)
    End Sub
    Private Sub txtHlavickaListka5_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        frmDistribut.txtDistributor_KeyPress(sender, e, Me)
    End Sub
    Private Sub btnUlozit_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles btnUlozit.KeyPress
        frmPredstavenia.btnPridaj_KeyPress(sender, e, Me)
    End Sub
    Private Sub btnZrusit_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles btnZrusit.KeyPress
        frmPredstavenia.btnPridaj_KeyPress(sender, e, Me)
    End Sub
    Private Sub txtFontNazPredst_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFontNazPredst.KeyPress
        frmPredstavenia.txtCena_KeyPress(sender, e, Me)
    End Sub
    Private Sub txtFontOst_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFontOst.KeyPress
        frmPredstavenia.txtCena_KeyPress(sender, e, Me)
    End Sub
    Private Sub txtFontKodList_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFontKodList.KeyPress
        frmPredstavenia.txtCena_KeyPress(sender, e, Me)
    End Sub
    Private Sub cmbFontOst2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmbFontOst2.KeyPress
        frmPredstavenia.cmbFilmy_KeyPress(sender, e, Me)
    End Sub
    Private Sub cmbFontNazPredst2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmbFontNazPredst2.KeyPress
        frmPredstavenia.cmbFilmy_KeyPress(sender, e, Me)
    End Sub
    Private Sub cmbFontKodList2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmbFontKodList2.KeyPress
        frmPredstavenia.cmbFilmy_KeyPress(sender, e, Me)
    End Sub

    Private Sub cmbFontNazPredst_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbFontNazPredst.SelectedIndexChanged
        Call ZobrazListok()
    End Sub
    Private Sub cmbFontNazPredst2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbFontNazPredst2.SelectedIndexChanged
        Call ZobrazListok()
    End Sub
    Private Sub txtFontNazPredst_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFontNazPredst.TextChanged
        Call ZobrazListok()
    End Sub
    Private Sub cmbFontOst_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbFontOst.SelectedIndexChanged
        Call ZobrazListok()
    End Sub
    Private Sub cmbFontOst2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbFontOst2.SelectedIndexChanged
        Call ZobrazListok()
    End Sub
    Private Sub txtFontOst_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFontOst.TextChanged
        Call ZobrazListok()
    End Sub
    Private Sub cmbFontKodList_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbFontKodList.SelectedIndexChanged
        Call ZobrazListok()
    End Sub
    Private Sub cmbFontKodList2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbFontKodList2.SelectedIndexChanged
        Call ZobrazListok()
    End Sub
    Private Sub txtFontKodList_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFontKodList.TextChanged
        Call ZobrazListok()
    End Sub
End Class