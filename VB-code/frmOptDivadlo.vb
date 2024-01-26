Public Class frmOptDivadlo

    Dim HASH As String
    Dim VizualUprava As Boolean = True

    Dim PremDatum As Boolean = False
    Dim PremCas As Boolean = False
    Dim PremNazov As Boolean = False
    Dim PremRad As Boolean = False
    Dim PremMiesta As Boolean = False
    Dim PremOsob As Boolean = False
    Dim PremCena As Boolean = False
    Dim PremKod As Boolean = False

    Private Sub frmOptDivadlo_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
1:      On Error GoTo Chyba
2:      'Zafarbi
3:      Call mdlColors.Skining(Me)
4:      If VizualUprava Then
5:          Call mdlColors.Sizing(Me)
6:          Me.CenterToParent()
7:          VizualUprava = False
8:      End If
9:
10:     picListok.Image = System.Drawing.Image.FromFile(frmSpracovavanie.Adresa2 + "obr\logo.jpg")
11:     lblDatum.Text = Today
12:
13:     cmbFontNazPredst.Items.Clear()
14:     cmbFontOst.Items.Clear()
15:     For i = 1 To FontFamily.Families.Count
16:         cmbFontNazPredst.Items.Add(FontFamily.Families.ElementAt(i - 1).Name)
17:         cmbFontOst.Items.Add(FontFamily.Families.ElementAt(i - 1).Name)
18:     Next i
19:
20:     'Nacita nastavenia listka
21:     Dim nElements As Short
22:     Dim aKey As Object
23:     Dim fnt(10) As String
24:     aKey = Split(frmSpracovavanie.LstkDivadloFont1, ";")
25:     nElements = UBound(aKey) - LBound(aKey) + 1
26:     For i = 0 To nElements - 1
27:         If aKey(i) <> "" Then
28:             fnt(i) = aKey(i)
29:         End If
30:     Next i
31:     cmbFontNazPredst.Text = fnt(0)
32:     txtFontNazPredst.Text = fnt(1)
33:     cmbFontNazPredst2.Text = fnt(2)
34:
35:     aKey = Split(frmSpracovavanie.LstkDivadloFont2, ";")
36:     nElements = UBound(aKey) - LBound(aKey) + 1
37:     For i = 0 To nElements - 1
38:         If aKey(i) <> "" Then
39:             fnt(i) = aKey(i)
40:         End If
41:     Next i
42:     cmbFontOst.Text = fnt(0)
43:     txtFontOst.Text = fnt(1)
44:     cmbFontOst2.Text = fnt(2)
45:     txtFontKodList.Text = fnt(3)
46:     cmbFontKodList2.Text = fnt(4)
47:
48:     aKey = Split(frmSpracovavanie.LstkDivadloPozicie1, ";")
49:     nElements = UBound(aKey) - LBound(aKey) + 1
50:     For i = 0 To nElements - 1
51:         If aKey(i) <> "" Then
52:             fnt(i) = aKey(i)
53:         End If
54:     Next i
55:     txtListokX.Text = fnt(0)
56:     txtListokY.Text = fnt(1)
57:     txtDatumX.Text = fnt(2)
58:     txtDatumY.Text = fnt(3)
59:     txtCasX.Text = fnt(4)
60:     txtCasY.Text = fnt(5)
61:     txtNazovX.Text = fnt(6)
62:     txtNazovY.Text = fnt(7)
63:     txtRadX.Text = fnt(8)
64:     txtRadY.Text = fnt(9)
65:
66:     aKey = Split(frmSpracovavanie.LstkDivadloPozicie2, ";")
67:     nElements = UBound(aKey) - LBound(aKey) + 1
68:     For i = 0 To nElements - 1
69:         If aKey(i) <> "" Then
70:             fnt(i) = aKey(i)
71:         End If
72:     Next i
73:     txtMiestaX.Text = fnt(0)
74:     txtMiestaY.Text = fnt(1)
75:     txtOsobX.Text = fnt(2)
76:     txtOsobY.Text = fnt(3)
77:     txtCenaX.Text = fnt(4)
78:     txtCenaY.Text = fnt(5)
79:     txtKodX.Text = fnt(6)
80:     txtKodY.Text = fnt(7)
81:     If fnt(8) = "0" Then cbxZobraz.Checked = False Else cbxZobraz.Checked = True
82:     'Vlozi nastavenia do textboxov
83:     picListok.Width = CInt(txtListokX.Text) * frmSpracovavanie.Zvacsenie
84:     picListok.Height = CInt(txtListokY.Text) * frmSpracovavanie.Zvacsenie
85:
        Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub btnUlozit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUlozit.Click
1:      On Error GoTo Chyba
2:      If cmbFontNazPredst.Text = "" Or txtFontNazPredst.Text = "" Or cmbFontNazPredst2.Text = "" Or cmbFontOst.Text = "" Or txtFontOst.Text = "" Or _
            cmbFontOst2.Text = "" Or txtFontKodList.Text = "" Or cmbFontKodList2.Text = "" Or txtListokX.Text = "" Or txtListokY.Text = "" Or _
            txtDatumX.Text = "" Or txtDatumY.Text = "" Or txtCasX.Text = "" Or txtCasY.Text = "" Or txtNazovX.Text = "" Or _
            txtNazovY.Text = "" Or txtRadX.Text = "" Or txtRadY.Text = "" Or txtMiestaX.Text = "" Or txtMiestaY.Text = "" Or _
            txtOsobX.Text = "" Or txtOsobY.Text = "" Or txtCenaX.Text = "" Or txtCenaY.Text = "" Or txtKodX.Text = "" Or _
            txtKodY.Text = "" Then MsgBox("Nie su vyplnene vsetky polia") : Exit Sub

3:      With frmSpracovavanie.NastaveniaBindingSource
4:          'Font Nazvu Predstavenia
5:          .Position = .Find("Option", "LstkDivadloFont1")
6:          .Current("setting") = cmbFontNazPredst.Text & ";" & txtFontNazPredst.Text & ";" & cmbFontNazPredst2.Text
7:          HASH = Module1.RetazecRiadku("Nastavenia", .Position)
8:          .Current("CRC") = Module1.HASHString(HASH)
9:          .EndEdit()
10:         frmSpracovavanie.NastaveniaTableAdapter.Update(frmSpracovavanie.DataSet1.Nastavenia)
11:         frmSpracovavanie.NastaveniaTableAdapter.Fill(frmSpracovavanie.DataSet1.Nastavenia)
12:         frmSpracovavanie.LstkDivadloFont1 = cmbFontNazPredst.Text & ";" & txtFontNazPredst.Text & ";" & cmbFontNazPredst2.Text
13:         'Font ostatnych veci
14:         .Position = .Find("Option", "LstkDivadloFont2")
15:         .Current("setting") = cmbFontOst.Text & ";" & txtFontOst.Text & ";" & cmbFontOst2.Text & ";" & txtFontKodList.Text & ";" & cmbFontKodList2.Text
16:         HASH = Module1.RetazecRiadku("Nastavenia", .Position)
17:         .Current("CRC") = Module1.HASHString(HASH)
18:         .EndEdit()
19:         frmSpracovavanie.NastaveniaTableAdapter.Update(frmSpracovavanie.DataSet1.Nastavenia)
20:         frmSpracovavanie.NastaveniaTableAdapter.Fill(frmSpracovavanie.DataSet1.Nastavenia)
21:         frmSpracovavanie.LstkDivadloFont2 = cmbFontOst.Text & ";" & txtFontOst.Text & ";" & cmbFontOst2.Text & ";" & txtFontKodList.Text & ";" & cmbFontKodList2.Text
22:         'Polohy prva polovica
23:         .Position = .Find("Option", "LstkDivadloPozicie1")
24:         .Current("setting") = txtListokX.Text & ";" & txtListokY.Text & ";" & txtDatumX.Text & ";" & txtDatumY.Text & ";" & txtCasX.Text & ";" & txtCasY.Text & ";" & txtNazovX.Text & ";" & txtNazovY.Text & ";" & txtRadX.Text & ";" & txtRadY.Text
25:         HASH = Module1.RetazecRiadku("Nastavenia", .Position)
26:         .Current("CRC") = Module1.HASHString(HASH)
27:         .EndEdit()
28:         frmSpracovavanie.NastaveniaTableAdapter.Update(frmSpracovavanie.DataSet1.Nastavenia)
29:         frmSpracovavanie.NastaveniaTableAdapter.Fill(frmSpracovavanie.DataSet1.Nastavenia)
30:         frmSpracovavanie.LstkDivadloPozicie1 = txtListokX.Text & ";" & txtListokY.Text & ";" & txtDatumX.Text & ";" & txtDatumY.Text & ";" & txtCasX.Text & ";" & txtCasY.Text & ";" & txtNazovX.Text & ";" & txtNazovY.Text & ";" & txtRadX.Text & ";" & txtRadY.Text
31:         'Polohy druha polovica
32:         .Position = .Find("Option", "LstkDivadloPozicie2")
33:         .Current("setting") = txtMiestaX.Text & ";" & txtMiestaY.Text & ";" & txtOsobX.Text & ";" & txtOsobY.Text & ";" & txtCenaX.Text & ";" & txtCenaY.Text & ";" & txtKodX.Text & ";" & txtKodY.Text
34:         If cbxZobraz.Checked Then .Current("setting") += ";1" Else .Current("setting") += ";0"
35:         HASH = Module1.RetazecRiadku("Nastavenia", .Position)
36:         .Current("CRC") = Module1.HASHString(HASH)
37:         .EndEdit()
38:         frmSpracovavanie.NastaveniaTableAdapter.Update(frmSpracovavanie.DataSet1.Nastavenia)
39:         frmSpracovavanie.NastaveniaTableAdapter.Fill(frmSpracovavanie.DataSet1.Nastavenia)
40:         frmSpracovavanie.LstkDivadloPozicie2 = txtMiestaX.Text & ";" & txtMiestaY.Text & ";" & txtOsobX.Text & ";" & txtOsobY.Text & ";" & txtCenaX.Text & ";" & txtCenaY.Text & ";" & txtKodX.Text & ";" & txtKodY.Text
41:         If cbxZobraz.Checked Then frmSpracovavanie.LstkDivadloPozicie2 += ";1" Else frmSpracovavanie.LstkDivadloPozicie2 += ";0"
42:     End With
43:
44:     'Call Module1.HASHdbKontrola()
45:     Call Module1.HASHSubor("Data\data.pac")
46:
47:     MsgBox("Detaily listka Divadlo zmenene.")
48:
        Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, "ZmenaFontDivadlo", ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub btnZrusit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnZrusit.Click
        Me.Close()
    End Sub

    Private Sub btnUlozit_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles btnUlozit.KeyPress
        frmPredstavenia.btnPridaj_KeyPress(sender, e, Me)
    End Sub
    Private Sub btnZrusit_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles btnZrusit.KeyPress
        frmPredstavenia.btnPridaj_KeyPress(sender, e, Me)
    End Sub

    Private Sub cmbFontNazPredst_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbFontNazPredst.SelectedIndexChanged
        On Error GoTo NeniFont
        Dim styl As FontStyle
        Dim Velkost As Integer
        lblNazPredst.Text = "Nazov Predstavenia"
        If cmbFontNazPredst2.Text = "Tucne" Then styl = FontStyle.Bold
        If cmbFontNazPredst2.Text = "Kurziva" Then styl = FontStyle.Italic
        If cmbFontNazPredst2.Text = "Normalne" Then styl = FontStyle.Regular
        Velkost = Format(CInt(txtFontNazPredst.Text) * frmSpracovavanie.Zvacsenie, "##0")
        lblNazPredst.Font = New Font(cmbFontNazPredst.Text, Velkost, styl)
        Exit Sub
NeniFont:
        lblNazPredst.Text = ""
    End Sub
    Private Sub txtFontNazPredst_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFontNazPredst.TextChanged
        cmbFontNazPredst_SelectedIndexChanged(sender, e)
    End Sub
    Private Sub cmbFontNazPredst2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbFontNazPredst2.SelectedIndexChanged
        cmbFontNazPredst_SelectedIndexChanged(sender, e)
    End Sub

    Private Sub cmbFontOst_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbFontOst.SelectedIndexChanged
        On Error GoTo NeniFont
        Dim styl, stylKod As FontStyle
        Dim Velkost1, Velkost2 As Integer

        lblDatum.Text = Today
        lblCas.Text = "20:00"
        lblCena.Text = "100E"
        lblKod.Text = "123456789012345"
        If cbxZobraz.Checked Then
            lblRad.Text = "Rad: 15."
            lblMiesta.Text = "Miesta: 3L - 2L"
            lblOsob.Text = "Osob: 10"
        Else
            lblRad.Text = "15."
            lblMiesta.Text = "3L - 2L"
            lblOsob.Text = "10"
        End If

        If cmbFontOst2.Text = "Tucne" Then styl = FontStyle.Bold
        If cmbFontOst2.Text = "Kurziva" Then styl = FontStyle.Italic
        If cmbFontOst2.Text = "Normalne" Then styl = FontStyle.Regular

        If cmbFontKodList2.Text = "Tucne" Then stylKod = FontStyle.Bold
        If cmbFontKodList2.Text = "Kurziva" Then stylKod = FontStyle.Italic
        If cmbFontKodList2.Text = "Normalne" Then stylKod = FontStyle.Regular

        Velkost1 = Format(CInt(txtFontOst.Text) * frmSpracovavanie.Zvacsenie, "##0")
        Velkost2 = Format(CInt(txtFontKodList.Text) * frmSpracovavanie.Zvacsenie, "##0")

        lblDatum.Font = New Font(cmbFontOst.Text, Velkost1, styl)
        lblCas.Font = New Font(cmbFontOst.Text, Velkost1, styl)
        lblRad.Font = New Font(cmbFontOst.Text, Velkost1, styl)
        lblMiesta.Font = New Font(cmbFontOst.Text, Velkost1, styl)
        lblCena.Font = New Font(cmbFontOst.Text, Velkost1, styl)
        lblOsob.Font = New Font(cmbFontOst.Text, Velkost1, styl)
        lblKod.Font = New Font(cmbFontOst.Text, Velkost2, stylKod)

        Exit Sub
NeniFont:
        lblDatum.Text = ""
        lblCas.Text = ""
        lblRad.Text = ""
        lblMiesta.Text = ""
        lblCena.Text = ""
        lblOsob.Text = ""
        lblKod.Text = ""
    End Sub
    Private Sub txtFontOst_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFontOst.TextChanged
        cmbFontOst_SelectedIndexChanged(sender, e)
    End Sub
    Private Sub cmbFontOst2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbFontOst2.SelectedIndexChanged
        cmbFontOst_SelectedIndexChanged(sender, e)
    End Sub
    Private Sub txtFontKodList_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFontKodList.TextChanged
        cmbFontOst_SelectedIndexChanged(sender, e)
    End Sub
    Private Sub cmbFontKodList2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbFontKodList2.SelectedIndexChanged
        cmbFontOst_SelectedIndexChanged(sender, e)
    End Sub
    Private Sub cbxZobraz_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxZobraz.CheckedChanged
        cmbFontOst_SelectedIndexChanged(sender, e)
    End Sub

    Private Sub lblDatum_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblDatum.MouseMove
        If PremDatum Then
            lblDatum.Left = MousePosition.X - Me.Location.X - lblDatum.Width / 2
            lblDatum.Top = MousePosition.Y - Me.Location.Y - lblDatum.Height / 2 - 24
            txtDatumX.Text = Format((lblDatum.Left - picListok.Left) / frmSpracovavanie.Zvacsenie, "##0")
            txtDatumY.Text = Format((lblDatum.Top - picListok.Top) / frmSpracovavanie.Zvacsenie, "##0")
        End If
    End Sub
    Private Sub lblCas_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblCas.MouseMove
        If PremCas Then
            lblCas.Left = MousePosition.X - Me.Location.X - lblCas.Width / 2
            lblCas.Top = MousePosition.Y - Me.Location.Y - lblCas.Height / 2 - 24
            txtCasX.Text = Format((lblCas.Left - picListok.Left) / frmSpracovavanie.Zvacsenie, "##0")
            txtCasY.Text = Format((lblCas.Top - picListok.Top) / frmSpracovavanie.Zvacsenie, "##0")
        End If
    End Sub
    Private Sub lblNazPredst_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblNazPredst.MouseMove
        If PremNazov Then
            lblNazPredst.Left = MousePosition.X - Me.Location.X - lblNazPredst.Width / 2
            lblNazPredst.Top = MousePosition.Y - Me.Location.Y - lblNazPredst.Height / 2 - 24
            txtNazovX.Text = Format((lblNazPredst.Left - picListok.Left) / frmSpracovavanie.Zvacsenie, "##0")
            txtNazovY.Text = Format((lblNazPredst.Top - picListok.Top) / frmSpracovavanie.Zvacsenie, "##0")
        End If
    End Sub
    Private Sub lblRad_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblRad.MouseMove
        If PremRad Then
            lblRad.Left = MousePosition.X - Me.Location.X - lblRad.Width / 2
            lblRad.Top = MousePosition.Y - Me.Location.Y - lblRad.Height / 2 - 24
            txtRadX.Text = Format((lblRad.Left - picListok.Left) / frmSpracovavanie.Zvacsenie, "##0")
            txtRadY.Text = Format((lblRad.Top - picListok.Top) / frmSpracovavanie.Zvacsenie, "##0")
        End If
    End Sub
    Private Sub lblMiesta_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblMiesta.MouseMove
        If PremMiesta Then
            lblMiesta.Left = MousePosition.X - Me.Location.X - lblMiesta.Width / 2
            lblMiesta.Top = MousePosition.Y - Me.Location.Y - lblMiesta.Height / 2 - 24
            txtMiestaX.Text = Format((lblMiesta.Left - picListok.Left) / frmSpracovavanie.Zvacsenie, "##0")
            txtMiestaY.Text = Format((lblMiesta.Top - picListok.Top) / frmSpracovavanie.Zvacsenie, "##0")
        End If
    End Sub
    Private Sub lblOsob_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblOsob.MouseMove
        If PremOsob Then
            lblOsob.Left = MousePosition.X - Me.Location.X - lblOsob.Width / 2
            lblOsob.Top = MousePosition.Y - Me.Location.Y - lblOsob.Height / 2 - 24
            txtOsobX.Text = Format((lblOsob.Left - picListok.Left) / frmSpracovavanie.Zvacsenie, "##0")
            txtOsobY.Text = Format((lblOsob.Top - picListok.Top) / frmSpracovavanie.Zvacsenie, "##0")
        End If
    End Sub
    Private Sub lblCena_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblCena.MouseMove
        If PremCena Then
            lblCena.Left = MousePosition.X - Me.Location.X - lblCena.Width / 2
            lblCena.Top = MousePosition.Y - Me.Location.Y - lblCena.Height / 2 - 24
            txtCenaX.Text = Format((lblCena.Left - picListok.Left) / frmSpracovavanie.Zvacsenie, "##0")
            txtCenaY.Text = Format((lblCena.Top - picListok.Top) / frmSpracovavanie.Zvacsenie, "##0")
        End If
    End Sub
    Private Sub lblKod_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblKod.MouseMove
        If PremKod Then
            lblKod.Left = MousePosition.X - Me.Location.X - lblKod.Width / 2
            lblKod.Top = MousePosition.Y - Me.Location.Y - lblKod.Height / 2 - 24
            txtKodX.Text = Format((lblKod.Left - picListok.Left) / frmSpracovavanie.Zvacsenie, "##0")
            txtKodY.Text = Format((lblKod.Top - picListok.Top) / frmSpracovavanie.Zvacsenie, "##0")
        End If
    End Sub

    Private Sub lblDatum_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblDatum.MouseDown
        PremDatum = True
    End Sub
    Private Sub lblDatum_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblDatum.MouseUp
        PremDatum = False
    End Sub
    Private Sub lblCas_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblCas.MouseDown
        PremCas = True
    End Sub
    Private Sub lblCas_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblCas.MouseUp
        PremCas = False
    End Sub
    Private Sub lblNazPredst_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblNazPredst.MouseDown
        PremNazov = True
    End Sub
    Private Sub lblNazPredst_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblNazPredst.MouseUp
        PremNazov = False
    End Sub
    Private Sub lblRad_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblRad.MouseDown
        PremRad = True
    End Sub
    Private Sub lblRad_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblRad.MouseUp
        PremRad = False
    End Sub
    Private Sub lblMiesta_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblMiesta.MouseDown
        PremMiesta = True
    End Sub
    Private Sub lblMiesta_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblMiesta.MouseUp
        PremMiesta = False
    End Sub
    Private Sub lblOsob_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblOsob.MouseDown
        PremOsob = True
    End Sub
    Private Sub lblOsob_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblOsob.MouseUp
        PremOsob = False
    End Sub
    Private Sub lblCena_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblCena.MouseDown
        PremCena = True
    End Sub
    Private Sub lblCena_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblCena.MouseUp
        PremCena = False
    End Sub
    Private Sub lblKod_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblKod.MouseDown
        PremKod = True
    End Sub
    Private Sub lblKod_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblKod.MouseUp
        PremKod = False
    End Sub

    Private Sub txtListokX_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtListokX.TextChanged
        If txtListokX.Text = "" Or txtListokY.Text = "" Then Exit Sub
        picListok.Width = CInt(txtListokX.Text) * frmSpracovavanie.Zvacsenie
        lblVelkost.Text = "Velkost: " & Format(txtListokX.Text / 39, "#0.0") & " x " & Format(txtListokY.Text / 39, "#0.0") & " cm"
    End Sub
    Private Sub txtListokY_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtListokY.TextChanged
        If txtListokX.Text = "" Or txtListokY.Text = "" Then Exit Sub
        picListok.Height = CInt(txtListokY.Text) * frmSpracovavanie.Zvacsenie
        lblVelkost.Text = "Velkost: " & Format(txtListokX.Text / 39, "#0.0") & " x " & Format(txtListokY.Text / 39, "#0.0") & " cm"
    End Sub
    Private Sub txtDatumX_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDatumX.TextChanged
        If txtDatumX.Text = "" Then Exit Sub
        lblDatum.Left = CInt(txtDatumX.Text) * frmSpracovavanie.Zvacsenie + picListok.Left
    End Sub
    Private Sub txtDatumY_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDatumY.TextChanged
        If txtDatumY.Text = "" Then Exit Sub
        lblDatum.Top = CInt(txtDatumY.Text) * frmSpracovavanie.Zvacsenie + picListok.Top
    End Sub
    Private Sub txtCasX_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCasX.TextChanged
        If txtCasX.Text = "" Then Exit Sub
        lblCas.Left = CInt(txtCasX.Text) * frmSpracovavanie.Zvacsenie + picListok.Left
    End Sub
    Private Sub txtCasY_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCasY.TextChanged
        If txtCasY.Text = "" Then Exit Sub
        lblCas.Top = CInt(txtCasY.Text) * frmSpracovavanie.Zvacsenie + picListok.Top
    End Sub
    Private Sub txtNazovX_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNazovX.TextChanged
        If txtNazovX.Text = "" Then Exit Sub
        lblNazPredst.Left = CInt(txtNazovX.Text) * frmSpracovavanie.Zvacsenie + picListok.Left
    End Sub
    Private Sub txtNazovY_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNazovY.TextChanged
        If txtNazovY.Text = "" Then Exit Sub
        lblNazPredst.Top = CInt(txtNazovY.Text) * frmSpracovavanie.Zvacsenie + picListok.Top
    End Sub
    Private Sub txtRadX_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRadX.TextChanged
        If txtRadX.Text = "" Then Exit Sub
        lblRad.Left = CInt(txtRadX.Text) * frmSpracovavanie.Zvacsenie + picListok.Left
    End Sub
    Private Sub txtRadY_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRadY.TextChanged
        If txtRadY.Text = "" Then Exit Sub
        lblRad.Top = CInt(txtRadY.Text) * frmSpracovavanie.Zvacsenie + picListok.Top
    End Sub
    Private Sub txtMiestaX_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMiestaX.TextChanged
        If txtMiestaX.Text = "" Then Exit Sub
        lblMiesta.Left = CInt(txtMiestaX.Text) * frmSpracovavanie.Zvacsenie + picListok.Left
    End Sub
    Private Sub txtMiestaY_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMiestaY.TextChanged
        If txtMiestaY.Text = "" Then Exit Sub
        lblMiesta.Top = CInt(txtMiestaY.Text) * frmSpracovavanie.Zvacsenie + picListok.Top
    End Sub
    Private Sub txtOsobX_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOsobX.TextChanged
        If txtOsobX.Text = "" Then Exit Sub
        lblOsob.Left = CInt(txtOsobX.Text) * frmSpracovavanie.Zvacsenie + picListok.Left
    End Sub
    Private Sub txtOsobY_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOsobY.TextChanged
        If txtOsobY.Text = "" Then Exit Sub
        lblOsob.Top = CInt(txtOsobY.Text) * frmSpracovavanie.Zvacsenie + picListok.Top
    End Sub
    Private Sub txtCenaX_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCenaX.TextChanged
        If txtCenaX.Text = "" Then Exit Sub
        lblCena.Left = CInt(txtCenaX.Text) * frmSpracovavanie.Zvacsenie + picListok.Left
    End Sub
    Private Sub txtCenaY_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCenaY.TextChanged
        If txtCenaY.Text = "" Then Exit Sub
        lblCena.Top = CInt(txtCenaY.Text) * frmSpracovavanie.Zvacsenie + picListok.Top
    End Sub
    Private Sub txtKodX_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtKodX.TextChanged
        If txtKodX.Text = "" Then Exit Sub
        lblKod.Left = CInt(txtKodX.Text) * frmSpracovavanie.Zvacsenie + picListok.Left
    End Sub
    Private Sub txtKodY_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtKodY.TextChanged
        If txtKodY.Text = "" Then Exit Sub
        lblKod.Top = CInt(txtKodY.Text) * frmSpracovavanie.Zvacsenie + picListok.Top
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
    Private Sub txtFontNazPredst_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFontNazPredst.KeyPress
        frmPredstavenia.txtCena_KeyPress(sender, e, Me)
    End Sub
    Private Sub txtFontOst_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFontOst.KeyPress
        frmPredstavenia.txtCena_KeyPress(sender, e, Me)
    End Sub
    Private Sub txtFontKodList_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFontKodList.KeyPress
        frmPredstavenia.txtCena_KeyPress(sender, e, Me)
    End Sub
    Private Sub txtListokX_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtListokX.KeyPress
        frmPredstavenia.txtCena_KeyPress(sender, e, Me)
    End Sub
    Private Sub txtListokY_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtListokY.KeyPress
        frmPredstavenia.txtCena_KeyPress(sender, e, Me)
    End Sub
    Private Sub txtDatumX_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDatumX.KeyPress
        frmPredstavenia.txtCena_KeyPress(sender, e, Me)
    End Sub
    Private Sub txtDatumY_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDatumY.KeyPress
        frmPredstavenia.txtCena_KeyPress(sender, e, Me)
    End Sub
    Private Sub txtCasX_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCasX.KeyPress
        frmPredstavenia.txtCena_KeyPress(sender, e, Me)
    End Sub
    Private Sub txtCasY_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCasY.KeyPress
        frmPredstavenia.txtCena_KeyPress(sender, e, Me)
    End Sub
    Private Sub txtNazovX_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtNazovX.KeyPress
        frmPredstavenia.txtCena_KeyPress(sender, e, Me)
    End Sub
    Private Sub txtNazovY_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtNazovY.KeyPress
        frmPredstavenia.txtCena_KeyPress(sender, e, Me)
    End Sub
    Private Sub txtRadX_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtRadX.KeyPress
        frmPredstavenia.txtCena_KeyPress(sender, e, Me)
    End Sub
    Private Sub txtRadY_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtRadY.KeyPress
        frmPredstavenia.txtCena_KeyPress(sender, e, Me)
    End Sub
    Private Sub txtMiestaX_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMiestaX.KeyPress
        frmPredstavenia.txtCena_KeyPress(sender, e, Me)
    End Sub
    Private Sub txtMiestaY_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMiestaY.KeyPress
        frmPredstavenia.txtCena_KeyPress(sender, e, Me)
    End Sub
    Private Sub txtOsobX_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtOsobX.KeyPress
        frmPredstavenia.txtCena_KeyPress(sender, e, Me)
    End Sub
    Private Sub txtOsobY_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtOsobY.KeyPress
        frmPredstavenia.txtCena_KeyPress(sender, e, Me)
    End Sub
    Private Sub txtCenaX_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCenaX.KeyPress
        frmPredstavenia.txtCena_KeyPress(sender, e, Me)
    End Sub
    Private Sub txtCenaY_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCenaY.KeyPress
        frmPredstavenia.txtCena_KeyPress(sender, e, Me)
    End Sub
    Private Sub txtKodX_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtKodX.KeyPress
        frmPredstavenia.txtCena_KeyPress(sender, e, Me)
    End Sub
    Private Sub txtKodY_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtKodY.KeyPress
        frmPredstavenia.txtCena_KeyPress(sender, e, Me)
    End Sub

    Private Sub frmOptDivadlo_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseMove
        ' lblVelkost.Text = MousePosition.X & " " & MousePosition.Y
    End Sub
End Class