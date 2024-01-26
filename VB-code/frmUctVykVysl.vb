Public Class frmUctVykVysl

    Dim VizualUprava As Boolean = True

    Private Structure Vysledky
        Dim Datum As Date
        Dim NazovFilmu As String
        Dim Trzba As Double
        Dim Predstaveni As Double
        Dim Divakov As Double
        Dim Zlav As Double
        Dim PevnePoz As Double
        Dim PercentPoz As Double
    End Structure

    Private Sub frmUctVykVysl_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
1:      On Error GoTo Chyba
2:      'Zafarbi
3:      Call mdlColors.Skining(Me)
4:      If VizualUprava Then
5:          Call mdlColors.Sizing(Me)
6:          Me.CenterToParent()
7:          VizualUprava = False
8:      End If
9:
10:     Dim i As Short
11:
12:     '************ SPRAVA DATA GRIDU*************************
13:     With DataGridView1
14:         .SuspendLayout()
15:         .DataSource = Nothing
16:         .AllowUserToAddRows = False
17:         .AllowUserToDeleteRows = False
18:         .AllowUserToResizeRows = False
19:         .AllowUserToResizeColumns = False
20:
21:         .AllowUserToOrderColumns = True
22:         .SelectionMode = DataGridViewSelectionMode.FullRowSelect
23:         .ReadOnly = True
24:         .MultiSelect = False
25:         .RowHeadersVisible = False
26:         .Columns.Clear()
27:         .Rows.Clear()
28:         ' setup columns
29:
30:         .Columns.Add("dtgrdDatum", "Den")
31:         .Columns(0).Width = (.Width - 20) * 0.07
32:         .Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
33:         .Columns.Add("dtgrdDen", "Nazov")
34:         .Columns(1).Width = (.Width - 20) * 0.19
35:         .Columns(1).SortMode = DataGridViewColumnSortMode.NotSortable
36:         .Columns.Add("dtgrdCas", "Trzba")
37:         .Columns(2).Width = (.Width - 20) * 0.08
38:         .Columns(2).SortMode = DataGridViewColumnSortMode.NotSortable
39:         .Columns.Add("dtgrdFilm", "Pr.")
40:         .Columns(3).Width = (.Width - 20) * 0.05
41:         .Columns(3).SortMode = DataGridViewColumnSortMode.NotSortable
42:         .Columns.Add("dtgrdSala", "Fond")
43:         .Columns(4).Width = (.Width - 20) * 0.06
44:         .Columns(4).SortMode = DataGridViewColumnSortMode.NotSortable
45:         .Columns.Add("dtgrdSala", "Div.")
46:         .Columns(5).Width = (.Width - 20) * 0.05
47:         .Columns(5).SortMode = DataGridViewColumnSortMode.NotSortable
48:         .Columns.Add("dtgrdSala", "Zlav")
49:         .Columns(6).Width = (.Width - 20) * 0.05
50:         .Columns(6).SortMode = DataGridViewColumnSortMode.NotSortable
51:         .Columns.Add("dtgrdSala", "Trzba -fon")
52:         .Columns(7).Width = (.Width - 20) * 0.08
53:         .Columns(7).SortMode = DataGridViewColumnSortMode.NotSortable
54:         .Columns.Add("dtgrdSala", "%")
55:         .Columns(8).Width = (.Width - 20) * 0.05
56:         .Columns(8).SortMode = DataGridViewColumnSortMode.NotSortable
57:         .Columns.Add("dtgrdSala", "Odvod")
58:         .Columns(9).Width = (.Width - 20) * 0.08
59:         .Columns(9).SortMode = DataGridViewColumnSortMode.NotSortable
60:         .Columns.Add("dtgrdSala", frmSpracovavanie.pDPH & "% DPH")
61:         .Columns(10).Width = (.Width - 20) * 0.08
62:         .Columns(10).SortMode = DataGridViewColumnSortMode.NotSortable
63:         .Columns.Add("dtgrdSala", "Odvod celkom")
64:         .Columns(11).Width = (.Width - 20) * 0.08
65:         .Columns(11).SortMode = DataGridViewColumnSortMode.NotSortable
66:         .Columns.Add("dtgrdSala", "Trzba -f -o")
67:         .Columns(12).Width = (.Width - 20) * 0.08
68:         .Columns(12).SortMode = DataGridViewColumnSortMode.NotSortable
69:
70:
71:         .ResumeLayout(True)
72:
73:     End With
74:
75:     'Naplni combobox z Filmami (nazvami predstaveni)
76:     With frmSpracovavanie.FilmyBindingSource
77:         cmbFilmy.Items.Clear()
78:         .MoveFirst()
79:         For i = 1 To .Count
80:             cmbFilmy.Items.Add(.Current("film"))
81:             .MoveNext()
82:         Next i
83:     End With
84:
85:     'Naplni combobox z distributormi
86:     With frmSpracovavanie.DistributoriBindingSource
87:         cmbDistributori.Items.Clear()
88:         .MoveFirst()
89:         For i = 1 To .Count
90:             cmbDistributori.Items.Add(.Current("Distributor"))
91:             .MoveNext()
92:         Next i
93:     End With
94:
95:     lblTrzba.Text = 0
96:     lblPr.Text = 0
97:     lblFK.Text = 0
98:     lblNavstev.Text = 0
99:     lblSlev.Text = 0
100:    lblTrzbaF.Text = 0
101:    lblTrzbaFO.Text = 0
102:    lblOdvod.Text = 0
103:    lblOdvodDPH.Text = 0
104:    lblOdvodSpolu.Text = 0
105:
106:    Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub btnZobrazit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnZobrazit.Click
        If frmSpracovavanie.ChybaFilmy Or frmSpracovavanie.ChybaPredstavenia Then MsgBox("Poskodene data. Vystup nemusi byt spravny.", MsgBoxStyle.Exclamation)
        If (CDate(dtpOd.Text) >= CDate("1.1.2011") And CDate(dtpDo.Text) <= CDate("1.1.2011")) Or _
                (CDate(dtpOd.Text) < CDate("1.1.2011") And CDate(dtpDo.Text) > CDate("1.1.2011")) Then
            MsgBox("Obe datumy by mali byt vacsie alebo mensie ako 1.1.2011 kvoli zmene DPH.")
        End If

1:      Dim i, i2, ipos, x As Integer
2:      Dim DatumDO, DatumOD, DatumPr, CasPr As Date
3:      Dim Trzb, FK, TrzbaF, TrzbaFO, Odvod, OdvodCel, DPH As Double
4:      Dim Pr, Div, Slev As Integer
5:      Dim s As String
6:      Dim Tabulka(2000) As Vysledky
7:      Dim Nahradene As Boolean = False
8:
9:      On Error GoTo Chyba
10:
11:     DatumOD = CDate(dtpOd.Text)
12:     DatumDO = CDate(dtpDo.Text)
13:     ipos = 0
14:     i2 = 1
15:
16:     DataGridView1.Rows.Clear()
17:
18:     '**********Vytvori tabulku s potrebnymi udajmi***********
19:     '1. Do tabulky vlozi predstavenia
20:     With frmSpracovavanie.PredstaveniaBindingSource
21:         .MoveFirst()
22:         For i = 1 To .Count
23:             If .Current("NazovFilmu") = cmbFilmy.Text Or cmbFilmy.Text = "" Then
24:                 'Ak triedit aj distributora tak najde v tabulke filmov distributora
25:                 If cmbDistributori.Text <> "" Then frmSpracovavanie.FilmyBindingSource.Position = frmSpracovavanie.FilmyBindingSource.Find("Film", .Current("NazovFilmu"))
26:                 If cmbDistributori.Text = "" Or cmbDistributori.Text = frmSpracovavanie.FilmyBindingSource.Current("Distributor") Then
27:                     If .Current("Predaj") <> 0 Then
28:                         DatumPr = Format(.Current("datum"), "dd.MM.yyyy")
29:                         CasPr = Format(.Current("datum"), "HH:mm")
30:                         If DatumOD <= DatumPr And DatumDO >= DatumPr Then
31:
32:                             s = .Current("Miesta")
33:                             Slev = 0
34:                             For x = 1 To Len(s)
35:                                 If Mid(s, x, 1) = "3" Then Slev += 1
36:                             Next x
37:
38:                             Nahradene = False
39:                             For x = 1 To i2
40:                                 If Tabulka(x).Datum = DatumPr And Tabulka(x).NazovFilmu = .Current("NazovFilmu") Then
41:                                     Tabulka(x).Trzba += .Current("TrzbaPredstavenia")
42:                                     Tabulka(x).Predstaveni += 1
43:                                     Tabulka(x).Divakov += .Current("Predaj")
44:                                     Tabulka(x).Zlav += Slev
45:                                     Nahradene = True
46:                                 End If
47:                             Next x
48:
49:                             If Nahradene = False Then
50:                                 Tabulka(i2).Datum = DatumPr
51:                                 Tabulka(i2).NazovFilmu = .Current("NazovFilmu")
52:                                 Tabulka(i2).Trzba = .Current("TrzbaPredstavenia")
53:                                 Tabulka(i2).Predstaveni = 1
54:                                 Tabulka(i2).Divakov = .Current("Predaj")
55:                                 Tabulka(i2).Zlav += Slev
56:                                 i2 = i2 + 1
57:                             End If
58:                         End If
59:                     Else
60:
61:                         'Ak Predstavenie nebolo odohrane
62:                         DatumPr = Format(.Current("datum"), "dd.MM.yyyy")
63:                         CasPr = Format(.Current("datum"), "HH:mm")
64:                         If DatumOD <= DatumPr And DatumDO >= DatumPr Then
65:                             Nahradene = False
66:                             For x = 1 To i2
67:                                 If Tabulka(x).Datum = DatumPr And Tabulka(x).NazovFilmu = .Current("NazovFilmu") Then Nahradene = True
68:                             Next x
69:                             If Nahradene = False Then
70:                                 Tabulka(i2).Datum = DatumPr
71:                                 Tabulka(i2).NazovFilmu = .Current("NazovFilmu")
72:                                 Tabulka(i2).Trzba = 0
73:                                 Tabulka(i2).Predstaveni = 0
74:                                 Tabulka(i2).Divakov = 0
75:                                 Tabulka(i2).Zlav += 0
76:                                 i2 = i2 + 1
77:                             End If
78:                         End If
79:
80:                     End If
81:                 End If
82:             End If
83:             .MoveNext()
84:         Next i
85:     End With
86:     ipos = i2 - 1
87:     '2. do tabulky si vlozi pozicovne
88:     With frmSpracovavanie.FilmyBindingSource
89:         .MoveFirst()
90:         For i2 = 1 To ipos
91:             i = .Find("Film", Tabulka(i2).NazovFilmu)
92:             .Position = i
93:             Tabulka(i2).PercentPoz = .Current("percentpozicovne")
94:             Tabulka(i2).PevnePoz = .Current("pevnepozicovne")
95:         Next i2
96:     End With
97:
98:     'Vlozi data z tabulky do datagridu
99:     With DataGridView1
100:        For i = 1 To ipos
101:            .Rows.Add()
102:            .Rows(.Rows.Count - 1).Cells(0).Value = Format(Tabulka(i).Datum, "dd.MM.")
103:            .Rows(.Rows.Count - 1).Cells(1).Value = Tabulka(i).NazovFilmu
104:            .Rows(.Rows.Count - 1).Cells(2).Value = Format(Tabulka(i).Trzba, "###0.00")
105:            .Rows(.Rows.Count - 1).Cells(3).Value = Tabulka(i).Predstaveni
106:            .Rows(.Rows.Count - 1).Cells(4).Value = Format(Tabulka(i).Trzba * frmSpracovavanie.PercentOdvoduDoFondov, "###0.##") 'Tabulka(i).Divakov * 0.03
107:            .Rows(.Rows.Count - 1).Cells(5).Value = Tabulka(i).Divakov
108:            .Rows(.Rows.Count - 1).Cells(6).Value = Tabulka(i).Zlav
109:            .Rows(.Rows.Count - 1).Cells(7).Value = Format(Tabulka(i).Trzba - (Tabulka(i).Trzba * frmSpracovavanie.PercentOdvoduDoFondov), "0.00") '- Tabulka(i).Divakov * 0.03, "0.00")
110:            .Rows(.Rows.Count - 1).Cells(8).Value = Tabulka(i).PercentPoz
                If InStr(LCase(Tabulka(i).NazovFilmu), "organizovane") <> 0 Or InStr(LCase(Tabulka(i).NazovFilmu), "organizované") <> 0 Then
112:                .Rows(.Rows.Count - 1).Cells(9).Value = Format(0, "###0.00")
                    .Rows(.Rows.Count - 1).Cells(10).Value = Format(0, "###0.00")
113:                .Rows(.Rows.Count - 1).Cells(11).Value = Format(0, "###0.00")
231:            Else
                    If Tabulka(i).PercentPoz = 0 Then
                        .Rows(.Rows.Count - 1).Cells(9).Value = Tabulka(i).PevnePoz
                    Else
                        .Rows(.Rows.Count - 1).Cells(9).Value = Format((Tabulka(i).Trzba - Tabulka(i).Divakov * 0.03) * Tabulka(i).PercentPoz / 100, "0.00")
                    End If
                    .Rows(.Rows.Count - 1).Cells(10).Value = Format(.Rows(.Rows.Count - 1).Cells(9).Value * frmSpracovavanie.pDPH / 100, "0.00")
                    .Rows(.Rows.Count - 1).Cells(11).Value = Format(CDbl(.Rows(.Rows.Count - 1).Cells(9).Value) + CDbl(.Rows(.Rows.Count - 1).Cells(10).Value), "0.00")
235:            End If
                .Rows(.Rows.Count - 1).Cells(12).Value = Format(CDbl(.Rows(.Rows.Count - 1).Cells(7).Value) - CDbl(.Rows(.Rows.Count - 1).Cells(11).Value), "0.00")
115:        Next i
116:    End With
117:
118:    Trzb = 0
119:    FK = 0
120:    TrzbaF = 0
121:    TrzbaFO = 0
122:    Odvod = 0
123:    Pr = 0
124:    Div = 0
125:    Slev = 0
126:    DPH = 0
127:    OdvodCel = 0
128:
129:    With DataGridView1
130:        For i = 1 To .RowCount
131:            Trzb += CDbl(.Rows(i - 1).Cells(2).Value)
132:            Pr += CDbl(.Rows(i - 1).Cells(3).Value)
133:            FK += CDbl(.Rows(i - 1).Cells(4).Value)
134:            Div += CDbl(.Rows(i - 1).Cells(5).Value)
135:            Slev += CDbl(.Rows(i - 1).Cells(6).Value)
136:            TrzbaF += CDbl(.Rows(i - 1).Cells(7).Value)
137:            Odvod += CDbl(.Rows(i - 1).Cells(9).Value)
138:            DPH += CDbl(.Rows(i - 1).Cells(10).Value)
139:            OdvodCel += CDbl(.Rows(i - 1).Cells(11).Value)
140:            TrzbaFO += CDbl(.Rows(i - 1).Cells(12).Value)
141:        Next i
142:    End With
143:
144:    lblTrzba.Text = Trzb
145:    lblPr.Text = Pr
146:    lblFK.Text = FK
147:    lblNavstev.Text = Div
148:    lblSlev.Text = Slev
149:    lblTrzbaF.Text = TrzbaF
150:    lblTrzbaFO.Text = TrzbaFO
151:    lblOdvod.Text = Odvod
152:    lblOdvodDPH.Text = DPH
153:    lblOdvodSpolu.Text = OdvodCel
154:
155:    Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub btnVytlacit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnVytlacit.Click
1:      If frmSpracovavanie.ChybaFilmy Or frmSpracovavanie.ChybaPredstavenia Then MsgBox("Poskodene data. Vystup nemusi byt spravny.", MsgBoxStyle.Exclamation) : Exit Sub
2:      frmMenu.Vytlacit = ""
3:      frmMenu.Velkost = 8
4:      Dim ODDELOVAC As String = StrDup(97, "─")
5:      Dim X, i As Integer
6:      Dim PomocRetazec, CRC As String
7:      Dim Hlavicka(17) As String
8:      Dim Tab(17) As Integer
9:      Dim DPH As Double
10:     On Error GoTo Chyba
11:
12:     Tab(0) = 0
13:     Tab(1) = 7
14:     Tab(2) = 23
15:     Tab(3) = 32
16:     Tab(4) = 36
17:     Tab(5) = 43
18:     Tab(6) = 48
19:     Tab(7) = 53
20:     Tab(8) = 62
21:     Tab(9) = 66
22:     Tab(10) = 74
23:     Tab(11) = 81
24:     Tab(12) = 89
25:
26:     Hlavicka(0) = "Den"
27:     Hlavicka(1) = "Nazov"
28:     Hlavicka(2) = "Trzba"
29:     Hlavicka(3) = "Pr."
30:     Hlavicka(4) = "Fond"
31:     Hlavicka(5) = "Div."
32:     Hlavicka(6) = "Zlav"
33:     Hlavicka(7) = "Trzba"
34:     Hlavicka(8) = "%"
35:     Hlavicka(9) = "Odvod"
36:     Hlavicka(10) = "DPH"
37:     Hlavicka(11) = "Odvod"
38:     Hlavicka(12) = "Trzba"
39:
40:     'Printer.FontSize = 13
41:     frmMenu.Vytlacit += "Vykaz vysledkov po dnoch - formular" & "  " & frmSpracovavanie.MenoKina & Chr(13) & Chr(10)
42:     frmMenu.Vytlacit += "Od: " & dtpOd.Text & "      Do: " & dtpDo.Text & Chr(13) & Chr(10)
43:     frmMenu.Vytlacit += frmSpracovavanie.MenoKina & " " & frmSpracovavanie.Mesto & Chr(13) & Chr(10)
44:     CRC = Now
45:     frmMenu.Vytlacit += "Dátum vytvorenia: " & CRC & Chr(13) & Chr(10)
46:     frmMenu.Vytlacit += Chr(13) & Chr(10)
47:
48:     'Printer.FontSize = 10
49:     'Hlavicka "│" &
50:     'frmMenu.Vytlacit += "---------1---------2---------3---------4---------5---------6---------7---------8---------9-------" & Chr(13) & Chr(10)
51:     'frmMenu.Vytlacit += "1234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567" & Chr(13) & Chr(10)
52:     frmMenu.Vytlacit += ODDELOVAC & Chr(13) & Chr(10)
53:     PomocRetazec = ""
54:     For i = 0 To 12
55:         PomocRetazec += StrDup((Tab(i) - Len(PomocRetazec)), " ")
56:         PomocRetazec += Hlavicka(i)
57:     Next i
58:     frmMenu.Vytlacit += PomocRetazec & Chr(13) & Chr(10)
59:
60:     Hlavicka(7) = "-f"
61:     Hlavicka(10) = frmSpracovavanie.pDPH & "%"
62:     Hlavicka(11) = "celkom"
63:     Hlavicka(12) = "-f -0"
64:     PomocRetazec = ""
65:     For i = 7 To 12
66:         If (Not i = 8 And Not i = 9) Then
67:             PomocRetazec += StrDup((Tab(i) - Len(PomocRetazec)), " ")
68:             PomocRetazec += Hlavicka(i)
69:         End If
70:     Next i
71:     frmMenu.Vytlacit += PomocRetazec & Chr(13) & Chr(10)
72:
73:     frmMenu.Vytlacit += ODDELOVAC & Chr(13) & Chr(10)
74:     frmMenu.Vytlacit += Chr(13) & Chr(10)
75:
76:     With DataGridView1
77:         For X = 1 To .Rows.Count
78:             PomocRetazec = ""
79:
80:             'Zmena oproti inym tabulkam, odstavec 1 skracuje dlzku nazvu filmu na 15 znakov
81:             For i = 0 To 12
82:                 PomocRetazec += StrDup((Tab(i) - Len(PomocRetazec)), " ")
83:                 If i = 1 Then
84:                     If Len(CStr(.Rows(X - 1).Cells(i).Value)) > 15 Then
85:                         PomocRetazec += LSet(CStr(.Rows(X - 1).Cells(i).Value), 14) & "…"
86:                     Else
87:                         PomocRetazec += CStr(.Rows(X - 1).Cells(i).Value)
88:                     End If
89:                 Else
90:                     PomocRetazec += CStr(.Rows(X - 1).Cells(i).Value)
91:                 End If
92:             Next i
93:             frmMenu.Vytlacit += PomocRetazec & Chr(13) & Chr(10)
94:
95:         Next X
96:     End With
97:
98:     frmMenu.Vytlacit += Chr(13) & Chr(10)
99:     frmMenu.Vytlacit += ODDELOVAC & Chr(13) & Chr(10)
100:
101:    PomocRetazec = ""
102:    PomocRetazec += "Spolu:"
103:    PomocRetazec += StrDup((Tab(2) - Len(PomocRetazec)), " ")
104:    PomocRetazec += lblTrzba.Text
105:    PomocRetazec += StrDup((Tab(3) - Len(PomocRetazec)), " ")
106:    PomocRetazec += lblPr.Text
107:    PomocRetazec += StrDup((Tab(4) - Len(PomocRetazec)), " ")
108:    PomocRetazec += lblFK.Text
109:    PomocRetazec += StrDup((Tab(5) - Len(PomocRetazec)), " ")
110:    PomocRetazec += lblNavstev.Text
111:    PomocRetazec += StrDup((Tab(6) - Len(PomocRetazec)), " ")
112:    PomocRetazec += lblSlev.Text
113:    PomocRetazec += StrDup((Tab(7) - Len(PomocRetazec)), " ")
114:    PomocRetazec += lblTrzbaF.Text
115:    PomocRetazec += StrDup((Tab(9) - Len(PomocRetazec)), " ")
116:    PomocRetazec += lblOdvod.Text
117:    PomocRetazec += StrDup((Tab(10) - Len(PomocRetazec)), " ")
118:    PomocRetazec += lblOdvodDPH.Text
119:    PomocRetazec += StrDup((Tab(11) - Len(PomocRetazec)), " ")
120:    PomocRetazec += lblOdvodSpolu.Text
121:    PomocRetazec += StrDup((Tab(12) - Len(PomocRetazec)), " ")
122:    PomocRetazec += lblTrzbaFO.Text
123:    frmMenu.Vytlacit += PomocRetazec & Chr(13) & Chr(10)
124:
125:    'frmMenu.Vytlacit += Chr(13) & Chr(10)
126:    'frmMenu.Vytlacit += "                                                       ───────────────────────────────" & Chr(13) & Chr(10)
127:    '106:    DPH = CDbl(lblOdvod.Text) * frmSpracovavanie.pDPH / 100
128:    '107:    DPH = Format(DPH, "0.00")
129:    '108:    frmMenu.Vytlacit += "                                                       DPH " & frmSpracovavanie.pDPH & "%                 " & DPH & Chr(13) & Chr(10)
130:    '109:    frmMenu.Vytlacit += "                                                       ───────────────────────────────" & Chr(13) & Chr(10)
131:    '110:    frmMenu.Vytlacit += "                                                       Celkom odvod:           " & DPH + CDbl(lblOdvod.Text) & Chr(13) & Chr(10)
132:
133:    frmMenu.Vytlacit += "CRC: " & Module1.HASHString(CRC & dtpOd.Text & dtpDo.Text & lblOdvodSpolu.Text) & Chr(13) & Chr(10)
134:
135:    'Printer.EndDoc()
136:
137:    frmMenu.Tlacit()

        Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub btnVytlacit_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles btnVytlacit.KeyPress
        frmPredstavenia.btnPridaj_KeyPress(sender, e, Me)
    End Sub
    Private Sub btnZobrazit_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles btnZobrazit.KeyPress
        frmPredstavenia.btnPridaj_KeyPress(sender, e, Me)
    End Sub
    Private Sub cmbDistributori_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmbDistributori.KeyPress
        frmPredstavenia.cmbFilmy_KeyPress(sender, e, Me)
    End Sub

    Private Sub dtpOd_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpOd.ValueChanged
        If dtpOd.Value > dtpDo.Value Then dtpDo.Value = dtpOd.Value
    End Sub
    Private Sub dtpDo_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpDo.ValueChanged
        If dtpOd.Value > dtpDo.Value Then dtpOd.Value = dtpDo.Value
    End Sub
End Class