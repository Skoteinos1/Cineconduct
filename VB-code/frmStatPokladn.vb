Public Class frmStatPokladn

    Dim VizualUprava As Boolean = True

    Private Structure PolozkyVyuctovania
        Dim Pokladnik As String
        Dim DenPredaja As Date
        Dim Suma As Double
        Dim Kusy As Integer
        Dim NormVstup As Integer
        Dim ZlavVstup As Integer
        Dim Storno As Double
        Dim StornoKs As Integer
    End Structure

    Private Sub frmStatPokladn_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
1:      Dim dDate As Date
2:
3:      'Zafarbi
4:      Call mdlColors.Skining(Me)
5:      If VizualUprava Then
6:          Call mdlColors.Sizing(Me)
7:          Me.CenterToParent()
8:          VizualUprava = False
9:      End If
10:
11:     If Month(Today) = 1 Then
12:         dDate = CDate("01." & 12 & "." & Year(Today) - 1)
13:     Else
14:         dDate = CDate("01." & Month(Today) - 1 & "." & Year(Today))
15:     End If
16:     dtpOd.Text = Format(dDate, "dd.MM.yyyy")
17:     dtpDo.Text = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, -1, DateSerial(Year(dDate), Month(dDate) + 1, 1)))
18:
19:     '************ SPRAVA DATA GRIDU*************************
20:     With DataGridView1
21:         .SuspendLayout()
22:         .DataSource = Nothing
23:         .AllowUserToAddRows = False
24:         .AllowUserToDeleteRows = False
25:         .AllowUserToResizeRows = False
26:         .AllowUserToResizeColumns = False
27:
28:         .AllowUserToOrderColumns = True
29:         .SelectionMode = DataGridViewSelectionMode.FullRowSelect
30:         .ReadOnly = True
31:         .MultiSelect = False
32:         .RowHeadersVisible = False
33:         .Columns.Clear()
34:         .Rows.Clear()
35:         ' setup columns
36:         .Columns.Add("dtgrdDatum", "Pokladnik")
37:         .Columns(0).Width = (.Width - 20) * 0.275
38:         '.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
39:         .Columns.Add("dtgrdTyp", "Datum")
40:         .Columns(1).Width = (.Width - 20) * 0.12
41:         '.Columns(1).SortMode = DataGridViewColumnSortMode.NotSortable
42:         .Columns.Add("dtgrdTyp", "Trzba (EUR)")
43:         .Columns(2).Width = (.Width - 20) * 0.1
44:         .Columns(2).SortMode = DataGridViewColumnSortMode.NotSortable
45:         .Columns.Add("dtgrdTyp", "Predane listky (ks)")
46:         .Columns(3).Width = (.Width - 20) * 0.1
47:         .Columns(3).SortMode = DataGridViewColumnSortMode.NotSortable
48:         .Columns.Add("dtgrdTyp", "Obycajne vstupy (Os)")
49:         .Columns(4).Width = (.Width - 20) * 0.1
50:         .Columns(4).SortMode = DataGridViewColumnSortMode.NotSortable
51:         .Columns.Add("dtgrdTyp", "Zlavnene vstupy (Os)")
52:         .Columns(5).Width = (.Width - 20) * 0.1
53:         .Columns(5).SortMode = DataGridViewColumnSortMode.NotSortable
54:         .Columns.Add("dtgrdTyp", "Storno (ks)")
55:         .Columns(6).Width = (.Width - 20) * 0.1
56:         .Columns(6).SortMode = DataGridViewColumnSortMode.NotSortable
57:         .Columns.Add("dtgrdTyp", "Storno (EUR)")
58:         .Columns(7).Width = (.Width - 20) * 0.1
59:         .Columns(7).SortMode = DataGridViewColumnSortMode.NotSortable
60:
61:         .ResumeLayout(True)
62:     End With
63:
64:     lblTrzba.Text = 0
65:     lblListKs.Text = 0
66:     lblOsob.Text = 0
67:     lblStorn.Text = 0
68:     lblStornKs.Text = 0
69:     lblZlav.Text = 0
70:
71:     With frmSpracovavanie.PristupBindingSource
72:         cmbOsoba.Items.Clear()
73:         For i = 0 To .Count - 1
74:             .Position = i
75:             cmbOsoba.Items.Add(Module1.Dekoduj(.Current("Login")))
76:         Next i
77:     End With

    End Sub

    Private Sub btnZobrazit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnZobrazit.Click
1:      On Error GoTo Chyba
        If frmSpracovavanie.ChybaListky Then MsgBox("Poskodene data. Vystup nemusi byt spravny.", MsgBoxStyle.Exclamation)
2:      'Zobrazi dnesnu trzbu, co sa predalo dnes a na akykolvek den
3:      Dim DatumOD, DatumDO, DatumTr As Date
4:      Dim Spolu, Storn As Double
5:      Dim LstKs, Osb, Zlv, StrnKs As Integer
6:      Dim StvListka As String = ""
7:      Dim CelkTrzba(500) As PolozkyVyuctovania
8:      Dim i3, i, ipos As Short
9:      Dim Pokladn As String
10:     Dim nElements As Short
11:     Dim aKey As Object
12:     Dim Nasiel As Boolean
13:     Cursor = Cursors.WaitCursor
        ProgressBar1.Visible = True

14:     DatumOD = CDate(dtpOd.Text)
15:     DatumDO = CDate(dtpDo.Text)
16:     Spolu = 0
17:     Storn = 0
18:     LstKs = 0
19:     Osb = 0
20:     Zlv = 0
21:     StrnKs = 0
22:     ipos = 1
23:     DataGridView1.RowCount = 0
24:
25:     With frmSpracovavanie.ListkyBindingSource
            .Filter = "DenPredaja >= '" & Format(DatumOD, "dd.MM.yyyy") & "' AND DenPredaja <= '" & Format(DatumDO, "dd.MM.yyyy") & "'"
            ProgressBar1.Maximum = .Count
26:         For i3 = 0 To .Count - 1
                ProgressBar1.Value = i3
27:             .Position = i3
28:             DatumTr = CDate(.Current("denpredaja"))
29:             If DatumTr >= DatumOD And DatumTr <= DatumDO Then
30:                 'skusi najst ci ma v zozname trzbu
31:
32:                 aKey = Split(.Current("Stav"), ";")
33:                 nElements = UBound(aKey) - LBound(aKey) + 1
34:                 StvListka = aKey(0)
35:                 If nElements = 1 Then Pokladn = "" Else Pokladn = aKey(1)
36:
37:                 If cmbOsoba.Text = "" Or cmbOsoba.Text = Pokladn Then
38:                     Nasiel = False
39:                     For i = 1 To ipos
40:                         If CelkTrzba(i).Pokladnik = Pokladn Then
41:                             If CDate(CelkTrzba(i).DenPredaja) = DatumTr Then
42:                                 If StvListka = "Predane" Or StvListka = "Bezhotov" Then
43:                                     CelkTrzba(i).Suma += .Current("Suma")
44:                                     CelkTrzba(i).Kusy += 1
45:                                     CelkTrzba(i).ZlavVstup += .Current("Zlavnenych")
46:                                     CelkTrzba(i).NormVstup += .Current("MiestoDo") - .Current("MiestoOd") + 1
47:                                 ElseIf StvListka = "Storno" Or StvListka = "BezhotStorn" Then
48:                                     CelkTrzba(i).Suma -= .Current("Suma")
49:                                     CelkTrzba(i).Kusy -= 1
50:                                     CelkTrzba(i).ZlavVstup -= .Current("Zlavnenych")
51:                                     CelkTrzba(i).Storno += .Current("Suma")
52:                                     CelkTrzba(i).StornoKs += 1
53:                                     CelkTrzba(i).NormVstup -= .Current("MiestoDo") - .Current("MiestoOd") + 1
54:                                 End If
55:                                 Nasiel = True
56:                             End If
57:                         End If
58:                     Next i
59:
60:                     If Nasiel = False Then
61:                         CelkTrzba(ipos).Pokladnik = Pokladn
62:                         CelkTrzba(ipos).DenPredaja = Format(CDate(DatumTr), "dd.MM.yyyy")
63:                         If StvListka = "Predane" Or StvListka = "Bezhotov" Then
64:                             CelkTrzba(ipos).Suma = .Current("Suma")
65:                             CelkTrzba(ipos).Kusy = 1
66:                             CelkTrzba(ipos).ZlavVstup = .Current("Zlavnenych")
67:                             CelkTrzba(ipos).NormVstup = .Current("MiestoDo") - .Current("MiestoOd") + 1
68:                         ElseIf StvListka = "Storno" Or StvListka = "BezhotStorn" Then
69:                             CelkTrzba(ipos).Suma = .Current("Suma") * -1
70:                             CelkTrzba(ipos).Kusy = -1
71:                             CelkTrzba(ipos).ZlavVstup = .Current("Zlavnenych") * -1
72:                             CelkTrzba(ipos).Storno = .Current("Suma")
73:                             CelkTrzba(ipos).StornoKs = 1
74:                             CelkTrzba(ipos).NormVstup = (.Current("MiestoDo") - .Current("MiestoOd") + 1) * -1
75:                         End If
76:                         ipos = ipos + 1
77:                     End If
78:                 End If
79:             End If
80:         Next i3
            .RemoveFilter()
81:     End With
82:
83:     'Naplni Datagrid
84:     With DataGridView1
85:         For i = 1 To ipos - 1
86:             .Rows.Add()
87:             .Rows(.Rows.Count - 1).Cells(0).Value = CelkTrzba(i).Pokladnik
88:             .Rows(.Rows.Count - 1).Cells(1).Value = Format(CDate(CelkTrzba(i).DenPredaja), "dd.MM.yy")
89:             .Rows(.Rows.Count - 1).Cells(2).Value = Format(CelkTrzba(i).Suma, "###0.00")
90:             .Rows(.Rows.Count - 1).Cells(3).Value = CelkTrzba(i).Kusy
91:             .Rows(.Rows.Count - 1).Cells(4).Value = CelkTrzba(i).NormVstup
92:             .Rows(.Rows.Count - 1).Cells(5).Value = CelkTrzba(i).ZlavVstup
93:             .Rows(.Rows.Count - 1).Cells(6).Value = CelkTrzba(i).StornoKs
94:             .Rows(.Rows.Count - 1).Cells(7).Value = Format(CelkTrzba(i).Storno, "###0.00")
95:         Next i
96:         'Spravi sucet dolu
97:         For i = 0 To .Rows.Count - 1
98:             Spolu += .Rows(i).Cells(2).Value
99:             LstKs += .Rows(i).Cells(3).Value
100:            Osb += .Rows(i).Cells(4).Value
101:            Storn += .Rows(i).Cells(7).Value
102:            StrnKs += .Rows(i).Cells(6).Value
103:            Zlv += .Rows(i).Cells(5).Value
104:        Next i
105:    End With
106:
107:    lblTrzba.Text = Format(Spolu, "##0.00")
108:    lblListKs.Text = LstKs
109:    lblOsob.Text = Osb
110:    lblStorn.Text = Format(Storn, "##0.00")
111:    lblStornKs.Text = StrnKs
112:    lblZlav.Text = Zlv
113:
114:    DataGridView1.Sort(DataGridView1.Columns(1), System.ComponentModel.ListSortDirection.Ascending)
115:    'DataGridView1.Sort(DataGridView1.Columns(0), System.ComponentModel.ListSortDirection.Ascending)
116:
        Cursor = Cursors.Default
        ProgressBar1.Visible = False

117:    Exit Sub
        Cursor = Cursors.Default
        ProgressBar1.Visible = False
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub btnTlacit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTlacit.Click
1:      On Error GoTo Chyba
        If frmSpracovavanie.ChybaListky Then MsgBox("Poskodene data. Vystup nemusi byt spravny.", MsgBoxStyle.Exclamation) : Exit Sub
2:      frmMenu.Vytlacit = ""
3:      frmMenu.Velkost = 9
4:      Dim ODDELOVAC As String = StrDup(84, "─")
5:      Dim X, i, i2, ipos As Integer
6:      Dim PomocRetazec As String
7:      Dim Hlavicka(10) As String
8:      Dim Tab(10) As Integer
9:      Dim Spolu(10) As Double
10:     Dim ViacPokladnikov As Boolean = False
11:     Dim PrvyRiadok As Boolean = True
12:     Dim Pokladnici(50) As String
13:     Dim nasiel As Boolean
14:
15:     Tab(0) = 0
16:     Tab(1) = 18
17:     Tab(2) = 28
18:     Tab(3) = 38
19:     Tab(4) = 47
20:     Tab(5) = 57
21:     Tab(6) = 67
22:     Tab(7) = 76
23:
24:     DataGridView1.Sort(DataGridView1.Columns(1), System.ComponentModel.ListSortDirection.Ascending)
25:     'Spravi si zoznam pokladnikov
26:     'Ano je to blbe riesenie na to aby boli datumy v spravnom poradi
27:     ipos = 1
28:     With DataGridView1
29:         For X = 0 To .RowCount - 1
30:             nasiel = False
31:             For i = 1 To ipos
32:                 If Pokladnici(i) = .Rows(X).Cells(0).Value Then
33:                     nasiel = True
34:                     Exit For
35:                 End If
36:             Next i
37:             If nasiel = False Then
38:                 Pokladnici(ipos) = .Rows(X).Cells(0).Value
39:                 ipos += 1
40:             End If
41:         Next X
42:     End With
43:
44:     'Printer.FontSize = 13
45:     frmMenu.Vytlacit += "Vyuctovanie pokladnikov pre " & frmSpracovavanie.MenoKina & Chr(13) & Chr(10)
46:     frmMenu.Vytlacit += Chr(13) & Chr(10)
47:     frmMenu.Vytlacit += "Od: " & dtpOd.Text & "      Do: " & dtpDo.Text & Chr(13) & Chr(10)
48:     frmMenu.Vytlacit += Chr(13) & Chr(10)
49:     frmMenu.Vytlacit += "Dátum vytvorenia: " & Now & Chr(13) & Chr(10)
50:     frmMenu.Vytlacit += Chr(13) & Chr(10)
51:
52:     'Printer.FontSize = 10
53:     'Hlavicka
54:     ' frmMenu.Vytlacit += "---------1---------2---------3---------4---------5---------6---------7---------8---------9---" & Chr(13) & Chr(10)
55:     ' frmMenu.Vytlacit += "123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123" & Chr(13) & Chr(10)
56:     frmMenu.Vytlacit += ODDELOVAC & Chr(13) & Chr(10)
57:
58:     Hlavicka(0) = "Pokladnik"
59:     Hlavicka(1) = "Datum"
60:     Hlavicka(2) = "  Trzba"
61:     Hlavicka(3) = "Predane"
62:     Hlavicka(4) = "Obycajne"
63:     Hlavicka(5) = "Zlavnene"
64:     Hlavicka(6) = "Storno"
65:     Hlavicka(7) = " Storno"
66:     PomocRetazec = ""
67:     For i = 0 To 7
68:         PomocRetazec += StrDup((Tab(i) - Len(PomocRetazec)), " ")
69:         PomocRetazec += Hlavicka(i)
70:     Next i
71:     frmMenu.Vytlacit += PomocRetazec & Chr(13) & Chr(10)
72:
73:     Hlavicka(3) = "listky"
74:     Hlavicka(4) = " vstupy"
75:     Hlavicka(5) = " vstupy"
76:     PomocRetazec = ""
77:     For i = 3 To 5
78:         PomocRetazec += StrDup((Tab(i) - Len(PomocRetazec)), " ")
79:         PomocRetazec += Hlavicka(i)
80:     Next i
81:     frmMenu.Vytlacit += PomocRetazec & Chr(13) & Chr(10)
82:
83:     Hlavicka(2) = "  (EUR)"
84:     Hlavicka(3) = " (ks)"
85:     Hlavicka(4) = "  (Os)"
86:     Hlavicka(5) = "  (Os)"
87:     Hlavicka(6) = " (ks)"
88:     Hlavicka(7) = " (EUR)"
89:     PomocRetazec = ""
90:     For i = 2 To 7
91:         PomocRetazec += StrDup((Tab(i) - Len(PomocRetazec)), " ")
92:         PomocRetazec += Hlavicka(i)
93:     Next i
94:     frmMenu.Vytlacit += PomocRetazec & Chr(13) & Chr(10)
95:
96:     With DataGridView1
97:         .Sort(.Columns(1), System.ComponentModel.ListSortDirection.Ascending)
98:         For i2 = 1 To ipos - 1
99:             PrvyRiadok = True
100:            For X = 0 To 9
101:                Spolu(X) = 0
102:            Next X
103:
104:            frmMenu.Vytlacit += ODDELOVAC & Chr(13) & Chr(10)
105:            For X = 0 To .Rows.Count - 1
106:                If CStr(.Rows(X).Cells(0).Value) = Pokladnici(i2) Then
107:                    PomocRetazec = ""
108:                    If PrvyRiadok Then PomocRetazec += LSet(Pokladnici(i2), 17) : PrvyRiadok = False
109:                    For i = 1 To 7
110:                        PomocRetazec += StrDup((Tab(i) - Len(PomocRetazec)), " ")
111:                        If i = 2 Or i = 7 Then PomocRetazec += StrDup((8 - Len(CStr(.Rows(X).Cells(i).Value))), " ")
112:                        If i = 4 Then PomocRetazec += StrDup((5 - Len(CStr(.Rows(X).Cells(i).Value))), " ")
113:                        If i = 3 Or i = 5 Or i = 6 Then PomocRetazec += StrDup((4 - Len(CStr(.Rows(X).Cells(i).Value))), " ")
114:                        PomocRetazec += CStr(.Rows(X).Cells(i).Value)
115:                        If i > 1 Then Spolu(i) += CDbl(.Rows(X).Cells(i).Value)
116:                    Next i
117:                    frmMenu.Vytlacit += PomocRetazec & Chr(13) & Chr(10)
118:                End If
119:
120:                If X = .RowCount - 1 Then
121:                    frmMenu.Vytlacit += ODDELOVAC & Chr(13) & Chr(10)
122:                    PomocRetazec = "Spolu:"
123:                    For i = 2 To 7
124:                        PomocRetazec += StrDup((Tab(i) - Len(PomocRetazec)), " ")
125:                        If i = 2 Or i = 7 Then
126:                            PomocRetazec += StrDup((8 - Len(Format(Spolu(i), "##0.00"))), " ") & Format(Spolu(i), "##0.00")
127:                        ElseIf i = 4 Then
128:                            PomocRetazec += StrDup((5 - Len(Format(Spolu(i), "##0"))), " ") & Format(Spolu(i), "##0")
129:                        Else
130:                            PomocRetazec += StrDup((4 - Len(Format(Spolu(i), "##0"))), " ") & Format(Spolu(i), "##0")
131:                        End If
132:                    Next i
133:                    frmMenu.Vytlacit += PomocRetazec & Chr(13) & Chr(10)
134:                    frmMenu.Vytlacit += Chr(13) & Chr(10)
135:                    frmMenu.Vytlacit += Chr(13) & Chr(10)
136:                End If
137:            Next X
138:        Next i2
139:
140:        If ipos > 2 Then
141:            frmMenu.Vytlacit += Chr(13) & Chr(10)
142:            PomocRetazec = "Spolu za vsetkych:"
143:            PomocRetazec += StrDup((Tab(2) - Len(PomocRetazec)), " ")
144:            PomocRetazec += StrDup((8 - Len(Format(CDbl(lblTrzba.Text), "##0.00"))), " ") & Format(CDbl(lblTrzba.Text), "##0.00")
145:            PomocRetazec += StrDup((Tab(3) - Len(PomocRetazec)), " ")
146:            PomocRetazec += StrDup((4 - Len(Format(CDbl(lblListKs.Text), "##0"))), " ") & Format(CDbl(lblListKs.Text), "##0")
147:            PomocRetazec += StrDup((Tab(4) - Len(PomocRetazec)), " ")
148:            PomocRetazec += StrDup((5 - Len(Format(CDbl(lblOsob.Text), "##0"))), " ") & Format(CDbl(lblOsob.Text), "##0")
149:            PomocRetazec += StrDup((Tab(5) - Len(PomocRetazec)), " ")
150:            PomocRetazec += StrDup((4 - Len(Format(CDbl(lblZlav.Text), "##0"))), " ") & Format(CDbl(lblZlav.Text), "##0")
151:            PomocRetazec += StrDup((Tab(6) - Len(PomocRetazec)), " ")
152:            PomocRetazec += StrDup((4 - Len(Format(CDbl(lblStornKs.Text), "##0"))), " ") & Format(CDbl(lblStornKs.Text), "##0")
153:            PomocRetazec += StrDup((Tab(7) - Len(PomocRetazec)), " ")
154:            PomocRetazec += StrDup((8 - Len(Format(CDbl(lblStorn.Text), "##0.00"))), " ") & Format(CDbl(lblStorn.Text), "##0.00")
155:
156:            frmMenu.Vytlacit += PomocRetazec & Chr(13) & Chr(10)
157:        End If
158:    End With
159:
160:    frmMenu.Tlacit()
161:
162:    Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub cmbOsoba_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmbOsoba.KeyPress
        frmPredstavenia.cmbFilmy_KeyPress(sender, e, Me)
    End Sub
    Private Sub btnTlacit_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles btnTlacit.KeyPress
        frmPredstavenia.btnPridaj_KeyPress(sender, e, Me)
    End Sub
    Private Sub btnZobrazit_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles btnZobrazit.KeyPress
        frmPredstavenia.btnPridaj_KeyPress(sender, e, Me)
    End Sub

    Private Sub dtpOd_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpOd.ValueChanged
        If dtpOd.Value > dtpDo.Value Then dtpDo.Value = dtpOd.Value
    End Sub
    Private Sub dtpDo_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpDo.ValueChanged
        If dtpOd.Value > dtpDo.Value Then dtpOd.Value = dtpDo.Value
    End Sub
End Class