Public Class frmStatRocVysl

    Dim VizualUprava As Boolean = True

    Private Structure Mesiace
        Dim Predstaveni As Short
        Dim Navstev As Short
        Dim Trzieb As Double
        Dim PocetStorna As Short
    End Structure

    Private Sub frmStatRocVysl_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
1:      On Error GoTo Chyba
2:      'Zafarbi
3:      Call mdlColors.Skining(Me)
4:      If VizualUprava Then
5:          Call mdlColors.Sizing(Me)
6:          Me.CenterToParent()
7:          VizualUprava = False
8:      End If
9:      '************ SPRAVA DATA GRIDU*************************
10:     With DataGridView1
11:         .SuspendLayout()
12:         .DataSource = Nothing
13:         .AllowUserToAddRows = False
14:         .AllowUserToDeleteRows = False
15:         .AllowUserToResizeRows = False
16:         .AllowUserToResizeColumns = False
17:
18:         .AllowUserToOrderColumns = True
19:         .SelectionMode = DataGridViewSelectionMode.FullRowSelect
20:         .ReadOnly = True
21:         .MultiSelect = False
22:         .RowHeadersVisible = False
23:         .Columns.Clear()
24:         .Rows.Clear()
25:         ' setup columns
26:
27:         .Columns.Add("dtgrdDatum", "Mesiac")
28:         .Columns(0).Width = (.Width - 20) / 8
29:         .Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
30:         .Columns.Add("dtgrdDen", "Predstavení")
31:         .Columns(1).Width = (.Width - 20) / 8
32:         .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
33:         .Columns(1).SortMode = DataGridViewColumnSortMode.NotSortable
34:         .Columns.Add("dtgrdCas", "Návstev.")
35:         .Columns(2).Width = (.Width - 20) / 8
36:         .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
37:         .Columns(2).SortMode = DataGridViewColumnSortMode.NotSortable
38:         .Columns.Add("dtgrdFilm", "Trzba (EUR)")
39:         .Columns(3).Width = (.Width - 20) / 8
40:         .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
41:         .Columns(3).SortMode = DataGridViewColumnSortMode.NotSortable
42:         .Columns.Add("dtgrdCena", "Priemer návstev.")
43:         .Columns(4).Width = (.Width - 20) / 8
44:         .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
45:         .Columns(4).SortMode = DataGridViewColumnSortMode.NotSortable
46:         .Columns.Add("dtgrdCenaZlav", "Priemerná trzba (EUR)")
47:         .Columns(5).Width = (.Width - 20) / 8
48:         .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
49:         .Columns(5).SortMode = DataGridViewColumnSortMode.NotSortable
50:         .Columns.Add("dtgrdSala2", "Priemerné vstupné (EUR)")
51:         .Columns(6).Width = (.Width - 20) / 8
52:         .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
53:         .Columns(6).SortMode = DataGridViewColumnSortMode.NotSortable
54:         .Columns.Add("dtgrdSala3", "Pocet Storna (EUR)")
55:         .Columns(7).Width = (.Width - 20) / 8
56:         .Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
57:         .Columns(7).SortMode = DataGridViewColumnSortMode.NotSortable
58:
59:         .ResumeLayout(True)
60:     End With
61:
62:     txtRok.Text = Year(Today) - 1
63:
        Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub btnVytlacit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnVytlacit.Click
1:      On Error GoTo Chyba

        If frmSpracovavanie.ChybaPredstavenia Then MsgBox("Poskodene data. Vystup nemusi byt spravny.", MsgBoxStyle.Exclamation) : Exit Sub

2:      If DataGridView1.RowCount < 2 Then MsgBox("Vyberte rok a kliknite na Zobrazit.") : Exit Sub
3:      frmMenu.Vytlacit = ""
4:      frmMenu.Velkost = 9
5:      Dim ODDELOVAC As String = StrDup(83, "─")
6:      Dim X, i As Integer
7:      Dim PomocRetazec As String
8:      Dim Hlavicka(10) As String
9:      Dim Tab(10) As Integer
        Dim Medzer(10) As Integer
10:     Dim CRC As String
11:
        Medzer(0) = 1
        Medzer(1) = 5
        Medzer(2) = 7
        Medzer(3) = 10
        Medzer(4) = 7
        Medzer(5) = 8
        Medzer(6) = 7
        Medzer(7) = 5

12:     Tab(0) = 0
13:     Tab(1) = 13
14:     Tab(2) = 25
15:     Tab(3) = 35
16:     Tab(4) = 46
17:     Tab(5) = 55
18:     Tab(6) = 66
19:     Tab(7) = 77
20:
21:     Hlavicka(0) = "Mesiac"
22:     Hlavicka(1) = "Predstaveni"
23:     Hlavicka(2) = "Návstev."
24:     Hlavicka(3) = "  Trzba"
25:     Hlavicka(4) = "Priemer"
26:     Hlavicka(5) = "Priemerná"
27:     Hlavicka(6) = "Priemerné"
28:     Hlavicka(7) = "Pocet"
29:
30:     'Printer.FontSize = 13
31:     frmMenu.Vytlacit += "Rocný statistický výkaz kina: " & frmSpracovavanie.MenoKina & Chr(13) & Chr(10)
32:     frmMenu.Vytlacit += "Za rok: " & txtRok.Text & Chr(13) & Chr(10)
33:     CRC = Now
34:     frmMenu.Vytlacit += "Vytvorené: " & CRC & Chr(13) & Chr(10)
35:     frmMenu.Vytlacit += Chr(13) & Chr(10)
36:
37:     'Printer.FontSize = 10
38:     'Hlavicka
39:     'frmMenu.Vytlacit += "---------1---------2---------3---------4---------5---------6---------7---------8---------9---" & Chr(13) & Chr(10)
40:     'frmMenu.Vytlacit += "123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123" & Chr(13) & Chr(10)
41:     frmMenu.Vytlacit += ODDELOVAC & Chr(13) & Chr(10)
42:     PomocRetazec = ""
43:     For i = 0 To 7
44:         PomocRetazec += StrDup((Tab(i) - Len(PomocRetazec)), " ")
45:         PomocRetazec += Hlavicka(i)
46:     Next i
47:     frmMenu.Vytlacit += PomocRetazec & Chr(13) & Chr(10)
48:
49:     Hlavicka(4) = "návst."
50:     Hlavicka(5) = "  trzba"
51:     Hlavicka(6) = " vstupné"
52:     Hlavicka(7) = "storna"
53:     PomocRetazec = ""
54:     For i = 4 To 7
55:         ' If i = 3 Or i = 4 Or i = 7 Then
56:         PomocRetazec += StrDup((Tab(i) - Len(PomocRetazec)), " ")
57:         PomocRetazec += Hlavicka(i)
58:         'End If
59:     Next i
60:     frmMenu.Vytlacit += PomocRetazec & Chr(13) & Chr(10)
61:
62:     Hlavicka(3) = "  (EUR)"
63:     Hlavicka(5) = "  (EUR)"
64:     Hlavicka(6) = "  (EUR)"
65:     Hlavicka(7) = "(EUR)"
66:     PomocRetazec = ""
67:     For i = 3 To 7
68:         If i <> 4 Then
69:             PomocRetazec += StrDup((Tab(i) - Len(PomocRetazec)), " ")
70:             PomocRetazec += Hlavicka(i)
71:         End If
72:     Next i
73:     frmMenu.Vytlacit += PomocRetazec & Chr(13) & Chr(10)
74:
75:     frmMenu.Vytlacit += ODDELOVAC & Chr(13) & Chr(10)
76:     frmMenu.Vytlacit += Chr(13) & Chr(10)
77:
78:     With DataGridView1
79:         For X = 1 To .Rows.Count
80:             PomocRetazec = ""
81:             For i = 0 To 7
82:                 PomocRetazec += StrDup((Tab(i) - Len(PomocRetazec)), " ")
                    If i <> 0 Then PomocRetazec += StrDup((Medzer(i) - Len(CStr(.Rows(X - 1).Cells(i).Value))), " ")
83:                 PomocRetazec += CStr(.Rows(X - 1).Cells(i).Value)
84:             Next i
85:             frmMenu.Vytlacit += PomocRetazec & Chr(13) & Chr(10)
86:         Next X
87:     End With
88:
89:     frmMenu.Vytlacit += Chr(13) & Chr(10)
90:     frmMenu.Vytlacit += ODDELOVAC & Chr(13) & Chr(10)
91:     frmMenu.Vytlacit += Chr(13) & Chr(10)
92:
94:     PomocRetazec = "Spolu:"
95:     PomocRetazec += StrDup((Tab(1) - Len(PomocRetazec)), " ")
        PomocRetazec += StrDup((Medzer(1) - Len(lblPredst.Text)), " ")
96:     PomocRetazec += lblPredst.Text
97:     PomocRetazec += StrDup((Tab(2) - Len(PomocRetazec)), " ")
        PomocRetazec += StrDup((Medzer(2) - Len(lblNavst.Text)), " ")
98:     PomocRetazec += lblNavst.Text
99:     PomocRetazec += StrDup((Tab(3) - Len(PomocRetazec)), " ")
        PomocRetazec += StrDup((Medzer(3) - Len(lblTrzb.Text)), " ")
100:    PomocRetazec += lblTrzb.Text
101:    PomocRetazec += StrDup((Tab(4) - Len(PomocRetazec)), " ")
        PomocRetazec += StrDup((Medzer(4) - Len(lblPNav.Text)), " ")
102:    PomocRetazec += lblPNav.Text
103:    PomocRetazec += StrDup((Tab(5) - Len(PomocRetazec)), " ")
        PomocRetazec += StrDup((Medzer(5) - Len(lblPTrzb.Text)), " ")
104:    PomocRetazec += lblPTrzb.Text
105:    PomocRetazec += StrDup((Tab(6) - Len(PomocRetazec)), " ")
        PomocRetazec += StrDup((Medzer(6) - Len(lblPVstup.Text)), " ")
106:    PomocRetazec += lblPVstup.Text
107:    PomocRetazec += StrDup((Tab(7) - Len(PomocRetazec)), " ")
        PomocRetazec += StrDup((Medzer(7) - Len(lblStorn.Text)), " ")
108:    PomocRetazec += lblStorn.Text
109:    frmMenu.Vytlacit += PomocRetazec & Chr(13) & Chr(10)
110:
111:    frmMenu.Vytlacit += Chr(13) & Chr(10)
112:    frmMenu.Vytlacit += "CRC: " & Module1.HASHString("roc" & CRC & txtRok.Text & lblTrzb.Text) & Chr(13) & Chr(10)
113:
114:    frmMenu.Tlacit()
115:
        Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub btnZobrazit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnZobrazit.Click
1:      On Error GoTo Chyba
2:      Dim i, X As Integer
3:      Dim RocnyVysledok(13) As Mesiace
4:      Dim DatumDO, DatumOD, dDate As Date
5:
        If frmSpracovavanie.ChybaPredstavenia Then MsgBox("Poskodene data. Vystup nemusi byt spravny.", MsgBoxStyle.Exclamation)

6:      lblPredst.Text = 0
7:      lblNavst.Text = 0
8:      lblTrzb.Text = 0
9:      lblPNav.Text = 0
10:     lblPTrzb.Text = 0
11:     lblPVstup.Text = 0
12:     lblStorn.Text = 0
13:     DataGridView1.Rows.Clear()
14:     ProgressBar1.Visible = True
15:     ProgressBar1.Maximum = frmSpracovavanie.PredstaveniaBindingSource.Count * 12
16:
17:     For X = 1 To 12
18:         DatumOD = Format(CDate("01." & X & "." & txtRok.Text), "dd.MM.yyyy")
19:         dDate = CDate(DatumOD)
20:         ' DatumDO = DateAdd(Microsoft.VisualBasic.DateInterval.Day, -1, DateSerial(Year(dDate), Month(dDate) + 1, 1))
21:         DatumDO = DateAdd(Microsoft.VisualBasic.DateInterval.Day, 0, DateSerial(Year(dDate), Month(dDate) + 1, 1))
22:
23:         With frmSpracovavanie.PredstaveniaBindingSource
24:             For i = 0 To .Count - 1
25:                 .Position = i
26:                 ProgressBar1.Value = i + (.Count * (X - 1))
27:                 If .Current("datum") >= DatumOD And .Current("datum") <= DatumDO Then
28:                     If .Current("predaj") <> 0 Then
29:                         RocnyVysledok(X).Predstaveni = RocnyVysledok(X).Predstaveni + 1
30:                         RocnyVysledok(X).Navstev = RocnyVysledok(X).Navstev + .Current("predaj")
31:                         RocnyVysledok(X).Trzieb = RocnyVysledok(X).Trzieb + .Current("trzbapredstavenia")
32:                     End If
33:                 End If
34:             Next i
35:         End With
36:
37:         'Obsahuje iba velke storno
38:         With frmSpracovavanie.StornoPredstaveniBindingSource
39:             For i = 0 To .Count - 1
40:                 .Position = i
41:                 If .Current("datum") >= DatumOD And .Current("datum") <= DatumDO Then
42:                     RocnyVysledok(X).PocetStorna += .Current("Osob")
43:                 End If
44:             Next i
45:         End With
46:
47:         With DataGridView1
48:             .Rows.Add()
49:             .Rows(.Rows.Count - 1).Cells(0).Value = Format(X, "00") & " / " & txtRok.Text
50:             .Rows(.Rows.Count - 1).Cells(1).Value = RocnyVysledok(X).Predstaveni
51:             .Rows(.Rows.Count - 1).Cells(2).Value = RocnyVysledok(X).Navstev
52:             .Rows(.Rows.Count - 1).Cells(3).Value = Format(RocnyVysledok(X).Trzieb, "###0.00")
53:
54:             If RocnyVysledok(X).Predstaveni = 0 Then
55:                 .Rows(.Rows.Count - 1).Cells(4).Value = 0
56:                 .Rows(.Rows.Count - 1).Cells(5).Value = 0
57:             Else
58:                 .Rows(.Rows.Count - 1).Cells(4).Value = Format((RocnyVysledok(X).Navstev / RocnyVysledok(X).Predstaveni), "####0")
59:                 '.Rows(.Rows.Count - 1).Cells(4).Value =  Format(.get_TextMatrix(X, 4), "####0"))
60:                 .Rows(.Rows.Count - 1).Cells(5).Value = Format((RocnyVysledok(X).Trzieb / RocnyVysledok(X).Predstaveni), "####0.00")
61:                 '.set_TextMatrix(X, 5, Format(.get_TextMatrix(X, 5), "####0.00"))
62:             End If
63:
64:             If RocnyVysledok(X).Navstev = 0 Then
65:                 .Rows(.Rows.Count - 1).Cells(6).Value = 0
66:             Else
67:                 .Rows(.Rows.Count - 1).Cells(6).Value = Format((RocnyVysledok(X).Trzieb / RocnyVysledok(X).Navstev), "####0.00")
68:                 '.set_TextMatrix(X, 6, Format(.get_TextMatrix(X, 6), "####0.00"))
69:             End If
70:
71:             .Rows(.Rows.Count - 1).Cells(7).Value = RocnyVysledok(X).PocetStorna
72:
73:             lblPredst.Text = CDbl(lblPredst.Text) + RocnyVysledok(X).Predstaveni
74:             lblNavst.Text = CDbl(lblNavst.Text) + RocnyVysledok(X).Navstev
75:             lblTrzb.Text = CDbl(lblTrzb.Text) + RocnyVysledok(X).Trzieb
76:             lblPNav.Text = CDbl(lblPNav.Text) + RocnyVysledok(X).Navstev
77:             lblPTrzb.Text = CDbl(lblPTrzb.Text) + RocnyVysledok(X).Trzieb
78:             lblPVstup.Text = CDbl(lblPVstup.Text) + RocnyVysledok(X).Trzieb
79:             lblStorn.Text = CDbl(lblStorn.Text) + RocnyVysledok(X).PocetStorna
80:         End With
81:     Next X
82:
83:     ProgressBar1.Visible = False
84:
85:     lblTrzb.Text = Format(CDbl(lblTrzb.Text), "###0.00")
86:     lblPNav.Text = Format(CDbl(lblPNav.Text) / CDbl(lblPredst.Text), "##0")
87:     lblPTrzb.Text = Format(CDbl(lblPTrzb.Text) / CDbl(lblPredst.Text), "###0.00")
88:     lblPVstup.Text = Format(CDbl(lblPVstup.Text) / CDbl(lblNavst.Text), "###0.00")
89:
        Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub txtRok_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtRok.KeyPress
        frmPredstavenia.txtDatum_KeyPress(sender, e, Me)
    End Sub

    Private Sub btnZobrazit_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles btnZobrazit.KeyPress
        frmPredstavenia.btnPridaj_KeyPress(sender, e, Me)
    End Sub

    Private Sub btnVytlacit_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles btnVytlacit.KeyPress
        frmPredstavenia.btnPridaj_KeyPress(sender, e, Me)
    End Sub

End Class