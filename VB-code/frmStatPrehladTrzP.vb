Public Class frmStatPrehladTrzP

    Dim VizualUprava As Boolean = True

    Private Structure PolozkyTrzby
        Dim DenPred As Date
        Dim StavTrz As String
        Dim DtPredst As Date
        Dim CsPredst As Date
        Dim Osob As Short
        Dim Zlavnen As Short
        Dim Suma As Double
    End Structure

    Private Sub frmStatPrehladTrz_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
1:      If frmSpracovavanie.ChybaListky Then MsgBox("Poskodene data. Vystup nemusi byt spravny.", MsgBoxStyle.Information)
2:      'Zafarbi
3:      Call mdlColors.Skining(Me)
4:      If VizualUprava Then
5:          Call mdlColors.Sizing(Me)
6:          Me.CenterToParent()
7:          VizualUprava = False
8:      End If
9:
10:     '************ SPRAVA DATA GRIDU*************************
11:     With DataGridView1
12:         .SuspendLayout()
13:         .DataSource = Nothing
14:         .AllowUserToAddRows = False
15:         .AllowUserToDeleteRows = False
16:         .AllowUserToResizeRows = False
17:         .AllowUserToResizeColumns = False
18:
19:         .AllowUserToOrderColumns = True
20:         .SelectionMode = DataGridViewSelectionMode.FullRowSelect
21:         .ReadOnly = True
22:         .MultiSelect = False
23:         .RowHeadersVisible = False
24:         .Columns.Clear()
25:         .Rows.Clear()
26:         ' setup columns
27:         .Columns.Add("dtgrdDatum", "Dátum")
28:         .Columns(0).Width = (.Width - 20) * 0.14
29:         .Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
30:         .Columns.Add("dtgrdTyp", "Typ")
31:         .Columns(1).Width = (.Width - 20) * 0.14
32:         '.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
33:         .Columns(1).SortMode = DataGridViewColumnSortMode.NotSortable
34:         .Columns.Add("dtgrdDen", "Den predstavenia")
35:         .Columns(2).Width = (.Width - 20) * 0.15
36:         '.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
37:         .Columns(2).SortMode = DataGridViewColumnSortMode.NotSortable
38:         .Columns.Add("dtgrdCas", "Cas predstavenia")
39:         .Columns(3).Width = (.Width - 20) * 0.15
40:         '.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
41:         .Columns(3).SortMode = DataGridViewColumnSortMode.NotSortable
42:         .Columns.Add("dtgrdOsob", "Osob")
43:         .Columns(4).Width = (.Width - 20) * 0.14
44:         .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
45:         .Columns(4).SortMode = DataGridViewColumnSortMode.NotSortable
46:         .Columns.Add("dtgrdZlavnenych", "Zlavnenych")
47:         .Columns(5).Width = (.Width - 20) * 0.14
48:         .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
49:         .Columns(5).SortMode = DataGridViewColumnSortMode.NotSortable
50:         .Columns.Add("dtgrdSuma", "Suma (EUR)")
51:         .Columns(6).Width = (.Width - 20) * 0.14
52:         .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
53:         .Columns(6).SortMode = DataGridViewColumnSortMode.NotSortable
54:         .ResumeLayout(True)
55:     End With
56:
57:     'zobrazi dnesny predaj
58:     Call btnTrzbaDna_Click(sender, e)
    End Sub

    Private Sub btnTrzbaDna_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTrzbaDna.Click
1:      On Error GoTo Chyba
2:      'Zobrazi dnesnu trzbu, co sa predalo dnes a na akykolvek den
3:      Dim DatumOD As Date
4:      Dim DatPreds, DatumTr, CasPreds As Date
5:      Dim Spolu, SpoluS As Double
6:      Dim ZlavP, OsobP, OsobS, ZlavS As Short
7:      Dim StvListka As String = ""
8:      Dim CelkTrzba(100) As PolozkyTrzby
9:      Dim i3, i, ipos As Short
10:     Dim aKey As Object
11:     Dim Nasiel As Boolean
12:
13:     DatumOD = CDate(dtpOd.Text)
14:     Spolu = 0
15:     ipos = 1
16:     DataGridView1.RowCount = 0
17:
18:     With frmSpracovavanie.ListkyBindingSource
19:         .Filter = "DenPredaja = '" & Format(DatumOD, "dd.MM.yyyy") & "'"
20:         For i3 = 0 To .Count - 1
                .Position = i3
21:             DatumTr = CDate(.Current("denpredaja"))
22:             If DatumTr = DatumOD Then
23:                 'skusi najst ci ma v zozname predstavenie
24:                 DatPreds = CDate(.Current("datumpredst"))
25:                 CasPreds = CDate(.Current("CasPredst"))
26:                 aKey = Split(.Current("Stav"), ";")
27:                 StvListka = aKey(0)
28:
29:                 Nasiel = False
30:                 For i = 1 To ipos
31:                     If CDate(CelkTrzba(i).DenPred) = DatumTr Then
32:                         If CelkTrzba(i).DtPredst = DatPreds And CelkTrzba(i).CsPredst = CasPreds And CelkTrzba(i).StavTrz = StvListka Then
33:                             CelkTrzba(i).Osob = CelkTrzba(i).Osob + System.Math.Abs(.Current("miestood") - .Current("miestodo")) + 1
34:                             CelkTrzba(i).Zlavnen = CelkTrzba(i).Zlavnen + .Current("zlavnenych")
35:                             CelkTrzba(i).Suma = CelkTrzba(i).Suma + .Current("Suma")
36:                             Nasiel = True
37:                         End If
38:                     End If
39:                 Next i
40:
41:                 If Nasiel = False Then
42:                     CelkTrzba(ipos).CsPredst = CasPreds
43:                     CelkTrzba(ipos).DtPredst = DatPreds
44:                     CelkTrzba(ipos).DenPred = DatumTr
45:                     CelkTrzba(ipos).StavTrz = StvListka
46:                     CelkTrzba(ipos).Zlavnen = .Current("zlavnenych")
47:                     CelkTrzba(ipos).Suma = .Current("Suma")
48:                     CelkTrzba(ipos).Osob = System.Math.Abs(.Current("miestood") - .Current("miestodo")) + 1
49:                     ipos = ipos + 1
50:                 End If
52:             End If
54:         Next i3
            .RemoveFilter()
55:     End With
56:
57:     'Naplni Datagrid
58:     With DataGridView1
59:         For i = 1 To ipos - 1
60:             .Rows.Add()
61:             .Rows(.Rows.Count - 1).Cells(0).Value = Format(CelkTrzba(i).DenPred, "dd.MM.yyyy")
62:             .Rows(.Rows.Count - 1).Cells(1).Value = CelkTrzba(i).StavTrz
63:             .Rows(.Rows.Count - 1).Cells(2).Value = Format(CelkTrzba(i).DtPredst, "dd.MM.yyyy")
64:             .Rows(.Rows.Count - 1).Cells(3).Value = Format(CelkTrzba(i).CsPredst, "HH:mm")
65:             .Rows(.Rows.Count - 1).Cells(4).Value = CelkTrzba(i).Osob
66:             .Rows(.Rows.Count - 1).Cells(5).Value = CelkTrzba(i).Zlavnen
67:             .Rows(.Rows.Count - 1).Cells(6).Value = CelkTrzba(i).Suma
68:         Next i
69:         'Spravi sucet dolu
70:         For i = 1 To .Rows.Count
71:             If .Rows(i - 1).Cells(1).Value = "Predane" Or .Rows(i - 1).Cells(1).Value = "Bezhotov" Then
72:                 OsobP = OsobP + .Rows(i - 1).Cells(4).Value
73:                 ZlavP = ZlavP + .Rows(i - 1).Cells(5).Value
74:                 Spolu = Spolu + .Rows(i - 1).Cells(6).Value
75:             ElseIf .Rows(i - 1).Cells(1).Value = "Storno" Or .Rows(i - 1).Cells(1).Value = "BezhotStorn" Then
76:                 OsobS = OsobS + .Rows(i - 1).Cells(4).Value
77:                 ZlavS = ZlavS + .Rows(i - 1).Cells(5).Value
78:                 SpoluS = SpoluS + .Rows(i - 1).Cells(6).Value
79:             End If
80:         Next i
81:     End With
82:
83:     lblSpoluP.Text = CStr(Spolu)
84:     lblSpoluS.Text = CStr(SpoluS)
85:     lblZlavP.Text = CStr(ZlavP)
86:     lblZlavS.Text = CStr(ZlavS)
87:     lblOsobP.Text = CStr(OsobP)
88:     lblOsobS.Text = CStr(OsobS)
89:
        Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub btnTrzbaNaDen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTrzbaNaDen.Click
1:      'Zobrazuje trzbu pripadajucu na dnesny den, NIE vytvorenu v dnesny den
2:      On Error GoTo Chyba
3:      Dim DatumOD As Date
4:      Dim DatPreds, DatumTr, CasPreds As Date
5:      Dim Spolu, SpoluS As Double
6:      Dim ZlavP, OsobP, OsobS, ZlavS As Short
7:      Dim StvListka As String = ""
8:      Dim CelkTrzba(500) As PolozkyTrzby
9:      Dim i2, i, ipos As Short
10:     Dim aKey As Object
11:     Dim Nasiel As Boolean
12:
13:     'If txtOd.Text = "" Then DatumOD = Today Else 
14:     DatumOD = CDate(dtpOd.Text)
15:     Spolu = 0
16:     ipos = 1
17:     DataGridView1.RowCount = 0
18:
19:     With frmSpracovavanie.ListkyBindingSource
            .Filter = "DatumPredst = '" & Format(DatumOD, "dd.MM.yyyy") & "'"
21:         For i3 = 0 To .Count - 1
                .Position = i3
22:             DatPreds = CDate(.Current("datumpredst"))
23:             If DatPreds = DatumOD Then
24:                 'skusi najst ci ma v zozname predstavenie
25:                 CasPreds = CDate(.Current("CasPredst"))
26:                 aKey = Split(.Current("Stav"), ";")
27:                 StvListka = aKey(0)
28:
29:                 Nasiel = False
30:                 For i = 1 To ipos
31:                     If CelkTrzba(i).DtPredst = DatPreds And CelkTrzba(i).CsPredst = CasPreds And CelkTrzba(i).StavTrz = StvListka Then
32:                         CelkTrzba(i).Osob = CelkTrzba(i).Osob + System.Math.Abs(.Current("miestood") - .Current("miestodo")) + 1
33:                         CelkTrzba(i).Zlavnen = CelkTrzba(i).Zlavnen + .Current("zlavnenych")
34:                         CelkTrzba(i).Suma = CelkTrzba(i).Suma + .Current("Suma")
35:                         Nasiel = True
36:                     End If
37:                 Next i
38:
39:                 If Nasiel = False Then
40:                     CelkTrzba(ipos).CsPredst = CasPreds
41:                     CelkTrzba(ipos).DtPredst = DatPreds
42:                     'CelkTrzba(ipos).DenPred = DatumTr
43:                     CelkTrzba(ipos).StavTrz = StvListka
44:                     CelkTrzba(ipos).Zlavnen = .Current("zlavnenych")
45:                     CelkTrzba(ipos).Suma = .Current("Suma")
46:                     CelkTrzba(ipos).Osob = System.Math.Abs(.Current("miestood") - .Current("miestodo")) + 1
47:                     ipos = ipos + 1
48:                 End If
49:             End If
51:         Next i3
            .RemoveFilter()
52:     End With
53:
54:     'Naplni Datagrid
55:     With DataGridView1
56:         For i = 1 To ipos - 1
57:             .Rows.Add()
58:             '.Rows(.Rows.Count - 1).Cells(0).Value = Format(CelkTrzba(i).DenPred, "dd.MM.yyyy")
59:             .Rows(.Rows.Count - 1).Cells(1).Value = CelkTrzba(i).StavTrz
60:             .Rows(.Rows.Count - 1).Cells(2).Value = Format(CelkTrzba(i).DtPredst, "dd.MM.yyyy")
61:             .Rows(.Rows.Count - 1).Cells(3).Value = Format(CelkTrzba(i).CsPredst, "HH:mm")
62:             .Rows(.Rows.Count - 1).Cells(4).Value = CelkTrzba(i).Osob
63:             .Rows(.Rows.Count - 1).Cells(5).Value = CelkTrzba(i).Zlavnen
64:             .Rows(.Rows.Count - 1).Cells(6).Value = CelkTrzba(i).Suma
65:         Next i
66:         'Spravi sucet dolu
67:         For i = 1 To .Rows.Count
68:             If .Rows(i - 1).Cells(1).Value = "Predane" Or .Rows(i - 1).Cells(1).Value = "Bezhotov" Then
69:                 OsobP = OsobP + .Rows(i - 1).Cells(4).Value
70:                 ZlavP = ZlavP + .Rows(i - 1).Cells(5).Value
71:                 Spolu = Spolu + .Rows(i - 1).Cells(6).Value
72:             ElseIf .Rows(i - 1).Cells(1).Value = "Storno" Or .Rows(i - 1).Cells(1).Value = "BezhotStorn" Then
73:                 OsobS = OsobS + .Rows(i - 1).Cells(4).Value
74:                 ZlavS = ZlavS + .Rows(i - 1).Cells(5).Value
75:                 SpoluS = SpoluS + .Rows(i - 1).Cells(6).Value
76:             End If
77:         Next i
78:     End With
79:
80:     lblSpoluP.Text = CStr(Spolu)
81:     lblSpoluS.Text = CStr(SpoluS)
82:     lblZlavP.Text = CStr(ZlavP)
83:     lblZlavS.Text = CStr(ZlavS)
84:     lblOsobP.Text = CStr(OsobP)
85:     lblOsobS.Text = CStr(OsobS)
86:
87:     Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub btnTrzbaNaDen_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles btnTrzbaNaDen.KeyPress
        frmPredstavenia.btnPridaj_KeyPress(sender, e, Me)
    End Sub

    Private Sub DataGridView1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DataGridView1.KeyPress
        frmPredstavenia.btnPridaj_KeyPress(sender, e, Me)
    End Sub

    Private Sub btnTrzbaDna_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles btnTrzbaDna.KeyPress
        frmPredstavenia.btnPridaj_KeyPress(sender, e, Me)
    End Sub
End Class