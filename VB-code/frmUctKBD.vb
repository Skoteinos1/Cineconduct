Public Class frmUctKBD

    Dim VizualUprava As Boolean = True
    Dim DatumDO, DatumOD As Date
    Dim ZobrazitFilmy(500) As VlastnostiFilmu
    Dim ipos As Integer

    Private Structure VlastnostiFilmu
        Dim MenoFilmu As String
        Dim PevnePoz As Double
        Dim PercentPoz As Double
        Dim NaklNaNav As Double
        Dim NaklNaPredst As Double
    End Structure

    Private Sub frmUctKsD_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
1:      On Error GoTo Chyba
2:      Dim i As Short
3:
4:      'Zafarbi
5:      Call mdlColors.Skining(Me)
6:      If VizualUprava Then
7:          Call mdlColors.Sizing(Me)
8:          Me.CenterToParent()
9:          VizualUprava = False
10:     End If
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
30:         .Columns.Add("dtgrdDatum", "Film")
31:         .Columns(0).Width = (.Width - 20) * 0.2
32:         .Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
33:         .Columns.Add("dtgrdDen", "Dátum")
34:         .Columns(1).Width = (.Width - 20) * 0.07
35:         .Columns(1).SortMode = DataGridViewColumnSortMode.NotSortable
36:         .Columns.Add("dtgrdCas", "Cas")
37:         .Columns(2).Width = (.Width - 20) * 0.07
38:         .Columns(2).SortMode = DataGridViewColumnSortMode.NotSortable
39:         .Columns.Add("dtgrdFilm", "Trzba (EUR)")
40:         .Columns(3).Width = (.Width - 20) * 0.08
41:         .Columns(3).SortMode = DataGridViewColumnSortMode.NotSortable
42:         .Columns.Add("dtgrdSala", "Návst.")
43:         .Columns(4).Width = (.Width - 20) * 0.07
44:         .Columns(4).SortMode = DataGridViewColumnSortMode.NotSortable
45:         .Columns.Add("dtgrdCena", "Storno (EUR)")
46:         .Columns(5).Width = (.Width - 20) * 0.07
47:         .Columns(5).SortMode = DataGridViewColumnSortMode.NotSortable
48:         .Columns.Add("dtgrdCenaZlav", "Storno návst.")
49:         .Columns(6).Width = (.Width - 20) * 0.07
50:         .Columns(6).SortMode = DataGridViewColumnSortMode.NotSortable
51:         .Columns.Add("dtgrdSala", "Trzba -Storno")
52:         .Columns(7).Width = (.Width - 20) * 0.07
53:         .Columns(7).SortMode = DataGridViewColumnSortMode.NotSortable
54:         .Columns.Add("dtgrdSala", "(%)")
55:         .Columns(8).Width = (.Width - 20) * 0.05
56:         .Columns(8).SortMode = DataGridViewColumnSortMode.NotSortable
57:         .Columns.Add("dtgrdSala", "Odvod (EUR)")
58:         .Columns(9).Width = (.Width - 20) * 0.08
59:         .Columns(9).SortMode = DataGridViewColumnSortMode.NotSortable
60:         .Columns.Add("dtgrdSala", "Odvod DPH")
61:         .Columns(10).Width = (.Width - 20) * 0.08
62:         .Columns(10).SortMode = DataGridViewColumnSortMode.NotSortable
63:         .Columns.Add("dtgrdSala", "Odvod spolu")
64:         .Columns(11).Width = (.Width - 20) * 0.09
65:         .Columns(11).SortMode = DataGridViewColumnSortMode.NotSortable
66:
67:         .ResumeLayout(True)
68:     End With
69:
70:     'Naplni combobox z distributormi
71:     With frmSpracovavanie.DistributoriBindingSource
72:         cmbDistributori.Items.Clear()
73:         .MoveFirst()
74:         For i = 1 To .Count
75:             cmbDistributori.Items.Add(.Current("Distributor"))
76:             .MoveNext()
77:         Next i
78:     End With
79:
80:     lblTrzba.Text = 0
81:     lblNavstev.Text = 0
82:     lblStrn.Text = 0
83:     lblStrnNav.Text = 0
84:     lblNakl.Text = 0
85:     lblOdvod.Text = 0
86:     lblOdvodDPH.Text = 0
87:     lblOdvodSpolu.Text = 0
88:     lblpDPH.Text = "Sadzba DPH: " & frmSpracovavanie.pDPH & "%"

        Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub btnZobrazit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnZobrazit.Click
        If frmSpracovavanie.ChybaListky Or frmSpracovavanie.ChybaPredstavenia Or frmSpracovavanie.ChybaFilmy Then MsgBox("Poskodene data. Vystup nemusi byt spravny.", MsgBoxStyle.Exclamation)
        If (CDate(dtpOd.Text) >= CDate("1.1.2011") And CDate(dtpDo.Text) <= CDate("1.1.2011")) Or _
        (CDate(dtpOd.Text) < CDate("1.1.2011") And CDate(dtpDo.Text) > CDate("1.1.2011")) Then
            MsgBox("Obe datumy by mali byt vacsie alebo mensie ako 1.1.2011 kvoli zmene DPH.")
        End If

1:      Dim x, i, i2 As Integer
2:      Dim DatumPr As Date
3:      Dim StrnNav, Strn, Naklad, Odvod, Trzba, Navstev, PomocDat, OdvodE, OdvodDPH As Double
4:      Dim MenoFilmu As String
5:      Dim nasiel As Boolean
6:
7:      On Error GoTo Chyba
8:      Cursor = Cursors.WaitCursor
9:      ProgressBar1.Visible = True
10:
11:     DataGridView1.Rows.Clear()
12:
13:     DatumOD = CDate(dtpOd.Text)
14:     PomocDat = CDate(dtpDo.Text).ToOADate + 1
15:     DatumDO = DateTime.FromOADate(PomocDat)
16:     i2 = 1
17:
18:     '**********Vytvori tabulku s vlastnostami filmov***********
19:     'Vsetkych od toho distributora aj ked sa nevykazu. Aby ked prezera predstavenia vedelo koho ma ignorovat.
20:     'Nevlozi tam filmy ktore urcite neboli premietane
21:     '1. do tabulky si vlozi pozicovne
22:     With frmSpracovavanie.FilmyBindingSource
23:         For x = 0 To .Count - 1
24:             .Position = x
25:             If DatumOD <= CDate(.Current("PoslednePredstavenie")) Then
26:                 If cmbDistributori.Text = .Current("distributor") Or cmbDistributori.Text = "" Then
27:                     ZobrazitFilmy(i2).MenoFilmu = .Current("film")
28:                     ZobrazitFilmy(i2).PercentPoz = .Current("percentpozicovne")
29:                     ZobrazitFilmy(i2).PevnePoz = .Current("pevnepozicovne")
30:                     i2 = i2 + 1
31:                 End If
32:             End If
33:         Next x
34:     End With
35:     ipos = i2 - 1
36:     '2. do tabulky si vlozi naklady na film
37:     With frmSpracovavanie.NakladFilmBindingSource
38:         For i2 = 1 To ipos
39:             x = .Find("nazovfilmu", ZobrazitFilmy(i2).MenoFilmu)
40:             If x <> -1 Then
41:                 ZobrazitFilmy(i2).NaklNaNav += .Current("naklnavstev")
42:                 ZobrazitFilmy(i2).NaklNaPredst += .Current("naklpredst")
43:             End If
44:         Next i2
45:     End With
46:
47:     '************Vytvori tabulku s predstaveniamy*************
48:     With frmSpracovavanie.PredstaveniaBindingSource
49:         For x = 0 To .Count - 1
50:             .Position = x
51:             For i2 = 1 To ipos
52:                 MenoFilmu = .Current("nazovfilmu")
53:                 If ZobrazitFilmy(i2).MenoFilmu = MenoFilmu Then
54:                     ' If .Current("Predaj") <> 0 Then
55:                     If DatumOD <= .Current("datum") And DatumDO >= .Current("datum") Then
56:                         DataGridView1.Rows.Add()
57:                         DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = ZobrazitFilmy(i2).MenoFilmu
58:                         DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = Format(CDate(.Current("Datum")), "dd.MM")
59:                         DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = Format(CDate(.Current("Datum")), "HH:mm")
60:                         DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = Format(.Current("trzbapredstavenia"), "##0.00")
61:                         DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(4).Value = .Current("predaj")
62:                     End If
63:                     'End IF
64:                 End If
65:             Next i2
66:         Next x
67:     End With
68:
69:     'Uprava o predpredaj a odstranenie predaja z minuleho mesiaca 
        If rdbListky.Checked Then
70:         With frmSpracovavanie.ListkyBindingSource
71:             DatumDO = CDate(dtpDo.Text)
                .Filter = "(DenPredaja >= '" & Format(DatumOD, "dd.MM.yyyy") & "' AND DenPredaja <= '" & Format(DatumDO, "dd.MM.yyyy") & "') OR (DatumPredst >= '" & Format(DatumOD, "dd.MM.yyyy") & "' AND DatumPredst <= '" & Format(DatumDO, "dd.MM.yyyy") & "')"
72:             ProgressBar1.Maximum = .Count
73:             For i = 0 To .Count - 1
74:                 ProgressBar1.Value = i
75:                 .Position = i
76:                 If .Current("denpredaja") >= DatumOD And .Current("denpredaja") <= DatumDO Then
77:                     nasiel = False
78:                     If .Current("datumpredst") > DatumDO Then
79:                         If InStr(.Current("Stav"), "Predane") <> 0 Or InStr(.Current("Stav"), "Bezhotov") <> 0 Then
80:                             For x = 0 To DataGridView1.RowCount - 1
81:                                 If DataGridView1.Rows(x).Cells(1).Value = Format(CDate(.Current("DatumPredst")), "dd.MM") And _
                                     DataGridView1.Rows(x).Cells(2).Value = Format(CDate(.Current("CasPredst")), "HH:mm") Then
83:                                     DataGridView1.Rows(x).Cells(3).Value += .Current("suma")
84:                                     DataGridView1.Rows(x).Cells(3).Value = Format(DataGridView1.Rows(x).Cells(3).Value, "##0.00")
85:                                     DataGridView1.Rows(x).Cells(4).Value += (CInt(.Current("MiestoDo")) - CInt(.Current("MiestoOd")) + 1)
86:                                     nasiel = True
87:                                     Exit For
88:                                 End If
89:                             Next x
90:                             If nasiel = False Then
91:                                 DataGridView1.Rows.Add()
92:                                 DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = Format(CDate(.Current("DatumPredst")), "dd.MM")
93:                                 DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = Format(CDate(.Current("CasPredst")), "HH:mm")
94:                                 DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = .Current("suma")
95:                                 DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(4).Value = (CInt(.Current("MiestoDo")) - CInt(.Current("MiestoOd")) + 1)
96:                             End If
97:                         ElseIf InStr(.Current("Stav"), "Storno") <> 0 Or InStr(.Current("Stav"), "BezhotStorn") <> 0 Then
98:                             For x = 0 To DataGridView1.RowCount - 1
99:                                 If DataGridView1.Rows(x).Cells(1).Value = Format(CDate(.Current("DatumPredst")), "dd.MM") And _
                                     DataGridView1.Rows(x).Cells(2).Value = Format(CDate(.Current("CasPredst")), "HH:mm") Then
101:                                    DataGridView1.Rows(x).Cells(3).Value -= .Current("suma")
102:                                    DataGridView1.Rows(x).Cells(3).Value = Format(DataGridView1.Rows(x).Cells(3).Value, "##0.00")
103:                                    DataGridView1.Rows(x).Cells(4).Value -= (CInt(.Current("MiestoDo")) - CInt(.Current("MiestoOd")) + 1)
104:                                    nasiel = True
105:                                    Exit For
106:                                End If
107:                            Next x
108:                            If nasiel = False Then
109:                                DataGridView1.Rows.Add()
110:                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = Format(CDate(.Current("DatumPredst")), "dd.MM")
111:                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = Format(CDate(.Current("CasPredst")), "HH:mm")
112:                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = .Current("suma") * -1
113:                                DataGridView1.Rows(x).Cells(3).Value = Format(DataGridView1.Rows(x).Cells(3).Value, "##0.00")
114:                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(4).Value = (CInt(.Current("MiestoDo")) - CInt(.Current("MiestoOd")) + 1) * -1
115:                            End If
116:                        End If
117:                    End If
118:                ElseIf .Current("denpredaja") < DatumOD _
                     And .Current("datumpredst") >= DatumOD And .Current("datumpredst") <= DatumDO Then
120:                    nasiel = False
121:                    If InStr(.Current("Stav"), "Predane") <> 0 Or InStr(.Current("Stav"), "Bezhotov") <> 0 Then
122:                        For x = 0 To DataGridView1.RowCount - 1
123:                            If DataGridView1.Rows(x).Cells(1).Value = Format(CDate(.Current("DatumPredst")), "dd.MM") And _
                                 DataGridView1.Rows(x).Cells(2).Value = Format(CDate(.Current("CasPredst")), "HH:mm") Then
125:                                DataGridView1.Rows(x).Cells(3).Value -= .Current("suma")
126:                                DataGridView1.Rows(x).Cells(3).Value = Format(DataGridView1.Rows(x).Cells(3).Value, "##0.00")
127:                                DataGridView1.Rows(x).Cells(4).Value -= (CInt(.Current("MiestoDo")) - CInt(.Current("MiestoOd")) + 1)
128:                                nasiel = True
129:                                Exit For
130:                            End If
131:                        Next x
132:                        If nasiel = False Then
133:                            DataGridView1.Rows.Add()
134:                            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = Format(CDate(.Current("DatumPredst")), "dd.MM")
135:                            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = Format(CDate(.Current("CasPredst")), "HH:mm")
136:                            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = .Current("suma") * -1
137:                            DataGridView1.Rows(x).Cells(3).Value = Format(DataGridView1.Rows(x).Cells(3).Value, "##0.00")
138:                            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(4).Value = (CInt(.Current("MiestoDo")) - CInt(.Current("MiestoOd")) + 1) * -1
139:                        End If
140:                    ElseIf InStr(.Current("Stav"), "Storno") <> 0 Or InStr(.Current("Stav"), "BezhotStorn") <> 0 Then
141:                        For x = 0 To DataGridView1.RowCount - 1
142:                            If DataGridView1.Rows(x).Cells(1).Value = Format(CDate(.Current("DatumPredst")), "dd.MM") And _
                                 DataGridView1.Rows(x).Cells(2).Value = Format(CDate(.Current("CasPredst")), "HH:mm") Then
144:                                DataGridView1.Rows(x).Cells(3).Value += .Current("suma")
145:                                DataGridView1.Rows(x).Cells(3).Value = Format(DataGridView1.Rows(x).Cells(3).Value, "##0.00")
146:                                DataGridView1.Rows(x).Cells(4).Value += (CInt(.Current("MiestoDo")) - CInt(.Current("MiestoOd")) + 1)
147:                                nasiel = True
148:                                Exit For
149:                            End If
150:                        Next x
151:                        If nasiel = False Then
152:                            DataGridView1.Rows.Add()
153:                            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = Format(CDate(.Current("DatumPredst")), "dd.MM")
154:                            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = Format(CDate(.Current("CasPredst")), "HH:mm")
155:                            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = .Current("suma")
156:                            DataGridView1.Rows(x).Cells(3).Value = Format(DataGridView1.Rows(x).Cells(3).Value, "##0.00")
158:                            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(4).Value = (CInt(.Current("MiestoDo")) - CInt(.Current("MiestoOd")) + 1)
159:                        End If
160:                    End If
161:                End If
162:            Next i
                .RemoveFilter()
163:        End With
164:
165:        'Zmaze riadky od ineho distributora
166:        With DataGridView1
ZmazatRiadok:
168:            For x = 0 To .RowCount - 1
169:                If .Rows(x).Cells(0).Value = "" Then
170:                    MenoFilmu = DataGridView1.Rows(x).Cells(1).Value & "." & Format(DatumOD, "yyyy") & " " & DataGridView1.Rows(x).Cells(2).Value
171:                    frmSpracovavanie.PredstaveniaBindingSource.Position = frmSpracovavanie.PredstaveniaBindingSource.Find("Datum", MenoFilmu)
172:                    MenoFilmu = frmSpracovavanie.PredstaveniaBindingSource.Current("NazovFilmu")
173:                    nasiel = False
174:                    For i = 1 To ipos
175:                        If MenoFilmu = ZobrazitFilmy(i).MenoFilmu Then
176:                            .Rows(x).Cells(0).Value = MenoFilmu
177:                            nasiel = True
178:                            Exit For
179:                        End If
180:                    Next i
181:                    If nasiel = False Then
182:                        .Rows.RemoveAt(x)
183:
184:                        GoTo ZmazatRiadok
185:                    End If
186:                End If
187:            Next x
188:        End With
        End If
189:
190:    '***********Dopocita Velke Storno***************
191:    With frmSpracovavanie.StornoPredstaveniBindingSource
192:        For i2 = 0 To DataGridView1.RowCount - 1
193:            nasiel = False
194:            For i = 0 To .Count - 1
195:                .Position = i
196:                If .Current("datum") = (DataGridView1.Rows(i2).Cells(1).Value & "." & Year(DatumDO)) And Format(CDate(.Current("cas")), "HH:mm") = DataGridView1.Rows(i2).Cells(2).Value Then
197:                    DataGridView1.Rows(i2).Cells(5).Value = Format(.Current("Suma"), "##0.00")
198:                    DataGridView1.Rows(i2).Cells(6).Value = .Current("Osob")
199:                    nasiel = True
200:                    Exit For
201:                End If
202:            Next i
203:            If nasiel = False Then
204:                DataGridView1.Rows(i2).Cells(5).Value = Format(0, "##0.00")
205:                DataGridView1.Rows(i2).Cells(6).Value = 0
206:            End If
207:        Next i2
208:    End With
209:
210:    'Vypocita zvysok
211:    With DataGridView1
212:        For i = 0 To .RowCount - 1
213:            For i2 = 1 To ipos
214:                If ZobrazitFilmy(i2).MenoFilmu = .Rows(i).Cells(0).Value Then
215:                    'Naklady vyriesit dodatocne
216:                    'Odvod = CDbl(DataGridView1.Rows(i2).Cells(3).Value) - CDbl(DataGridView1.Rows(i2).Cells(5).Value) - CDbl(DataGridView1.Rows(i2).Cells(7).Value)
217:                    Odvod = (CDbl(.Rows(i).Cells(3).Value) - CDbl(.Rows(i).Cells(5).Value))
218:                    .Rows(i).Cells(7).Value = Format(Odvod, "##0.00")
219:
220:                    If ZobrazitFilmy(i2).PevnePoz <> 0 Then
221:                        Odvod = ZobrazitFilmy(i2).PevnePoz
222:                    Else
223:                        .Rows(i).Cells(8).Value = ZobrazitFilmy(i2).PercentPoz
224:                        Odvod = Odvod * (ZobrazitFilmy(i2).PercentPoz / 100)
225:                    End If
226:
227:                    If InStr(LCase(.Rows(i).Cells(0).Value), "organizovane") <> 0 Or InStr(LCase(.Rows(i).Cells(0).Value), "organizované") <> 0 Then
228:                        .Rows(i).Cells(9).Value = Format(0, "##0.00")
229:                        .Rows(i).Cells(10).Value = Format(0, "###0.00")
230:                        .Rows(i).Cells(11).Value = Format(0, "###0.00")
231:                    Else
232:                        .Rows(i).Cells(9).Value = Format(Odvod, "##0.00")
233:                        .Rows(i).Cells(10).Value = Format(Odvod * frmSpracovavanie.pDPH / 100, "###0.00")
234:                        .Rows(i).Cells(11).Value = Format(Odvod + Odvod * (frmSpracovavanie.pDPH / 100), "###0.00")
235:                    End If
236:                    Exit For
237:                End If
238:            Next i2
239:        Next i
240:    End With
241:
242:    'Sucet na koniec
243:    With DataGridView1
244:        Strn = 0
245:        StrnNav = 0
246:        Naklad = 0
247:        Odvod = 0
248:        Trzba = 0
249:        Navstev = 0
250:        OdvodDPH = 0
251:        OdvodE = 0
252:        For i = 0 To .RowCount - 1
253:            Trzba += CDbl(.Rows(i).Cells(3).Value)
254:            Navstev += CDbl(.Rows(i).Cells(4).Value)
255:            Strn += CDbl(.Rows(i).Cells(5).Value)
256:            StrnNav += CDbl(.Rows(i).Cells(6).Value)
257:            Naklad += CDbl(.Rows(i).Cells(7).Value)
258:            OdvodE += CDbl(.Rows(i).Cells(9).Value)
259:            OdvodDPH += CDbl(.Rows(i).Cells(10).Value)
260:            Odvod += CDbl(.Rows(i).Cells(11).Value)
261:        Next i
262:    End With
263:    lblTrzba.Text = Format(Trzba, "##0.00")
264:    lblNavstev.Text = Navstev
265:    lblStrn.Text = Format(Strn, "##0.00")
266:    lblStrnNav.Text = StrnNav
267:    lblNakl.Text = Format(Naklad, "##0.00")
268:    lblOdvod.Text = Format(OdvodE, "##0.00")
269:    lblOdvodDPH.Text = Format(OdvodDPH, "##0.00")
270:    lblOdvodSpolu.Text = Format(Odvod, "##0.00")
271:
272:    'Ak je v zozname viac distributorov
273:    If cmbDistributori.Text = "" And DataGridView1.RowCount > 0 Then
274:        MenoFilmu = ""
275:        i2 = 1
276:        Dim Distributori(100) As String
277:        With DataGridView1
278:            For x = 0 To .RowCount - 1
279:                frmSpracovavanie.FilmyBindingSource.Position = frmSpracovavanie.FilmyBindingSource.Find("Film", .Rows(x).Cells(0).Value)
280:                Distributori(0) = frmSpracovavanie.FilmyBindingSource.Current("Distributor")
281:                For i = 1 To i2
282:                    If Distributori(i) = Distributori(0) Then Exit For
283:                    If i = i2 Then
284:                        Distributori(i2) = Distributori(0)
285:                        i2 += 1
286:                        MenoFilmu += Distributori(0) & Chr(13) & Chr(10)
287:                    End If
288:                Next i
289:            Next x
290:        End With
291:        MsgBox("Zoznam obsahuje tychto Distributorov:" & Chr(13) & Chr(10) & MenoFilmu)
292:    End If
293:
294:    ProgressBar1.Visible = False
295:    Cursor = Cursors.Default
        Exit Sub
Chyba:
        Cursor = Cursors.Default
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub btnVytlacit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnVytlacit.Click
1:      If frmSpracovavanie.ChybaListky Or frmSpracovavanie.ChybaPredstavenia Or frmSpracovavanie.ChybaFilmy Then MsgBox("Poskodene data. Vystup nemusi byt spravny.", MsgBoxStyle.Exclamation) : Exit Sub
        frmMenu.Vytlacit = ""
2:      frmMenu.Velkost = 9
3:      Dim ODDELOVAC As String = StrDup(95, "─")
4:      Dim ODDELOVAC2 As String = StrDup(16, " ")
5:      Dim X, i, i2, j, i3, ipos As Integer
6:      Dim PomocRetazec, CRC As String
7:      Dim Hlavicka(15) As String
8:      Dim Tab(15) As Integer
9:      Dim Distributori(50) As String
10:     Dim Distr As String
11:     Dim sucty(13) As Double
12:     Dim Filmy(150) As String
13:     Dim nasiel As Boolean
14:     Dim PrvyRiadok As Boolean
15:     Dim Vysledky(15) As Double
16:
17:     ODDELOVAC2 += StrDup((95 - Len(ODDELOVAC2)), "─")
18:
19:     Vysledky(1) = CDbl(lblTrzba.Text)
20:     Vysledky(2) = CDbl(lblNavstev.Text)
21:     Vysledky(3) = CDbl(lblStrn.Text)
22:     Vysledky(4) = CDbl(lblStrnNav.Text)
23:     Vysledky(5) = CDbl(lblNakl.Text)
24:     Vysledky(6) = CDbl(lblOdvod.Text)
25:     Vysledky(7) = CDbl(lblOdvodDPH.Text)
26:     Vysledky(8) = CDbl(lblOdvodSpolu.Text)
27:
28:     On Error GoTo Chyba
29:     'Zisti kolkych distributorov ma vytlacit
30:     i2 = 1
31:     If cmbDistributori.Text = "" Then
32:         With DataGridView1
33:             For X = 0 To .RowCount - 1
34:                 frmSpracovavanie.FilmyBindingSource.Position = frmSpracovavanie.FilmyBindingSource.Find("Film", .Rows(X).Cells(0).Value)
35:                 Distr = frmSpracovavanie.FilmyBindingSource.Current("Distributor")
36:                 For i = 1 To i2
37:                     If Distributori(i) = Distr Then Exit For
38:                     If i = i2 Then
39:                         Distributori(i2) = Distr
40:                         i2 += 1
41:                     End If
42:                 Next i
43:             Next X
44:         End With
45:         i2 -= 1
46:     End If
47:
48:     For j = 1 To i2
49:         ipos = 0
50:         If cmbDistributori.Text = "" Or i2 <> 1 Then
51:             cmbDistributori.Text = Distributori(j)
52:             btnZobrazit_Click(sender, e)
53:         End If
54:
55:         Tab(0) = 0
56:         Tab(1) = 16
57:         Tab(2) = 22
58:         Tab(3) = 28
59:         Tab(4) = 37
60:         Tab(5) = 44
61:         Tab(6) = 51
62:         Tab(7) = 57
63:         Tab(8) = 67
64:         Tab(9) = 70
65:         Tab(10) = 78
66:         Tab(11) = 87
67:
68:         Hlavicka(0) = "Film"
69:         Hlavicka(1) = "Dátum"
70:         Hlavicka(2) = "Cas"
71:         Hlavicka(3) = " Trzba"
72:         Hlavicka(4) = "Návst."
73:         Hlavicka(5) = "Storno"
74:         Hlavicka(6) = "Storno"
75:         Hlavicka(7) = "  Trzba"
76:         Hlavicka(8) = "%"
77:         Hlavicka(9) = "  Odvod"
78:         Hlavicka(10) = "  Odvod"
79:         Hlavicka(11) = "  Odvod"
80:
81:         'Printer.FontSize = 13
82:         If rdbListky.Checked Then frmMenu.Vytlacit += "Vyuctovanie podla predaja listkov za " Else frmMenu.Vytlacit += "Vyuctovanie podla predstaveni za "
            frmMenu.Vytlacit += frmSpracovavanie.MenoKina & Chr(13) & Chr(10)
83:         frmMenu.Vytlacit += "Od: " & dtpOd.Text & "      Do: " & dtpDo.Text & Chr(13) & Chr(10)
84:         frmMenu.Vytlacit += "Pre distributora: " & cmbDistributori.Text & Chr(13) & Chr(10)
85:         CRC = Now
86:         frmMenu.Vytlacit += "Dátum vytvorenia: " & CRC & Chr(13) & Chr(10)
87:         frmMenu.Vytlacit += Chr(13) & Chr(10)
88:
89:         'Printer.FontSize = 10
90:         'Hlavicka
91:         ' frmMenu.Vytlacit += "---------1---------2---------3---------4---------5---------6---------7---------8---------9---" & Chr(13) & Chr(10)
92:         'frmMenu.Vytlacit += "123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123" & Chr(13) & Chr(10)
93:         frmMenu.Vytlacit += ODDELOVAC & Chr(13) & Chr(10)
94:         PomocRetazec = ""
95:         For i = 0 To 11
96:             PomocRetazec += StrDup((Tab(i) - Len(PomocRetazec)), " ")
97:             PomocRetazec += Hlavicka(i)
98:         Next i
99:         frmMenu.Vytlacit += PomocRetazec & Chr(13) & Chr(10)
100:
101:        Hlavicka(6) = "návst."
102:        Hlavicka(7) = " -Storno"
103:        PomocRetazec = ""
104:        For i = 6 To 7
105:            PomocRetazec += StrDup((Tab(i) - Len(PomocRetazec)), " ")
106:            PomocRetazec += Hlavicka(i)
107:        Next i
108:        frmMenu.Vytlacit += PomocRetazec & Chr(13) & Chr(10)
109:
110:        Hlavicka(3) = " (EUR)"
111:        Hlavicka(5) = "(EUR)"
112:        Hlavicka(7) = "  (EUR)"
113:        Hlavicka(9) = "  (EUR)"
114:        Hlavicka(10) = "   DPH"
115:        Hlavicka(11) = "  spolu"
116:
117:        PomocRetazec = ""
118:        For i = 3 To 11
119:            If (Not i = 4 And Not i = 6 And Not i = 8) Then
120:                PomocRetazec += StrDup((Tab(i) - Len(PomocRetazec)), " ")
121:                PomocRetazec += Hlavicka(i)
122:            End If
123:        Next i
124:        frmMenu.Vytlacit += PomocRetazec & Chr(13) & Chr(10)
125:
126:        'Zisti si mena filmov na tlacenie
127:        For i = 0 To 149
128:            Filmy(i) = ""
129:        Next i
130:        With DataGridView1
131:            For X = 0 To .RowCount - 1
132:                nasiel = False
133:                For i = 0 To ipos
134:                    If Filmy(i) = .Rows(X).Cells(0).Value Then
135:                        nasiel = True
136:                        Exit For
137:                    End If
138:                Next i
139:                If nasiel = False Then
140:                    Filmy(ipos) = .Rows(X).Cells(0).Value
141:                    ipos += 1
142:                End If
143:            Next X
144:
145:            DataGridView1.Sort(DataGridView1.Columns(1), System.ComponentModel.ListSortDirection.Ascending)
146:            For i3 = 0 To ipos - 1
147:                PrvyRiadok = True
148:                For X = 0 To 12
149:                    sucty(X) = 0
150:                Next X
151:
152:                For X = 0 To .Rows.Count - 1
153:                    If CStr(.Rows(X).Cells(0).Value) = Filmy(i3) Then
154:                        PomocRetazec = ""
155:                        If PrvyRiadok Then
156:                            frmMenu.Vytlacit += ODDELOVAC & Chr(13) & Chr(10)
157:                            If Len(Filmy(i3)) > 15 Then
158:                                PomocRetazec += LSet(Filmy(i3), 12) & "..."
159:                            Else
160:                                PomocRetazec += Filmy(i3)
161:                            End If
162:                            PrvyRiadok = False
163:                        End If
164:
165:                        For i = 1 To 11
166:                            PomocRetazec += StrDup((Tab(i) - Len(PomocRetazec)), " ")
167:                            If i = 3 Or i = 7 Or i >= 9 Then PomocRetazec += StrDup((8 - Len(CStr(.Rows(X).Cells(i).Value))), " ")
168:                            If i = 4 Or i = 6 Then PomocRetazec += StrDup((4 - Len(CStr(.Rows(X).Cells(i).Value))), " ")
169:                            If i = 5 Then PomocRetazec += StrDup((6 - Len(CStr(.Rows(X).Cells(i).Value))), " ")
170:                            PomocRetazec += CStr(.Rows(X).Cells(i).Value)
171:                            If i = 2 Then sucty(i) += 1
172:                            If i > 2 Then sucty(i) += CDbl(.Rows(X).Cells(i).Value)
173:                        Next i
174:                        frmMenu.Vytlacit += PomocRetazec & Chr(13) & Chr(10)
175:                    End If
176:
177:                    If X = .RowCount - 1 Then
178:                        frmMenu.Vytlacit += ODDELOVAC2 & Chr(13) & Chr(10)
179:                        PomocRetazec = "Spolu:"
180:                        For i = 2 To 11
181:                            If i <> 8 Then
182:                                PomocRetazec += StrDup((Tab(i) - Len(PomocRetazec)), " ")
183:                                If i = 3 Or i = 7 Or i >= 9 Then PomocRetazec += StrDup((8 - Len(Format(sucty(i), "##0.00"))), " ")
184:                                If i = 4 Or i = 6 Then PomocRetazec += StrDup((4 - Len(Format(sucty(i), "##0"))), " ")
185:                                If i = 5 Then PomocRetazec += StrDup((6 - Len(Format(sucty(i), "##0.00"))), " ")
186:                                If i = 3 Or i = 5 Or i >= 7 Then
187:                                    PomocRetazec += Format(sucty(i), "##0.00")
188:                                Else
189:                                    PomocRetazec += CStr(sucty(i))
190:                                End If
191:                            End If
192:                        Next i
193:                        frmMenu.Vytlacit += PomocRetazec & Chr(13) & Chr(10)
194:                        frmMenu.Vytlacit += Chr(13) & Chr(10)
195:                        'frmMenu.Vytlacit += Chr(13) & Chr(10)
196:                    End If
197:
198:                Next X
199:            Next i3
200:        End With
201:
202:        frmMenu.Vytlacit += Chr(13) & Chr(10)
203:        frmMenu.Vytlacit += ODDELOVAC & Chr(13) & Chr(10)
204:        frmMenu.Vytlacit += Chr(13) & Chr(10)
205:
206:        PomocRetazec = ""
207:        PomocRetazec += "Spolu:"
208:        PomocRetazec += StrDup((Tab(3) - Len(PomocRetazec)), " ")
209:        PomocRetazec += StrDup((8 - Len(Format(CDbl(lblTrzba.Text), "##0.00"))), " ") & Format(CDbl(lblTrzba.Text), "##0.00")
210:        PomocRetazec += StrDup((Tab(4) - Len(PomocRetazec)), " ")
211:        PomocRetazec += StrDup((4 - Len(Format(CDbl(lblNavstev.Text), "##0"))), " ") & Format(CDbl(lblNavstev.Text), "##0")
212:        PomocRetazec += StrDup((Tab(5) - Len(PomocRetazec)), " ")
213:        PomocRetazec += StrDup((6 - Len(Format(CDbl(lblStrn.Text), "##0.00"))), " ") & Format(CDbl(lblStrn.Text), "##0.00")
214:        PomocRetazec += StrDup((Tab(6) - Len(PomocRetazec)), " ")
215:        PomocRetazec += StrDup((4 - Len(Format(CDbl(lblStrnNav.Text), "##0"))), " ") & Format(CDbl(lblStrnNav.Text), "##0")
216:        PomocRetazec += StrDup((Tab(7) - Len(PomocRetazec)), " ")
217:        PomocRetazec += StrDup((8 - Len(Format(CDbl(lblNakl.Text), "##0.00"))), " ") & Format(CDbl(lblNakl.Text), "##0.00")
218:        PomocRetazec += StrDup((Tab(9) - Len(PomocRetazec)), " ")
219:        PomocRetazec += StrDup((8 - Len(Format(CDbl(lblOdvod.Text), "##0.00"))), " ") & Format(CDbl(lblOdvod.Text), "##0.00")
220:        PomocRetazec += StrDup((Tab(10) - Len(PomocRetazec)), " ")
221:        PomocRetazec += StrDup((8 - Len(Format(CDbl(lblOdvodDPH.Text), "##0.00"))), " ") & Format(CDbl(lblOdvodDPH.Text), "##0.00")
222:        PomocRetazec += StrDup((Tab(11) - Len(PomocRetazec)), " ")
223:        PomocRetazec += StrDup((8 - Len(Format(CDbl(lblOdvodSpolu.Text), "##0.00"))), " ") & Format(CDbl(lblOdvodSpolu.Text), "##0.00")
224:        frmMenu.Vytlacit += PomocRetazec & Chr(13) & Chr(10)
225:        frmMenu.Vytlacit += Chr(13) & Chr(10)
226:
227:        frmMenu.Vytlacit += "CRC: " & Module1.HASHString(CRC & dtpOd.Text & dtpDo.Text & lblOdvod.Text) & Chr(13) & Chr(10)
228:
229:        'Vytlaci storno
230:        If CDbl(lblStrn.Text) <> 0 Then
231:            Tab(1) = 27
232:            Tab(2) = 39
233:            Tab(3) = 46
234:            Tab(4) = 52
235:            Tab(5) = 60
236:            frmMenu.Vytlacit += Chr(13) & Chr(10)
237:            frmMenu.Vytlacit += Chr(13) & Chr(10)
238:            frmMenu.Vytlacit += "Storno predstavení:" & Chr(13) & Chr(10)
239:            frmMenu.Vytlacit += ODDELOVAC & Chr(13) & Chr(10)
240:
241:            Hlavicka(0) = "Film"
242:            Hlavicka(1) = "Dátum"
243:            Hlavicka(2) = "Cas"
244:            Hlavicka(3) = "Suma"
245:            Hlavicka(4) = "Navst."
246:            Hlavicka(5) = "Dovod"
247:            PomocRetazec = ""
248:            For i = 0 To 5
249:                PomocRetazec += StrDup((Tab(i) - Len(PomocRetazec)), " ")
250:                PomocRetazec += Hlavicka(i)
251:            Next i
252:            frmMenu.Vytlacit += PomocRetazec & Chr(13) & Chr(10)
253:            frmMenu.Vytlacit += ODDELOVAC & Chr(13) & Chr(10)
254:
255:            For X = 1 To ipos
256:                With frmSpracovavanie.StornoPredstaveniBindingSource
257:                    .MoveFirst()
258:                    For i = 1 To .Count
259:                        If .Current("film") = ZobrazitFilmy(X).MenoFilmu Then
260:                            PomocRetazec = ""
261:                            PomocRetazec += .Current("film")
262:                            PomocRetazec += StrDup((Tab(1) - Len(PomocRetazec)), " ")
263:                            PomocRetazec += .Current("datum")
264:                            PomocRetazec += StrDup((Tab(2) - Len(PomocRetazec)), " ")
265:                            PomocRetazec += Format(CDate(.Current("cas")), "HH:mm")
266:                            PomocRetazec += StrDup((Tab(3) - Len(PomocRetazec)), " ")
267:                            PomocRetazec += CStr(.Current("Suma"))
268:                            PomocRetazec += StrDup((Tab(4) - Len(PomocRetazec)), " ")
269:                            PomocRetazec += CStr(.Current("Osob"))
270:                            PomocRetazec += StrDup((Tab(5) - Len(PomocRetazec)), " ")
271:                            PomocRetazec += .Current("dovod")
272:                            frmMenu.Vytlacit += PomocRetazec & Chr(13) & Chr(10)
273:                        End If
274:                        .MoveNext()
275:                    Next i
276:                End With
277:            Next X
278:            frmMenu.Vytlacit += ODDELOVAC & Chr(13) & Chr(10)
279:        End If
280:
281:        'Vytlaci naklady
282:        'Blbost aby netlacilo naklady
283:        If lblNakl.Text = "BBC" Then
284:            'If CDbl(lblNakl.Text) <> 0 Then
285:            Tab(1) = 27
286:            Tab(2) = 41
287:            Tab(3) = 55
288:            frmMenu.Vytlacit += Chr(13) & Chr(10)
289:            frmMenu.Vytlacit += Chr(13) & Chr(10)
290:            frmMenu.Vytlacit += "Náklady predstavení:" & Chr(13) & Chr(10)
291:            frmMenu.Vytlacit += ODDELOVAC & Chr(13) & Chr(10)
292:
293:            Hlavicka(0) = "Film"
294:            Hlavicka(1) = "Náklad na"
295:            Hlavicka(2) = "Náklad na"
296:            Hlavicka(3) = "Názov nákladu"
297:            PomocRetazec = ""
298:            For i = 0 To 3
299:                PomocRetazec += StrDup((Tab(i) - Len(PomocRetazec)), " ")
300:                PomocRetazec += Hlavicka(i)
301:            Next i
302:            frmMenu.Vytlacit += PomocRetazec & Chr(13) & Chr(10)
303:
304:            Hlavicka(0) = ""
305:            Hlavicka(1) = "návstevníka"
306:            Hlavicka(2) = "predstavenie"
307:            PomocRetazec = ""
308:            For i = 0 To 2
309:                PomocRetazec += StrDup((Tab(i) - Len(PomocRetazec)), " ")
310:                PomocRetazec += Hlavicka(i)
311:            Next i
312:            frmMenu.Vytlacit += PomocRetazec & Chr(13) & Chr(10)
313:
314:            Hlavicka(0) = ""
315:            Hlavicka(1) = "(EUR)"
316:            Hlavicka(2) = "(EUR)"
317:            PomocRetazec = ""
318:            For i = 0 To 2
319:                PomocRetazec += StrDup((Tab(i) - Len(PomocRetazec)), " ")
320:                PomocRetazec += Hlavicka(i)
321:            Next i
322:            frmMenu.Vytlacit += PomocRetazec & Chr(13) & Chr(10)
323:            frmMenu.Vytlacit += ODDELOVAC & Chr(13) & Chr(10)
324:
325:            For X = 1 To ipos
326:                With frmSpracovavanie.NakladFilmBindingSource
327:                    .MoveFirst()
328:                    For i = 1 To .Count
329:                        If .Current("nazovfilmu") = ZobrazitFilmy(X).MenoFilmu Then
330:                            PomocRetazec = ""
331:                            PomocRetazec += .Current("nazovfilmu")
332:                            PomocRetazec += StrDup((Tab(1) - Len(PomocRetazec)), " ")
333:                            PomocRetazec += CStr(.Current("naklnavstev"))
334:                            PomocRetazec += StrDup((Tab(2) - Len(PomocRetazec)), " ")
335:                            PomocRetazec += CStr(.Current("naklpredst"))
336:                            PomocRetazec += StrDup((Tab(3) - Len(PomocRetazec)), " ")
337:                            PomocRetazec += .Current("menonakl")
338:                            frmMenu.Vytlacit += PomocRetazec & Chr(13) & Chr(10)
339:                        End If
340:                        .MoveNext()
341:                    Next i
342:                End With
343:            Next X
344:            frmMenu.Vytlacit += ODDELOVAC & Chr(13) & Chr(10)
345:        End If
346:
347:        frmMenu.Vytlacit += Chr(13) & Chr(10) & Chr(13) & Chr(10)
348:    Next j
349:
350:    If i2 > 1 Then
351:        'Sucet za vsetkych na konci ak ich tlaci viac
352:        frmMenu.Vytlacit += Chr(13) & Chr(10)
353:        frmMenu.Vytlacit += ODDELOVAC & Chr(13) & Chr(10)
354:        frmMenu.Vytlacit += Chr(13) & Chr(10)
355:        PomocRetazec = ""
356:        PomocRetazec += "Spolu za vsetkych:"
357:        PomocRetazec += StrDup((Tab(3) - Len(PomocRetazec)), " ")
358:        PomocRetazec += StrDup((8 - Len(Format(Vysledky(1), "##0.00"))), " ") & Format(Vysledky(1), "##0.00")
359:        PomocRetazec += StrDup((Tab(4) - Len(PomocRetazec)), " ")
360:        PomocRetazec += StrDup((4 - Len(Format(Vysledky(2), "##0"))), " ") & Format(Vysledky(2), "##0")
361:        PomocRetazec += StrDup((Tab(5) - Len(PomocRetazec)), " ")
362:        PomocRetazec += StrDup((6 - Len(Format(Vysledky(3), "##0.00"))), " ") & Format(Vysledky(3), "##0.00")
363:        PomocRetazec += StrDup((Tab(6) - Len(PomocRetazec)), " ")
364:        PomocRetazec += StrDup((4 - Len(Format(Vysledky(4), "##0"))), " ") & Format(Vysledky(4), "##0")
365:        PomocRetazec += StrDup((Tab(7) - Len(PomocRetazec)), " ")
366:        PomocRetazec += StrDup((8 - Len(Format(Vysledky(5), "##0.00"))), " ") & Format(Vysledky(5), "##0.00")
367:        PomocRetazec += StrDup((Tab(9) - Len(PomocRetazec)), " ")
368:        PomocRetazec += StrDup((8 - Len(Format(Vysledky(6), "##0.00"))), " ") & Format(Vysledky(6), "##0.00")
369:        PomocRetazec += StrDup((Tab(10) - Len(PomocRetazec)), " ")
370:        PomocRetazec += StrDup((8 - Len(Format(Vysledky(7), "##0.00"))), " ") & Format(Vysledky(7), "##0.00")
371:        PomocRetazec += StrDup((Tab(11) - Len(PomocRetazec)), " ")
372:        PomocRetazec += StrDup((8 - Len(Format(Vysledky(8), "##0.00"))), " ") & Format(Vysledky(8), "##0.00")
373:        frmMenu.Vytlacit += PomocRetazec & Chr(13) & Chr(10)
374:        frmMenu.Vytlacit += Chr(13) & Chr(10)
375:        frmMenu.Vytlacit += ODDELOVAC
376:    End If
377:
378:    'Printer.EndDoc()
379:
380:    If frmMenu.Vytlacit <> "" Then frmMenu.Tlacit()
381:
382:    Exit Sub
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
    Private Sub rdbListky_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdbListky.CheckedChanged
        DataGridView1.Rows.Clear()
    End Sub
    Private Sub rdbPredstavenia_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdbPredstavenia.CheckedChanged
        DataGridView1.Rows.Clear()
    End Sub

    Private Sub dtpOd_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpOd.ValueChanged
        If dtpOd.Value > dtpDo.Value Then dtpDo.Value = dtpOd.Value
    End Sub
    Private Sub dtpDo_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpDo.ValueChanged
        If dtpOd.Value > dtpDo.Value Then dtpOd.Value = dtpDo.Value
    End Sub
End Class