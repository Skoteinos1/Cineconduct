Public Class frmStatSucty

    Dim VizualUprava As Boolean = True

    Private Structure Sucty
        Dim Nazov As String
        Dim Filmy As Double
        Dim Predstavenia As Double
        Dim Listky As Double
        Dim PredstOs As Integer
        Dim ListkyOs As Integer
        Dim Datum As Date
        Dim Zlav As Integer
    End Structure

    Private Structure Listok
        Dim CisloListka As String
        Dim CasPredst As Date
        Dim cenaN As Integer
        Dim CenaZ As Integer
        Dim OsobN As Integer
        Dim OsobZ As Integer
        Dim CenaList As Integer
        Dim CenaTeor As Integer
        Dim Stav As String
    End Structure


    Private Sub frmStatSucty_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
1:      On Error GoTo Chyba
2:      'Zafarbi
3:      Call mdlColors.Skining(Me)
4:      If VizualUprava Then
5:          Call mdlColors.Sizing(Me)
6:          Me.CenterToParent()
7:          VizualUprava = False
8:      End If

        'Naplni combobox z Filmamy
62:     With frmSpracovavanie.FilmyBindingSource
63:         cmbFilmy.Items.Clear()
65:         For i = 0 To .Count - 1
66:             .Position = i
67:             cmbFilmy.Items.Add(.Current("Film"))
68:         Next i
69:     End With

        Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub btnZobrazit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnZobrazit.Click
        On Error GoTo Chyba
        If frmSpracovavanie.ChybaListky Or frmSpracovavanie.ChybaPredstavenia Or frmSpracovavanie.ChybaFilmy Then MsgBox("Poskodene data. Vystup nemusi byt spravny.", MsgBoxStyle.Exclamation)
        ' ---------- Spravny tvar Datagridu --------------
        With DataGridView1
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

            If rdbVsetko.Checked Then
25:             ' setup columns
27:             .Columns.Add("dtgrdDatum", "Názov")
28:             .Columns(0).Width = (.Width - 20) * 0.2
29:             .Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
30:             .Columns.Add("dtgrdDen", "Filmy")
31:             .Columns(1).Width = (.Width - 20) * 0.35
32:             .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
33:             .Columns(1).SortMode = DataGridViewColumnSortMode.NotSortable
34:             .Columns.Add("dtgrdCas", "Predstavenia")
35:             .Columns(2).Width = (.Width - 20) * 0.15
36:             .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
37:             .Columns(2).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Add("dtgrdCas", "Lístky")
                .Columns(3).Width = (.Width - 20) * 0.15
                .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(3).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Add("dtgrdCas", "Stav")
                .Columns(4).Width = (.Width - 20) * 0.15
                .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(4).SortMode = DataGridViewColumnSortMode.NotSortable
            ElseIf rdbPredstaveni.Checked Then
                If cmbFIlmy.Text = "" Then MsgBox("Musite vybrat film!") : Exit Sub
                ' setup columns
                .Columns.Add("dtgrdDatum", "Film")
                .Columns(0).Width = (.Width - 20) * 0.28
                .Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Add("dtgrdDen", "Datum")
                .Columns(1).Width = (.Width - 20) * 0.2
                .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(1).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Add("dtgrdCas", "Preddst. os")
                .Columns(2).Width = (.Width - 20) * 0.1
                .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(2).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Add("dtgrdCas", "Preddst. trzb")
                .Columns(3).Width = (.Width - 20) * 0.1
                .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(3).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Add("dtgrdCas", "Lístky os")
                .Columns(4).Width = (.Width - 20) * 0.1
                .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(4).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Add("dtgrdCas", "Lístky trzb")
                .Columns(5).Width = (.Width - 20) * 0.1
                .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(5).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Add("dtgrdCas", "Stav")
                .Columns(6).Width = (.Width - 20) * 0.12
                .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(6).SortMode = DataGridViewColumnSortMode.NotSortable
            ElseIf rdbListky.Checked Then
                If cmbFIlmy.Text = "" Then MsgBox("Musite vybrat film!") : Exit Sub
                ' setup columns
                .Columns.Add("dtgrdDatum", "Listok")
                .Columns(0).Width = (.Width - 20) * 0.2
                .Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Add("dtgrdDen", "Datum")
                .Columns(1).Width = (.Width - 20) * 0.2
                .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(1).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Add("dtgrdCas", "Cena N")
                .Columns(2).Width = (.Width - 20) * 0.08
                .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(2).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Add("dtgrdCas", "Cena Z")
                .Columns(3).Width = (.Width - 20) * 0.08
                .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(3).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Add("dtgrdCas", "Osob")
                .Columns(4).Width = (.Width - 20) * 0.08
                .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(4).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Add("dtgrdCas", "Zlav")
                .Columns(5).Width = (.Width - 20) * 0.08
                .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(5).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Add("dtgrdCas", "Cena L")
                .Columns(6).Width = (.Width - 20) * 0.08
                .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(6).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Add("dtgrdCas", "Cena S")
                .Columns(7).Width = (.Width - 20) * 0.08
                .Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(7).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns.Add("dtgrdCas", "Stav")
                .Columns(8).Width = (.Width - 20) * 0.08
                .Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(8).SortMode = DataGridViewColumnSortMode.NotSortable
            End If
            .ResumeLayout(True)
        End With

        Dim i, X, ipos As Integer
        Dim Vysledky(500) As Sucty

        DataGridView1.Rows.Clear()
        ProgressBar1.Visible = True
        ProgressBar1.Value = 0
        ipos = 0

        If rdbVsetko.Checked Then
            'Odkontroluje vsetky predstavenia, filmy, listky dokopy
            ProgressBar1.Maximum = frmSpracovavanie.FilmyBindingSource.Count + frmSpracovavanie.PredstaveniaBindingSource.Count + frmSpracovavanie.ListkyBindingSource.Count
            With frmSpracovavanie.FilmyBindingSource
                For i = 0 To .Count - 1
                    ProgressBar1.Value += 1
                    .Position = i
                    Vysledky(i).Nazov = .Current("Film")
                    Vysledky(i).Filmy = .Current("TrzbaFilmu")
                Next i
                ipos = .Count - 1
            End With

            With frmSpracovavanie.PredstaveniaBindingSource
                For i = 0 To .Count - 1
                    .Position = i
                    ProgressBar1.Value += 1
                    For X = 0 To ipos
                        If .Current("NazovFilmu") = Vysledky(X).Nazov Then
                            Vysledky(X).Predstavenia += .Current("TrzbaPredstavenia")
                            Exit For
                        End If
                    Next X
                Next i
            End With

            With frmSpracovavanie.ListkyBindingSource
                For i = 0 To .Count - 1
                    .Position = i
                    ProgressBar1.Value += 1
                    frmSpracovavanie.PredstaveniaBindingSource.Position = frmSpracovavanie.PredstaveniaBindingSource.Find("Datum", .Current("DatumPredst") & " " & .Current("CasPredst"))
                    For X = 0 To ipos
                        If frmSpracovavanie.PredstaveniaBindingSource.Current("NazovFilmu") = Vysledky(X).Nazov Then
                            If InStr(.Current("Stav"), "Storn") <> 0 Then
                                Vysledky(X).Listky -= .Current("suma")
                            Else
                                Vysledky(X).Listky += .Current("suma")
                            End If
                            Exit For
                        End If
                    Next X
                Next i
            End With

            For X = 0 To ipos
                With DataGridView1
                    .Rows.Add()
                    .Rows(.Rows.Count - 1).Cells(0).Value = Vysledky(X).Nazov
                    .Rows(.Rows.Count - 1).Cells(1).Value = Format(Vysledky(X).Filmy, "##0.00")
                    .Rows(.Rows.Count - 1).Cells(2).Value = Format(Vysledky(X).Predstavenia, "##0.00")
                    .Rows(.Rows.Count - 1).Cells(3).Value = Format(Vysledky(X).Listky, "##0.00")
                    If Format(Vysledky(X).Filmy, "##0.00") = Format(Vysledky(X).Predstavenia, "##0.00") And _
                         Format(Vysledky(X).Filmy, "##0.00") = Format(Vysledky(X).Listky, "##0.00") Then
                        .Rows(.Rows.Count - 1).Cells(4).Value = "OK"
                    Else
                        .Rows(.Rows.Count - 1).Cells(4).Value = "CHYBA"
                    End If
                End With
            Next X

        ElseIf rdbPredstaveni.Checked Then
            'kontroluje len jeden film po predstaveniach
            ProgressBar1.Maximum = +frmSpracovavanie.PredstaveniaBindingSource.Count + frmSpracovavanie.ListkyBindingSource.Count + 1
            With frmSpracovavanie.PredstaveniaBindingSource
                For i = 0 To .Count - 1
                    .Position = i
                    ProgressBar1.Value += 1
                    If .Current("NazovFilmu") = cmbFIlmy.Text Then
                        Vysledky(ipos).Datum = .Current("Datum")
                        Vysledky(ipos).Predstavenia += .Current("TrzbaPredstavenia")
                        Vysledky(ipos).PredstOs += .Current("Predaj")
                        ipos += 1
                    End If
                Next i
            End With

            With frmSpracovavanie.ListkyBindingSource
                For i = 0 To .Count - 1
                    .Position = i
                    ProgressBar1.Value += 1
                    frmSpracovavanie.PredstaveniaBindingSource.Position = frmSpracovavanie.PredstaveniaBindingSource.Find("Datum", .Current("DatumPredst") & " " & .Current("CasPredst"))

                    If frmSpracovavanie.PredstaveniaBindingSource.Current("NazovFilmu") = cmbFIlmy.Text Then
                        For X = 0 To ipos

                            If Vysledky(X).Datum = .Current("DatumPredst") & " " & .Current("CasPredst") Then

                                If InStr(.Current("Stav"), "Storn") <> 0 Then
                                    Vysledky(X).Listky -= .Current("suma")
                                    Vysledky(X).ListkyOs -= (.Current("MiestoDo") - .Current("MiestoOd") + 1)
                                Else
                                    Vysledky(X).Listky += .Current("suma")
                                    Vysledky(X).ListkyOs += (.Current("MiestoDo") - .Current("MiestoOd") + 1)
                                End If
                                Exit For

                            End If
                        Next X
                    End If
                Next i
            End With


            For X = 0 To ipos - 1
                With DataGridView1
                    .Rows.Add()
                    .Rows(.Rows.Count - 1).Cells(1).Value = Vysledky(X).Datum
                    .Rows(.Rows.Count - 1).Cells(2).Value = Vysledky(X).PredstOs
                    .Rows(.Rows.Count - 1).Cells(3).Value = Format(Vysledky(X).Predstavenia, "##0.00")
                    .Rows(.Rows.Count - 1).Cells(4).Value = Format(Vysledky(X).ListkyOs, "##0.00")
                    .Rows(.Rows.Count - 1).Cells(5).Value = Format(Vysledky(X).Listky, "##0.00")

                    If Format(Vysledky(X).Listky, "##0.00") = Format(Vysledky(X).Predstavenia, "##0.00") And _
                         Format(Vysledky(X).ListkyOs, "##0.00") = Format(Vysledky(X).PredstOs, "##0.00") Then
                        .Rows(.Rows.Count - 1).Cells(6).Value = "OK"
                    Else
                        .Rows(.Rows.Count - 1).Cells(6).Value = "CHYBA"
                    End If
                End With
            Next X

        ElseIf rdbListky.Checked Then
            Dim Listky(2000) As Listok

            With frmSpracovavanie.PredstaveniaBindingSource
                .Filter = "NazovFilmu = '" & cmbFIlmy.Text & "'"
                For i = 0 To .Count - 1
                    .Position = i
                    Vysledky(i).Datum = .Current("Datum")
                    Vysledky(i).Listky = .Current("CenaListka")
                    Vysledky(i).Zlav = .Current("CenaZlavListka")
                Next i
                ipos = i
                .RemoveFilter()
            End With

            Dim hod, den As Date
            Dim jpos As Integer = 0
            Dim hod2 As String

            With frmSpracovavanie.ListkyBindingSource
                jpos = 1
                For i = 0 To ipos - 1
                    den = Format(Vysledky(i).Datum, "dd.MM.yyyy")
                    hod = Format(Vysledky(i).Datum, "HH:mm")
                    hod2 = "30.12.1899 " & hod

                    .Filter = "DatumPredst = '" & den & "' AND CasPredst = '" & hod2 & "'"

                    For j = 0 To .Count - 1
                        .Position = j
                        If .Current("Kod") <> Listky(jpos - 1).CisloListka Then
                            Listky(jpos).CisloListka = .Current("Kod")
                            Listky(jpos).CasPredst = .Current("DatumPredst") & " " & .Current("CasPredst")
                            Listky(jpos).cenaN = Vysledky(i).Listky
                            Listky(jpos).CenaZ = Vysledky(i).Zlav
                            Listky(jpos).OsobN = .Current("MiestoDo") - .Current("MiestoOd") + 1
                            Listky(jpos).OsobZ = .Current("Zlavnenych")
                            Listky(jpos).CenaList = .Current("Suma")
                            If InStr(.Current("Stav"), "Storn") <> 0 Then
                                Listky(jpos).Stav = "Storno"
                            End If
                            jpos += 1
                        Else
                            If Listky(jpos).Stav <> "Storno" Then
                                If InStr(.Current("Stav"), "Storn") <> 0 Then Listky(jpos - 1).Stav = "Storno"
                            End If
                        End If
                    Next j
                Next i
                .RemoveFilter()
            End With

            i = 0

            For X = 1 To jpos
                With DataGridView1

                    Listky(X).CenaTeor = Listky(X).cenaN * Listky(X).OsobN - (Listky(X).cenaN - Listky(X).CenaZ) * Listky(X).OsobZ

                    If Listky(X).CenaTeor = Listky(X).CenaList Then
                        '    .Rows(.Rows.Count - 1).Cells(8).Value = "OK"
                        i += 1
                    Else
                        .Rows.Add()
                        .Rows(.Rows.Count - 1).Cells(0).Value = Listky(X).CisloListka
                        .Rows(.Rows.Count - 1).Cells(1).Value = Listky(X).CasPredst
                        .Rows(.Rows.Count - 1).Cells(2).Value = Listky(X).cenaN
                        .Rows(.Rows.Count - 1).Cells(3).Value = Listky(X).CenaZ
                        .Rows(.Rows.Count - 1).Cells(4).Value = Listky(X).OsobN
                        .Rows(.Rows.Count - 1).Cells(5).Value = Listky(X).OsobZ
                        .Rows(.Rows.Count - 1).Cells(6).Value = Listky(X).CenaList
                        .Rows(.Rows.Count - 1).Cells(7).Value = Listky(X).CenaTeor
                        If Listky(X).Stav <> "" Then
                            .Rows(.Rows.Count - 1).Cells(8).Value = Listky(X).Stav
                        Else
                            .Rows(.Rows.Count - 1).Cells(8).Value = "CHYBA"
                        End If
                    End If
                End With
            Next X
            MsgBox("Pocet listkov v poriadku: " & i)
        End If

        ProgressBar1.Visible = False

        Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub
End Class