Imports System.IO

Public Class frmLogy

    Dim VizualUprava As Boolean = True

    Private Sub frmLogy_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Zafarbi
        Call mdlColors.Skining(Me)
        If VizualUprava Then
            Call mdlColors.Sizing(Me)
            Me.CenterToParent()
            VizualUprava = False
        End If

        '************ SPRAVA DATA GRIDU*************************
        With DataGridView1
            .SuspendLayout()
            .DataSource = Nothing
            .AllowUserToAddRows = False
            .AllowUserToDeleteRows = False
            .AllowUserToResizeRows = False
            .AllowUserToResizeColumns = False

            .AllowUserToOrderColumns = True
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .ReadOnly = True
            .MultiSelect = False
            .RowHeadersVisible = False
            .Columns.Clear()
            .Rows.Clear()

            .Columns.Add("dtgrdMeno", "Datum")
            .Columns(0).Width = (.Width - 20) * 0.2
            .Columns.Add("dtgrdUdaje", "Log")
            .Columns(1).Width = (.Width - 20) * 0.8
            .ResumeLayout(True)
        End With
    End Sub

    Private Sub btnZobrazit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnZobrazit.Click



        Using sr As New StreamReader(frmSpracovavanie.Adresa & "log.txt")
            Dim line As String

            Dim Polozky(20000) As String
            Dim nElements As Short
            Dim aKey As Object
            Dim s1, s2 As String
            Dim DatumDO, DatumOD As Date
            Dim x As Integer = 0

            line = sr.ReadToEnd()
            'Console.WriteLine(line)

            aKey = Split(line, Chr(13) & Chr(10))
            nElements = UBound(aKey) - LBound(aKey) + 1

            For i = 0 To nElements - 1
                If aKey(i) <> "" Then
                    Polozky(i) = aKey(i)
                End If
            Next i

            DatumOD = CDate(dtpOd.Text)
            DatumDO = CDate(dtpDo.Text)

            'Naplni Datagrid
            With DataGridView1
                For i = 1 To nElements - 1

                    s1 = Trim(LSet(Polozky(i), 19))
                    s2 = Trim(Mid(Polozky(i), 22))

                    'Try
                    'datum = CDate(s1)

                    'If datum >= DatumOD And datum < DatumDO Then
                    .Rows.Add()
                    .Rows(.Rows.Count - 1).Cells(0).Value = s1 'datum
                    .Rows(.Rows.Count - 1).Cells(1).Value = s2
                    'End If
                    'Catch ex As Exception
                    '.Rows(.Rows.Count - 1).Cells(0).Value = s1
                    'x += 1
                    'End Try
                Next i


            End With

            MsgBox(x)

        End Using



        'Catch ex As Exception
        'Console.WriteLine("The file could not be read:")
        'Console.WriteLine(ex.Message)
        'MsgBox("The file could not be read: " & ex.Message)
        'End Try
    End Sub
End Class