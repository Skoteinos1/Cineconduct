Imports System.IO

Public Class frmZavierka

    Dim VizualUprava As Boolean = True

    Public TrzbaPokladnika As Double
    Public DnesnaTrzba2 As Double
    Public DnesStorno2 As Double
    Public DnesBezhot As Double
    Public DnesBezhStor2 As Double
    Public PokladnBezhot As Double

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        frmZavierka_Disposed(sender, e)
    End Sub

    Public Sub frmZavierka_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
1:      If frmLogin.Rights <> "A" Then On Error Resume Next
2:      '  Me.Hide()
3:      ' frmVybPredst.Hide()
4:
5:      AboutBox1 = Nothing 'frmAbout2 = Nothing 'frmAbout2.Dispose()
        frmCommand = Nothing
6:      frmDistribut = Nothing
7:      frmDistributUprav = Nothing
8:      frmFilmy = Nothing
9:      frmFilmyNaklady = Nothing
10:     frmFilmyUprav = Nothing
11:     frmLicencia = Nothing
12:     frmLogin = Nothing
        frmOptBlocek = Nothing
13:     frmOptDivadlo = Nothing
14:     frmOptions = Nothing
15:     frmPocStav = Nothing
16:     frmPredaj = Nothing
        frmPredajObr = Nothing
21:     frmPredpisTrzby = Nothing
22:     frmPredstavenia = Nothing
23:     frmPredstaveniaUprav = Nothing
24:     'frmPredstPrehl = Nothing
25:     frmPredstStorn = Nothing
26:     frmPresad = Nothing
27:     frmRezerv = Nothing
28:     frmRezervPredaj = Nothing
29:     frmSplash = Nothing
30:     'frmStatLogy = Nothing
        'frmStatPokladn = Nothing
31:     'frmStatPoradF = Nothing
32:     'frmStatPrehladTrz = Nothing
33:     frmStatPrehladTrzP = Nothing
34:     frmStatRocVysl = Nothing
        frmStatSucty = Nothing
35:     'frmStatVytazSal = Nothing
36:     frmStorno = Nothing
37:     frmUctDBD = Nothing
38:     frmUctDsD = Nothing
39:     frmUctKBD = Nothing
40:     frmUctKsD = Nothing
41:     frmUctVykVysl = Nothing
42:     frmVybPredst = Nothing
43:     frmVydat = Nothing
44:
45:     'DataEnvironment1.Connection1.Close()
        If frmSpracovavanie.Zaloha Then
            FileCopy(frmSpracovavanie.Adresa2 & "Zalohy\data2.pac", frmSpracovavanie.Adresa2 & "Zalohy\data3.pac")
46:         FileCopy(frmSpracovavanie.Adresa & "Data\data.pac", frmSpracovavanie.Adresa2 & "Zalohy\data2.pac")
        End If
47:     Call Module1.HASHSubor("Zalohy\data2.pac")
48:     Call Module1.HASHSubor("Data\data.pac")
49:
        If Directory.Exists(frmSpracovavanie.Adresa2 & "Foto") Then
            On Error Resume Next
            Dim j As Integer = 50
            If j < frmSpracovavanie.NazovObrazku Then j = frmSpracovavanie.NazovObrazku
            For i = 0 To j + 5
                If File.Exists(frmSpracovavanie.Adresa2 & "Foto\" & i & ".jpg") Then File.Delete(frmSpracovavanie.Adresa2 & "Foto\" & i & ".jpg")
            Next i
        End If

50:     frmSpracovavanie = Nothing
51:     End
    End Sub

    Private Sub frmZavierka_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
1:      'Zafarbi
2:      On Error GoTo Chyba
3:      Call mdlColors.Skining(Me)
4:      If VizualUprava Then
5:          Call mdlColors.Sizing(Me)
6:          Me.CenterToParent()
7:          VizualUprava = False
8:      End If
9:
10:     lblPociatStav.Text = CStr(frmPocStav.PocStav)
11:     lblDnesTrzba.Text = CStr(DnesnaTrzba2)
12:     lblDnesStrn.Text = CStr(DnesStorno2)
13:     lblSpolu.Text = CStr(frmPocStav.PocStav + DnesnaTrzba2 - DnesStorno2 + DnesBezhot - DnesBezhStor2)
14:     lblUserName.Text = frmLogin.UserName
15:     lblTrzbPoklad.Text = TrzbaPokladnika
16:     lblDnesBezhot.Text = DnesBezhot
17:     lblPokladBezhot.Text = PokladnBezhot
        lblDnesBezhotStor.Text = DnesBezhStor2

        Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub
End Class