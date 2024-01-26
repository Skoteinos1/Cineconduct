Public Class frmPresad

    Dim VizualUprava As Boolean = True

    Public DatumPresad As Date
    Public CasPresad As Date
    Public SalaPresad As Integer
    Public MstOdPresad As Integer
    Public MstDoPresad As Integer
    Public NazMstPresad As String
    Public ZlavPresad As Integer
    Public SumaPresad As Double

    Private Sub frmPresad_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtCisListk.Text = ""
        gbxStorno.Visible = False
        btnPresad.Enabled = False
        btnOK.Enabled = False

        'Zafarbi
        Call mdlColors.Skining(Me)
        If VizualUprava Then
            Call mdlColors.Sizing(Me)
            Me.CenterToParent()
            VizualUprava = False
        End If
    End Sub

    Private Sub txtCisListk_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCisListk.KeyPress
        frmPredstavenia.txtCena_KeyPress(sender, e, Me)
    End Sub

    Private Sub txtCisListk_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCisListk.TextChanged
        btnPresad.Enabled = False
        gbxStorno.Visible = False

        If txtCisListk.Text = "" Then btnOK.Enabled = False Else btnOK.Enabled = True
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
1:      If frmSpracovavanie.SietVerzia Then frmSpracovavanie.ListkyTableAdapter.Fill(frmSpracovavanie.DataSet1.Listky)
2:      On Error GoTo Chyba
3:      Dim Nasiel As Boolean = False
4:      Dim NasielStorno As Boolean = False
5:      Dim DatPredst, CasPredstav As Date
6:      Dim i As Integer
7:
8:      With frmSpracovavanie.ListkyBindingSource
9:          i = .Find("Kod", txtCisListk.Text)
10:         If i <> -1 Then
11:             'Zisti ci je a ci bol stornovany
12:             .Position = i
13:             For i = 0 To .Count - 1
14:                 If .Current("kod") = txtCisListk.Text Then
15:                     If InStr(.Current("Stav"), "Storn") <> 0 Then
16:                         NasielStorno = True
17:                         Exit For
18:                     ElseIf InStr(.Current("Stav"), "Predane") <> 0 Or InStr(.Current("Stav"), "Bezhotov") <> 0 Then
19:                         Nasiel = True
20:                     End If
21:                 End If
22:                 If .Position = .Count - 1 Then Exit For
23:                 .Position += 1
24:             Next i
25:         End If
26:
27:         If Nasiel And NasielStorno = False Then
28:             .Position = .Find("Kod", txtCisListk.Text)
29:             'Zisti ci predstavenie bolo
30:             DatPredst = Format(CDate(.Current("Datumpredst")), "dd.MM.yyyy")
31:             CasPredstav = Format(CDate(.Current("Caspredst")), "HH:mm")
32:
33:             If SumaPresad <> .Current("Suma") Then MsgBox("Ceny listkov sa nezhoduju.") : Exit Sub
34:             If (MstDoPresad - MstOdPresad) <> (.Current("miestodo") - .Current("miestoOd")) Then MsgBox("Pocty vstupov sa nezhoduju.") : Exit Sub
35:             'If ZlavPresad <> .Current("Zlavnenych") Then MsgBox("Pocty zlavnenych vstupov sa nezhoduju.") : Exit Sub
36:
37:             If DatPredst >= System.DateTime.FromOADate(Today.ToOADate - 8) Then
38:                 'odblokuje storno
39:                 gbxStorno.Visible = True
40:                 btnPresad.Enabled = True
41:
42:                 'Zobrazi info v textboxoch
43:                 lblKod.Text = .Current("kod")
44:                 lblDatumPredst.Text = Format(CDate(.Current("Datumpredst")), "dd.MM.yyyy")
45:                 lblCasPredst.Text = Format(CDate(.Current("Caspredst")), "HH:mm")
46:                 lblOsob.Text = CInt(.Current("miestodo")) - CInt(.Current("miestood")) + 1
47:                 lblZlavnenych.Text = .Current("Zlavnenych")
48:                 lblSuma.Text = .Current("Suma")
49:                 lblMiesta.Text = .Current("NazovMiest")
50:
51:                 If InStr(.Current("Stav"), "Predane") <> 0 Then
52:                     frmVybPredst.TlacitBezhotov = False
53:                 ElseIf InStr(.Current("Stav"), "Bezhotov") <> 0 Then
54:                     frmVybPredst.TlacitBezhotov = True
55:                 End If
56:
57:             Else
58:                 MsgBox("Predstavenie uz bolo." & Chr(13) & "Nie je mozné presadit.", MsgBoxStyle.Exclamation)
59:             End If
60:         Else
61:             MsgBox("Lístok sa nenachádza v pamäti, alebo uz bol stornovaný." & Chr(13) & "Nie je mozné presadit.", MsgBoxStyle.Exclamation)
62:         End If
63:     End With
64:
        Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Private Sub btnPresad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPresad.Click
1:      If frmSpracovavanie.SietVerzia Then
2:          frmSpracovavanie.PredstaveniaTableAdapter.Fill(frmSpracovavanie.DataSet1.Predstavenia)
3:          frmSpracovavanie.FilmyTableAdapter.Fill(frmSpracovavanie.DataSet1.Filmy)
4:          frmSpracovavanie.ListkyTableAdapter.Fill(frmSpracovavanie.DataSet1.Listky)
5:      End If
6:      On Error GoTo Chyba
7:
8:      Dim p2, p1, p3, PosPreds, PosFlm As Short
9:      Dim s, s1, s2, HASH As String
10:     'Dim Poloha As Integer
11:
12:     'Vyhlada spravne zaznamy
13:     PosPreds = frmSpracovavanie.PredstaveniaBindingSource.Find("Datum", lblDatumPredst.Text & " " & lblCasPredst.Text)
14:     If PosPreds <> -1 Then frmSpracovavanie.PredstaveniaBindingSource.Position = PosPreds Else MsgBox("Nenasiel som predstavenie!", MsgBoxStyle.Critical) : Exit Sub
15:
16:     PosFlm = frmSpracovavanie.FilmyBindingSource.Find("Film", frmSpracovavanie.PredstaveniaBindingSource.Current("NazovFilmu"))
17:     If PosFlm <> -1 Then frmSpracovavanie.FilmyBindingSource.Position = PosFlm Else MsgBox("Nenasiel som film!", MsgBoxStyle.Critical) : Exit Sub
18:
19:     If MsgBox("Naozaj presadit lístok?", MsgBoxStyle.YesNo + MsgBoxStyle.Question) = MsgBoxResult.No Then txtCisListk.Text = "" : Exit Sub
20:
21:     Module1.WriteLog(Chr(13) & Chr(10) & Now & " - Presadenie listka (Os:" & lblOsob.Text & " C:" & lblSuma.Text & "E " & ")")
22:
23:     With frmSpracovavanie.PredstaveniaBindingSource
24:         'Zmenit 2 na 1, odpocita trzbu a predaj
25:         s = .Current("miesta")
26:         p2 = CInt(frmSpracovavanie.ListkyBindingSource.Current("miestood"))
27:         p3 = CInt(frmSpracovavanie.ListkyBindingSource.Current("miestodo"))
28:         s1 = LSet(s, p2 - 1)
29:         s2 = Mid(s, p3 + 1)
30:         s = s1 & StrDup(p3 - p2 + 1, "1") & s2
31:         .Current("miesta") = s
32:         .Current("Predaj") = CDbl(.Current("Predaj")) - CDbl(lblOsob.Text)
33:         .Current("TrzbaPredstavenia") = CDbl(.Current("TrzbaPredstavenia")) - CDbl(lblSuma.Text)
34:         HASH = Module1.RetazecRiadku("Predstavenia", .Position)
35:         .Current("crc") = Module1.HASHString(HASH)
36:         Module1.UpdatePredstavenia(PosPreds)
37:     End With
38:
39:     With frmSpracovavanie.FilmyBindingSource
40:         .Current("Predanelistky") = CDbl(.Current("Predanelistky")) - CDbl(lblOsob.Text)
41:         .Current("TrzbaFilmu") = CDbl(.Current("TrzbaFilmu")) - CDbl(lblSuma.Text)
42:         HASH = Module1.RetazecRiadku("Filmy", PosFlm)
43:         .Current("crc") = Module1.HASHString(HASH)
44:         Module1.UpdateFilmy(PosFlm)
45:
46:         frmSpracovavanie.FilmyBindingSource.Position = frmVybPredst.PozFilm
47:         .Current("predanelistky") += (MstDoPresad - MstOdPresad + 1)
48:         .Current("TrzbaFilmu") += SumaPresad
49:         'Vytvori HASH pre kontrolu upravy riadku
50:         HASH = Module1.RetazecRiadku("Filmy", frmVybPredst.PozFilm)
51:         .Current("crc") = Module1.HASHString(HASH)
52:         Module1.UpdateFilmy(frmVybPredst.PozFilm)
53:         .Position = frmVybPredst.PozFilm
54:     End With
55:
56:     '****************Presadi*****************
57:     frmSpracovavanie.PredstaveniaBindingSource.Position = frmVybPredst.PozPredstav
58:
59:     'Tabulka predstaveni
60:     s = frmSpracovavanie.PredstaveniaBindingSource.Current("miesta")
61:     'Da 2ky na predane miesta a 3ky na zlavy
62:     s1 = LSet(s, MstOdPresad - 1)
63:     s2 = Mid(s, MstDoPresad + 1)
64:     s = s1 & StrDup(ZlavPresad, "3") & StrDup(MstDoPresad - MstOdPresad + 1 - ZlavPresad, "2") & s2
65:
66:     'Ulozi predaj
67:     With frmSpracovavanie.PredstaveniaBindingSource
68:         .Current("miesta") = s
69:         .Current("predaj") += (MstDoPresad - MstOdPresad + 1)
70:         .Current("trzbapredstavenia") += SumaPresad
71:         'Vytvori HASH pre kontrolu upravy riadku
72:         HASH = Module1.RetazecRiadku("Predstavenia", frmVybPredst.PozPredstav)
73:         .Current("crc") = Module1.HASHString(HASH)
74:         Module1.UpdatePredstavenia(frmVybPredst.PozPredstav)
75:         .Position = frmVybPredst.PozPredstav
76:         If s <> .Current("miesta") Then MsgBox("Chyba pri ukladani obsadenych miest. Obratte sa na Programatora.", MsgBoxStyle.Critical) : Exit Sub
77:     End With
78:
79:     'Upravi data v tabulke listkov
80:     With frmSpracovavanie.ListkyBindingSource
81:         .Current("datumpredst") = DatumPresad
82:         .Current("Caspredst") = CasPresad
83:         .Current("sala") = SalaPresad
84:         .Current("miestood") = MstOdPresad
85:         .Current("miestodo") = MstDoPresad
86:         .Current("nazovmiest") = NazMstPresad
87:         .Current("crc") = ""
88:         HASH = Module1.RetazecRiadku("Listky", .Position) '.Find("crc", ""))
89:         .Current("CRC") = Module1.HASHString(HASH)
90:         Module1.UpdateListky(.Position)
91:     End With
92:
93:     'Call Module1.HASHdbKontrola()
94:     Call Module1.HASHSubor("Data\data.pac")
95:
96:     'Vytlaci listok
97:     frmVybPredst.TlacitKodListka = lblKod.Text
98:     Call frmVybPredst.TlacitListok()
99:     Module1.WriteLog("   OK")
100:    frmSpracovavanie.Zaloha = True
101:    MsgBox("Skartujte stary listok.")
102:    Me.Close()
103:
        Exit Sub
Chyba:
        Chyby.Chyba(Me.Name, sender.ToString, ErrorToString(), Err.Number, Err.Erl)
    End Sub

   
End Class