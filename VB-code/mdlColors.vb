Module mdlColors
    Sub Skining(ByRef frm As System.Windows.Forms.Form)
1:      On Error GoTo Chyba
2:      'Zafarbi dany form
3:      If frmSpracovavanie.FSchema = "Windows" Then Exit Sub
4:      Dim ctrl As System.Windows.Forms.Control
5:
6:      System.Drawing.ColorTranslator.FromOle(CInt(frmSpracovavanie.BckClr))
7:      frm.BackgroundImageLayout = 3
8:      frm.BackgroundImage = System.Drawing.Image.FromFile(frmSpracovavanie.Pozadie)
9:
10:     For Each ctrl In frm.Controls
11:         If TypeOf ctrl Is System.Windows.Forms.Button Then ctrl.BackColor = System.Drawing.ColorTranslator.FromOle(CInt(frmSpracovavanie.BckClr))
12:         If TypeOf ctrl Is System.Windows.Forms.GroupBox Then ctrl.BackColor = System.Drawing.ColorTranslator.FromOle(CInt(frmSpracovavanie.BckClr))
13:         If TypeOf ctrl Is System.Windows.Forms.RadioButton Then ctrl.BackColor = System.Drawing.ColorTranslator.FromOle(CInt(frmSpracovavanie.BckClr))
14:         If TypeOf ctrl Is System.Windows.Forms.Label Then ctrl.BackColor = Color.Transparent
            If TypeOf ctrl Is System.Windows.Forms.CheckBox Then ctrl.BackColor = Color.Transparent
15:     Next ctrl
16:
        Exit Sub
Chyba:
        Chyby.Chyba("mdlColors", "Skining", ErrorToString(), Err.Number, Err.Erl)
    End Sub

    Sub Sizing(ByRef frm As System.Windows.Forms.Form)
1:      On Error GoTo Chyba
2:      'Upravuje velkosti (aby Okno na obrazovke 1024x768 nebolo pixel v rohu)
3:
4:      Dim ctrl As System.Windows.Forms.Control
5:      Dim ctrlGroupB As System.Windows.Forms.Control
6:      Dim ctrlGroupB2 As System.Windows.Forms.Control
7:      Dim ctrlTabP As System.Windows.Forms.Control
8:
9:      frm.Height *= frmSpracovavanie.Zvacsenie
10:     frm.Width *= frmSpracovavanie.Zvacsenie
11:
12:     For Each ctrl In frm.Controls
13:         ctrl.Left *= frmSpracovavanie.Zvacsenie
14:         ctrl.Top *= frmSpracovavanie.Zvacsenie
15:         ctrl.Height *= frmSpracovavanie.Zvacsenie
16:         ctrl.Width *= frmSpracovavanie.Zvacsenie
17:         ctrl.Font = New Font(ctrl.Font.Name, ctrl.Font.Size * frmSpracovavanie.Zvacsenie, ctrl.Font.Style, ctrl.Font.Unit)
18:         If TypeOf ctrl Is System.Windows.Forms.GroupBox Then
19:             For Each ctrlGroupB In ctrl.Controls
20:                 ctrlGroupB.Left *= frmSpracovavanie.Zvacsenie
21:                 ctrlGroupB.Top *= frmSpracovavanie.Zvacsenie
22:                 ctrlGroupB.Height *= frmSpracovavanie.Zvacsenie
23:                 ctrlGroupB.Width *= frmSpracovavanie.Zvacsenie
24:                 If TypeOf ctrlGroupB Is System.Windows.Forms.GroupBox Then
25:                     For Each ctrlGroupB2 In ctrlGroupB.Controls
26:                         ctrlGroupB2.Left *= frmSpracovavanie.Zvacsenie
27:                         ctrlGroupB2.Top *= frmSpracovavanie.Zvacsenie
28:                         ctrlGroupB2.Height *= frmSpracovavanie.Zvacsenie
29:                         ctrlGroupB2.Width *= frmSpracovavanie.Zvacsenie
30:                     Next ctrlGroupB2
31:                 End If
32:             Next ctrlGroupB
33:         End If
34:         If TypeOf ctrl Is System.Windows.Forms.TabControl Then
35:             For Each ctrlTabP In ctrl.Controls
36:                 ctrlTabP.Left *= frmSpracovavanie.Zvacsenie
37:                 ctrlTabP.Top *= frmSpracovavanie.Zvacsenie
38:                 ctrlTabP.Height *= frmSpracovavanie.Zvacsenie
39:                 ctrlTabP.Width *= frmSpracovavanie.Zvacsenie
40:                 If TypeOf ctrlTabP Is System.Windows.Forms.TabPage Then
41:                     For Each ctrlGroupB In ctrlTabP.Controls
42:                         ctrlGroupB.Left *= frmSpracovavanie.Zvacsenie
43:                         ctrlGroupB.Top *= frmSpracovavanie.Zvacsenie
44:                         ctrlGroupB.Height *= frmSpracovavanie.Zvacsenie
45:                         ctrlGroupB.Width *= frmSpracovavanie.Zvacsenie
46:                         If TypeOf ctrlGroupB Is System.Windows.Forms.GroupBox Then
47:                             For Each ctrlGroupB2 In ctrlGroupB.Controls
48:                                 ctrlGroupB2.Left *= frmSpracovavanie.Zvacsenie
49:                                 ctrlGroupB2.Top *= frmSpracovavanie.Zvacsenie
50:                                 ctrlGroupB2.Height *= frmSpracovavanie.Zvacsenie
51:                                 ctrlGroupB2.Width *= frmSpracovavanie.Zvacsenie
52:                             Next ctrlGroupB2
53:                         End If
54:                     Next ctrlGroupB
55:                 End If
56:             Next ctrlTabP
57:         End If
58:     Next ctrl
59:
60:     'Overi ci nemanipulovali s datumom do minulosti
61:     '  If frmSpracovavanie.DatumSpustenia > Now Then
62:     'MsgBox("Zistena manipulacia s datumom.", MsgBoxStyle.Critical)
63:     'End
64:     'End If
65:
66:     'System.Windows.Forms.Screen()
67:     'Simply display a messagebox with the resolution of the computers screen.
68:     'MsgBox(Screen.PrimaryScreen.Bounds.Size.ToString)
69:     ' Screen.PrimaryScreen.Bounds.Size.
70:     'If you had multiple displays, and you wanted to know the resolution of the display that contains your form, this code would do that for you.
71:     'MsgBox(Screen.FromControl(Me).Bounds.Size.ToString)
72:     'Use the MY shortcut in VB 05/08 to get the sacreen size...
73:     'MsgBox(My.Computer.Screen.Bounds.Size.ToString)
74:
75:     Exit Sub
Chyba:
        Chyby.Chyba("mdlColors", "Skining", ErrorToString(), Err.Number, Err.Erl)
    End Sub
End Module
