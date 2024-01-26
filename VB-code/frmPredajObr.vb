Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.IO

Public Class frmPredajObr

    Dim VizualUprava As Boolean = True
    Dim Cas As Integer
    Public host As System.Windows.Forms.Form

    Private FarbaZakl As String = CStr(&H8080) '&HC000&
    Private FarbaPredane As String = CStr(&H8080FF) '&H8080FF  '  &HC0&
    Private FarbaZlava As String = CStr(&HC0E0FF)
    Private FarbaVyber As String = CStr(&HFFC0C0) ' &H80FFFF
    Private FarbaRezerv As Integer = RGB(111, 49, 152)
    Private FarbaSluzob As Integer = RGB(230, 230, 230)

    Private Sub frmPredajObr_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Zafarbi
5:      Call mdlColors.Skining(Me)
6:      If VizualUprava Then
7:          Call mdlColors.Sizing(Me)
8:          Me.CenterToScreen()
9:          VizualUprava = False
10:     End If

        'Farby legendy
48:     btnVolVz.BackColor = System.Drawing.ColorTranslator.FromOle(CInt(FarbaZakl))
49:     btnPredVz.BackColor = System.Drawing.ColorTranslator.FromOle(CInt(FarbaPredane))
50:     btnRezVz.BackColor = System.Drawing.ColorTranslator.FromOle(CInt(FarbaRezerv))
51:     btnZlvVz.BackColor = System.Drawing.ColorTranslator.FromOle(CInt(FarbaZlava))

        If Directory.Exists(frmSpracovavanie.Adresa2 & "Foto") = False Then
            Directory.CreateDirectory(frmSpracovavanie.Adresa2 & "Foto")
        End If

    End Sub

    Private Sub ZmenaObrazku()
        On Error Resume Next
        Dim SubINI As String = ""
        Timer1.Enabled = False
        Cas = 0

        picSala.Image = Nothing

        If host.Name = "frmPredaj" Then
            SubINI = "Data\map1.ini"
12:     ElseIf host.Name = "frmPredaj2" Then
            SubINI = "Data\map2.ini"
17:     ElseIf host.Name = "frmPredaj3" Then
            SubINI = "Data\map3.ini"
22:     ElseIf host.Name = "frmPredaj4" Then
            SubINI = "Data\map4.ini"
27:     ElseIf host.Name = "frmPredaj5" Then
            SubINI = "Data\map5.ini"
32:     End If

        Dim objIniFile As New IniFile(frmSpracovavanie.Adresa2 & SubINI)

        Dim FotoHeight, FotoWidth, zLeftu, zTopu, PozFLeft, PozFTop As Integer
        FotoHeight = objIniFile.GetInteger("Foto", "FotoHeight", 0)
        FotoWidth = objIniFile.GetInteger("Foto", "FotoWidth", 0)
        PozFLeft = objIniFile.GetInteger("Foto", "PozFLeft", 290)
        PozFTop = objIniFile.GetInteger("Foto", "PozFTop", 20)
        zLeftu = objIniFile.GetInteger("Foto", "zLeftu", 0) * frmSpracovavanie.Zvacsenie + host.Left
        zTopu = objIniFile.GetInteger("Foto", "zTopu", 0) * frmSpracovavanie.Zvacsenie + host.Top

        If FotoHeight = 0 Then
            picSala.Height = host.Height
            picSala.Width = host.Width
            MsgBox("Pre lepsie fungovanie fotografovania mapy saly volajte na cislo 0944 068 485.", MsgBoxStyle.Information)
        Else
            picSala.Height = FotoHeight * frmSpracovavanie.Zvacsenie
            picSala.Width = FotoWidth * frmSpracovavanie.Zvacsenie
        End If
        gbxFoto.Height = picSala.Height + 25 * frmSpracovavanie.Zvacsenie
        gbxFoto.Width = picSala.Width + 20 * frmSpracovavanie.Zvacsenie
        picSala.Left = 10 * frmSpracovavanie.Zvacsenie
        picSala.Top = 15 * frmSpracovavanie.Zvacsenie
        gbxFoto.Top = (PozFTop - 10) * frmSpracovavanie.Zvacsenie
        gbxFoto.Left = (PozFLeft - 10) * frmSpracovavanie.Zvacsenie

        Dim bmpSS As Bitmap
        Dim gfxSS As Graphics
        bmpSS = New Bitmap(picSala.Width, picSala.Height, PixelFormat.Format32bppArgb)
        gfxSS = Graphics.FromImage(bmpSS)
        gfxSS.CopyFromScreen(zLeftu, zTopu, 0, 0, host.Size)
        bmpSS.Save(frmSpracovavanie.Adresa2 & "Foto\" & frmSpracovavanie.NazovObrazku & ".jpg", ImageFormat.Jpeg)

        picSala.Image = System.Drawing.Image.FromFile(frmSpracovavanie.Adresa2 & "Foto\" & frmSpracovavanie.NazovObrazku & ".jpg")

        If frmSpracovavanie.NazovObrazku > 2 Then System.IO.File.Delete(frmSpracovavanie.Adresa2 & "Foto\" & frmSpracovavanie.NazovObrazku - 2 & ".jpg")

        frmSpracovavanie.NazovObrazku += 1

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Cas += 1
        If Cas > 3 Then
            Call ZmenaObrazku()
        End If
    End Sub
End Class