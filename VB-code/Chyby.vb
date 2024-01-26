Module Chyby
    Sub Chyba(ByRef Frm As String, ByVal Proc As String, ByVal Erro As String, ByVal ErroCis As Long, ByVal ErroLine As Integer)
        On Error Resume Next
        Dim s As String = ""

        s = Now & Chr(13) & Chr(10)
        s += "Form:         " & Frm & Chr(13) & Chr(10)
        s += "Aplikacia:    " & My.Application.Info.Title & " " & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & My.Application.Info.Version.Build & Chr(13) & Chr(10)
        s += "Sub:          " & Proc & Chr(13) & Chr(10)
        s += "Error:        " & Erro & Chr(13) & Chr(10)
        s += "Cislo chyby:  " & ErroCis & Chr(13) & Chr(10)
        s += "Cislo riadku: " & ErroLine & Chr(13) & Chr(10)

        If s = frmSpracovavanie.ChybaVystup Then Exit Sub Else frmSpracovavanie.ChybaVystup = s

        FileOpen(1, frmSpracovavanie.Adresa2 & "errorlog.pac", OpenMode.Append)
        'WriteLine(1, Now)
        'WriteLine(1, "Form:         " & Frm)
        'WriteLine(1, "Aplikacia:    " & My.Application.Info.Title & " " & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & My.Application.Info.Version.Build)
        'WriteLine(1, "Sub:          " & Proc)
        'WriteLine(1, "Error:        " & Erro)
        'WriteLine(1, "Cislo chyby:  " & ErroCis)
        'WriteLine(1, "Cislo riadku: " & ErroLine, Chr(10))
        WriteLine(1, s & Chr(13) & Chr(10))
        FileClose(1)

        Module1.WriteLog(Chr(13) & Chr(10) & Now & " - Padnutie softveru")

        'Call Module1.HASHdbKontrola("all")
        'Call Module1.HASHSubor("Data\data.pac")

        If InStr(Erro, "you are connected to the server") <> 0 Then
            frmSpracovavanie.Timer2_Tick()
        Else
            If frmLogin.UserName = "Programator" Then MsgBox(s)
            MsgBox("Doslo k chybe v programe. Na adresu pnagy11@gmail.com zaslite subor errorlog.pac" & Chr(13) & Chr(10) & Erro, MsgBoxStyle.Critical)
        End If

    End Sub

    Sub OpravaLogu()
        Dim x As Integer
        Dim HASH As String
        'Kontroluje ci logu bolo pridelene CRC
        'ALE ak v databaze zmazu CRC a cas odhlasenia daju rovnaky ako cas prihlasenia aj to opravi
        With frmSpracovavanie.LogBindingSource
            x = .Find("crc", "")
            If x = -1 Then Exit Sub
            .Position = x
            If .Current("Prihlasenie") = .Current("Odhlasenie") Then
                HASH = Module1.RetazecRiadku("Log", x)
                .Current("CRC") = Module1.HASHString(HASH)
                Module1.UpdateLog(x)
            End If
        End With
        Call Module1.HASHSubor("Data\data.pac")
    End Sub
End Module
