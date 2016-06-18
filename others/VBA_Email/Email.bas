Attribute VB_Name = "Email"
Option Explicit

Private Sub 藉由Gmail寄信()
    '====================================================================================================================================================
    '在使用Gmail寄信前，請先到Google的「低安全性應用程式」網頁中，設定安全性較低的應用程式存取權限作開啟，如果沒有做這個動作，是無法經由Gmail做寄信的動作
    'https://myaccount.google.com/security?pli=1
    '由於需使用到CDO物件，在編寫VBA程式碼前，須先設定引用"Microsoft CDO for Windows 2000 Library"
    '====================================================================================================================================================
    Dim Mail As New Message
    Dim config As Configuration
    Set config = Mail.Configuration

    config(cdoSMTPAuthenticate) = cdoBasic
    config(cdoSMTPUseSSL) = True                        '設定SSL加密傳送
    config(cdoSMTPServer) = "smtp.gmail.com"            '設定smtp主機
    config(cdoSMTPServerPort) = 25                      '設定stmp port，預設為25，也可使用465
    config(cdoSendUsingMethod) = cdoSendUsingPort
    config(cdoSendUserName) = "jay.cc.hsieh@gmail.com"  '填寫您的gmail郵件位址
    config(cdoSendPassword) = ""           '填寫上述郵件位址的使用者密碼
    config.Fields.Update

    With Mail       '寄件對象
        .To = "jay.cc.hsieh@gmail.com"                  '收件者
        .From = "jay.cc.hsieh@gmail.com"                '寄件者
        .CC = ""                                        '附件收件者
        .Subject = "VBA透過Gmail寄mail"                 '信件主旨
        .HTMLBody = "測試內容"                          '信件內容
        .BodyPart.Charset = "utf-8"                     '內容編碼
        .HTMLBodyPart.Charset = "utf-8"                 '網頁編碼，針對網路信箱以utf-8編碼方式才能見到內容的部分作處理，如Hotmail信箱
        '.AddAttachment "檔案路徑"                      '附件存放位置

        On Error Resume Next
        .Send                                           '開始船送mail
        If Err.Number <> 0 Then
            MsgBox Err.Description, vbCritical, "信件無法寄出"
            Exit Sub
        Else
             MsgBox "信件已寄出", vbInformation, "信件寄出狀態"
        End If
    End With
End Sub

Sub 藉由Hotmail寄信()
    '============================================================================================
    '由於需使用到CDO物件，在編寫VBA程式碼前，須先設定引用"Microsoft CDO for Windows 2000 Library"
    '============================================================================================
    Dim Mail As CDO.Message
    Set Mail = New CDO.Message
    With Mail.Configuration.Fields
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp-mail.outlook.com"
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "jay_hsieh@livemail.tw"
        .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = ""
        .Update
    End With

    With Mail
        .Subject = "VBA透過Hotmail寄mail"
        .From = "jay_hsieh@livemail.tw"
        .To = "jay.cc.hsieh@gmail.com"
        .CC = ""
        .HTMLBody = "測試內容"
        .BodyPart.Charset = "utf-8"
        .HTMLBodyPart.Charset = "utf-8"
        .Send
    End With
    MsgBox "信件已寄出", vbInformation, "寄出"

    Set Mail = Nothing
End Sub

Sub 藉由Yahoo寄信()
    Dim Mail As CDO.Message
    Dim sURL As String
    Set Mail = New CDO.Message
    sURL = "http://schemas.microsoft.com/cdo/configuration/"
    With Mail.Configuration.Fields
        .Item(sURL & "smtpusessl") = True
        .Item(sURL & "smtpauthenticate") = 1
        .Item(sURL & "smtpserver") = "smtp.mail.yahoo.com"
        .Item(sURL & "smtpserverport") = 25
        .Item(sURL & "sendusing") = 2
        .Item(sURL & "sendusername") = ""
        .Item(sURL & "sendpassword") = ""
        .Update
    End With

    With Mail
        .Subject = "VBA透Yahoo寄mail"
        .From = ""
        .To = "jay.cc.hsieh@gmail.com"
        .CC = ""
        .HTMLBody = "測試內容"
        .BodyPart.Charset = "utf-8"
        .HTMLBodyPart.Charset = "utf-8"
        .Send
    End With

    MsgBox "信件已寄出", vbInformation, "寄出"

    Set Mail = Nothing
End Sub

