Attribute VB_Name = "Email"
Option Explicit

Private Sub �ǥ�Gmail�H�H()
    '====================================================================================================================================================
    '�b�ϥ�Gmail�H�H�e�A�Х���Google���u�C�w�������ε{���v�������A�]�w�w���ʸ��C�����ε{���s���v���@�}�ҡA�p�G�S�����o�Ӱʧ@�A�O�L�k�g��Gmail���H�H���ʧ@
    'https://myaccount.google.com/security?pli=1
    '�ѩ�ݨϥΨ�CDO����A�b�s�gVBA�{���X�e�A�����]�w�ޥ�"Microsoft CDO for Windows 2000 Library"
    '====================================================================================================================================================
    Dim Mail As New Message
    Dim config As Configuration
    Set config = Mail.Configuration

    config(cdoSMTPAuthenticate) = cdoBasic
    config(cdoSMTPUseSSL) = True                        '�]�wSSL�[�K�ǰe
    config(cdoSMTPServer) = "smtp.gmail.com"            '�]�wsmtp�D��
    config(cdoSMTPServerPort) = 25                      '�]�wstmp port�A�w�]��25�A�]�i�ϥ�465
    config(cdoSendUsingMethod) = cdoSendUsingPort
    config(cdoSendUserName) = "jay.cc.hsieh@gmail.com"  '��g�z��gmail�l���}
    config(cdoSendPassword) = ""           '��g�W�z�l���}���ϥΪ̱K�X
    config.Fields.Update

    With Mail       '�H���H
        .To = "jay.cc.hsieh@gmail.com"                  '�����
        .From = "jay.cc.hsieh@gmail.com"                '�H���
        .CC = ""                                        '���󦬥��
        .Subject = "VBA�z�LGmail�Hmail"                 '�H��D��
        .HTMLBody = "���դ��e"                          '�H�󤺮e
        .BodyPart.Charset = "utf-8"                     '���e�s�X
        .HTMLBodyPart.Charset = "utf-8"                 '�����s�X�A�w������H�c�Hutf-8�s�X�覡�~�ਣ�줺�e�������@�B�z�A�pHotmail�H�c
        '.AddAttachment "�ɮ׸��|"                      '����s���m

        On Error Resume Next
        .Send                                           '�}�l��email
        If Err.Number <> 0 Then
            MsgBox Err.Description, vbCritical, "�H��L�k�H�X"
            Exit Sub
        Else
             MsgBox "�H��w�H�X", vbInformation, "�H��H�X���A"
        End If
    End With
End Sub

Sub �ǥ�Hotmail�H�H()
    '============================================================================================
    '�ѩ�ݨϥΨ�CDO����A�b�s�gVBA�{���X�e�A�����]�w�ޥ�"Microsoft CDO for Windows 2000 Library"
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
        .Subject = "VBA�z�LHotmail�Hmail"
        .From = "jay_hsieh@livemail.tw"
        .To = "jay.cc.hsieh@gmail.com"
        .CC = ""
        .HTMLBody = "���դ��e"
        .BodyPart.Charset = "utf-8"
        .HTMLBodyPart.Charset = "utf-8"
        .Send
    End With
    MsgBox "�H��w�H�X", vbInformation, "�H�X"

    Set Mail = Nothing
End Sub

Sub �ǥ�Yahoo�H�H()
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
        .Subject = "VBA�zYahoo�Hmail"
        .From = ""
        .To = "jay.cc.hsieh@gmail.com"
        .CC = ""
        .HTMLBody = "���դ��e"
        .BodyPart.Charset = "utf-8"
        .HTMLBodyPart.Charset = "utf-8"
        .Send
    End With

    MsgBox "�H��w�H�X", vbInformation, "�H�X"

    Set Mail = Nothing
End Sub

