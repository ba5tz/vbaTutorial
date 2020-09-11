'
' ---------------------------------------------------------
'| YouTube Channel      : Http://youtube.com/andisetiadii  |
' ---------------------------------------------------------
'   _             _ _   __      _   _           _ _
'  /_\  _ __   __| (_) / _\ ___| |_(_) __ _  __| (_)
' //_\\| '_ \ / _` | | \ \ / _ \ __| |/ _` |/ _` | |
'/  _  \ | | | (_| | | _\ \  __/ |_| | (_| | (_| | |
'\_/ \_/_| |_|\__,_|_| \__/\___|\__|_|\__,_|\__,_|_|
'
'           Auth    : Andi Setiadi
'           Date    : 11 September 2020
'           About   : Kirim Email dengan Akun Gmail dari Ms. Excel

Sub KirimGmail()
Dim CDO_Mail As Object
Dim CDO_Config As Object
Dim SMTP_Config As Variant
Dim Schema As String

On Error GoTo kesalahan
Set CDO_Mail = CreateObject("CDO.Message")
Set CDO_Config = CreateObject("CDO.configuration")
CDO_Config.Load -1

Set SMTP_Config = CDO_Config.Fields
Schema = "http://schemas.microsoft.com/cdo/configuration/"

With SMTP_Config
    .Item(Schema & "sendusing") = 2 'untuk port
    .Item(Schema & "smtpserver") = "smtp.gmail.com"
    .Item(Schema & "smtpserverport") = 465
    .Item(Schema & "smtpauthenticate") = 1
    .Item(Schema & "sendusername") = "xxxxxx@gmail.com"   'diisi dengan Alamat Email Gmail
    .Item(Schema & "sendpassword") = "password"           'diisi dengan Password Gmail
    .Item(Schema & "smtpusessl") = True
    .Update
End With

With CDO_Mail
    .configuration = CDO_Config
    
    .Subject = "Kirim Dari Excel"
    .From = "xxxxxx@gmail.com"    'diisi dengan Alamat Email Gmail
    .to = "email@tujuan.com"      'diisi dengan email tujuan
    .CC = ""
    .bcc = ""
    .textbody = "Hallo, email ini dikirim dari Excel"  'Isi Pesan
    .send
End With
Exit Sub

kesalahan:
If Err.Description <> "" Then MsgBox Err.Description
End Sub
