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
'           Date    : 21 September 2020
'           About   : Kirim Email dengan Outlook dari Ms. Excel
'
' # Early Binding => silahkan tambahkan Microsoft Outlook Object Library Pada References
'

Sub KirimEmail()
Dim Outl As Outlook.Application
Dim Msg As Outlook.MailItem

On Error GoTo kesalahan

Set Outl = New Outlook.Application
Set Msg = Outl.CreateItem(0) 

With Msg
    .To = Sheet1.Range("C2").Value
    .cc = Sheet1.Range("C3").Value
    .bcc = Sheet1.Range("C4").Value
    .Subject = Sheet1.Range("C5").Value
    .Attachments.add "LokasiFile"
    .Body = Sheet1.Range("C6").Value
    .send
End With

MsgBox "Email Sudah Berhasil terkirim"
Exit Sub

kesalahan:
MsgBox "Email Gagal dikirim " & vbNewLine & "Error :" & Err.Description

End Sub
