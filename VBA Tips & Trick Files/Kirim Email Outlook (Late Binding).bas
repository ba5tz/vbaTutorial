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

Sub KirimEmail()
Dim Outl As Object   'Deklarasi Object
Dim Msg As Object

'Set Object Outlook
Set Outl = CreateObject("Outlook.Application")
Set Msg = Outl.CreateItem(0)

' isi Item
With Msg
  .To = Sheet1.Range("B2").Value  
  .CC = Sheet1.Range("B3").Value
  .bcc = Sheet1.Range("B4").Value
  .Subject = Sheet1.Range("B5").Value
  .Body = Sheet1.Range("B6").Value
  .Attachments.Add "lokasiFile"  'jika ada
  .Send
End with
End Sub
