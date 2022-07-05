Sub SHKirim_Click()
Dim KirimJson As String
Dim title As String
Dim body As String
Dim request As Object
Dim URL As String, Token As String

URL = "https://api.pushbullet.com/v2/pushes"
  Token = "Token"

title = Sheet1.Range("D3").Value
body = Sheet1.Range("D4").Value

KirimJson = "{""type"":""note"",""title"":""" & title & """,""body"":""" & body & """}"

Set request = CreateObject("MSXML2.XMLHTTP")

request.Open "POST", URL, False

request.setrequestheader "Autorization", "Bearer " & Token
request.setrequestheader "Content-type", "application/json"

request.send (KirimJson)

End Sub
