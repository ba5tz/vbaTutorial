'Fungsi Untuk Parse JSON
'------------------------------------------------------------------
Function GetJadwal(tanggal as string) As Dictionary
  Dim objHTTP As Object

  Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
      URL = "https://api.pray.zone/v2/times/day.json?city=tasikmalaya&date=" & tanggal

  objHTTP.Open "GET", URL, False
  objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
  objHTTP.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
  objHTTP.Send

  Dim JSON As Object
  Set JSON = JsonConverter.ParseJson(objHTTP.responseText)
  Set GetJadwal = JSON("results")("datetime")(1)("times")
End Function
      

'UserForm1 Module
'------------------------------------------------------------------
Private Sub UserForm_Initialize()
  Dim Jadwal As New Dictionary

  Set Jadwal = GetJadwal(format(date, "YYYY-MM-DD"))
  LBImsak = Jadwal("Imsak")
  LBShubuh = Jadwal("Fajr")
  LBDzuhur = Jadwal("Dhuhr")
  LBAshar = Jadwal("Asr")
  LBMaghrib = Jadwal("Maghrib")
  LbIsya = Jadwal("Isha")
End Sub


Private Sub UserForm_Activate()
  Berhenti = False
  Do Until Berhenti
      LBJam.Caption = Format(Time, "hh:mm:ss am/pm")
      LBTanggal.Caption = WorksheetFunction.Text(Date, "[$-0421] DDDD, DD MMMM YYYY")
      DoEvents
  Loop
  End Sub

  Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  Berhenti = True
End Sub
