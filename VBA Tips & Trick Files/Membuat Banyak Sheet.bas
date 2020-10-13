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
'           Date    : 13 Oktober 2020
'           About   : Membuat Banyak Sheet dengan Nama dari Range

Sub BuatSheet()
Dim ArrSheet As Variant
Dim Baris As Long
Dim Sht as Variant

Baris = Sheet1.Range("A1").End(xlDown).Row                                          'Menghitung Baris Akhir
ArrSheet = Application.Transpose(Sheet1.Range("A2:A" & Baris))                      'Array Nama Sheet yang akan dibuat
For Each Sht In ArrSheet                                                            'Looping Array
    If CheckSheet(Sht) Then Sheets.Add(after:=Sheets(Sheets.Count)).Name = Sht
Next

End Sub

Function CheckSheet(Sht As Variant) As Boolean   'Fungsi Untuk Mengecek Nama Sheet sudah digunakan atau belum
Dim Sh As Worksheet
On Error Resume Next
Set Sh = Sheets(Sht)                            'jika ada akan menjadi object Sh jika belum ada akan Error (Object Nothing)
CheckSheet = Sh Is Nothing            
Err.Clear
End Function
