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
'           Author    : Andi Setiadi
'           Update    : 10 Oktober 2021
'           About     : Shortcut Shift + F3 Override
'
'       Cara Penggunaan
'       1. Simpan Script ke Masing-Masing Module
'       2. Tutup dan buka kembali Excel untuk mendapatkan effect

'---------------------------------------------------------
'Override Default SHift + F3 Excel
'Simpan di ThisWorkbook Module
'---------------------------------------------------------
Private Sub Workbook_Open()
Application.OnKey "+{F3}", "GantiCase"
End Sub


'---------------------------------------------------------
'Simpan Dalam Standar Module
'---------------------------------------------------------
Public Sub GantiCase()
Dim Ch1 As String
Dim Ch2 As String

Ch1 = Left(Selection(1, 1).Value, 1)
Ch2 = Mid(Selection(1, 1).Value, 2, 1)

For Each cell In Selection
    If Asc(Ch1) < 91 And Asc(Ch2) < 91 Then          
        cell.Value = LCase(cell.Value)
    ElseIf Asc(Ch1) > 91 And Asc(Ch2) > 91 Then    
        cell.Value = StrConv(cell.Value, vbProperCase)
    Else
        cell.Value = UCase(cell.Value)
    End If
Next cell
End Sub
