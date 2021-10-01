'Override Default SHift + F3 Excel
'Simpan di ThisWorkbook Module
'---------------------------------------------------------
Private Sub Workbook_Open()
Application.OnKey "+{F3}", "GantiCase"
End Sub



'Simpan Dalam Standar Module
'---------------------------------------------------------
Public Sub GantiCase()
Dim Ch1 As String
Dim Ch2 As String

Ch1 = Left(Selection(1, 1).Value, 1)
Ch2 = Mid(Selection(1, 1).Value, 2, 1)

For Each cell In Selection
    If Asc(Ch1) < 91 And Asc(Ch2) < 91 Then          'jika isi hurup besar
        cell.Value = LCase(cell.Value)
    ElseIf Asc(Ch1) > 91 And Asc(Ch2) > 91 Then     'jika isi, hurup kecil
        cell.Value = StrConv(cell.Value, vbProperCase)
    Else
        cell.Value = UCase(cell.Value)
    End If
Next cell
End Sub
