Function JOINIF(CriteriaRange As Range, _
                Criteria As Variant, _
                GabungRange As Range, _
                Optional Delimiter As String = ",") As Variant

    Dim j As Long
    Dim TempString As String: TempString = ""

    On Error GoTo Kesalahan
    
    If CriteriaRange.Count <> GabungRange.Count Then
        JOINIF = CVErr(xlErrRef)
        Exit Function
    End If

    For j = 1 To CriteriaRange.Count
        If CriteriaRange.Cells(j).Value = Criteria Then
            TempString = TempString & Delimiter & GabungRange.Cells(j).Value
        End If
    Next j

    If Not TempString = "" Then
        TempString = Mid(TempString, Len(Delimiter) + 1)
    End If

    JOINIF = TempString
    Exit Function

Kesalahan:
JOINIF = CVErr(xlErrValue)

End Function
