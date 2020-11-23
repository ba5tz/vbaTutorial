
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
'           Date    : 23 November  2020
'           About   : Game Control Sederhana

Dim Obj As Shape                                          'Deklarasi Global
Dim Gas As Boolean

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
Set Obj = Shapes("Mobil")                                 'Deklarasi Object Shape

If Not Intersect(Target, Range("C2")) Is Nothing Then     'Untuk Maju
    If Not Gas Then
        Gas = True
        Call maju
    Else
        Gas = False
    End If
ElseIf Not Intersect(Target, Range("B3")) Is Nothing Then   'Untuk Belok kiri
    Obj.IncrementRotation -10
ElseIf Not Intersect(Target, Range("D3")) Is Nothing Then   'Untuk Belok Kanan
    Obj.IncrementRotation 10
ElseIf Not Intersect(Target, Range("C4")) Is Nothing Then   'Untuk Mundur
    If Not Gas Then
        Gas = True
        Call Mundur
    Else
        Gas = False
    End If
End If
Range("C3").Select
End Sub

Sub Maju()
GasMaju:
If Obj.Rotation >= 0 And Obj.Rotation < 180 Then
    If Obj.Rotation = 90 Then
        Obj.IncrementLeft 3
    Else
        Obj.IncrementLeft ((90 - Abs(90 - Obj.Rotation)) / 30)
    End If
Else
    If Obj.Rotation = 270 Then
        Obj.IncrementLeft -3
    Else
        Obj.IncrementLeft -((90 - Abs(270 - Obj.Rotation)) / 30)
    End If
End If

If Obj.Rotation > 270 And Obj.Rotation < 90 Then
    If Obj.Rotation = 0 Then
        Obj.IncrementTop -3
    Else
        If Obj.Rotation < 90 Then
            Obj.IncrementTop -((90 - Abs(0 - Obj.Rotation)) / 30)
        Else
            Obj.IncrementTop -((90 - Abs(360 - Obj.Rotation)) / 30)
        End If
    End If
Else
    If Obj.Rotation = 180 Then
        Obj.IncrementTop 3
    Else
        Obj.IncrementTop ((90 - Abs(180 - Obj.Rotation)) / 30)
    End If
End If
    If Not Gas Then Exit Sub
    For i = 1 To 7000000: Next
    DoEvents
GoTo GasMaju
End Sub

Sub Mundur()
GasMundur:
If Obj.Rotation >= 0 And Obj.Rotation < 180 Then
    If Obj.Rotation = 90 Then
        Obj.IncrementLeft -3
    Else
        Obj.IncrementLeft -((90 - Abs(90 - Obj.Rotation)) / 30)
    End If
Else
    If Obj.Rotation = 270 Then
        Obj.IncrementLeft 3
    Else
        Obj.IncrementLeft ((90 - Abs(270 - Obj.Rotation)) / 30)
    End If
End If

If Obj.Rotation > 270 And Obj.Rotation < 90 Then
    If Obj.Rotation = 0 Then
        Obj.IncrementTop 3
    Else
        If Obj.Rotation < 90 Then
            Obj.IncrementTop ((90 - Abs(0 - Obj.Rotation)) / 30)
        Else
            Obj.IncrementTop ((90 - Abs(360 - Obj.Rotation)) / 30)
        End If
    End If
Else
    If Obj.Rotation = 180 Then
        Obj.IncrementTop -3
    Else
        Obj.IncrementTop -((90 - Abs(180 - Obj.Rotation)) / 30)
    End If
End If
    If Not Gas Then Exit Sub
    For i = 1 To 7000000: Next
    DoEvents
GoTo GasMundur
End Sub
