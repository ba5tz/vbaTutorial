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
'           Date    : 30 Oktober 2021
'           About   : Hightlight Cell

'===--------------***-----------------===
'   Simpan di Sheet Module
'===--------------***-----------------===

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
Dim Areanya As Range
Dim Brs As Long, Kol As Long

Set Areanya = Range("B2:Q21")


If Not Intersect(Target, Areanya) Is Nothing Then
    
    Brs = Target.Row
    Kol = Target.Column
    
    Range("A2:A21").Interior.Color = 6299648
    Range("B1:Q1").Interior.Color = 6299648
    
    Areanya.Interior.ColorIndex = 0
    Range(Cells(2, Kol), Cells(21, Kol)).Interior.Color = 16750848
    Range(Cells(Brs, 2), Cells(Brs, "Q")).Interior.Color = 16763904
    
    Cells(1, Kol).Interior.Color = 49407
    Cells(Brs, 1).Interior.Color = 49407
    Target.Interior.Color = vbYellow
End If
End Sub
