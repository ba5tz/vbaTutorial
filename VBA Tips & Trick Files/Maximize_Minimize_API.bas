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
'           Date    : 14 Juli 2021
'           About   :
'
' Cara Penggunaan
' 1. Simpan API dan Function SetMaxMin ke dalam Module
' 2. Pada Userform yang ingin ditambahkan Max dan Min tambhakan pada Initialize script 
'    SetMaxMin Me.caption


#If VBA7 And Win64 Then
    '64 bit
    Private Declare PtrSafe Function SetWindowLong Lib "user32.dll" Alias _
        "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function FindWindow Lib "user32.dll" Alias _
        "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Function GetWindowLong Lib "user32.dll" Alias _
        "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
#Else
    '32bit
    Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" _
        (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" _
        (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetWindowLong Lib "user32.dll" Alias _
        "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
#End If

Private Const GWL_STYLE = -16
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000

Function SetMaxMin(xCaption As String)
Dim hwnd As Long
Dim stylelama As Long

hwnd = FindWindow("ThunderDFrame", xCaption)
stylelama = GetWindowLong(hwnd, GWL_STYLE)
SetWindowLong hwnd, GWL_STYLE, stylelama Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX
End Function
