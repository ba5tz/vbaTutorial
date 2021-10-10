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
'           About     : Transparent Userform
'
'       Cara Penggunaan
'       1. Simpan script Module
'       2. untuk penggunaanya panggil BuatTransparent me.caption
'

#If Win64 And VBA7 Then
Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32.dll" _
  (ByVal hwnd As Long, _
  ByVal crKey As Long, _
  ByVal bAlpha As Byte, _
  ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function FindWindow Lib "user32.dll" Alias _
    "FindWindowA" (ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long
    
Private Declare PtrSafe Function SetWindowLong Lib "user32.dll" Alias _
    "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long
    
Private Declare PtrSafe Function GetWindowLong Lib "user32.dll" Alias _
    "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
#Else
    '32 Bit
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" _
  (ByVal hwnd As Long, _
  ByVal crKey As Long, _
  ByVal bAlpha As Byte, _
  ByVal dwFlags As Long) As Long

Private Declare Function FindWindow Lib "user32.dll" Alias _
    "FindWindowA" (ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long
    
Private Declare Function SetWindowLong Lib "user32.dll" Alias _
    "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long
    
Private Declare Function GetWindowLong Lib "user32.dll" Alias _
    "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
#End If


Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
Private Const GWL_EXSTYLE = -20
Private Const WS_EX_LAYERED = &H80000

Public Sub BuatTransparent(xCaption As String)
Dim hwnd As Long
Dim SStyle As Long

hwnd = FindWindow("ThunderDFrame", xCaption)
SStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
SetWindowLong hwnd, GWL_EXSTYLE, SStyle Or WS_EX_LAYERED
SetLayeredWindowAttributes hwnd, 0, 200, LWA_ALPHA
End Sub
