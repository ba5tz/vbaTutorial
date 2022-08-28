Attribute VB_Name = "ModTaskIcon"
Option Explicit

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type


Private Type PROPERTYKEY
    fmtid As GUID
    pid As Long
End Type


#If VBA7 Then
  
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Function DispCallFunc Lib "oleAut32.dll" (ByVal pvInstance As LongPtr, ByVal offsetinVft As LongPtr, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As LongPtr, ByRef retVAR As Variant) As Long
    Private Declare PtrSafe Sub SetLastError Lib "kernel32.dll" (ByVal dwErrCode As Long)
    Private Declare PtrSafe Function CoCreateInstance Lib "ole32" (ByRef rclsid As GUID, ByVal pUnkOuter As LongPtr, ByVal dwClsContext As Long, ByRef riid As GUID, ByRef ppv As LongPtr) As Long
    Private Declare PtrSafe Function ExtractIconA Lib "Shell32.dll" (ByVal hInst As LongPtr, ByVal lpszExeFileName As String, ByVal nIconIndex As LongPtr) As Long
    Private Declare PtrSafe Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pIUnk As IUnknown, ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function CLSIDFromString Lib "ole32" (ByVal OleStringCLSID As LongPtr, ByRef cGUID As Any) As Long
    Private Declare PtrSafe Function SHGetPropertyStoreForWindow Lib "Shell32.dll" (ByVal hwnd As LongPtr, ByRef riid As GUID, ByRef ppv As LongPtr) As Long
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
    Private Declare PtrSafe Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As LongPtr, ByVal lpString As String, ByVal hData As LongPtr) As Long
    Private Declare PtrSafe Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As LongPtr, ByVal lpString As String) As LongPtr
    Private Declare PtrSafe Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As LongPtr, ByVal lpString As String) As LongPtr
    Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function BringWindowToTop Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As LongPtr) As Long
    Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
  
#Else

    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    Private Declare Function DispCallFunc Lib "oleAut32.dll" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As Long, ByRef retVAR As Variant) As Long
    Private Declare Sub SetLastError Lib "kernel32.dll" (ByVal dwErrCode As Long)
    Private Declare Function CoCreateInstance Lib "ole32" (ByRef rclsid As GUID, ByVal pUnkOuter As Long, ByVal dwClsContext As Long, ByRef riid As GUID, ByRef ppv As Long) As Long
    Private Declare Function ExtractIconA Lib "Shell32.dll" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
    Private Declare Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pIUnk As IUnknown, ByVal hwnd As Long) As Long
    Private Declare Function CLSIDFromString Lib "ole32" (ByVal OleStringCLSID As Long, ByRef cGUID As Any) As Long
    Private Declare Function SHGetPropertyStoreForWindow Lib "Shell32.dll" (ByVal hwnd As Long, ByRef riid As GUID, ByRef ppv As Long) As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
    Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
    Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
    Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
    Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

#End If

Public Sub ChangeExcelIcon _
        (Optional ByVal Reset As Boolean = False, _
        Optional ByVal IconGambar As StdPicture, _
        Optional ByVal IconFile As String, _
        Optional ByVal FileIconIndex As Long = 0, _
        Optional taskcaption As String _
    )


    Const VT_LPWSTR = 31
  
    #If Win64 Then
        Const vTblOffsetFac_32_64 = 2
        Dim hVbe As LongLong, pPstore As LongLong, pTBarList As LongLong
        Dim PV(0 To 2) As LongLong
      
        PV(0) = VT_LPWSTR: PV(1) = StrPtr("Dummy")
    #Else
        Const vTblOffsetFac_32_64 = 1
        Dim hVbe As Long, pPstore As Long, pTBarList As Long
        Dim PV(0 To 3) As Long
      
        PV(0) = VT_LPWSTR: PV(2) = StrPtr("Dummy")
    #End If


    Const IPropertyKey_SetValue = 24 * vTblOffsetFac_32_64
    Const IPropertyKey_Commit = 28 * vTblOffsetFac_32_64
    Const ITASKLIST3_HrInit = 12 * vTblOffsetFac_32_64
    Const ITASKLIST3_AddTab = 16 * vTblOffsetFac_32_64
    Const ITASKLIST3_DeleteTab = 20 * vTblOffsetFac_32_64
    Const ITASKLIST3_ActivateTab = 24 * vTblOffsetFac_32_64
    Const ITASKLIST3_Settaskcaption = 76 * vTblOffsetFac_32_64


    Const IID_PropertyStore = "{886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99}"
    Const IID_PropertyKey = "{9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3}"
    Const CLSID_TASKLIST = "{56FDF344-FD6D-11D0-958A-006097C9A090}"
    Const IID_TASKLIST3 = "{EA1AFB91-9E28-4B86-90E9-9E9F8A5EEFAF}"

    Const CLSCTX_INPROC_SERVER = &H1
    Const S_OK = 0
    Const CC_STDCALL = 4

    Const GWL_STYLE = (-16)
    Const WS_MINIMIZEBOX = &H20000
    Const GWL_HWNDPARENT = (-8)

    Dim tClsID As GUID, tIID As GUID, tPK As PROPERTYKEY
  
    If Reset Then
        Call addicon(True)
    ElseIf Not IconGambar Is Nothing Then
        Call addicon(, IconGambar, , FileIconIndex)
    ElseIf Len(IconFile) Then
        Call addicon(, , IconFile, FileIconIndex)
    End If
    
    Call CLSIDFromString(StrPtr(IID_PropertyStore), tIID)
    If SHGetPropertyStoreForWindow(Application.hwnd, tIID, pPstore) = S_OK Then
        Call CLSIDFromString(StrPtr(IID_PropertyKey), tPK)
        tPK.pid = 5 ':  PV(0) = VT_LPWSTR: PV(1) = StrPtr("Dummy")
        Call vtblCall(pPstore, IPropertyKey_SetValue, vbLong, CC_STDCALL, VarPtr(tPK), VarPtr(PV(0))) 'SetValue Method
        Call vtblCall(pPstore, IPropertyKey_Commit, vbLong, CC_STDCALL) ' Commit Method
        Call CLSIDFromString(StrPtr(CLSID_TASKLIST), tClsID)
        Call CLSIDFromString(StrPtr(IID_TASKLIST3), tIID)
        
        If CoCreateInstance(tClsID, 0, CLSCTX_INPROC_SERVER, tIID, pTBarList) = S_OK Then
            SetProp Application.hwnd, "pTBarList", pTBarList
            Call vtblCall(pTBarList, ITASKLIST3_HrInit, vbLong, CC_STDCALL) 'HrInit Method
            Call vtblCall(pTBarList, ITASKLIST3_DeleteTab, vbLong, CC_STDCALL, Application.hwnd) 'DeleteTab Method
            Call vtblCall(pTBarList, ITASKLIST3_AddTab, vbLong, CC_STDCALL, Application.hwnd) 'AddTab Method
            Call vtblCall(pTBarList, ITASKLIST3_ActivateTab, vbLong, CC_STDCALL, Application.hwnd) 'ActivateTab Method
            
            If Len(taskcaption) Then
                Call vtblCall(pTBarList, ITASKLIST3_Settaskcaption, vbLong, CC_STDCALL, Application.hwnd, StrPtr(taskcaption)) 'Settaskcaption Method
            End If
            
            If Reset Then
                     Call vtblCall(pTBarList, ITASKLIST3_Settaskcaption, vbLong, CC_STDCALL, Application.hwnd, StrPtr(vbNullString)) 'Settaskcaption Method
            End If
            
                hVbe = FindWindow("wndclass_desked_gsk", vbNullString)
                If IsWindowVisible(hVbe) Then
                    Call SetProp(Application.hwnd, "hVbe", hVbe)
                    Call ShowWindow(hVbe, 0)
                    Call vtblCall(pTBarList, ITASKLIST3_DeleteTab, vbLong, CC_STDCALL, hVbe) 'DeleteTab Method
                End If
        End If
    End If
    
    Call SetForegroundWindow(Application.hwnd): Call BringWindowToTop(Application.hwnd)
  
End Sub


Private Sub addicon(Optional ByVal Reset As Boolean, Optional ByVal IconGambar As StdPicture, Optional ByVal IconFile As String, Optional ByVal Index As Long = 0)
  
    #If Win64 Then
        Dim hIcon As LongPtr
    #Else
        Dim hIcon As Long
    #End If
  
    Const WM_SETICON = &H80
    Const ICON_SMALL = 0
    Const ICON_BIG = 1

    Dim N As Long, S As String
        
    If Not IconGambar Is Nothing Then
        hIcon = IconGambar.Handle
        Call SendMessage(Application.hwnd, WM_SETICON, ICON_SMALL, ByVal hIcon)
        Call SendMessage(Application.hwnd, WM_SETICON, ICON_BIG, ByVal hIcon)
    ElseIf Len(IconFile) Then
        If Dir(IconFile, vbNormal) = vbNullString Then
            Exit Sub
        End If
        N = InStrRev(IconFile, ".")
        S = LCase(Mid(IconFile, N + 1))
        Select Case S
            Case "exe", "ico", "dll"
            Case Else
                Err.Raise 5
        End Select
        If Application.hwnd = 0 Then
            Exit Sub
        End If
        
        hIcon = ExtractIconA(0, IconFile, Index)
        If hIcon <> 0 Then
            Call SendMessage(Application.hwnd, WM_SETICON, ICON_SMALL, ByVal hIcon)
            Call SendMessage(Application.hwnd, WM_SETICON, ICON_BIG, ByVal hIcon)
        End If
        
    ElseIf Reset Then
         hIcon = ExtractIconA(0, Application.Path & "\Excel.exe", 0)
        If hIcon <> 0 Then
            Call SendMessage(Application.hwnd, WM_SETICON, ICON_SMALL, ByVal hIcon)
            Call SendMessage(Application.hwnd, WM_SETICON, ICON_BIG, ByVal hIcon)
        End If
    End If
    
    Call DrawMenuBar(Application.hwnd)
    DeleteObject hIcon
  
End Sub


#If Win64 Then
    Private Function vtblCall(ByVal InterfacePointer As LongLong, ByVal VTableOffset As Long, ByVal FunctionReturnType As Long, ByVal CallConvention As Long, ParamArray FunctionParameters() As Variant) As Variant

    Dim vParamPtr() As LongLong
#Else
    Private Function vtblCall(ByVal InterfacePointer As Long, ByVal VTableOffset As Long, ByVal FunctionReturnType As Long, ByVal CallConvention As Long, ParamArray FunctionParameters() As Variant) As Variant

    Dim vParamPtr() As Long
#End If

    If InterfacePointer = 0& Or VTableOffset < 0& Then Exit Function
    If Not (FunctionReturnType And &HFFFF0000) = 0& Then Exit Function

    Dim pIndex As Long, pCount As Long
    Dim vParamType() As Integer
    Dim vRtn As Variant, vParams() As Variant

    vParams() = FunctionParameters()
    pCount = Abs(UBound(vParams) - LBound(vParams) + 1&)
    If pCount = 0& Then
        ReDim vParamPtr(0 To 0)
        ReDim vParamType(0 To 0)
    Else
        ReDim vParamPtr(0 To pCount - 1&)
        ReDim vParamType(0 To pCount - 1&)
        For pIndex = 0& To pCount - 1&
            vParamPtr(pIndex) = VarPtr(vParams(pIndex))
            vParamType(pIndex) = VarType(vParams(pIndex))
        Next
    End If

    pIndex = DispCallFunc(InterfacePointer, VTableOffset, CallConvention, FunctionReturnType, pCount, _
    vParamType(0), vParamPtr(0), vRtn)
    If pIndex = 0& Then
        vtblCall = vRtn
    Else
        SetLastError pIndex
    End If

End Function




