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
'           About   : Upload File ke web Melalui FTP

Private Const FTP_TRANSFER_TYPE_UNKNOWN     As Long = 0
Private Const INTERNET_FLAG_RELOAD          As Long = &H80000000
 
Private Declare Function InternetOpenA Lib "wininet.dll" ( _
    ByVal sAgent As String, _
    ByVal lAccessType As Long, _
    ByVal sProxyName As String, _
    ByVal sProxyBypass As String, _
    ByVal lFlags As Long) As Long
 
Private Declare Function InternetConnectA Lib "wininet.dll" ( _
    ByVal hInternetSession As Long, _
    ByVal sServerName As String, _
    ByVal nServerPort As Long, _
    ByVal sUsername As String, _
    ByVal sPassword As String, _
    ByVal lService As Long, _
    ByVal lFlags As Long, _
    ByVal lcontext As Long) As Long
 
Private Declare Function FtpPutFileA _
   Lib "wininet.dll" _
       (ByVal hFtpSession As Long, _
        ByVal lpszLocalFile As String, _
        ByVal lpszRemoteFile As String, _
        ByVal dwFlags As Long, _
        ByVal dwContext As Long) As Boolean
 
Private Declare Function InternetCloseHandle Lib "wininet" ( _
    ByVal hInet As Long) As Long
 
Sub FtpUpload(ByVal strLocalFile As String, ByVal strRemoteFile As String, ByVal strHost As String, ByVal lngPort As Long, ByVal strUser As String, ByVal strPass As String)
    Dim Status As Boolean
    Dim hOpen   As Long
    Dim hConn   As Long
 
    hOpen = InternetOpenA("FTPGET", 1, vbNullString, vbNullString, 1)
    hConn = InternetConnectA(hOpen, strHost, lngPort, strUser, strPass, 1, 0, 2)
    Status = FtpPutFileA(hConn, strLocalFile, strRemoteFile, FTP_TRANSFER_TYPE_UNKNOWN Or INTERNET_FLAG_RELOAD, 0) 
    If Status Then
        Debug.Print "Upload Success"
    Else
        Debug.Print "Upload Fail"
    End If
 
    'Close connections
    InternetCloseHandle hConn
    InternetCloseHandle hOpen
 
End Sub
 
'Pengunaaan
'--------------------
Sub TestUpload
FtpUpload "C:\LokadiFile\Nama Fiile.txt", "//Download/Nama file.txt", _
            "192.168.0.100", 21, "username", "password"
End Sub
