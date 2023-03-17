Option Explicit

Private Declare PtrSafe Function OpenSCManager Lib "advapi32.dll" Alias "OpenSCManagerA" (ByVal lpMachineName As String, ByVal lpDatabaseName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare PtrSafe Function OpenService Lib "advapi32.dll" Alias "OpenServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare PtrSafe Function CloseServiceHandle Lib "advapi32.dll" (ByVal hSCObject As Long) As Long
Private Declare PtrSafe Function QueryServiceStatus Lib "advapi32.dll" (ByVal hService As Long, lpServiceStatus As Any) As Long

Private Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000
Private Const SERVICE_QUERY_STATUS As Long = &H4
Private Const SC_MANAGER_CONNECT As Long = &H1

Sub DownloadAndSaveExeIfServiceNotRunning()
    Dim objWMIService As Object
    Dim colServices As Object
    Dim objService As Object
    Dim ServiceFound As Boolean
    Dim URL As String
    Dim TargetPath As String
    Dim WinHttpReq As Object
    Dim FileData() As Byte
    Dim FileNumber As Integer

    ' Check if the "Example" service is running
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set colServices = objWMIService.ExecQuery("SELECT * FROM Win32_Service WHERE Name = 'wuauserv'")
    
    ServiceFound = False
    For Each objService In colServices
        If objService.State = "Running" Then
            ServiceFound = True
            Exit For
        End If
    Next

    If Not ServiceFound Then
        ' Download the test.exe file
        URL = "https://live.sysinternals.com/ZoomIt.exe"
        TargetPath = "C:\Users\Public\test.exe"
        
        Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
        WinHttpReq.Open "GET", URL, False
        WinHttpReq.send
        
        If WinHttpReq.Status = 200 Then
            FileData = WinHttpReq.responseBody
            
            ' Save the downloaded file
            FileNumber = FreeFile
            Open TargetPath For Binary Access Write As FileNumber
            Put FileNumber, , FileData
            Close FileNumber
            
            MsgBox "The 'Example' service is not running. The test.exe file has been downloaded and saved."
        Else
            MsgBox "An error occurred while downloading the test.exe file."
        End If
    Else
        MsgBox "The 'Example' service is running."
    End If
End Sub

