' Retrieves weather information using PowerShell and displays the results in
' Notepad. This code launches a child PowerShell process and accesses the
' internet.
Option Explicit

' Declare Sleep function from Windows API
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Sub GetWeatherAndTime()
    Dim weatherAndTime As String
    Dim fileName As String
    Dim filePath As String
    Dim fileNumber As Integer
    Dim shell As Object
    Dim command As String
    
    ' Build and run the PowerShell command to get current weather and time
    command = "powershell -Command ""(Invoke-WebRequest 'https://wttr.in/London?format=%C+%t+%w+%T').Content"""
    Set shell = CreateObject("WScript.Shell")
    weatherAndTime = shell.Exec(command).StdOut.ReadAll
    
    ' Save the retrieved information to a file
    fileName = "WeatherAndTime.txt"
    filePath = "C:\Users\Public\" & fileName
    fileNumber = FreeFile
    
    On Error GoTo ErrorHandler
    ' Write the data to disk
    Open filePath For Output As fileNumber
    Print #fileNumber, "In the background, I was able to connect externally and fetch info about weather and current time and save it in the file on disc: " & filePath
    Print #fileNumber, weatherAndTime
    Close fileNumber
    On Error GoTo 0
    
    ' Pause briefly before showing the file
    Sleep 3000
    
    ' Display the file using Notepad
    shell.Run "notepad.exe " & filePath, 1, True
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred while trying to save the file. Please ensure you have the necessary permissions.", vbCritical, "Error"
    On Error GoTo 0
End Sub
