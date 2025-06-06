' Creates a batch file that downloads an image and sets it as the desktop
' wallpaper by updating the registry. This downloads content from the
' internet and modifies user settings.
Sub CreateAndRunBatchFileChangeWallpaper()
    Dim filePath As String
    Dim fileNum As Integer
    Dim imageURL As String
    Dim batchCommands As Variant
    Dim i As Integer
    
    ' Path for the temporary batch file
    filePath = "C:\Users\Public\ChangeWallpaper.bat"
    fileNum = FreeFile()
    ' Image to download and set as wallpaper
    imageURL = "https://images.unsplash.com/photo-1501594907352-04cda38ebc29" ' Replace with desired image URL
    
    ' Commands to download the image and change the wallpaper
    batchCommands = Array("@echo off", _
        "powershell -Command ""(New-Object System.Net.WebClient).DownloadFile('" & imageURL & "', 'C:\Users\Public\NewWallpaper.jpg')""", _
        "reg add ""HKEY_CURRENT_USER\Control Panel\Desktop"" /v Wallpaper /t REG_SZ /d ""C:\Users\Public\NewWallpaper.jpg"" /f", _
        "RUNDLL32.EXE USER32.DLL,UpdatePerUserSystemParameters")
    
    ' Delete any existing batch file
    On Error Resume Next
    Kill filePath
    On Error GoTo 0

    ' Write the commands to disk
    Open filePath For Output As #fileNum
        For i = LBound(batchCommands) To UBound(batchCommands)
            Print #fileNum, batchCommands(i)
        Next i
    Close #fileNum

    ' Run the batch file silently
    On Error Resume Next
    Shell filePath, vbHide
    On Error GoTo 0

    ' Inform the user that the batch file was executed
    MsgBox "Attempted to create and run a batch file to change the desktop wallpaper. If ASR rules are in place, this action should be blocked.", vbInformation
End Sub
