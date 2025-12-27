Sub Azuriraj_Click()
    '
    ' Azuriraj Click Macro
    ' Calls Python script to update tournament participants
    '
    
    Dim scriptPath As String
    Dim workbookPath As String
    Dim tournamentFolder As String
    Dim command As String
    
    ' Get the current workbook directory (this is the tournament folder)
    workbookPath = ThisWorkbook.Path
    tournamentFolder = workbookPath
    
    ' Build path to Python script (assuming code folder is in parent directory)
    scriptPath = workbookPath & "\..\code\azuriraj_ucesnike.py"
    
    ' Save the current workbook to ensure data is up to date
    ThisWorkbook.Save
    
    ' Try direct Python execution (bypasses cmd.exe)
    command = "python3 """ & scriptPath & """ """ & tournamentFolder & """"
    
    ' Run the Python script with multiple fallback methods
    On Error Resume Next
    
    ' Method 1: Direct Shell call
    Shell command, vbHide
    
    If Err.Number <> 0 Then
        Err.Clear
        ' Method 2: WScript.Shell with working directory
        Dim wsh As Object
        Set wsh = CreateObject("WScript.Shell")
        wsh.CurrentDirectory = workbookPath & "\.."
        wsh.Run command, 1, False
    End If
    
    On Error GoTo 0
    
End Sub
