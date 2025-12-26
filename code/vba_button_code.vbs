Sub AzurirajUcesnike()
    '
    ' Azuriraj Ucesnike Macro
    ' Calls Python script to update tournament participants
    '
    
    Dim scriptPath As String
    Dim workbookPath As String
    Dim tournamentFolder As String
    Dim command As String
    Dim result As Integer
    
    ' Get the current workbook directory (this is the tournament folder)
    workbookPath = ThisWorkbook.Path
    tournamentFolder = workbookPath
    
    ' Build path to Python script (assuming code folder is in parent directory)
    scriptPath = workbookPath & "\..\code\azuriraj_ucesnike.py"
    
    ' Check if Python script exists
    If Dir(scriptPath) = "" Then
        MsgBox "Python script not found at: " & scriptPath, vbCritical, "Error"
        Exit Sub
    End If
    
    ' Show confirmation dialog
    result = MsgBox("Da li želite da ažurirate učesnike turnira?" & vbCrLf & vbCrLf & _
                   "Ovo će preneti sve plaćene učesnike u turnirsku tabelu.", _
                   vbYesNo + vbQuestion, "Ažuriranje učesnika")
    
    If result = vbNo Then
        Exit Sub
    End If
    
    ' Save the current workbook to ensure data is up to date
    ThisWorkbook.Save
    
    ' Build command to run Python script with tournament folder as parameter
    command = "python """ & scriptPath & """ """ & Chr(34) & tournamentFolder & Chr(34) & """"
    
    ' Show status message
    Application.StatusBar = "Ažuriranje učesnika u toku..."
    Application.ScreenUpdating = False
    
    ' Run the Python script
    result = Shell("cmd /c cd /d """ & workbookPath & "\.."" && " & command, vbNormalFocus)
    
    ' Wait a moment for the script to complete
    Application.Wait (Now + TimeValue("0:00:03"))
    
    ' Reset status
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    ' Show completion message
    MsgBox "Ažuriranje učesnika je završeno!" & vbCrLf & vbCrLf & _
           "Proverite konzolu za detalje o izvršavanju.", _
           vbInformation, "Završeno"
    
End Sub
