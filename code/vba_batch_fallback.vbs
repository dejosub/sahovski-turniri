Sub Azuriraj_Click_Batch_Fallback()
    '
    ' Alternative: Create batch file and ask user to run it
    '
    
    Dim workbookPath As String
    Dim tournamentFolder As String
    Dim batchFile As String
    Dim fileNum As Integer
    
    ' Get paths
    workbookPath = ThisWorkbook.Path
    tournamentFolder = workbookPath
    batchFile = workbookPath & "\run_tournament_update.bat"
    
    ' Save workbook
    ThisWorkbook.Save
    
    ' Create batch file
    fileNum = FreeFile
    Open batchFile For Output As fileNum
    Print #fileNum, "@echo off"
    Print #fileNum, "echo Updating tournament participants..."
    Print #fileNum, "cd /d """ & workbookPath & "\.."""
    Print #fileNum, "python3 ""code\azuriraj_ucesnike.py"" """ & tournamentFolder & """"
    Print #fileNum, "echo."
    Print #fileNum, "echo Update complete! Press any key to close..."
    Print #fileNum, "pause > nul"
    Close fileNum
    
    ' Inform user
    MsgBox "Security policy prevents direct execution." & vbCrLf & vbCrLf & _
           "A batch file has been created at:" & vbCrLf & _
           batchFile & vbCrLf & vbCrLf & _
           "Please run this file to update participants.", _
           vbInformation, "Manual Execution Required"
    
End Sub
