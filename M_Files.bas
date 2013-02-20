Attribute VB_Name = "M_Files"
Option Explicit

Public Sub call_WriteFullPathOfTextFilesInFolderToActiveWks(control As IRibbonControl)
    WriteFullPathOfTextFilesInFolderToActiveWks
End Sub

Public Sub WriteFullPathOfTextFilesInFolderToActiveWks()
    Dim FSO As IWshRuntimeLibrary.FileSystemObject
    Dim objFolder As IWshRuntimeLibrary.Folder
    Dim sFolderPath As String
    Dim rng As Range
    Dim objFileDialog As FileDialog
    
    Set objFileDialog = Application.FileDialog(msoFileDialogFolderPicker)
    
    objFileDialog.AllowMultiSelect = False
    objFileDialog.Show
    sFolderPath = objFileDialog.SelectedItems.Item(1)
    
    Set FSO = New IWshRuntimeLibrary.FileSystemObject
    Set objFolder = FSO.GetFolder(sFolderPath)
    
    Dim oFile As IWshRuntimeLibrary.File
    Set rng = selection
    For Each oFile In objFolder.Files
        If oFile.Type = "TXT-Datei" Then
            rng.value = oFile.Path
            Set rng = rng.Offset(1)
        End If
    Next oFile
End Sub

Public Function CountLinesInTextFile(sFilePath As String) As Long
    Dim FSO As IWshRuntimeLibrary.FileSystemObject
    Dim txsTextStream As IWshRuntimeLibrary.TextStream
    Dim lNumberOfLine As Long
    
    Set FSO = New IWshRuntimeLibrary.FileSystemObject
    Set txsTextStream = FSO.OpenTextFile(sFilePath, ForReading)
    
    lNumberOfLine = 0
    Do Until txsTextStream.AtEndOfStream
        txsTextStream.ReadLine
        lNumberOfLine = lNumberOfLine + 1
    Loop
    txsTextStream.Close
    
    CountLinesInTextFile = lNumberOfLine
End Function
