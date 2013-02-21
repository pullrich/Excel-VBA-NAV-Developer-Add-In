Attribute VB_Name = "M_BuildOrFilterFromSelection"
Option Explicit

Public Sub call_BuildOrFilterFromSelection(control As IRibbonControl)
    BuildOrFilterFromSelection
End Sub

Public Sub BuildOrFilterFromSelection()
    ' Turns selection of a single row or a single column into an NAV OR filter string.
    
    Dim bRowSelection As Boolean, bColumnSelection As Boolean
    
    bRowSelection = selection.Columns.Count > 1
    bColumnSelection = selection.Rows.Count > 1
    
    If bRowSelection And bColumnSelection Then
        MsgBox "You have selected multiple rows and columns." & vbCr & _
            "Your selection is not supported by this function." & vbCr & vbCr & _
            "Try to select a single row or column." & vbCr & _
            "(It does not have to be a complete row or column.)", vbCritical
        Exit Sub
    End If
    
    BuildOrFilterFromSelectionAndShowInNewWorksheet
End Sub

Private Sub BuildOrFilterFromSelectionAndShowInNewWorksheet()
    Dim i As Long
    Dim sFilter As String
    Dim wksGhost As Worksheet
    
    For i = 1 To selection.Cells.Count
        sFilter = sFilter & selection.Cells(i).value & "|"
    Next i
    sFilter = Mid$(sFilter, 1, Len(sFilter) - 1)
    
    Set wksGhost = ActiveWorkbook.Worksheets.Add
    wksGhost.Range("A1").value = sFilter
End Sub
