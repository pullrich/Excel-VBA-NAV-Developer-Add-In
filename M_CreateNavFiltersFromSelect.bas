Attribute VB_Name = "M_CreateNavFiltersFromSelect"
Option Explicit

Public Const ACTIVE_FILTER_COLUMN_COLOR As Long = 969993 'RGB(9, 205, 14)

Sub call_CreateNavisionObjectFiltersFromSelection(control As IRibbonControl)
    Objects_CreateNavisionObjectFiltersFromSelection
End Sub

Sub Objects_CreateNavisionObjectFiltersFromSelection()
' <description>
' This macro creates a new sheet and inserts into it NAV filter strings by
' object type.
' </description>

    Dim filterSheet As Worksheet
    Dim oAllObjects As New Collection
    
    CheckSelectedAreas
    AddObjectIdsFromSelectedAreasToTheirCollection oAllObjects
        
    Sheets.Add After:=Sheets(Sheets.Count)
    Set filterSheet = Worksheets(Sheets.Count)
    
    WriteObjectFiltersToSheetVertically oAllObjects, filterSheet
End Sub

Private Sub CheckSelectedAreas()
    Dim selectionVariant As Variant
    Dim selection As Range
    
    For Each selectionVariant In Application.selection.Areas
        Set selection = selectionVariant
        If selection.Columns.Count <> 2 Then
            Err.Raise vbObjectError + 1, _
                Description:="Your selection must include cells in exactly two columns." & vbCrLf & _
                "Selected columns: " & selection.Columns.Count
        End If
    Next selectionVariant
End Sub

Private Sub AddObjectIdsFromSelectedAreasToTheirCollection(ByRef oAllObjects)
    Dim Tables As New Collection
    Dim Forms As New Collection
    Dim Reports As New Collection
    Dim Dataports As New Collection
    Dim Codeunits As New Collection
    Dim XMLports As New Collection
    Dim MenuSuites As New Collection
    Dim Pages As New Collection
    
    oAllObjects.Add Tables, "Tables"
    oAllObjects.Add Forms, "Forms"
    oAllObjects.Add Reports, "Reports"
    oAllObjects.Add Dataports, "Dataports"
    oAllObjects.Add Codeunits, "Codeunits"
    oAllObjects.Add XMLports, "XMLports"
    oAllObjects.Add MenuSuites, "MenuSuites"
    oAllObjects.Add Pages, "Pages"
    
    Dim varSelectionArea As Variant
    Dim oCurrentArea As Range
    
    For Each varSelectionArea In Application.selection.Areas
        Set oCurrentArea = varSelectionArea
        AddObjectIdsFromCurrentAreaToTheirCollection oAllObjects, oCurrentArea
    Next varSelectionArea
End Sub

Private Sub AddObjectIdsFromCurrentAreaToTheirCollection(ByRef oAllObjects, ByRef oCurrentSelection)
    Dim rowIdx As Long
    Dim lastRowInSelection As Long
    Dim realLastRow As Long
    Dim lastRowToUse As Long
    
    lastRowInSelection = oCurrentSelection.Rows.Count
    realLastRow = GetRealLastRow
    
    If lastRowInSelection > realLastRow Then
        lastRowToUse = realLastRow
    Else
        lastRowToUse = lastRowInSelection
    End If
    
    For rowIdx = 1 To lastRowToUse
        If Not oCurrentSelection.Cells(rowIdx, 1).EntireRow.Hidden Then
            Select Case oCurrentSelection.Cells(rowIdx, 1)
                Case 1, "Table", "Tabelle"
                    oAllObjects("Tables").Add oCurrentSelection.Cells(rowIdx, 2)
                Case 2, "Form", "Formular"
                    oAllObjects("Forms").Add oCurrentSelection.Cells(rowIdx, 2)
                Case 3, "Report", "Bericht"
                    oAllObjects("Reports").Add oCurrentSelection.Cells(rowIdx, 2)
                Case 4, "Dataport"
                    oAllObjects("Dataports").Add oCurrentSelection.Cells(rowIdx, 2)
                Case 5, "Codeunit"
                    oAllObjects("Codeunits").Add oCurrentSelection.Cells(rowIdx, 2)
                Case 6, "XMLport"
                    oAllObjects("XMLports").Add oCurrentSelection.Cells(rowIdx, 2)
                Case 7, "MenuSuite"
                    oAllObjects("MenuSuites").Add oCurrentSelection.Cells(rowIdx, 2)
                Case 8, "Page", "Seite"
                    oAllObjects("Pages").Add oCurrentSelection.Cells(rowIdx, 2)
            End Select
        End If
    Next rowIdx
    
End Sub

Private Function GetRealLastRow() As Long
    Dim realLastRow As Long
    
    Range("A1").Select
    On Error Resume Next
    realLastRow = Cells.Find("*", Range("A1"), xlFormulas, , xlByRows, xlPrevious).row
    GetRealLastRow = realLastRow
End Function

Private Sub WriteObjectFiltersToSheetVertically(ByRef AllObjectsCollection As Collection, ByRef filterSheet As Worksheet)
    Dim rowIdx As Long
    Dim columnIdx As Long
    
    Dim objectCollectionVariant As Variant
    Dim objectCollection As Collection
    Dim typeStringCollection As Collection
    
    Dim currentTypeString As String
    Dim currentObjectTypeIndex As Integer
    
    currentObjectTypeIndex = 0
    
    Set typeStringCollection = GetFilterSheetColumnDescriptions
    
    rowIdx = 1
    columnIdx = 1
    For Each objectCollectionVariant In AllObjectsCollection
        Set objectCollection = objectCollectionVariant
        currentObjectTypeIndex = currentObjectTypeIndex + 1
        currentTypeString = typeStringCollection.Item(currentObjectTypeIndex)
        
        WriteObjectFilterForTypeToSheetVertically currentTypeString, objectCollection, filterSheet, rowIdx, columnIdx
        rowIdx = rowIdx + 1
        
    Next objectCollectionVariant
    
    filterSheet.Columns("A:A").ColumnWidth = 14
End Sub

Private Sub WriteObjectFilterForTypeToSheetVertically( _
ByVal typeString As String, ByRef concreteTypeCollection As Collection, ByRef filterSheet As Worksheet, ByRef currentRowIdx As Long, ByVal startColumnIdx As Long)
    Dim columnIdx As Long
    Dim entry As Variant
    
    columnIdx = startColumnIdx
    
    ' Write header row
    filterSheet.Cells(currentRowIdx, columnIdx) = typeString
    filterSheet.Cells(currentRowIdx, columnIdx + 1) = "(Total Objects in Filters: " & CStr(concreteTypeCollection.Count) & ")"
    filterSheet.Cells(currentRowIdx, columnIdx).EntireRow.Font.Bold = True
    filterSheet.Cells(currentRowIdx, columnIdx).EntireRow.Font.Size = 14
    If concreteTypeCollection.Count <> 0 Then
        filterSheet.Cells(currentRowIdx, columnIdx).EntireRow.Font.Color = ACTIVE_FILTER_COLUMN_COLOR
    Else
        Exit Sub
    End If

    ' Write filter strings
    currentRowIdx = currentRowIdx + 1
    For Each entry In concreteTypeCollection
        If Len(filterSheet.Cells(currentRowIdx, columnIdx)) = 0 Then
            filterSheet.Cells(currentRowIdx, columnIdx) = CStr(entry)
        Else
            If Len(CStr(filterSheet.Cells(currentRowIdx, columnIdx)) + "|" + CStr(entry)) <= NAV_FILTER_LEN Then
                filterSheet.Cells(currentRowIdx, columnIdx) = CStr(filterSheet.Cells(currentRowIdx, columnIdx)) + "|" + CStr(entry)
            Else
                currentRowIdx = currentRowIdx + 1
                filterSheet.Cells(currentRowIdx, columnIdx) = filterSheet.Cells(currentRowIdx, columnIdx) + CStr(entry)
            End If
        End If
    Next entry
    
End Sub

Private Function GetFilterSheetColumnDescriptions() As Collection
    Dim descriptionColl As Collection
    Set descriptionColl = New Collection
    
    descriptionColl.Add "Table"
    descriptionColl.Add "Form"
    descriptionColl.Add "Report"
    descriptionColl.Add "Dataport"
    descriptionColl.Add "Codeunit"
    descriptionColl.Add "XMLport"
    descriptionColl.Add "MenuSuite"
    descriptionColl.Add "Page"
    
    Set GetFilterSheetColumnDescriptions = descriptionColl
End Function
