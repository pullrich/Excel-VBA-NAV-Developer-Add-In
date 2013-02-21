Attribute VB_Name = "M_ImportWorksheet"
'Type | Option: Table,Form,Report,Dataport,Codeunit,XMLport,MenuSuite,System,FieldNumber
'No. | LongInt

'New Object | Bool: No,Yes
'Warning | Bool: 0,1
'Action | Option: Create,Replace,Delete,Skip,Merge: Existing<-New,Merge: New<-Existing

'Existing Object Name | Text
'Existing Object Modified | Bool: No,Yes
'Existing Object Version List | Text
'Existing Object Size | LongInt
'Existing Object Date | Date: DD.MM.YYYY
'Existing Object Time | Time: HH:MM.SS
'Existing Object Changed | Bool: No,Yes | Ist "Yes" wenn Modified Haken gesetzt oder sich die Versionsliste unterscheidet! (siehe NAV Hilfe)

'Name
'New Object Modified | Bool: No,Yes
'New Object Version List | Text
'New Object Size | LongInt
'New Object Date | Date: DD.MM.YYYY
'New Object Time | Time: HH:MM.SS
'New Object Changed | Bool: No,Yes


' Welche Informationen sollen ermittelt werden?
' 1. Welches neue Objekt ist älter als das existierende? Datum und Zeit vergleichen.
' Das würde bedeuten, dass das neue Objekt zunächst nicht eingespielt werden dürfte.
' Es müsste zuerst gemergt werden.
' Neue Spalte "New Object Older".
' 2. Welches neue Objekt unterscheidet sich irgendwie vom existierenden?
' Datum, Zeit, Name, Version List, Size vergleichen. Unterscheidet sich irgendetwas
' dann ist das neue Objekt anders. Es sollte eine neue Spalte mit einem Bool Wert
' für "New Object Differs" und weitere Spalten pro Feldvergleich geben.
' "New Object Date": Option: Equal,Older,Younger
' "New Object Time": Option: Equal,Older,Younger
' "New Object Size": Option: Equal,Smaller,Bigger
' "New Object Name Different": Bool: No,Yes
' "New Object Version List Different": Bool: No,Yes

' Infomaterial:
' http://support.microsoft.com/kb/291308 - Umgang mit Ranges

Option Explicit

Sub call_BuildObjectComparison(control As IRibbonControl)
    ImportWorksheet_BuildObjectComparison
End Sub

Sub ImportWorksheet_BuildObjectComparison()

    Dim foundColumnHeaders As Collection
    Set foundColumnHeaders = New Collection
    
    Dim invalidColumnHeaders As Collection
    Set invalidColumnHeaders = New Collection
    
    Dim columnPosition As Collection
    Set columnPosition = New Collection
    
    Dim RowColumn(2) As Integer
    
    
    Dim cellValue As String
    Dim RowNo, ColumnNo As Integer
    
    Dim colAndAppear As ColumnAppearances
   
    RowNo = 1
    For ColumnNo = 1 To 19
        cellValue = Application.activeSheet.Cells(RowNo, ColumnNo)
        If IsValidHeaderText(cellValue) Then
            If IsHeaderInFoundColumnHeaders(cellValue, foundColumnHeaders) Then
                IncrementAppearance cellValue, foundColumnHeaders
            Else
                Set colAndAppear = New ColumnAppearances
                colAndAppear.ColumnName = cellValue
                colAndAppear.Appearances = 1
                foundColumnHeaders.Add colAndAppear
                RowColumn(0) = RowNo
                RowColumn(1) = ColumnNo
                columnPosition.Add RowColumn, Key:=cellValue
            End If

        Else
            invalidColumnHeaders.Add cellValue
            MsgBox (cellValue & " ist ungültig!")
        End If
    Next ColumnNo
    
    Dim missingColumnsColl As Collection
    Set missingColumnsColl = GetMissingColumnsCollection(foundColumnHeaders)
    
    Dim correctColumns As String
    Dim invalidColumns As String
    Dim missingColumns As String
    
    Dim entry As Variant
    Dim cAA As ColumnAppearances
    
    For Each entry In foundColumnHeaders
        Set cAA = entry
        correctColumns = correctColumns & cAA.ColumnName & " - " & CStr(cAA.Appearances) & vbNewLine
    Next entry
 
    For Each entry In invalidColumnHeaders
        invalidColumns = invalidColumns & entry & vbNewLine
    Next entry
    
    For Each entry In missingColumnsColl
        missingColumns = missingColumns & entry & vbNewLine
    Next entry
    
    If foundColumnHeaders.Count <> 19 Then
        MsgBox ("Korrekt vorgefunden:" & vbNewLine & correctColumns & vbNewLine & _
            "Ungültige Einträge:" & vbNewLine & invalidColumns & vbNewLine & _
            "Vermisste Spalten:" & vbNewLine & missingColumns)
        Exit Sub
    End If
    
    RowColumn(0) = 0
    RowColumn(1) = 0
    
    InsertColumnHeadersAt 1, 20
    
    
    Dim LastRow As Integer
    LastRow = Application.activeSheet.UsedRange.SpecialCells(XlCellType.xlCellTypeLastCell).row
    CreateColumnComparison columnPosition, LastRow, 20
    
    MsgBox ("Aufbau der Vergleichsspalten abgeschlossen.")
End Sub

Private Function IsValidHeaderText(value As String) As Boolean
    Dim hdrColl As Collection
    Set hdrColl = GetExpectedColumnsCollection()
    Dim entry As Variant
    
    IsValidHeaderText = False
    
    For Each entry In hdrColl
        If value = entry Then
            IsValidHeaderText = True
            Exit Function
        End If
    Next entry
End Function

Public Function GetExpectedColumnsCollection() As Collection
    Dim impWrkshtColumnHdrs As Collection
    Set impWrkshtColumnHdrs = New Collection
    
    impWrkshtColumnHdrs.Add ("Type")
    impWrkshtColumnHdrs.Add ("No.")
    impWrkshtColumnHdrs.Add ("Existing Object Name")
    impWrkshtColumnHdrs.Add ("Name")
    impWrkshtColumnHdrs.Add ("New Object")
    impWrkshtColumnHdrs.Add ("New Object Changed")
    impWrkshtColumnHdrs.Add ("Existing Object Changed")
    impWrkshtColumnHdrs.Add ("Warning")
    impWrkshtColumnHdrs.Add ("Action")
    impWrkshtColumnHdrs.Add ("Existing Object Modified")
    impWrkshtColumnHdrs.Add ("Existing Object Version List")
    impWrkshtColumnHdrs.Add ("New Object Modified")
    impWrkshtColumnHdrs.Add ("New Object Version List")
    impWrkshtColumnHdrs.Add ("Existing Object Size")
    impWrkshtColumnHdrs.Add ("New Object Size")
    impWrkshtColumnHdrs.Add ("Existing Object Date")
    impWrkshtColumnHdrs.Add ("New Object Date")
    impWrkshtColumnHdrs.Add ("Existing Object Time")
    impWrkshtColumnHdrs.Add ("New Object Time")

    Set GetExpectedColumnsCollection = impWrkshtColumnHdrs
End Function

Public Function IsHeaderInFoundColumnHeaders(header As String, found As Collection) As Boolean
    
    Dim entry As Variant
    
    IsHeaderInFoundColumnHeaders = False
    
    For Each entry In found
        Dim cAA As ColumnAppearances
        Set cAA = entry
        If header = cAA.ColumnName Then
            IsHeaderInFoundColumnHeaders = True
            Exit Function
        End If
    Next entry

End Function

Private Function GetMissingColumnsCollection(found As Collection) As Collection
    Dim expected, missing As Collection
    Dim cAA As ColumnAppearances
    
    Dim entry, entry2 As Variant
    Dim expectedColumn As String
    Dim headerFound As Boolean
    
    Set expected = GetExpectedColumnsCollection
    Set missing = New Collection
    
    For Each entry In expected
        expectedColumn = entry
        headerFound = False
        
        For Each entry2 In found
            Set cAA = entry2
            If expectedColumn = cAA.ColumnName Then
                headerFound = True
            End If
        Next entry2
        
        If Not headerFound Then
            missing.Add expectedColumn
        End If
    Next entry
        
    Set GetMissingColumnsCollection = missing
End Function

Public Function IncrementAppearance(header As String, found As Collection)
    Dim entry As Variant
    
    For Each entry In found
        Dim cAA As ColumnAppearances
        Set cAA = entry
        If header = cAA.ColumnName Then
            cAA.Appearances = cAA.Appearances + 1
            Exit Function
        End If
    Next entry
End Function

Private Sub InsertColumnHeadersAt(row As Integer, column As Integer)
    Dim newColumnHeaders As Collection
    Set newColumnHeaders = GetNewColumnHeaders()
    
    Dim entry As Variant
    
    For Each entry In newColumnHeaders
        Application.activeSheet.Cells(row, column).value = entry
        column = column + 1
    Next entry
End Sub

Public Function GetNewColumnHeaders() As Collection
    Dim impDifferingColumns As Collection
    Set impDifferingColumns = New Collection
    
    impDifferingColumns.Add ("Time differs")
    impDifferingColumns.Add ("Name differs")
    impDifferingColumns.Add ("Date differs")
    impDifferingColumns.Add ("Modified differs")
    impDifferingColumns.Add ("Version list differs")
    impDifferingColumns.Add ("Size differs")
    impDifferingColumns.Add ("Object changed differs")
    impDifferingColumns.Add ("Any difference")
    impDifferingColumns.Add ("Any difference except Size")
    
    Set GetNewColumnHeaders = impDifferingColumns
    
End Function

Public Function CreateColumnComparison(columnPosition As Collection, lastLine As Integer, startColumn As Integer)
    Dim newColumns As Collection
    Set newColumns = GetNewColumnHeaders
    
    Dim ColumnLeft, ColumnRight As Integer
    Dim entry As Variant
    Dim CurrentLineNo As Integer
    Dim FirstComparingColumn As Integer
    
    FirstComparingColumn = startColumn
    
    For Each entry In newColumns
        Select Case entry
            Case "Time differs"
                ColumnLeft = columnPosition.Item("Existing Object Time")(1)
                ColumnRight = columnPosition.Item("New Object Time")(1)
            Case "Date differs"
                ColumnLeft = columnPosition.Item("Existing Object Date")(1)
                ColumnRight = columnPosition.Item("New Object Date")(1)
            Case "Name differs"
                ColumnLeft = columnPosition.Item("Existing Object Name")(1)
                ColumnRight = columnPosition.Item("Name")(1)
            Case "Size differs"
                ColumnLeft = columnPosition.Item("Existing Object Size")(1)
                ColumnRight = columnPosition.Item("New Object Size")(1)
            Case "Version list differs"
                ColumnLeft = columnPosition.Item("Existing Object Version List")(1)
                ColumnRight = columnPosition.Item("New Object Version List")(1)
            Case "Modified differs"
                ColumnLeft = columnPosition.Item("Existing Object Modified")(1)
                ColumnRight = columnPosition.Item("New Object Modified")(1)
            Case "Object changed differs"
                ColumnLeft = columnPosition.Item("Existing Object Changed")(1)
                ColumnRight = columnPosition.Item("New Object Changed")(1)
        End Select
        
        CurrentLineNo = 2
        
        Select Case entry
            Case "Any difference"
                CurrentLineNo = 2
                activeSheet.Range(Cells(CurrentLineNo, startColumn), Cells(lastLine, startColumn)).Formula = _
                    GetForumlaForAnyDifference(FirstComparingColumn, newColumns, CurrentLineNo)
            Case "Any difference except Size"
                CurrentLineNo = 2
                activeSheet.Range(Cells(CurrentLineNo, startColumn), Cells(lastLine, startColumn)).Formula = _
                    GetForumlaForAnyDifferenceExceptSize(FirstComparingColumn, newColumns, CurrentLineNo)
            Case Else
                CurrentLineNo = 2
                activeSheet.Range(Cells(CurrentLineNo, startColumn), Cells(lastLine, startColumn)).Formula = _
                    "=" & Cells(CurrentLineNo, ColumnLeft).Address(RowAbsolute:=False, ColumnAbsolute:=False) & _
                    "<>" & Cells(CurrentLineNo, ColumnRight).Address(RowAbsolute:=False, ColumnAbsolute:=False)
        End Select
        
        startColumn = startColumn + 1
    Next entry
End Function

Public Function GetForumlaForAnyDifference(FirstComparingColumnNo As Integer, comparingColumnsCollection As Collection, CurrentRowNo As Integer)
    Dim FormulaForColumnAnyDifference, FormulaContent As String
    Dim CurrentColumnNo As Integer
    
    FormulaForColumnAnyDifference = "=OR("
    
    For CurrentColumnNo = FirstComparingColumnNo To FirstComparingColumnNo + comparingColumnsCollection.Count - 3
        If FormulaContent <> "" Then
            FormulaContent = FormulaContent & "," '!Achtung, dass darf nicht einfach das derzeit gültige Elementtrennzeichen sein, sondern muss in Excel VBA als Trennzeichen in der Formel unterstützt werden!
        End If
        FormulaContent = FormulaContent & Cells(CurrentRowNo, CurrentColumnNo).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    Next CurrentColumnNo
    
    FormulaForColumnAnyDifference = FormulaForColumnAnyDifference & FormulaContent & ")"
    
    GetForumlaForAnyDifference = FormulaForColumnAnyDifference
End Function

Public Function GetForumlaForAnyDifferenceExceptSize(FirstComparingColumnNo As Integer, comparingColumnsCollection As Collection, CurrentRowNo As Integer)
    Dim FormulaForColumnAnyDifference, FormulaContent As String
    Dim CurrentColumnNo As Integer
    
    FormulaForColumnAnyDifference = "=OR("
    
    For CurrentColumnNo = FirstComparingColumnNo To FirstComparingColumnNo + comparingColumnsCollection.Count - 3
        'Debug.Print (CurrentColumnNo + 1 - FirstComparingColumnNo)
        If comparingColumnsCollection.Item(CurrentColumnNo + 1 - FirstComparingColumnNo) <> "Size differs" Then
            If FormulaContent <> "" Then
                FormulaContent = FormulaContent & "," '!Achtung, dass darf nicht einfach das derzeit gültige Elementtrennzeichen sein, sondern muss in Excel VBA als Trennzeichen in der Formel unterstützt werden!
            End If
            FormulaContent = FormulaContent & Cells(CurrentRowNo, CurrentColumnNo).Address(RowAbsolute:=False, ColumnAbsolute:=False)
        End If
    Next CurrentColumnNo
    
    FormulaForColumnAnyDifference = FormulaForColumnAnyDifference & FormulaContent & ")"
    
    GetForumlaForAnyDifferenceExceptSize = FormulaForColumnAnyDifference
End Function

