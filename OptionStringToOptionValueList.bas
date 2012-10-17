Attribute VB_Name = "OptionStringToOptionValueList"
Sub OptionStringToOptionValueList()
'
' OptionStringToOptionValueList Makro
' Turns an OptionString in the current cell into an OptionValue list.
'

    Dim cellValue As String
    cellValue = ActiveCell.Value
    
    Dim optionValues As Variant
    optionValues = Split(cellValue, ",")
    
    Dim currVal As Integer
    Dim newCellValue As String
    For currVal = 0 To UBound(optionValues)
        newCellValue = newCellValue & currVal & "=" & optionValues(currVal)
        If currVal <> UBound(optionValues) Then
            newCellValue = newCellValue & vbCrLf
        End If
    Next currVal
    
    ActiveCell.Value = newCellValue
End Sub
