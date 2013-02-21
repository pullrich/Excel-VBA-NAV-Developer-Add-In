Attribute VB_Name = "M_Convert"
Option Explicit 'Variables must be defined explicitly (by Dim, ReDim).

Public Const NAV_FILTER_LEN As Integer = 250

Public Enum NavObjectTypes
    TABLE = 1
    FORM = 2
    REPORT = 3
    DATAPORT = 4
    CODEUNIT = 5
    XMLPORT = 6
    MenuSuite = 7
    PAGE = 8
End Enum

'Callback for Button
Sub call_numberToObjectType(control As IRibbonControl)
    Objects_NumbersToObjectTypes
End Sub


Public Sub Objects_NumbersToObjectTypes()
    Dim n As Long
    Dim UserSelection As Range
    Dim cellValue As String
    Dim navObjType As NavObjectTypes
    
    Set UserSelection = Application.selection
    
    For n = 1 To UserSelection.Rows.Count
        cellValue = UserSelection.Cells(n, 1)
        navObjType = GetObjType(cellValue)
        cellValue = GetDefaultTypeName(navObjType)
        If Not cellValue = "" Then
            UserSelection.Cells(n, 1) = cellValue
        End If
    Next n
End Sub

Sub call_ObjectTypeStringToNumber(control As IRibbonControl)
    Objects_ObjectTypesToNumbers
End Sub

Public Sub Objects_ObjectTypesToNumbers()
    Dim n As Long
    Dim UserSelection As Range
    Dim cellValue As String
    Dim navObjType As NavObjectTypes
    
    Set UserSelection = Application.selection
    
    For n = 1 To UserSelection.Rows.Count
        cellValue = UserSelection.Cells(n, 1)
        navObjType = GetObjType(cellValue)
        cellValue = CStr(navObjType)
        If Not cellValue = "0" Then
            UserSelection.Cells(n, 1) = cellValue
        End If
    Next n
End Sub

Public Function GetObjType(Id As String) As NavObjectTypes
    Select Case Id
        Case "1", "Table", "Tabelle"
            GetObjType = NavObjectTypes.TABLE
            Exit Function
        Case "2", "Form", "Formular"
            GetObjType = NavObjectTypes.FORM
            Exit Function
        Case "3", "Report", "Bericht"
            GetObjType = NavObjectTypes.REPORT
            Exit Function
        Case "4", "Dataport"
            GetObjType = NavObjectTypes.DATAPORT
            Exit Function
        Case "5", "Codeunit"
            GetObjType = NavObjectTypes.CODEUNIT
            Exit Function
        Case "6", "XMLport"
            GetObjType = NavObjectTypes.XMLPORT
            Exit Function
        Case "7", "MenuSuite"
            GetObjType = NavObjectTypes.MenuSuite
            Exit Function
        Case "8", "Page"
            GetObjType = NavObjectTypes.PAGE
            Exit Function
    End Select
End Function

Public Function GetDefaultTypeName(nType As NavObjectTypes) As String
    Select Case nType
        Case NavObjectTypes.TABLE
            GetDefaultTypeName = "Table"
            Exit Function
        Case NavObjectTypes.FORM
            GetDefaultTypeName = "Form"
            Exit Function
        Case NavObjectTypes.REPORT
            GetDefaultTypeName = "Report"
            Exit Function
        Case NavObjectTypes.DATAPORT
            GetDefaultTypeName = "Dataport"
            Exit Function
        Case NavObjectTypes.CODEUNIT
            GetDefaultTypeName = "Codeunit"
            Exit Function
        Case NavObjectTypes.XMLPORT
            GetDefaultTypeName = "XMLport"
            Exit Function
        Case NavObjectTypes.MenuSuite
            GetDefaultTypeName = "MenuSuite"
            Exit Function
        Case NavObjectTypes.PAGE
            GetDefaultTypeName = "Page"
            Exit Function
    End Select
End Function

