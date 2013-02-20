Attribute VB_Name = "M_VBA_Tests"
Option Explicit
' This module only contains functions which test VBA behavior.

Private Sub MakeNewArrays()
    ' Test: Assign one array to an index in a different array.
    Dim i As Integer
    Dim obary(1) As Variant
    Dim intary() As Integer
    
    ReDim intary(1)
    intary(0) = 1
    intary(1) = 2
    obary(0) = intary
    
    ReDim intary(2)
    intary(0) = 3
    intary(1) = 4
    intary(2) = 5
    obary(1) = intary

    
    'intarytest = obary(0)
    Debug.Print CStr(obary(0)(0))
    Debug.Print CStr(obary(0)(1))
    
    Debug.Print CStr(obary(1)(0))
    Debug.Print CStr(obary(1)(1))
    Debug.Print CStr(obary(1)(2))
    
    Debug.Print ""
    Debug.Print UBound(obary(0))
    Debug.Print UBound(obary(1))
    
    ' Success: Different arrays are available in "obary".
    
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, _
        ByVal Target As Excel.Range)
    Application.StatusBar = Sh.Name & ":" & Target.Address
End Sub


Sub TestArray()
    Dim Equivalents As Variant
    Equivalents = Array("1", "Table", "Tabelle")
    Debug.Print Equivalents(2)
End Sub

Sub Test_RGB()
    Debug.Print CStr(RGB(9, 205, 14))
End Sub
