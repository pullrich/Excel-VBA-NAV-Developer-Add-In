Attribute VB_Name = "M_Checks_NavTable"
Option Explicit

Sub test_NavTable()
    Dim testTable As New NavTable
    
    Debug.Print testTable.DefaultStringRepr
    Debug.Print testTable.DefaultIntegerRepr
    Debug.Print testTable.IsReprByString("TABELLE")
End Sub
