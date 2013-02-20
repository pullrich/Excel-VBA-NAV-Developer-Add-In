VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NavTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private iDefaultIntegerRepr As Integer
Private sDefaultStringRepr As String
Private vEquivalentStrings As Variant

Property Get DefaultIntegerRepr() As Integer
    DefaultIntegerRepr = iDefaultIntegerRepr
End Property

Property Get DefaultStringRepr() As String
    DefaultStringRepr = sDefaultStringRepr
End Property

Public Function IsReprByString(sRepr As String) As Boolean
    Dim i As Integer
    
    For i = LBound(vEquivalentStrings) To UBound(vEquivalentStrings)
        If LCase(sRepr) = LCase(vEquivalentStrings(i)) Then
            IsReprByString = True
            Exit Function
        End If
    Next i
    
    IsReprByString = False
End Function

Private Sub Class_Initialize()
    iDefaultIntegerRepr = 1
    sDefaultStringRepr = "Table"
    vEquivalentStrings = Array(CStr(iDefaultIntegerRepr), "table", "tabelle")
End Sub

