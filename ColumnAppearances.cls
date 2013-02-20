VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColumnAppearances"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pColumnName As String
Private pAppearances As Integer

Public Property Get ColumnName() As String
    ColumnName = pColumnName
End Property

Public Property Let ColumnName(value As String)
    pColumnName = value
End Property

Public Property Get Appearances() As Integer
    Appearances = pAppearances
End Property

Public Property Let Appearances(value As Integer)
    pAppearances = value
End Property


