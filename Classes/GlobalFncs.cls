VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GlobalFncs"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Sub Class_Initialize()
  Call InitColors
End Sub

Public Function ColorToRGB(Color As TColor) As Long
  If Color < 0 Then
    ColorToRGB = GetSysColor(Color And &HFF&)
  Else
    ColorToRGB = Color
  End If
End Function

Public Function ColorToString(Color As TColor) As String
  If Not ColorToIdent(Color, ColorToString) Then ColorToString = Hex$(Color)
    'FmtStr(Result, '%s%.8x', [HexDisplayPrefix, Color]);
End Function

Function StringToColor(S As String) As TColor
  If Not IdentToColor(S, CLng(StringToColor)) Then StringToColor = CLng(S)
End Function


