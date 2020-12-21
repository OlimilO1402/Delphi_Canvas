VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'in AXdll Instancing: 5 - MultiUse
Public X As Long
Public Y As Long
Public dX As Long
Public dY As Long

Public Sub AssignPt(ByVal SrcPoint As ClsPoint)
  dX = SrcPoint.X - Me.X
  dY = SrcPoint.Y - Me.Y
  Me.X = SrcPoint.X
  Me.Y = SrcPoint.Y
End Sub
Public Sub AssignXY(ByVal LngX As Long, ByVal LngY As Long)
  dX = LngX - Me.X
  dY = LngY - Me.Y
  Me.X = LngX
  Me.Y = LngY
End Sub

