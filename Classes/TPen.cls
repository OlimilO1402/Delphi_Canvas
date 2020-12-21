VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TPen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'in AXdll Instancing: 5 - MultiUse
Private mHandle As HPEN 'Long
Private mColor As TColor 'Long
Private mMode As TPenMode
Private mStyle As TPenStyle
'Private mEndCap As TEndCapStyle
Private mWidth As Long
'Private Declare Function CreatePenIndirect Lib "gdi32.dll" (ByRef lpLogPen As LOGPEN) As Long
Public Event OnChange(Sender As TPen)

Private Sub Class_Initialize()
  SetDefaultPen
End Sub
Private Sub SetDefaultPen()
  mHandle = 0
  mColor = clBlack
  mMode = pmCopy
  mStyle = psSolid
  mWidth = 1
End Sub
Public Property Get Handle() As HPEN
Dim mLogPen As LOGPEN
'Dim mLogBrush As LOGBRUSH
try: On Error GoTo Except
  If mHandle = 0 Then
    With mLogPen
      .lopnColor = ColorToRGB(mColor)
      .lopnStyle = mStyle 'Or mEndCap
      .lopnWidth.X = mWidth
      .lopnWidth.Y = 0 'mEndCap
    End With
    'mLogBrush.lbColor = mColor
    'mLogBrush.lbHatch = 0&
    'mLogBrush.lbStyle = 0&
    'ExtCreatePen Lib "gdi32.dll"
    '(ByVal dwPenStyle As Long,
    ' ByVal dwWidth As Long,
    ' ByRef lplb As LOGBRUSH,
    ' ByVal dwStyleCount As Long,
    ' ByRef lpStyle As Long) As Long
    mHandle = CreatePenIndirect(mLogPen)
    'mHandle = ExtCreatePen(mStyle, mWidth, mLogBrush, 0&, mStyle) '10& ????
  End If
  Handle = mHandle
  Exit Property
Except:
  Handle = 0
End Property
Public Property Let Handle(LngVal As HPEN)
  If mHandle <> LngVal Then
    Call DeleteHandle '
    mHandle = LngVal
  End If
End Property

Public Property Get Color() As TColor
  Color = mColor
End Property
Public Property Let Color(LngVal As TColor)
  If mColor <> LngVal Then
    mColor = LngVal
    Call DeleteHandle 'mHandle = 0
    RaiseEvent OnChange(Me)
  End If
End Property

Public Property Get Mode() As TPenMode
  Mode = mMode
End Property
Public Property Let Mode(LngVal As TPenMode)
  If mMode <> LngVal Then
    mMode = LngVal
    Call DeleteHandle 'mHandle = 0
    RaiseEvent OnChange(Me)
  End If
End Property

Public Property Get Style() As TPenStyle
  Style = mStyle
End Property
Public Property Let Style(LngVal As TPenStyle)
  If mStyle <> LngVal Then
    mStyle = LngVal
    Call DeleteHandle 'mHandle = 0
    RaiseEvent OnChange(Me)
  End If
End Property

'Public Property Get EndCapStyle() As EEndCapStyle
'  EndCapStyle = mEndCap
'End Property
'Public Property Let EndCapStyle(LngVal As EEndCapStyle)
'  If mEndCap <> LngVal Then
'    mEndCap = LngVal
'    mHandle = 0
'    RaiseEvent OnChange(Me)
'  End If
'End Property

Public Property Get Width() As Long
  Width = mWidth
End Property
Public Property Let Width(LngVal As Long)
  If mWidth <> LngVal Then
    mWidth = LngVal
    Call DeleteHandle 'mHandle = 0
    RaiseEvent OnChange(Me)
  End If
End Property

Public Sub Assign(SrcPen As TPen)
  Me.Color = SrcPen.Color
  'Me.Handle = SrcPen.Handle
  Me.Mode = SrcPen.Mode
  Me.Style = SrcPen.Style
  Me.Width = SrcPen.Width
  'Me.EndCapStyle = SrcPen.EndCapStyle
End Sub

Private Sub DeleteHandle()
  If mHandle <> 0 Then
    Call DeleteObject(mHandle)
    mHandle = 0
  End If
End Sub

Private Sub Class_Terminate()
'  Call DeleteObject(mHandle)
'  mHandle = 0
  Call DeleteHandle '
End Sub

