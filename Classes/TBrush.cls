VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TBrush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'Private Declare Function CreateBrushIndirect Lib "gdi32.dll" (ByRef lpLogBrush As LOGBRUSH) As Long
'in AXdll Instancing: 5 - MultiUse
Private mHandle As HBRUSH
Private mColor As TColor 'Long
Private mStyle As TBrushStyle
Private mBitmap As StdPicture
Public Event OnChange(Sender As TBrush)

Private Sub Class_Initialize()
  SetDefaultBrush
End Sub
Private Sub SetDefaultBrush()
  mHandle = 0
  mColor = clWhite
  mStyle = bsSolid
  Set mBitmap = Nothing
End Sub
Public Property Get Handle() As HBRUSH
Dim mLogBrush As LOGBRUSH

try: On Error GoTo Except
  If mHandle = 0 Then
    With mLogBrush
      If Not mBitmap Is Nothing Then
        .lbStyle = BS_PATTERN
        'Brush.Bitmap.HandleType = bmDDB
        .lbHatch = mBitmap.Handle
      Else
        .lbHatch = 0
        Select Case mStyle
        Case bsSolid: .lbStyle = BS_SOLID
        Case bsClear: .lbStyle = BS_HOLLOW
        Case Else
          .lbStyle = BS_HATCHED
          .lbHatch = mStyle - 2 'Ord(bsHorizontal)
        End Select
      End If
      .lbColor = ColorToRGB(mColor)
    End With
    mHandle = CreateBrushIndirect(mLogBrush)
  End If
  Handle = mHandle
  Exit Property
Except:
  Handle = 0
End Property
Public Property Let Handle(LngVal As HBRUSH)
  If mHandle <> LngVal Then
    Call DeleteHandle '
    mHandle = LngVal
    'RaiseEvent OnChange
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

Public Property Get Style() As TBrushStyle
  Style = mStyle
End Property
Public Property Let Style(LngVal As TBrushStyle)
  If mStyle <> LngVal Then
    mStyle = LngVal
    Call DeleteHandle 'mHandle = 0
    RaiseEvent OnChange(Me)
  End If
End Property

Public Property Get Bitmap() As StdPicture 'ClsBitmap
  Set Bitmap = mBitmap
End Property
Public Property Set Bitmap(Bmp As StdPicture) 'ClsBitmap)
  Set mBitmap = Bmp
  Call DeleteHandle 'mHandle = 0
  RaiseEvent OnChange(Me)
End Property

Public Sub Assign(SrcBrush As TBrush)
  'If Not mBitmap Is Nothing Then
    Set Me.Bitmap = SrcBrush.Bitmap
  'End If
  Me.Color = SrcBrush.Color
  'Me.Handle = SrcBrush.Handle
  Me.Style = SrcBrush.Style
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
