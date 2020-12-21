VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'in AXdll Instancing: 5 - MultiUse
Private mHandle As HFONT
Private mColor As TColor ' Long
Private mPixelsPerInch As Long
Private mCharSet As Byte ' TCharset = 0..255
Private mHeight As Long
Private mName As String
Private mPitch As TFontPitch 'Zeichenabstand variabel oder fest
Private mStyle As TFontStyles 'Long
Public Event Changed(Sender As TFont)

Private Sub Class_Initialize()
  SetDefaultFont
End Sub
Private Sub SetDefaultFont()
  mHandle = 0
  mColor = clWindowText ' Long
  mPixelsPerInch = GetScreenLogPixels 'ppi 'Screen.TwipsperpixelY
  mCharSet = DEFAULT_CHARSET ' TCharset = 0..255
  Size = 8
  mName = "MS Sans Serif"
  mPitch = fpDefault 'Zeichenabstand variabel oder fest
  mStyle = 0
End Sub
'Public Type LOGFONT
'  lfHeight As Long
'  lfWidth As Long
'  lfEscapement As Long
'  lfOrientation As Long
'  lfWeight As Long
'  lfItalic As Byte
'  lfUnderline As Byte
'  lfStrikeOut As Byte
'  lfCharSet As Byte
'  lfOutPrecision As Byte
'  lfClipPrecision As Byte
'  lfQuality As Byte
'  lfPitchAndFamily As Byte
'  lfFaceName As String * LF_FACESIZE
'End Type
Public Property Get Handle() As HFONT
Dim mLogFont As LOGFONT
Dim H As Long
Dim i As Long, PChar As String
try: On Error GoTo Except
  If mHandle = 0 Then
    With mLogFont
      If mHandle = 0 Then
      '.lfHeight = -MulDiv(.Size, GetDeviceCaps(GetDC(GetActiveWindow). LOGPIXELSY), 72)
        .lfHeight = Height '(mHeight * -20) / Screen.TwipsPerPixelY
        'Windows expects the font size in Pixels and to be negative
        .lfWidth = 0        '; { have font mapper choose }
        .lfEscapement = 0   '; { only straight fonts } '1800 = 180°
        .lfOrientation = 0  '; { no rotation }
        If (mStyle And fsBold) Then .lfWeight = FW_BOLD Else: .lfWeight = FW_NORMAL
        .lfItalic = CByte(mStyle And fsItalic)
        .lfUnderline = CByte(mStyle And fsUnderline)
        .lfStrikeOut = CByte(mStyle And fsStrikeOut)
        .lfCharSet = CByte(mCharSet)
        .lfQuality = CByte(DEFAULT_QUALITY)
        '{ Everything else as default }
        .lfOutPrecision = CByte(OUT_DEFAULT_PRECIS)
        .lfClipPrecision = CByte(CLIP_DEFAULT_PRECIS)
        .lfPitchAndFamily = CByte(mPitch)
        .lfFaceName = mName & Chr$(0) 'Str2ByA(mName) '
        mHandle = CreateFontIndirect(mLogFont)
      End If
    End With
  End If
  Handle = mHandle
  Exit Property
Except:
  Handle = 0
End Property
Public Property Let Handle(LngVal As HFONT)
  Call DeleteHandle '
  mHandle = LngVal
End Property

Public Property Get PixelsPerInch() As Long
  PixelsPerInch = mPixelsPerInch
End Property
Public Property Let PixelsPerInch(LngVal As Long)
  mPixelsPerInch = LngVal
  Call DeleteHandle 'mHandle = 0
  RaiseEvent Changed(Me)
End Property

Public Property Get Charset() As Byte
  Charset = mCharSet
End Property
Public Property Let Charset(BytVal As Byte)
  mCharSet = BytVal
  Call DeleteHandle 'mHandle = 0
  RaiseEvent Changed(Me)
End Property

Public Property Get Color() As TColor
  Color = mColor
End Property
Public Property Let Color(LngVal As TColor)
  mColor = LngVal
  Call DeleteHandle 'mHandle = 0
  RaiseEvent Changed(Me)
End Property

Public Property Get Height() As Long
  Height = mHeight
End Property
Public Property Let Height(LngVal As Long)
  mHeight = LngVal
  Call DeleteHandle 'mHandle = 0
  RaiseEvent Changed(Me)
End Property

Public Property Get Name() As String
  Name = mName
End Property
Public Property Let Name(StrVal As String)
  mName = StrVal
  Call DeleteHandle 'mHandle = 0
  RaiseEvent Changed(Me)
End Property

Public Property Get Pitch() As TFontPitch
  Pitch = mPitch
End Property
Public Property Let Pitch(LngVal As TFontPitch)
  mPitch = LngVal
  Call DeleteHandle 'mHandle = 0
  RaiseEvent Changed(Me)
End Property

Public Property Get Size() As Long
  Size = -mHeight * Screen.TwipsPerPixelY / 20 'mSize
End Property
Public Property Let Size(LngVal As Long)
  mHeight = -LngVal * 20 / Screen.TwipsPerPixelY
  Call DeleteHandle 'mHandle = 0
  RaiseEvent Changed(Me)
End Property

Public Property Get Style() As TFontStyles 'Long
  Style = mStyle
End Property
Public Property Let Style(LngVal As TFontStyles) 'Long)
  mStyle = LngVal
  Call DeleteHandle 'mHandle = 0
End Property

Public Sub Assign(SrcFont As TFont)
  Me.Charset = SrcFont.Charset
  Me.Color = SrcFont.Color
  Me.Height = SrcFont.Height
  Me.Name = SrcFont.Name
  Me.Pitch = SrcFont.Pitch
  Me.PixelsPerInch = SrcFont.PixelsPerInch
  Me.Size = SrcFont.Size
  Me.Style = SrcFont.Style
  'Handle braucht nicht, da sowieso zu null gesetzt wird
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
  DeleteHandle
End Sub
