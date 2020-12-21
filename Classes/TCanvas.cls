VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TCanvas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'in AXdll Instancing: 5 - MultiUse
Private mHandle As HDC
Private mState As TCanvasStates
Private WithEvents mBrush As TBrush
Attribute mBrush.VB_VarHelpID = -1
Private WithEvents mFont As TFont
Attribute mFont.VB_VarHelpID = -1
Private WithEvents mPen As TPen
Attribute mPen.VB_VarHelpID = -1
Private mPenPos As TPoint 'die letzte Position
Private mCopyMode As TCopyMode 'in CopyRect
Private mCanvasOrientation As TCanvasOrientation
Private mTextFlags As Long
Public Event OnChange(Sender As TCanvas)
Public Event OnChanging(Sender As TCanvas)

Private Sub Class_Initialize() '.Create
  mHandle = 0
  mState = 0
  mCopyMode = cmSrcCopy
  StockBrush = GetStockObject(HOLLOW_BRUSH)
  StockFont = GetStockObject(SYSTEM_FONT)
  StockPen = GetStockObject(BLACK_PEN)
  'StockIcon = LoadIcon(0, IDI_APPLICATION)
  mPenPos.X = 0
  mPenPos.Y = 0
  Set mBrush = New TBrush
  Set mFont = New TFont
  Set mPen = New TPen
End Sub

Public Property Get Handle() As HDC
  Call Changing
  Call RequiredState(csAllValid)
  Handle = mHandle
End Property
Public Property Let Handle(HndVal As HDC)
  If mHandle <> HndVal Then
    If mHandle <> 0 Then
      Call DeselectHandles
      Call GetPenPos(mPenPos.X, mPenPos.Y)
      mHandle = 0
      Call ExcludeState(mState, csHandleValid)
    End If
    If HndVal <> 0 Then
      Call IncludeState(mState, csHandleValid)
      mHandle = HndVal
      Call SetPenPos(mPenPos.X, mPenPos.Y)
    End If
  End If
End Property

Public Property Get Brush() As TBrush
  Set Brush = mBrush
End Property
Public Property Set Brush(ABrush As TBrush)
  Set mBrush = ABrush
End Property

Public Property Get Font() As TFont
  Set Font = mFont
End Property
Public Property Set Font(AFont As TFont)
  Set mFont = AFont
End Property

Public Property Get Pen() As TPen
  Set Pen = mPen
End Property
Public Property Set Pen(APen As TPen)
  Set mPen = APen
End Property

Public Property Get CopyMode() As TCopyMode
  CopyMode = mCopyMode
End Property
Public Property Let CopyMode(LngVal As TCopyMode)
  mCopyMode = LngVal
End Property

Public Property Get CanvasOrientation() As TCanvasOrientation
Dim Point As TPoint
  CanvasOrientation = coLeftToRight
  If (mTextFlags And ETO_RTLREADING) <> 0 Then
    Call Windows.GetWindowOrgEx(mHandle, Point)
    If Point.X <> 0 Then CanvasOrientation = coRightToLeft
  End If
End Property
Public Property Let CanvasOrientation(co As TCanvasOrientation)
  mCanvasOrientation = co
End Property

Public Property Get TextFlags() As Long
  TextFlags = mTextFlags
End Property
Public Property Let TextFlags(LngVal As Long)
  mTextFlags = LngVal
End Property

Public Property Get Pixels(X As Long, Y As Long) As TColor
  Call RequiredState(csHandleValid)
  Pixels = Windows.GetPixel(mHandle, X, Y)
End Property
Public Property Let Pixels(X As Long, Y As Long, AColor As TColor)
  Call Changing
  Call RequiredState(csHandleValid Or csPenValid)
  Call Windows.SetPixel(mHandle, X, Y, ColorToRGB(AColor))
  Call Changed
End Property
Public Sub Farbverlauf2D(StrtX As Long, StepX As Long, EndeX As Long, StrtY As Long, StepY As Long, EndeY As Long, StepClX As Long, StepClY As Long, StrtColor As TColor)
Dim i As Long, j As Long, AColor As TColor, StepCl As TColor
  
  Call Changing
  Call RequiredState(csHandleValid Or csPenValid)
  AColor = StrtColor
  'StepCl = AColor
  For j = StrtY To EndeY Step StepY
    For i = StrtX To EndeX Step StepX
      Call Windows.SetPixel(mHandle, i, j, ColorToRGB(AColor))
      StepCl = StepCl + StepClX
      AColor = AColor + StepCl 'Or StepClX Or StepClY
    Next
    StepCl = StepCl + StepClY
  Next
  Call Changed
End Sub

Public Sub GetPenPos(ByRef X As Long, ByRef Y As Long)
  Call RequiredState(csHandleValid)
  Call Windows.GetCurrentPositionEx(mHandle, mPenPos)
  X = mPenPos.X
  Y = mPenPos.Y
End Sub
Public Sub SetPenPos(X As Long, Y As Long)
  mPenPos.X = X
  mPenPos.Y = Y
  Call MoveTo(X, Y)
End Sub

Public Property Get PenPos() As TPoint
  Call RequiredState(csHandleValid)
  Call Windows.GetCurrentPositionEx(mHandle, mPenPos)
  'X = mPenPos.X
  'Y = mPenPos.Y
  PenPos = mPenPos
End Property
Public Property Let PenPos(aPt As TPoint)
  mPenPos.X = aPt.X
  mPenPos.Y = aPt.Y
  Call MoveTo(mPenPos.X, mPenPos.Y)
End Property


Public Property Get ClipRect() As Rect
'ByRef Left As Long, ByRef Top As Long, ByRef Right As Long, ByRef Bottom As Long)
  RequiredState (csHandleValid)
  Call GetClipBox(mHandle, ClipRect)
End Property

Private Sub Changing()
  'If Not OnChanging Is Nothing Then
  RaiseEvent OnChanging(Me)
End Sub
Private Sub Changed()
  'If Not OnChange Is Nothing Then
  RaiseEvent OnChange(Me)
End Sub

Private Sub IncludeState(ByRef LngSta As Long, Valid As Long)
  'ist Bit noch nicht drin?
  If Not (LngSta Or Valid) Then  'ja, dann reinnehmen
    LngSta = LngSta + Valid      'LngSta = (LngSta + Valid) 'Or ist auch OK
  End If
End Sub
Private Sub ExcludeState(ByRef LngSta As Long, Valid As Long)
  'ist Bit schon drin?
  If (LngSta Or Valid) Then  'ja, dann rausnehmen
    LngSta = LngSta - Valid  'LngSta = (LngSta - Valid) halt bei And wird Bool-Verknüpft, also True oder False
  End If
End Sub
Friend Sub RequiredState(ReqState As TCanvasStates)
Dim NeededState As TCanvasStates
  NeededState = ReqState - mState
  If NeededState <> 0 Then
    If (NeededState And csHandleValid) Then
      Call CreateHandle ' 'If FHandle = 0 Then  raise EInvalidOperation.CreateRes(@SNoCanvasHandle)
    End If
    If (NeededState And csFontValid) Then
      CreateFont
    End If
    If (NeededState And csPenValid) Then
      CreatePen
    End If
    If (NeededState And csBrushValid) Then
      CreateBrush
    End If
    mState = mState + NeededState
  End If
End Sub

Public Sub mBrush_OnChange(Sender As TBrush)
  If (mState And csBrushValid) Then
    Call ExcludeState(mState, csBrushValid)
    Call SelectObject(mHandle, StockBrush)
  End If
End Sub
Public Sub mFont_OnChange(Sender As TFont)
  If (mState And csFontValid) Then
    Call ExcludeState(mState, csFontValid)
    Call SelectObject(mHandle, StockFont)
  End If
End Sub
Private Sub mPen_OnChange(Sender As TPen)
  If (mState And csPenValid) Then
    Call ExcludeState(mState, csPenValid)
    Call SelectObject(mHandle, StockPen)
  End If
End Sub

Private Sub CreateHandle()
  'leer
End Sub
Private Sub CreateBrush()
  Call UnrealizeObject(mBrush.Handle)
  Call SelectObject(mHandle, mBrush.Handle)
  If (mBrush.Style And bsSolid) Then
    Call SetBkColor(mHandle, ColorToRGB(mBrush.Color))
    Call SetBkMode(mHandle, OPAQUE)
  Else
    '{ Win95 doesn't draw brush hatches if bkcolor = brush color }
    '{ Since bkmode is transparent, nothing should use bkcolor anyway }
    Call SetBkColor(mHandle, Not ColorToRGB(mBrush.Color))
    Call SetBkMode(mHandle, TRANSPARENT)
  End If
End Sub
Private Sub CreateFont()
  Call SelectObject(mHandle, mFont.Handle)
  Call SetTextColor(mHandle, ColorToRGB(mFont.Color))
End Sub
Private Sub CreatePen()
  Call SelectObject(mHandle, mPen.Handle)
  Call SetROP2(mHandle, mPen.Mode)      'PenModes(Pen.Mode))
End Sub

Private Sub DeselectHandles()
  If (mHandle <> 0) And ((mState - (csPenValid Or csBrushValid Or csFontValid)) <> mState) Then
    Call SelectObject(mHandle, StockBrush)
    Call SelectObject(mHandle, StockFont)
    Call SelectObject(mHandle, StockPen)
    mState = mState And (csBrushValid Or csFontValid Or csPenValid)
  End If
End Sub

Public Sub Arc(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, X3 As Long, Y3 As Long, X4 As Long, Y4 As Long)
  Call Changing
  RequiredState (csHandleValid Or csBrushValid Or csPenValid)
  Call ArcXY(mHandle, X1, Y1, X2, Y2, X3, Y3, X4, Y4)
  Call Changed
End Sub

Public Sub Chord(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, X3 As Long, Y3 As Long, X4 As Long, Y4 As Long)
  Call Changing
  Call RequiredState(csHandleValid Or csBrushValid Or csPenValid)
  Call ChordXY(mHandle, X1, Y1, X2, Y2, X3, Y3, X4, Y4)
  Call Changed
End Sub
' Dst: Destination, wohin soll kopiert werden (in diesen Canvas)
' Src: Source, was soll von woher soll kopiert werden
Public Sub CopyRectA(DstL As Long, DstT As Long, DstW As Long, DstH As Long, SrcCanvas As TCanvas, SrcL As Long, SrcT As Long, SrcW As Long, SrcH As Long)
  Call Changing
  Call RequiredState(csHandleValid Or csBrushValid Or csFontValid)
  Call SrcCanvas.RequiredState(csHandleValid Or csBrushValid)
  Call StretchBlt(mHandle, DstL, DstT, DstW, DstH, _
         SrcCanvas.Handle, SrcL, SrcT, SrcW, SrcH, mCopyMode)
  Call Changed
End Sub
Public Sub CopyRect(Dest As Rect, ByRef Canvas As TCanvas, Source As Rect)
  Call Changing
  Call RequiredState(csHandleValid Or csFontValid Or csBrushValid)
  Call Canvas.RequiredState(csHandleValid Or csBrushValid)
  Call StretchBlt(mHandle, Dest.Left, Dest.Top, Dest.Right - Dest.Left, _
        Dest.Bottom - Dest.Top, Canvas.Handle, Source.Left, Source.Top, _
        Source.Right - Source.Left, Source.Bottom - Source.Top, mCopyMode)
  Call Changed
End Sub
'Public sub Draw(X as long, Y as long, Graphic as  ClsGraphic)
'  If (not( Graphic is nothing)) And (Not( Graphic.Empty)) Then
'    call Changing
'    RequiredState([csHandleValid]);
'    SetBkColor(FHandle, ColorToRGB(FBrush.Color));
'    SetTextColor(FHandle, ColorToRGB(FFont.Color));
'    Graphic.Draw(me, Rect(X, Y, X + Graphic.Width, Y + Graphic.Height));
'    CAll Changed
'  end;
'End Sub

Public Sub DrawFocusRectA(Left As Long, Top As Long, Right As Long, Bottom As Long)
  Call Changing
  Call RequiredState(csHandleValid Or csBrushValid)
  Call DrawFocusRectP(mHandle, VarPtr(Left))
  Call Changed
End Sub
Public Sub DrawFocusRect(Rec As Rect)
  Call Changing
  Call RequiredState(csHandleValid Or csBrushValid)
  Call DrawFocusRectP(mHandle, VarPtr(Rec.Left))
  Call Changed
End Sub

Public Sub Ellipse(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long)
  Call Changing
  RequiredState (csHandleValid Or csPenValid Or csBrushValid)
  Call Windows.EllipseXY(mHandle, X1, Y1, X2, Y2)
  Call Changed
End Sub

Public Sub EllipseR(Rec As Rect)
  Call Ellipse(Rec.Left, Rec.Top, Rec.Right, Rec.Bottom)
End Sub

Public Sub FillRect(Left As Long, Top As Long, Right As Long, Bottom As Long)
Dim Rec As Rect
  Rec.Left = Left:   Rec.Top = Top
  Rec.Right = Right: Rec.Bottom = Bottom
  Call Changing
  Call RequiredState(csHandleValid Or csBrushValid)
  Call Windows.FillRectA(mHandle, Rec, Brush.Handle)
  Call Changed
End Sub

Public Sub FloodFill(X As Long, Y As Long, Color As TColor, FillStyle As TFillStyle)
  Call Changing
  Call RequiredState(csHandleValid Or csBrushValid)
  Call ExtFloodFill(mHandle, X, Y, Color, FillStyle)
  Call Changed
End Sub

Public Sub FrameRect(Left As Long, Top As Long, Right As Long, Bottom As Long)
Dim Rec As Rect
  Rec.Left = Left:   Rec.Top = Top
  Rec.Right = Right: Rec.Bottom = Bottom
  Call Changing
  Call RequiredState(csHandleValid Or csBrushValid)
  Call FrameRectA(mHandle, Rec, Brush.Handle)
  Call Changed
End Sub

Public Function HandleAllocated() As Boolean
  HandleAllocated = (mHandle <> 0)
End Function

Public Sub MoveTo(X As Long, Y As Long)
Dim P As TPoint
  Call RequiredState(csHandleValid)
  Call MoveToEx(mHandle, X, Y, P)
End Sub
Public Sub LineTo(X As Long, Y As Long)
  Call Changing
  Call RequiredState(csHandleValid Or csBrushValid Or csPenValid)
  Call LineToXY(mHandle, X, Y)
  Call Changed
End Sub

Public Sub Pie(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, X3 As Long, Y3 As Long, X4 As Long, Y4 As Long)
  Call Changing
  Call RequiredState(csHandleValid Or csBrushValid Or csPenValid)
  Call PieXY(mHandle, X1, Y1, X2, Y2, X3, Y3, X4, Y4)
  Call Changed
End Sub

Public Sub PolygonP(VarPtrFirstPoint As Long, nCount As Long)
  Call Changing
  Call RequiredState(csHandleValid Or csBrushValid Or csPenValid)
  Call PolygonPA(mHandle, VarPtrFirstPoint, nCount)
  Call Changed
End Sub

Public Sub PolygonCol(PtCol As Collection)
Dim Points() As TPoint
Dim Pt As ClsPoint
Dim i As Long
  ReDim Points(1 To PtCol.Count)
  For i = 1 To PtCol.Count
    Set Pt = PtCol(i)
    Points(i).X = Pt.X
    Points(i).Y = Pt.Y
  Next
  Call Changing
  Call RequiredState(csHandleValid Or csBrushValid Or csPenValid)
  Call Polygon(mHandle, Points(1), UBound(Points))
  Call Changed
End Sub

Public Sub PolylineP(VarPtrFirstPoint As Long, nCount As Long)
  Call Changing
  Call RequiredState(csHandleValid Or csBrushValid Or csPenValid)
  Call PolylinePA(mHandle, VarPtrFirstPoint, nCount)
  Call Changed
End Sub

Public Sub PolylineCol(PtCol As Collection)
Dim Points() As TPoint
Dim Pt As ClsPoint
Dim i As Long
  ReDim Points(1 To PtCol.Count)
  For i = 1 To PtCol.Count
    Set Pt = PtCol(i)
    Points(i).X = Pt.X
    Points(i).Y = Pt.Y
  Next
  Call Changing
  Call RequiredState(csHandleValid Or csBrushValid Or csPenValid)
  Call Polyline(mHandle, Points(1), UBound(Points))
  'Call Polyline(mHandle, PtCol(1), PtCol.Count)
  Call Changed
End Sub

Public Sub PolyBezierP(VarPtrFirstPoint As Long, nCount As Long)
  Call Changing
  Call RequiredState(csHandleValid Or csBrushValid Or csPenValid)
  Call PolyBezierPA(mHandle, VarPtrFirstPoint, nCount)
  Call Changed
End Sub

Public Sub PolyBezierCol(PtCol As Collection)
Dim Points() As TPoint
Dim Pt As ClsPoint
Dim i As Long
  ReDim Points(1 To PtCol.Count)
  For i = 1 To PtCol.Count
    Set Pt = PtCol(i)
    Points(i).X = Pt.X
    Points(i).Y = Pt.Y
  Next
  Call Changing
  Call RequiredState(csHandleValid Or csBrushValid Or csPenValid)
  Call PolyBezier(mHandle, Points(0), UBound(Points) + 1)
  Call Changed
End Sub

Public Sub PolyBezierToP(VarPtrFirstPoint As Long, nCount As Long)
  Call Changing
  Call RequiredState(csHandleValid Or csBrushValid Or csPenValid)
  Call PolyBezierToPA(mHandle, VarPtrFirstPoint, nCount)
  Call Changed
End Sub

Public Sub PolyBezierToCol(PtCol As Collection)
Dim Points() As TPoint, Pt As ClsPoint, i As Long
  ReDim Points(1 To PtCol.Count)
  For i = 1 To PtCol.Count
    Set Pt = PtCol(i)
    Points(i).X = Pt.X
    Points(i).Y = Pt.Y
  Next
  Call Changing
  Call RequiredState(csHandleValid Or csBrushValid Or csPenValid)
  Call PolyBezierTo(mHandle, Points(0), UBound(Points) + 1)
  Call Changed
End Sub

Public Sub RectangleXY(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long)
  Call Changing
  Call RequiredState(csHandleValid Or csBrushValid Or csPenValid)
  Call RectangleAXY(mHandle, X1, Y1, X2, Y2)
  Call Changed
End Sub

Public Sub Rectangle(Rec As Rect)
  Call RectangleXY(Rec.Left, Rec.Top, Rec.Right, Rec.Bottom)
End Sub

Public Sub Refresh()
  Call DeselectHandles
End Sub

Public Sub RoundRect(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, X3 As Long, Y3 As Long)
  Call Changing
  Call RequiredState(csHandleValid Or csBrushValid Or csPenValid)
  Call RoundRectXY(mHandle, X1, Y1, X2, Y2, X3, Y3)
  Call Changed
End Sub

Public Sub TextOut(ByVal X As Long, ByVal Y As Long, ByVal Text As String)
Dim Distance() As Long, i As Long
Dim Rec As Rect, W As Long, H As Long
  If Text = "" Then Exit Sub
  'W = TextWidth(Text)
  'H = TextHeight(Text)
  Rec.Left = 0: Rec.Right = 0 'X + W
  Rec.Top = 0:  Rec.Bottom = 0 'Y + H
  Call Changing
  Call RequiredState(csHandleValid Or csBrushValid Or csFontValid)
  If mCanvasOrientation = coRightToLeft Then Call Inc(X, W + 1)
  ReDim Distance(Len(Text) - 1)
  For i = 0 To UBound(Distance)
    Distance(i) = TextWidth(Mid(Text, i + 1, 1)) '+ 1
  Next
  Call ExtTextOut(mHandle, X, Y, mTextFlags, ByVal 0&, Text, Len(Text), ByVal 0&) 'Distance(0))
  Call MoveTo(X + W, Y)
  Call Changed
End Sub

Public Sub TextRect(Rec As Long, X As Long, Y As Long, ByVal Text As String)  'Rect As Rect
Dim Options As Long ', Rec As Rect
  Call Changing
  Call RequiredState(csHandleValid Or csFontValid Or csBrushValid)
  Options = ETO_CLIPPED Or mTextFlags
  If Brush.Style <> bsClear Then Options = Options Or ETO_OPAQUE
  If ((mTextFlags And ETO_RTLREADING) <> 0) And (mCanvasOrientation = coRightToLeft) Then
    Call Inc(X, TextWidth(Text) + 1)
  End If
  Call ExtTextOut(mHandle, X, Y, Options, Rec, Text, Len(Text), vbNullChar)
  Call Changed
End Sub

Public Sub TextExtent(ByVal Text As String, ByRef cX As Long, ByRef cY As Long)
Dim TS As Size
  Call RequiredState(csHandleValid Or csFontValid)
  TS.cX = 0:   TS.cY = 0
  Call GetTextExtentPoint32(mHandle, Text, Len(Text), TS)
  cX = TS.cX:  cY = TS.cY
End Sub
Public Function TextWidth(ByVal Text As String) As Long
Dim H As Long
  Call TextExtent(Text, TextWidth, H)
End Function
Public Function TextHeight(ByVal Text As String) As Long
Dim W As Long
  Call TextExtent(Text, W, TextHeight)
End Function

'Public Function GetAllFonts(ByRef cmb As ComboBox)
'  cmb.Clear
'  Call GetAllFont(mHandle, cmb)
'End Function

'Public Sub GetAllFonts(ByVal hDC As Long, cboBox As ComboBox)
'    Set cboFonts = cboBox'
'
'    cboFonts.Clear
'    lf.lfCharSet = DEFAULT_CHARSET
'    EnumFontFamiliesEx hDC, lf, AddressOf EnumFontFamExProc, 0&, 0&
'End Sub

'Private Function EnumFontFamExProc(ByRef lpElfe As ENUMLOGFONTEX, ByVal lpntme As Long, ByVal FontType As Long, ByVal lParam As Long) As Long
'Dim FaceName As String
'  ByteArray2String lpElfe.elfLogFont.lfFaceName, FaceName
'  If Not LastFont = FaceName Then cboFonts.AddItem FaceName
'  mLastFont = FaceName
'
'  EnumFontFamExProc = 1
'End Function

'Private Function ByteArray2String(ByteArray() As Byte, OutputString As String)
'    Dim i As Long
'
'    OutputString = ""
'    For i = 0 To UBound(ByteArray)
'        If ByteArray(i) = 0 Then Exit For
'        OutputString = OutputString & Chr(ByteArray(i))
'    Next i
'End Function

'Es gibt noch keine ClsGraphic
'Public Sub StretchDraw(ByVal Rec As Rect, Graphic As ClsGraphic)
'  If Not Graphic Is Nothing Then
'    Call Changing
'    Call RequiredState(csAllValid)
'    Call Draw(Self, Rect)
'    Call Changed
'  End If
'End Sub

Private Sub Class_Terminate()
  Call DeleteObject(mHandle)
  mHandle = 0
  Set mBrush = Nothing
  Set mFont = Nothing
  Set mPen = Nothing
End Sub

