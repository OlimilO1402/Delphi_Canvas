Attribute VB_Name = "ModSomeSubNFuncs"
Option Explicit
'Private cboFonts As ComboBox
'Private LastFont As String
'nur für GetAllFonts: wurde schon definiert

Public Colors(0 To 51) As TIdentMapEntry

Public Sub InitColors()
'   = Array( _
'    IME(clBlack, "clBlack"), IME(clMaroon, "clMaroon"), IME(clGreen, "clGreen"), IME(clOlive, "clOlive"), _
'    IME(clNavy, "clNavy"), IME(clPurple, "clPurple"), IME(clTeal, "clTeal"), IME(clGray, "clGray"), _
'    IME(clSilver, "clSilver"), IME(clRed, "clRed"), IME(clLime, "clLime"), IME(clYellow, "clYellow"), _
'    IME(clBlue, "clBlue"), IME(clFuchsia, "clFuchsia"), IME(clAqua, "clAqua"), IME(clWhite, "clWhite"), _
'    IME(clMoneyGreen, "clMoneyGreen"), IME(clSkyBlue, "clSkyBlue"), IME(clCream, "clCream"), IME(clMedGray, "clMedGray"), _
'    IME(clActiveBorder, "clActiveBorder"), IME(clActiveCaption, "clActiveCaption"), IME(clAppWorkSpace, "clAppWorkSpace"), IME(clBackground, "clBackground"), _
'    IME(clBtnFace, "clBtnFace"), IME(clBtnHighlight, "clBtnHighlight"), IME(clBtnShadow, "clBtnShadow"), IME(clBtnText, "clBtnText"), _
'    IME(clCaptionText, "clCaptionText"), IME(clDefault, "clDefault"), IME(clGradientActiveCaption, "clGradientActiveCaption"), IME(clGradientInactiveCaption, "clGradientInactiveCaption"), _
'    IME(clGrayText, "clGrayText"), IME(clHighlight, "clHighlight"), IME(clHighlightText, "clHighlightText"), IME(clHotLight, "clHotLight"), _
'    IME(clInactiveBorder, "clInactiveBorder"), IME(clInactiveCaption, "clInactiveCaption"), IME(clInactiveCaptionText, "clInactiveCaptionText"), IME(clInfoBk, "clInfoBk"), _
'    IME(clInfoText, "clInfoText"), IME(clMenu, "clMenu"), IME(clMenuBar, "clMenuBar"), IME(clMenuHighlight, "clMenuHighlight"), _
'    IME(clMenuText, "clMenuText"), IME(clNone, "clNone"), IME(clScrollBar, "clScrollBar"), IME(cl3DDkShadow, "cl3DDkShadow"), _
'    IME(cl3DLight, "cl3DLight"), IME(clWindow, "clWindow"), IME(clWindowFrame, "clWindowFrame"), IME(clWindowText, "clWindowText"))

  Colors(0) = IME(clBlack, "clBlack")
  Colors(1) = IME(clMaroon, "clMaroon")
  Colors(2) = IME(clGreen, "clGreen")
  Colors(3) = IME(clOlive, "clOlive")
  Colors(4) = IME(clNavy, "clNavy")
  Colors(5) = IME(clPurple, "clPurple")
  Colors(6) = IME(clTeal, "clTeal")
  Colors(7) = IME(clGray, "clGray")
  Colors(8) = IME(clSilver, "clSilver")
  Colors(9) = IME(clRed, "clRed")
  Colors(10) = IME(clLime, "clLime")
  Colors(11) = IME(clYellow, "clYellow")
  Colors(12) = IME(clBlue, "clBlue")
  Colors(13) = IME(clFuchsia, "clFuchsia")
  Colors(14) = IME(clAqua, "clAqua")
  Colors(15) = IME(clWhite, "clWhite")
  Colors(16) = IME(clMoneyGreen, "clMoneyGreen")
  Colors(17) = IME(clSkyBlue, "clSkyBlue")
  Colors(18) = IME(clCream, "clCream")
  Colors(19) = IME(clMedGray, "clMedGray")
  Colors(20) = IME(clActiveBorder, "clActiveBorder")
  Colors(21) = IME(clActiveCaption, "clActiveCaption")
  Colors(22) = IME(clAppWorkSpace, "clAppWorkSpace")
  Colors(23) = IME(clBackground, "clBackground")
  Colors(24) = IME(clBtnFace, "clBtnFace")
  Colors(25) = IME(clBtnHighlight, "clBtnHighlight")
  Colors(26) = IME(clBtnShadow, "clBtnShadow")
  Colors(27) = IME(clBtnText, "clBtnText")
  Colors(28) = IME(clCaptionText, "clCaptionText")
  Colors(29) = IME(clDefault, "clDefault")
  Colors(30) = IME(clGradientActiveCaption, "clGradientActiveCaption")
  Colors(31) = IME(clGradientInactiveCaption, "clGradientInactiveCaption")
  Colors(32) = IME(clGrayText, "clGrayText")
  Colors(33) = IME(clHighlight, "clHighlight")
  Colors(34) = IME(clHighlightText, "clHighlightText")
  Colors(35) = IME(clHotLight, "clHotLight")
  Colors(36) = IME(clInactiveBorder, "clInactiveBorder")
  Colors(37) = IME(clInactiveCaption, "clInactiveCaption")
  Colors(38) = IME(clInactiveCaptionText, "clInactiveCaptionText")
  Colors(39) = IME(clInfoBk, "clInfoBk")
  Colors(40) = IME(clInfoText, "clInfoText")
  Colors(41) = IME(clMenu, "clMenu")
  Colors(42) = IME(clMenuBar, "clMenuBar")
  Colors(43) = IME(clMenuHighlight, "clMenuHighlight")
  Colors(44) = IME(clMenuText, "clMenuText")
  Colors(45) = IME(clNone, "clNone")
  Colors(46) = IME(clScrollBar, "clScrollBar")
  Colors(47) = IME(cl3DDkShadow, "cl3DDkShadow")
  Colors(48) = IME(cl3DLight, "cl3DLight")
  Colors(49) = IME(clWindow, "clWindow")
  Colors(50) = IME(clWindowFrame, "clWindowFrame")
  Colors(51) = IME(clWindowText, "clWindowText")
End Sub

Private Function IME(ColorVal As TColor, StrName As String) As TIdentMapEntry
  With IME
    .Value = ColorVal
    .Name = StrName
  End With
End Function

Public Function ColorToRGB(Color As TColor) As Long
  If Color < 0 Then
    ColorToRGB = GetSysColor(Color And &HFF&)
  Else
    ColorToRGB = Color
  End If
End Function

'procedure GetColorValues(Proc: TGetStrProc);
'Dim I As Long
'  for I = lbound(Colors) to ubound(Colors) do Proc(Colors[I].Name);
'End Sub

Function ColorToIdent(Color As Long, StrIdent As String) As Boolean
  ColorToIdent = IntToIdent(Color, StrIdent, Colors)
End Function

Public Function IdentToColor(StrIdent As String, Color As Long) As Boolean
  IdentToColor = IdentToInt(StrIdent, Color, Colors)
End Function

Private Function IdentToInt(StrIdent As String, iInt As Long, Map) As Boolean
Dim i As Long
  For i = LBound(Map) To UBound(Map)
    If StrComp(Map(i).Name, StrIdent, vbTextCompare) Then
      IdentToInt = True
      iInt = Map(i).Value
      Exit Function
    End If
  Next
  IdentToInt = False
End Function

Private Function IntToIdent(iInt As Long, StrIdent As String, Map) As Boolean
Dim i As Long
  For i = LBound(Map) To UBound(Map)
    If Map(i).Value = iInt Then
      IntToIdent = True
      StrIdent = Map(i).Name
      Exit Function
    End If
  Next
  IntToIdent = False
End Function


'##########################################################################
Public Function LngToBinStr(ByVal LngVal As Long, Optional ByVal MaxPow As Variant, Optional ByVal SepBy As Variant) As String
'wandelt eine Long Integer in einen String der Binärzahlen darstellt um.
Dim i As Long ', x As Long
  If IsMissing(MaxPow) Then MaxPow = 30: MaxPow = Min(MaxPow, 30)
  If IsMissing(SepBy) Then SepBy = 4:    SepBy = Min(SepBy, 30)
  If SepBy = 0 Then SepBy = MaxPow
  For i = MaxPow - 1 To 0 Step -1  ' maxpow -1: weil 2^0 = 1
    LngToBinStr = LngToBinStr + CStr(Abs(CBool(LngVal And Get2Pow(i))))
    If (i Mod SepBy) = 0 Then LngToBinStr = LngToBinStr + " "
  Next
End Function
Public Function Get2Pow(ByVal iPower As Integer) As Long
Dim i As Integer: Get2Pow = 1 'Get2Pow = 2 ^ iPower
  For i = 1 To iPower: Get2Pow = 2 * Get2Pow: Next
End Function
'verwende And um herauszufinden ob ein Bit in einer BinZahl gesetzt ist 0111001010011001
'verwende Or  um ein Bit in einer BinZahl zu setzen 0111001010011001
Public Function Min(ByVal Lng1 As Long, ByVal Lng2 As Long) As Long
  If Lng1 < Lng2 Then Min = Lng1 Else Min = Lng2
End Function
Public Function Max(ByVal Lng1 As Long, ByVal Lng2 As Long) As Long
  If Lng1 > Lng2 Then Max = Lng1 Else Max = Lng2
End Function

Public Function XTwipToX(TwipX As Single) As Long
'Wandelt einen x-Twip in Screenkoordinaten um
  XTwipToX = TwipX \ Screen.TwipsPerPixelX
End Function
Public Function YTwipToY(TwipY As Single) As Long
'Wandelt einen y-Twip in Screenkoordinaten um
  YTwipToY = TwipY \ Screen.TwipsPerPixelY
End Function

Public Function XToTwipX(X As Single) As Long
'Wandelt eine Screenkoordinate in einen x-Twip um
  XToTwipX = X * Screen.TwipsPerPixelX
End Function
Public Function YToTwipY(Y As Single) As Long
'Wandelt eine Screenkoordinate in einen y-Twip um
  YToTwipY = Y * Screen.TwipsPerPixelY
End Function

Public Sub Inc(ByRef LngVal As Long, Optional AddVal = 1)
  'If IsMissing(AddVal) Then AddVal = 1
  LngVal = LngVal + AddVal
End Sub
Public Sub Dec(ByRef LngVal As Long, Optional DifVal = 1)
  'If IsMissing(DifVal) Then DifVal = 1
  LngVal = LngVal - DifVal
End Sub

Public Function GetScreenLogPixels() As Long
Dim DC As HDC
  DC = GetDC(0)
  GetScreenLogPixels = GetDeviceCaps(DC, LOGPIXELSY)
  Call ReleaseDC(0, DC)
End Function

'Public Sub GetAllFont(ByVal hDC As Long, cboBox As ComboBox)
'Dim lf As tLOGFONT
'  Set cboFonts = cboBox
'  cboFonts.Clear
'  lf.lfCharSet = DEFAULT_CHARSET
'  EnumFontFamiliesEx hDC, lf, AddressOf EnumFontFamExProc, 0&, 0&
'End Sub
'
'Public Function EnumFontFamExProc(ByRef lpElfe As ENUMLOGFONTEX, ByVal lpntme As Long, ByVal FontType As Long, ByVal lParam As Long) As Long
'Dim FaceName As String
'  FaceName = ByA2Str(lpElfe.elfLogFont.lfFaceName)
'  If Not LastFont = FaceName Then cboFonts.AddItem FaceName
'  LastFont = FaceName
'  EnumFontFamExProc = 1
'End Function

Public Function ByA2Str(ByteArray() As Byte) As String
Dim i As Long
  For i = LBound(ByteArray) To UBound(ByteArray)
    If ByteArray(i) = 0 Then Exit For
    ByA2Str = ByA2Str & Chr(ByteArray(i))
  Next i
End Function

'Public Function Str2ByA(StrVal As String, ByteArray() As Byte) 'As String
'Dim i As Long
'  ReDim ByteArray(0 To Len(StrVal) + 1)
'  For i = LBound(ByteArray) To UBound(ByteArray) - 1
'    ByteArray(i) = Mid(StrVal, i, 1) '0 Then Exit For
'    'ByA2Str = ByA2Str & Chr(ByteArray(i))
'  Next i
'  ByteArray(i) = 0&
'End Function


