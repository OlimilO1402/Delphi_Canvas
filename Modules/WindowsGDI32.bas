Attribute VB_Name = "Windows"
Option Explicit
'UDTypes kann man über zwei Arten an API-Funktionen übergeben
'z.B. ein UD-Type hat den Namen MyUDType
'entweder
'ByRef lpMUDT As MyUDType
'oder über VarPtr(myUDT) und dann die API-Deklaration abändern in
'ByVal lpMUDT As Long
Public Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Public Declare Function FreeResource Lib "kernel32.dll" (ByVal hResData As Long) As Long
Public Declare Function InterlockedIncrement Lib "kernel32.dll" (ByRef lpAddend As Long) As Long
Public Declare Function InterlockedDecrement Lib "kernel32.dll" (ByRef lpAddend As Long) As Long
Public Declare Function GetCurrentThreadId Lib "kernel32.dll" () As Long
Public Declare Function GetCurrentThread Lib "kernel32.dll" () As Long
Public Declare Function AbortDoc Lib "gdi32.dll" (ByVal hhdc As HDC) As Long
Public Declare Function AbortPath Lib "gdi32.dll" (ByVal hhdc As HDC) As Long
Public Declare Function AddFontMemResourceEx Lib "gdi32.dll" (ByRef pvoid As Any, ByVal dword As Long, ByRef DESIGNVECTOR, ByRef pDword As Long) As Long
'Public Declare Function AddFontResource Lib "gdi32.dll" Alias "AddFontResourceA" ()
'Public Declare Function AddFontResourceA Lib "gdi32.dll" ()
'Public Declare Function AddFontResourceW Lib "gdi32.dll" ()
'Public Declare Function AddFontResourceEx Lib "gdi32.dll" Alias "AddFontResourceExA" ()
'Public Declare Function AddFontResourceExA Lib "gdi32.dll" ()
'Public Declare Function AddFontResourceExW Lib "gdi32.dll" ()
'Public Declare Function AlphaBlend Lib "msimg32.dll" ()
'Public Declare Function AlphaDIBBlend Lib "msimg32.dll" ()
Public Declare Function AngleArc Lib "gdi32.dll" (ByVal hhdc As HDC, ByVal X As Long, ByVal Y As Long, ByVal dwRadius As Long, ByVal eStartAngle As Single, ByVal eSweepAngle As Single) As Long
'Public Declare Function AnimatePalette Lib "gdi32.dll" ()
Public Declare Function ArcXY Lib "gdi32.dll" Alias "Arc" (ByVal hhdc As HDC, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
'Public Declare Function ArcTo Lib "gdi32.dll" ()
'Public Declare Function BeginPath Lib "gdi32.dll" ()
'Public Declare Function BitBlt Lib "gdi32.dll" ()
'Public Declare Function CancelDC Lib "gdi32.dll" ()
'Public Declare Function CheckColorsInGamut Lib "gdi32.dll" ()
'Public Declare Function ChoosePixelFormat Lib "gdi32.dll" ()
Public Declare Function ChordXY Lib "gdi32.dll" Alias "Chord" (ByVal hhdc As HDC, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
'Public Declare Function CloseEnhMetaFile Lib "gdi32.dll" ()
'Public Declare Function CloseFigure Lib "gdi32.dll" ()
'Public Declare Function CloseMetaFile Lib "gdi32.dll" ()
'Public Declare Function ColorCorrectPalette Lib "gdi32.dll" ()
'Public Declare Function ColorMatchToTarget Lib "gdi32.dll" ()
'Public Declare Function CombineRgn Lib "gdi32.dll" ()
'Public Declare Function CombineTransform Lib "gdi32.dll" ()
'Public Declare Function CopyEnhMetaFile Lib "gdi32.dll" Alias "CopyEnhMetaFileA" ()
'Public Declare Function CopyEnhMetaFileA Lib "gdi32.dll" ()
'Public Declare Function CopyEnhMetaFileW Lib "gdi32.dll" ()
'Public Declare Function CopyMetaFile Lib "gdi32.dll" Alias "CopyMetaFileA" ()
'Public Declare Function CopyMetaFileA Lib "gdi32.dll" ()
'Public Declare Function CopyMetaFileW Lib "gdi32.dll" ()
'Public Declare Function CreateBitmap Lib "gdi32.dll" ()
'Public Declare Function CreateBitmapIndirect Lib "gdi32.dll" ()
Public Declare Function CreateBrushIndirect Lib "gdi32.dll" (ByRef lpLogBrush As LOGBRUSH) As Long
'Public Declare Function CreateColorSpace Lib "gdi32.dll" Alias "CreateColorSpaceA" ()
'Public Declare Function CreateColorSpaceA Lib "gdi32.dll" ()
'Public Declare Function CreateColorSpaceW Lib "gdi32.dll" ()
'Public Declare Function CreateCompatibleBitmap Lib "gdi32.dll" ()
'Public Declare Function CreateCompatibleDC Lib "gdi32.dll" ()
''''Public Declare Function CreateDC Lib "gdi32.dll" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, ByRef lpInitData As DEVMODE) As Long
'Public Declare Function CreateDCA Lib "gdi32.dll" ()
'Public Declare Function CreateDCW Lib "gdi32.dll" ()
'Public Declare Function CreateDIBPatternBrush Lib "gdi32.dll" ()
'Public Declare Function CreateDIBPatternBrushPt Lib "gdi32.dll" ()
'Public Declare Function CreateDIBSection Lib "gdi32.dll" ()
'Public Declare Function CreateDIBitmap Lib "gdi32.dll" ()
'Public Declare Function CreateDiscardableBitmap Lib "gdi32.dll" ()
'Public Declare Function CreateEllipticRgn Lib "gdi32.dll" ()
'Public Declare Function CreateEllipticRgnIndirect Lib "gdi32.dll" ()
'Public Declare Function CreateEnhMetaFile Lib "gdi32.dll" Alias "CreateEnhMetaFileA" ()
'Public Declare Function CreateEnhMetaFileA Lib "gdi32.dll" ()
'Public Declare Function CreateEnhMetaFileW Lib "gdi32.dll" ()
Public Declare Function CreateFont Lib "gdi32.dll" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal S As Long, ByVal c As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
'Public Declare Function CreateFontA Lib "gdi32.dll" ()
'Public Declare Function CreateFontW Lib "gdi32.dll" ()
Public Declare Function CreateFontIndirect Lib "gdi32.dll" Alias "CreateFontIndirectA" (ByRef lpLogFont As LOGFONT) As Long
'Public Declare Function CreateFontIndirectA Lib "gdi32.dll" ()
'Public Declare Function CreateFontIndirectW Lib "gdi32.dll" ()
''''Public Declare Function CreateFontIndirectEx Lib "gdi32.dll" Alias "CreateFontIndirectExA" (ByRef ENUMLOGFONTEXDVA As ENUMLOGFONTEXDVA) As Long
'Public Declare Function CreateFontIndirectExA Lib "gdi32.dll" ()
'Public Declare Function CreateFontIndirectExW Lib "gdi32.dll" ()
'Public Declare Function CreateHalftonePalette Lib "gdi32.dll" ()
Public Declare Function CreateHatchBrush Lib "gdi32.dll" (ByVal nIndex As Long, ByVal crColor As Long) As Long
'Public Declare Function CreateIC Lib "gdi32.dll" Alias "CreateICA" ()
'Public Declare Function CreateICA Lib "gdi32.dll" ()
'Public Declare Function CreateICW Lib "gdi32.dll" ()
Public Declare Function CreateMetaFile Lib "gdi32.dll" Alias "CreateMetaFileA" (ByVal lpString As String) As Long
'Public Declare Function CreateMetaFileA Lib "gdi32.dll" ()
'Public Declare Function CreateMetaFileW Lib "gdi32.dll" ()
'Public Declare Function CreatePalette Lib "gdi32.dll" ()
'Public Declare Function CreatePatternBrush Lib "gdi32.dll" ()
Public Declare Function CreatePen Lib "gdi32.dll" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function CreatePenIndirect Lib "gdi32.dll" (ByRef lpLogPen As LOGPEN) As Long
'Public Declare Function CreatePolyPolygonRgn Lib "gdi32.dll" ()
'Public Declare Function CreatePolygonRgn Lib "gdi32.dll" ()
'Public Declare Function CreateRectRgn Lib "gdi32.dll" ()
'Public Declare Function CreateRectRgnIndirect Lib "gdi32.dll" ()
'Public Declare Function CreateRoundRectRgn Lib "gdi32.dll" ()
'Public Declare Function CreateScalableFontResource Lib "gdi32.dll" Alias "CreateScalableFontResourceA" ()
'Public Declare Function CreateScalableFontResourceA Lib "gdi32.dll" ()
'Public Declare Function CreateScalableFontResourceW Lib "gdi32.dll" ()
'Public Declare Function CreateSolidBrush Lib "gdi32.dll" ()
Public Declare Function DPtoLP Lib "gdi32.dll" (ByVal HDC As Long, ByRef lpPoint As TPoint, ByVal nCount As Long) As Long
'Public Declare Function DeleteColorSpace Lib "gdi32.dll" ()
'Public Declare Function DeleteDC Lib "gdi32.dll" ()
'Public Declare Function DeleteEnhMetaFile Lib "gdi32.dll" ()
'Public Declare Function DeleteMetaFile Lib "gdi32.dll" ()
Public Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
'Public Declare Function DescribePixelFormat Lib "gdi32.dll" ()
'Public Declare Function DeviceCapabilities Lib "gdi32.dll" Alias "DeviceCapabilitiesA" ()
''Public Declare Function DeviceCapabilitiesA Lib "gdi32.dll" ()
'Public Declare Function DeviceCapabilitiesW Lib "gdi32.dll" ()
'Public Declare Function DeviceCapabilitiesEx Lib "gdi32.dll" alias "DeviceCapabilities" }
'Public Declare Function DrawEscape Lib "gdi32.dll" ()
Public Declare Function DrawFocusRectP Lib "user32.dll" Alias "DrawFocusRect" (ByVal hhdc As HDC, ByVal lpRect As Long) As Long
Public Declare Function EllipseXY Lib "gdi32.dll" Alias "Ellipse" (ByVal hhdc As HDC, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'Public Declare Function EndDoc Lib "gdi32.dll" ()
'Public Declare Function EndPage Lib "gdi32.dll" ()
'Public Declare Function EndPath Lib "gdi32.dll" ()
'Public Declare Function EnumEnhMetaFile Lib "gdi32.dll" ()
'Public Declare Function EnumFontFamilies Lib "gdi32.dll" Alias "EnumFontFamiliesA" ()
'Public Declare Function EnumFontFamiliesA Lib "gdi32.dll" ()
'Public Declare Function EnumFontFamiliesW Lib "gdi32.dll" ()
Public Declare Function EnumFontFamiliesEx Lib "gdi32.dll" Alias "EnumFontFamiliesExA" (ByVal hhdc As HDC, lpLogFont As tLOGFONT, ByVal lpEnumFontProc As Long, ByVal lParam As Long, ByVal dw As Long) As Long
'Public Declare Function EnumFontFamiliesExA Lib "gdi32.dll" ()
'Public Declare Function EnumFontFamiliesExW Lib "gdi32.dll" ()
'Public Declare Function EnumFonts Lib "gdi32.dll" Alias "EnumFontsA" ()
'Public Declare Function EnumFontsA Lib "gdi32.dll" ()
'Public Declare Function EnumFontsW Lib "gdi32.dll" ()
'Public Declare Function EnumICMProfiles Lib "gdi32.dll" Alias "EnumICMProfilesA" ()
'Public Declare Function EnumICMProfilesA Lib "gdi32.dll" ()
'Public Declare Function EnumICMProfilesW Lib "gdi32.dll" ()
'Public Declare Function EnumMetaFile Lib "gdi32.dll" ()
'Public Declare Function EnumObjects Lib "gdi32.dll" ()
'Public Declare Function EqualRgn Lib "gdi32.dll" ()
'Public Declare Function Escape Lib "gdi32.dll" ()
'Public Declare Function ExcludeClipRect Lib "gdi32.dll" ()
Public Declare Function ExtCreatePen Lib "gdi32.dll" (ByVal dwPenStyle As Long, ByVal dwWidth As Long, ByRef lplb As LOGBRUSH, ByVal dwStyleCount As Long, ByRef lpStyle As Long) As Long
'Public Declare Function ExtCreateRegion Lib "gdi32.dll" ()
'Public Declare Function ExtEscape Lib "gdi32.dll" ()
Public Declare Function ExtFloodFill Lib "gdi32.dll" (ByVal hhdc As HDC, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
'Public Declare Function ExtSelectClipRgn Lib "gdi32.dll" ()
Public Declare Function ExtTextOut Lib "gdi32.dll" Alias "ExtTextOutA" (ByVal hhdc As HDC, ByVal X As Long, ByVal Y As Long, ByVal wOptions As Long, ByRef lpRect As Any, ByVal lpString As String, ByVal nCount As Long, ByRef lpDx As Long) As Long
''Public Declare Function ExtTextOutA Lib "gdi32.dll" ()
'Public Declare Function ExtTextOutW Lib "gdi32.dll" ()
'Public Declare Function FillPath Lib "gdi32.dll" ()
Public Declare Function FillRectA Lib "user32.dll" Alias "FillRect" (ByVal hhdc As HDC, ByRef lpRect As Rect, ByVal HBRUSH As Long) As Long
'Public Declare Function FillRgn Lib "gdi32.dll" ()
'Public Declare Function FlattenPath Lib "gdi32.dll" ()
'Public Declare Function FloodFill Lib "gdi32.dll" ()
Public Declare Function FrameRectA Lib "user32.dll" Alias "FrameRect" (ByVal hhdc As HDC, ByRef lpRect As Rect, ByVal HBRUSH As Long) As Long
'Public Declare Function FrameRgn Lib "gdi32.dll" ()
'pPublic Declare Function GdiComment Lib "gdi32.dll" ()
'Public Declare Function GdiFlush Lib "gdi32.dll" ()
'Public Declare Function GdiGetBatchLimit Lib "gdi32.dll" ()
'Public Declare Function GdiSetBatchLimit Lib "gdi32.dll" ()
'Public Declare Function GetArcDirection Lib "gdi32.dll" ()
'Public Declare Function GetAspectRatioFilterEx Lib "gdi32.dll" ()
'Public Declare Function GetBitmapBits Lib "gdi32.dll" ()
'Public Declare Function GetBitmapDimensionEx Lib "gdi32.dll" ()
'Public Declare Function GetBkColor Lib "gdi32.dll" ()
'Public Declare Function GetBkMode Lib "gdi32.dll" ()
'Public Declare Function GetBoundsRect Lib "gdi32.dll" ()
'Public Declare Function GetBrushOrgEx Lib "gdi32.dll" ()
'Public Declare Function GetCharABCWidths Lib "gdi32.dll" Alias "GetCharABCWidthsA" ()
'Public Declare Function GetCharABCWidthsA Lib "gdi32.dll" ()
'Public Declare Function GetCharABCWidthsW Lib "gdi32.dll" ()
'Public Declare Function GetCharABCWidthsI Lib "gdi32.dll" ()
'Public Declare Function GetCharABCWidthsFloat Lib "gdi32.dll" Alias "GetCharABCWidthsFloatA" ()
'Public Declare Function GetCharABCWidthsFloatA Lib "gdi32.dll" ()
'Public Declare Function GetCharABCWidthsFloatW Lib "gdi32.dll" ()
'Public Declare Function GetCharWidth32 Lib "gdi32.dll" Alias "GetCharWidth32A" ()
'Public Declare Function GetCharWidth32A Lib "gdi32.dll" ()
'Public Declare Function GetCharWidth32W Lib "gdi32.dll" ()
'Public Declare Function GetCharWidth Lib "gdi32.dll" Alias "GetCharWidthA" ()
'Public Declare Function GetCharWidthA Lib "gdi32.dll" ()
'Public Declare Function GetCharWidthW Lib "gdi32.dll" ()
'Public Declare Function GetCharWidthFloat Lib "gdi32.dll" Alias "GetCharWidthFloatA" ()
'Public Declare Function GetCharWidthFloatA Lib "gdi32.dll" ()
'Public Declare Function GetCharWidthFloatW Lib "gdi32.dll" ()
'Public Declare Function GetCharWidthI Lib "gdi32.dll" ()
'Public Declare Function GetCharacterPlacement Lib "gdi32.dll" Alias "GetCharacterPlacementA" ()
'Public Declare Function GetCharacterPlacementA Lib "gdi32.dll" ()
'Public Declare Function GetCharacterPlacementW Lib "gdi32.dll" ()
Public Declare Function GetClipBox Lib "gdi32.dll" (ByVal hhdc As HDC, ByRef lpRect As Rect) As Long
'Public Declare Function GetClipRgn Lib "gdi32.dll" ()
'Public Declare Function GetColorAdjustment Lib "gdi32.dll" ()
'Public Declare Function GetColorSpace Lib "gdi32.dll" ()
'Public Declare Function GetCurrentObject Lib "gdi32.dll" ()
Public Declare Function GetCurrentPositionEx Lib "gdi32.dll" (ByVal hhdc As HDC, ByRef lpPoint As TPoint) As Long
Public Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
'Public Declare Function GetDCBrushColor Lib "gdi32.dll" ()
Public Declare Function GetDCPenColor Lib "gdi32.dll" (ByVal hhdc As HDC) As Long
'Public Declare Function GetDCOrgEx Lib "gdi32.dll" ()
'Public Declare Function GetDIBColorTable Lib "gdi32.dll" ()
'Public Declare Function GetDIBits Lib "gdi32.dll" ()
Public Declare Function GetDeviceCaps Lib "gdi32.dll" (ByVal hhdc As HDC, ByVal nIndex As Long) As Long '
'Public Declare Function GetDeviceGammaRamp Lib "gdi32.dll" ()
'Public Declare Function GetEnhMetaFile Lib "gdi32.dll" Alias "GetEnhMetaFileA" ()
''Public Declare Function GetEnhMetaFileA Lib "gdi32.dll" ()
'Public Declare Function GetEnhMetaFileW Lib "gdi32.dll" ()
Public Declare Function GetEnhMetaFileBits Lib "gdi32.dll" (ByVal hemf As Long, ByVal cbBuffer As Long, ByRef lpbBuffer As Byte) As Long
'Public Declare Function GetEnhMetaFileDescription Lib "gdi32.dll" Alias "GetEnhMetaFileDescriptionA" ()
'Public Declare Function GetEnhMetaFileDescriptionA Lib "gdi32.dll" ()
'Public Declare Function GetEnhMetaFileDescriptionW Lib "gdi32.dll" ()
''''Public Declare Function GetEnhMetaFileHeader Lib "gdi32.dll" (ByVal hemf As Long, ByVal cbBuffer As Long, ByRef lpemh As ENHMETAHEADER) As Long
''''Public Declare Function GetEnhMetaFilePaletteEntries Lib "gdi32.dll" (ByVal hemf As Long, ByVal cEntries As Long, ByRef lppe As PALETTEENTRY) As Long
''''Public Declare Function GetEnhMetaFilePixelFormat Lib "gdi32.dll" (ByVal henhmetafile As Long, ByVal uint As Long, ByRef PIXELFORMATDESCRIPTOR As PIXELFORMATDESCRIPTOR) As Long
'Public Declare Function GetFontData Lib "gdi32.dll" ()
'Public Declare Function GetFontLanguageInfo Lib "gdi32.dll" ()
'Public Declare Function GetFontUnicodeRanges Lib "gdi32.dll" ()
'Public Declare Function GetGlyphIndices Lib "gdi32.dll" Alias "GetGlyphIndicesA" ()
'Public Declare Function GetGlyphIndicesA Lib "gdi32.dll" ()
'Public Declare Function GetGlyphIndicesW Lib "gdi32.dll" ()
'Public Declare Function GetGlyphOutline Lib "gdi32.dll" Alias "GetGlyphOutlineA" ()
'Public Declare Function GetGlyphOutlineA Lib "gdi32.dll" ()
'Public Declare Function GetGlyphOutlineW Lib "gdi32.dll" ()
'Public Declare Function GetGraphicsMode Lib "gdi32.dll" ()
'Public Declare Function GetICMProfile Lib "gdi32.dll" Alias "GetICMProfileA" ()
'Public Declare Function GetICMProfileA Lib "gdi32.dll" ()
'Public Declare Function GetICMProfileW Lib "gdi32.dll" ()
'Public Declare Function GetKerningPairs Lib "gdi32.dll" ()
'Public Declare Function GetLogColorSpace Lib "gdi32.dll" Alias "GetLogColorSpaceA" ()
'Public Declare Function GetLogColorSpaceA Lib "gdi32.dll" ()
'Public Declare Function GetLogColorSpaceW Lib "gdi32.dll" ()
Public Declare Function GetMapMode Lib "gdi32.dll" (ByVal hhdc As HDC) As Long
Public Declare Function GetMetaFile Lib "gdi32.dll" Alias "GetMetaFileA" (ByVal lpFileName As String) As Long
'Public Declare Function GetMetaFileA Lib "gdi32.dll" ()
'Public Declare Function GetMetaFileW Lib "gdi32.dll" ()
Public Declare Function GetMetaFileBitsEx Lib "gdi32.dll" (ByVal hMF As Long, ByVal nSize As Long, ByRef lpvData As Any) As Long
Public Declare Function GetMetaRgn Lib "gdi32.dll" (ByVal hhdc As HDC, ByVal hRgn As Long) As Long
'Public Declare Function GetMiterLimit Lib "gdi32.dll" ()
'Public Declare Function GetNearestColor Lib "gdi32.dll" ()
'Public Declare Function GetNearestPaletteIndex Lib "gdi32.dll" ()
'Public Declare Function GetObject Lib "gdi32.dll" Alias "GetObjectA" ()
'Public Declare Function GetObjectA Lib "gdi32.dll" ()
'Public Declare Function GetObjectW Lib "gdi32.dll" ()
'Public Declare Function GetObjectType Lib "gdi32.dll" ()
'Public Declare Function GetOutlineTextMetrics Lib "gdi32.dll" Alias "GetOutlineTextMetricsA" ()
'Public Declare Function GetOutlineTextMetricsA Lib "gdi32.dll" ()
'Public Declare Function GetOutlineTextMetricsW Lib "gdi32.dll" ()
'Public Declare Function GetPaletteEntries Lib "gdi32.dll" ()
'Public Declare Function GetPath Lib "gdi32.dll" ()
Public Declare Function GetPixel Lib "gdi32.dll" (ByVal hhdc As HDC, ByVal X As Long, ByVal Y As Long) As Long
'Public Declare Function GetPixelFormat Lib "gdi32.dll" ()
'Public Declare Function GetPolyFillMode Lib "gdi32.dll" ()
'Public Declare Function GetROP2 Lib "gdi32.dll" ()
'Public Declare Function GetRasterizerCaps Lib "gdi32.dll" ()
'Public Declare Function GetRegionData Lib "gdi32.dll" ()
'Public Declare Function GetRgnBox Lib "gdi32.dll" ()
Public Declare Function GetStockObject Lib "gdi32.dll" (ByVal nIndex As Long) As Long
'Public Declare Function GetStretchBltMode Lib "gdi32.dll" ()
'Public Declare Function GetSystemPaletteEntries Lib "gdi32.dll" ()
'Public Declare Function GetSystemPaletteUse Lib "gdi32.dll" ()
'Public Declare Function GetTextAlign Lib "gdi32.dll" ()
'Public Declare Function GetTextCharacterExtra Lib "gdi32.dll" ()
'Public Declare Function GetTextCharset Lib "gdi32.dll" ()
'Public Declare Function GetTextCharsetInfo Lib "gdi32.dll" ()
'Public Declare Function GetTextColor Lib "gdi32.dll" ()
Public Declare Function GetTextExtentExPoint Lib "gdi32.dll" Alias "GetTextExtentExPointA" (ByVal hhdc As HDC, ByVal lpszStr As String, ByVal cchString As Long, ByVal nMaxExtent As Long, ByRef lpnFit As Long, ByRef alpDx As Long, ByRef lpSize As Size) As Long
'Public Declare Function GetTextExtentExPointA Lib "gdi32.dll" ()
'Public Declare Function GetTextExtentExPointW Lib "gdi32.dll" ()
Public Declare Function GetTextExtentExPointI Lib "gdi32.dll" (ByVal hhdc As HDC, ByRef lpword As Integer, ByValt As Long, ByValt As Long, ByRef lpint As Long, ByRef lpint As Long, ByRef lpSize As Size) As Long
Public Declare Function GetTextExtentPoint32 Lib "gdi32.dll" Alias "GetTextExtentPoint32A" (ByVal hhdc As HDC, ByVal lpsz As String, ByVal cbString As Long, ByRef lpSize As Size) As Long
'Public Declare Function GetTextExtentPoint32A Lib "gdi32.dll" ()
'Public Declare Function GetTextExtentPoint32W Lib "gdi32.dll" ()
Public Declare Function GetTextExtentPoint Lib "gdi32.dll" Alias "GetTextExtentPointA" (ByVal hhdc As HDC, ByVal lpszString As String, ByVal cbString As Long, ByRef lpSize As Size) As Long
'Public Declare Function GetTextExtentPointA Lib "gdi32.dll" ()
'Public Declare Function GetTextExtentPointW Lib "gdi32.dll" ()
Public Declare Function GetTextExtentPointI Lib "gdi32.dll" (ByVal hhdc As HDC, ByRef lpword As Integer, ByValt As Long, ByRef lpSize As Size) As Long
'Public Declare Function GetTextFace Lib "gdi32.dll" Alias "GetTextFaceA" ()
'Public Declare Function GetTextFaceA Lib "gdi32.dll" ()
'Public Declare Function GetTextFaceW Lib "gdi32.dll" ()
'Public Declare Function GetTextMetrics Lib "gdi32.dll" Alias "GetTextMetricsA" ()
'Public Declare Function GetTextMetricsA Lib "gdi32.dll" ()
'Public Declare Function GetTextMetricsW Lib "gdi32.dll" ()
'Public Declare Function GetViewportExtEx Lib "gdi32.dll" ()
'Public Declare Function GetViewportOrgEx Lib "gdi32.dll" ()
'Public Declare Function GetWinMetaFileBits Lib "gdi32.dll" ()
'Public Declare Function GetWindowExtEx Lib "gdi32.dll" ()
Public Declare Function GetWindowOrgEx Lib "gdi32.dll" (ByVal hhdc As HDC, ByRef lpPoint As TPoint) As Long
'Public Declare Function GetWorldTransform Lib "gdi32.dll" ()
'Public Declare Function GradientFill Lib "msimg32.dll" ()
'Public Declare Function IntersectClipRect Lib "gdi32.dll" ()
'Public Declare Function InvertRgn Lib "gdi32.dll" ()
'Public Declare Function LPtoDP Lib "gdi32.dll" (ByVal hhdc As hdc, ByRef lpPoint As POINTAPI, ByVal nCount As Long) As Long
'Public Declare Function LineDDA Lib "gdi32.dll" ()
Public Declare Function LineToXY Lib "gdi32.dll" Alias "LineTo" (ByVal hhdc As HDC, ByVal X As Long, ByVal Y As Long) As Long
'Public Declare Function MaskBlt Lib "gdi32.dll" ()
'Public Declare Function ModifyWorldTransform Lib "gdi32.dll" ()
Public Declare Function MoveToEx Lib "gdi32.dll" (ByVal hhdc As HDC, ByVal X As Long, ByVal Y As Long, ByRef lpPoint As TPoint) As Long
'Public Declare Function OffsetClipRgn Lib "gdi32.dll" ()
'Public Declare Function OffsetRgn Lib "gdi32.dll" ()
Public Declare Function OffsetViewportOrg Lib "gdi32.dll" Alias "OffsetViewportOrgEx" (ByVal hndDC As HDC, ByVal nX As Long, ByVal nY As Long, ByRef lpPoint As TPoint) As Long
Public Declare Function OffsetWindowOrg Lib "gdi32.dll" (ByVal hhdc As HDC, ByVal nX As Long, ByVal nY As Long, ByRef lpPoint As TPoint) As Long
'Public Declare Function PaintRgn Lib "gdi32.dll" ()
'Public Declare Function PatBlt Lib "gdi32.dll" ()
'Public Declare Function PathToRegion Lib "gdi32.dll" ()
Public Declare Function PieXY Lib "gdi32.dll" Alias "Pie" (ByVal hhdc As HDC, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
'Public Declare Function PlayEnhMetaFile Lib "gdi32.dll" ()
'Public Declare Function PlayEnhMetaFileRecord Lib "gdi32.dll" ()
'Public Declare Function PlayMetaFile Lib "gdi32.dll" ()
'Public Declare Function PlayMetaFileRecord Lib "gdi32.dll" ()
'Public Declare Function PlgBlt Lib "gdi32.dll" ()

Public Declare Function PolyBezierPA Lib "gdi32.dll" Alias "PolyBezier" (ByVal hhdc As HDC, ByVal lppt As Long, ByVal nCount As Long) As Long
Public Declare Function PolyBezier Lib "gdi32.dll" (ByVal hhdc As HDC, ByRef lppt As TPoint, ByVal nCount As Long) As Long

Public Declare Function PolyBezierToPA Lib "gdi32.dll" Alias "PolyBezierTo" (ByVal hhdc As HDC, ByVal lppt As Long, ByVal nCount As Long) As Long
Public Declare Function PolyBezierTo Lib "gdi32.dll" (ByVal hhdc As HDC, ByRef lppt As TPoint, ByVal nCount As Long) As Long

Public Declare Function PolyDrawPA Lib "gdi32.dll" Alias "PolyDraw" (ByVal hhdc As HDC, ByVal lppt As Long, ByRef lpbTypes As Byte, ByVal nCount As Long) As Long
Public Declare Function PolyDraw Lib "gdi32.dll" (ByVal hhdc As HDC, ByRef lppt As TPoint, ByRef lpbTypes As Byte, ByVal nCount As Long) As Long

Public Declare Function PolyPolygonPA Lib "gdi32.dll" Alias "PolyPolygon" (ByVal hhdc As HDC, ByVal lppt As Long, ByRef lpPolyCounts As Long, ByVal nCount As Long) As Long
Public Declare Function PolyPolygon Lib "gdi32.dll" (ByVal hhdc As HDC, ByRef lppt As TPoint, ByRef lpPolyCounts As Long, ByVal nCount As Long) As Long

Public Declare Function PolyPolylinePA Lib "gdi32.dll" Alias "PolyPolyline" (ByVal hhdc As HDC, ByVal lppt As Long, ByRef lpdwPolyPoints As Long, ByVal nCount As Long) As Long
Public Declare Function PolyPolyline Lib "gdi32.dll" (ByVal hhdc As HDC, ByRef lppt As TPoint, ByRef lpdwPolyPoints As Long, ByVal nCount As Long) As Long

Public Declare Function PolyTextOut Lib "gdi32.dll" Alias "PolyTextOutA" (ByVal hhdc As HDC, ByRef pptxt As POLYTEXT, ByRef cStrings As Long) As Long

'Public Declare Function PolyTextOutA Lib "gdi32.dll" ()
'Public Declare Function PolyTextOutW Lib "gdi32.dll" ()

Public Declare Function PolygonPA Lib "gdi32.dll" Alias "Polygon" (ByVal hhdc As HDC, ByVal lpPoint As Long, ByVal nCount As Long) As Long
Public Declare Function Polygon Lib "gdi32.dll" (ByVal hhdc As HDC, ByRef lpPoint As TPoint, ByVal nCount As Long) As Long

Public Declare Function PolylinePA Lib "gdi32.dll" Alias "Polyline" (ByVal hhdc As HDC, ByVal lpPoint As Long, ByVal nCount As Long) As Long
Public Declare Function Polyline Lib "gdi32.dll" (ByVal hhdc As HDC, ByRef lpPoint As TPoint, ByVal nCount As Long) As Long

Public Declare Function PolylineToPA Lib "gdi32.dll" Alias "PolylineTo" (ByVal hhdc As HDC, ByVal lpPoint As Long, ByVal nCount As Long) As Long
Public Declare Function PolylineTo Lib "gdi32.dll" (ByVal hhdc As HDC, ByRef lppt As TPoint, ByVal nCount As Long) As Long

Public Declare Function PtInRegion Lib "gdi32.dll" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function PtVisible Lib "gdi32.dll" (ByVal hhdc As HDC, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function RealizePalette Lib "gdi32.dll" (ByVal hhdc As HDC) As Long
Public Declare Function RectInRegion Lib "gdi32.dll" (ByVal hRgn As Long, ByRef lpRect As Rect) As Long
Public Declare Function RectVisible Lib "gdi32.dll" (ByVal hhdc As HDC, ByRef lpRect As Rect) As Long
Public Declare Function RectangleAXY Lib "gdi32.dll" Alias "Rectangle" (ByVal hhdc As HDC, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hhdc As HDC) As Long
'Public Declare Function RemoveFontMemResourceEx Lib "gdi32.dll" ()
'Public Declare Function RemoveFontResource Lib "gdi32.dll" Alias "RemoveFontResourceA" ()
'Public Declare Function RemoveFontResourceA Lib "gdi32.dll" ()
'Public Declare Function RemoveFontResourceW Lib "gdi32.dll" ()
'Public Declare Function RemoveFontResourceEx Lib "gdi32.dll" Alias "RemoveFontResourceExA" ()
'Public Declare Function RemoveFontResourceExA Lib "gdi32.dll" ()
'Public Declare Function RemoveFontResourceExW Lib "gdi32.dll" ()
'Public Declare Function ResetDC Lib "gdi32.dll" Alias "ResetDCA" ()
'Public Declare Function ResetDCA Lib "gdi32.dll" ()
'Public Declare Function ResetDCW Lib "gdi32.dll" ()
'Public Declare Function ResizePalette Lib "gdi32.dll" ()
'Public Declare Function RestoreDC Lib "gdi32.dll" ()
Public Declare Function RoundRectXY Lib "gdi32.dll" Alias "RoundRect" (ByVal hhdc As HDC, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
'Public Declare Function SaveDC Lib "gdi32.dll" ()
'Public Declare Function ScaleViewportExtEx Lib "gdi32.dll" ()
'Public Declare Function ScaleWindowExtEx Lib "gdi32.dll" ()
'Public Declare Function SelectClipPath Lib "gdi32.dll" ()
'Public Declare Function SelectClipRgn Lib "gdi32.dll" ()
Public Declare Function SelectObject Lib "gdi32.dll" (ByVal hhdc As HDC, ByVal hObject As Long) As Long
'Public Declare Function SelectPalette Lib "gdi32.dll" ()
'Public Declare Function SetAbortProc Lib "gdi32.dll" ()
'Public Declare Function SetArcDirection Lib "gdi32.dll" ()
'Public Declare Function SetBitmapBits Lib "gdi32.dll" ()
'Public Declare Function SetBitmapDimensionEx Lib "gdi32.dll" ()
Public Declare Function SetBkColor Lib "gdi32.dll" (ByVal hhdc As HDC, ByVal crColor As Long) As Long
Public Declare Function SetBkMode Lib "gdi32.dll" (ByVal hhdc As HDC, ByVal nBkMode As Long) As Long

'Public Declare Function SetBoundsRect Lib "gdi32.dll" ()
'Public Declare Function SetBrushOrgEx Lib "gdi32.dll" ()
'Public Declare Function SetColorAdjustment Lib "gdi32.dll" ()
'Public Declare Function SetColorSpace Lib "gdi32.dll" ()
'Public Declare Function SetDCBrushColor Lib "gdi32.dll" ()
Public Declare Function SetDCPenColor Lib "gdi32.dll" (ByVal hhdc As HDC, ByVal colorref As Long) As Long
'Public Declare Function SetDIBColorTable Lib "gdi32.dll" ()
'Public Declare Function SetDIBits Lib "gdi32.dll" ()
'Public Declare Function SetDIBitsToDevice Lib "gdi32.dll" ()
'Public Declare Function SetDeviceGammaRamp Lib "gdi32.dll" ()
'Public Declare Function SetEnhMetaFileBits Lib "gdi32.dll" ()
Public Declare Function SetGraphicsMode Lib "gdi32.dll" (ByVal hhdc As HDC, ByVal iMode As Long) As Long
'Public Declare Function SetICMMode Lib "gdi32.dll" ()
'Public Declare Function SetICMProfile Lib "gdi32.dll" Alias "SetICMProfileA" ()
'Public Declare Function SetICMProfileA Lib "gdi32.dll" ()
'Public Declare Function SetICMProfileW Lib "gdi32.dll" ()
Public Declare Function SetMapMode Lib "gdi32.dll" (ByVal hhdc As HDC, ByVal nMapMode As Long) As Long
'Public Declare Function SetMapperFlags Lib "gdi32.dll" ()
'Public Declare Function SetMetaFileBitsEx Lib "gdi32.dll" ()
'Public Declare Function SetMetaRgn Lib "gdi32.dll" ()
'Public Declare Function SetMiterLimit Lib "gdi32.dll" ()
'Public Declare Function SetPaletteEntries Lib "gdi32.dll" ()
Public Declare Function SetPixel Lib "gdi32.dll" (ByVal hhdc As HDC, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
'Public Declare Function SetPixelFormat Lib "gdi32.dll" ()
'Public Declare Function SetPixelV Lib "gdi32.dll" ()
'Public Declare Function SetPolyFillMode Lib "gdi32.dll" ()
Public Declare Function SetROP2 Lib "gdi32.dll" (ByVal hhdc As HDC, ByVal nDrawMode As Long) As Long

'Public Declare Function SetRectRgn Lib "gdi32.dll" ()
'Public Declare Function SetStretchBltMode Lib "gdi32.dll" ()
'Public Declare Function SetSystemPaletteUse Lib "gdi32.dll" ()
'Public Declare Function SetTextAlign Lib "gdi32.dll" ()
Public Declare Function SetTextColor Lib "gdi32.dll" (ByVal hhdc As HDC, ByVal crColor As Long) As Long
'Public Declare Function SetTextCharacterExtra Lib "gdi32.dll" ()
'Public Declare Function SetTextJustification Lib "gdi32.dll" ()
'Public Declare Function SetViewportExtEx Lib "gdi32.dll" ()
'Public Declare Function SetViewportOrgEx Lib "gdi32.dll" ()
'Public Declare Function SetWinMetaFileBits Lib "gdi32.dll" ()
'Public Declare Function SetWindowExtEx Lib "gdi32.dll" ()
'Public Declare Function SetWindowOrgEx Lib "gdi32.dll" ()
'Public Declare Function SetWorldTransform Lib "gdi32.dll" ()
'Public Declare Function StartDoc Lib "gdi32.dll" Alias "StartDocA" ()
'Public Declare Function StartDocA Lib "gdi32.dll" ()
'Public Declare Function StartDocW Lib "gdi32.dll" ()
'Public Declare Function StartPage Lib "gdi32.dll" ()
Public Declare Function StretchBlt Lib "gdi32.dll" (ByVal hhdc As HDC, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
'Public Declare Function StretchDIBits Lib "gdi32.dll" ()
'Public Declare Function StrokeAndFillPath Lib "gdi32.dll" ()
'Public Declare Function StrokePath Lib "gdi32.dll" ()
'Public Declare Function SwapBuffers Lib "gdi32.dll" ()
'Public Declare Function TextOut Lib "gdi32.dll" Alias "TextOutA" ()
'Public Declare Function TextOutA Lib "gdi32.dll" ()
'Public Declare Function TextOutW Lib "gdi32.dll" ()
'Public Declare Function TranslateCharsetInfo Lib "gdi32.dll" ()
'Public Declare Function TransparentBlt Lib "msimg32.dll" ()
'Public Declare Function TransparentDIBits Lib "gdi32.dll" ()
Public Declare Function UnrealizeObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
'Public Declare Function UpdateColors Lib "gdi32.dll" ()
'Public Declare Function UpdateICMRegKey Lib "gdi32.dll" Alias "UpdateICMRegKeyA" ()
'Public Declare Function UpdateICMRegKeyA Lib "gdi32.dll" ()
'Public Declare Function UpdateICMRegKeyW Lib "gdi32.dll" ()
'Public Declare Function WidenPath Lib "gdi32.dll" ()

