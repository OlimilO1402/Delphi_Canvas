Attribute VB_Name = "Graphics"
Option Explicit
'in AXdll Instancing: 2 - PublicNotCreatable '6 - GlobalMultiUse
'Public Type TPoint
'  X As Long
'  Y As Long
'End Type
'Public Type POINTAPI
'  X As Long
'  Y As Long
'End Type
Public Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Type TRect
    TopLeft As TPoint
    BottomRight As TPoint
End Type

Public Type Size
  cX As Long
  cY As Long
End Type
Public Type TSize
  cX As Long
  cY As Long
End Type
Public Type POLYTEXT
    X As Long
    Y As Long
    n As Long
    lpStr As String
    uiFlags As Long
    rcl As Rect
    pdx As Long
End Type

Public Type LOGPEN
  lopnStyle As Long
  lopnWidth As TPoint
  lopnColor As Long
End Type

Public Type LOGBRUSH
  lbStyle As Long
  lbColor As Long
  lbHatch As Long
End Type


'für TBrush '#############################################################################
Public Enum HBRUSH
  [_]
End Enum
Public Enum TBrushStyle
  bsSolid = 0       'BS_SOLID
  bsClear = 1       'BS_HOLLOW
  bsHorizontal = 2   'HS_HORIZONTAL
  bsVertical = 3     'HS_VERTICAL
  bsFDiagonal = 4    'HS_FDIAGONAL
  bsBDiagonal = 5    'HS_BDIAGONAL
  bsCross = 6        'HS_CROSS
  bsDiagCross = 7    'HS_DIAGCROSS
End Enum
Public Enum TFillStyle
  fsBorder = 0  'FLOODFILLBORDER
  fsSurface = 1 'FLOODFILLSURFACE
End Enum
Public Enum TFillMode
  fmAlternate = 0
  fmWinding = 1
End Enum
Public Type TBrushData
  Handle As HBRUSH
  Color As TColor
  Bitmap As Long 'TBitmap
  Style As TBrushStyle
End Type
'für TCanvas '#############################################################################
Public Enum HDC
  [_]
End Enum
Public Enum TCanvasStates
  csHandleValid = 1
  csFontValid = 2
  csPenValid = 4
  csBrushValid = 8
  csAllValid = 15 'csHandleValid Or csFontValid Or csPenValid Or csBrushValid
End Enum
Public Enum TCanvasState
  [_]
End Enum
Public Enum TCanvasOrientation
  coLeftToRight
  coRightToLeft
'  coTopToBottom
'  coBottomToTop
End Enum

Public Enum TCopyMode
  cmBlackness = &H42       'BLACKNESS
  cmDstInvert = &H550009   'DSTINVERT
  cmMergeCopy = &HC000CA   'MERGECOPY
  cmMergePaint = &HBB0226  'MERGEPAINT
  cmNotSrcCopy = &H330008  'NOTSRCCOPY
  cmNotSrcErase = &H1100A6 'NOTSRCERASE
  cmPatCopy = &HF00021     'PATCOPY
  cmPatInvert = &H5A0049   'PATINVERT
  cmPatPaint = &HFB0A09    'PATPAINT
  cmSrcAnd = &H8800C6      'SRCAND
  cmSrcCopy = &HCC0020     'SRCCOPY
  cmSrcErase = &H440328    'SRCERASE
  cmSrcInvert = &H660046   'SRCINVERT
  cmSrcPaint = &HEE0086    'SRCPAINT
  cmWhiteness = &HFF0062   'WHITENESS
End Enum

'für TFont '#############################################################################
Public Enum HFONT
  [_]
End Enum
'API-GDI FontStyle
Public Enum TFontStyle 'mehrere sind möglich
  fsBold = 1
  fsItalic = 2
  fsUnderline = 4
  fsStrikeOut = 8
End Enum
'  {$NODEFINE TFontStyle}
'  TFontStyles = set of TFontStyle;
Public Enum TFontStyles
  [_] '[fsBold, fsItalic, fsUnderline, fsStrikeOut]
End Enum
Public Enum TFontStylesBase
  [_] '[fsBold, fsItalic, fsUnderline, fsStrikeOut]
End Enum
'TFontCharset = 0..255;
'TFontDataName = string[LF_FACESIZE - 1];
Public Enum TFontPitch 'nur eines, entweder oder ist möglich
  fpDefault = 0  'DEFAULT_PITCH
  fpFixed = 1    'FIXED_PITCH
  fpVariable = 2 'VARIABLE_PITCH
End Enum
Public Type TFontData
  Handle As HFONT
  Height As Long
  Pitch As TFontPitch
  Style As TFontStylesBase
  Charset As Byte 'TFontCharset
  Name As String 'TFontDataName
End Type

'für TPen  '#############################################################################
Public Enum HPEN
  [_]
End Enum
Public Enum TPenStyle
  psSolid = 0        'PS_SOLID
  psDash = 1         'PS_DASH
  psDot = 2          'PS_DOT
  psDashDot = 3      'PS_DASHDOT
  psDashDotDot = 4   'PS_DASHDOTDOT
  psClear = 5        'PS_NULL
  psInsideFrame = 6  'PS_INSIDEFRAME
End Enum
Public Enum TPenMode
  pmBlack = 1        'R2_BLACK       As Long = 1
  pmNotMerge = 2     'R2_NOTMERGEPEN As Long = 2
  pmMaskNotPen = 3   'R2_MASKNOTPEN  As Long = 3
  pmNotCopy = 4      'R2_NOTCOPYPEN  As Long = 4
  pmMaskPenNot = 5   'R2_MASKPENNOT  As Long = 5
  pmNot = 6          'R2_NOT         As Long = 6
  pmXor = 7          'R2_XORPEN      As Long = 7
  pmNotMask = 8      'R2_NOTMASKPEN  As Long = 8
  pmMask = 9         'R2_MASKPEN     As Long = 9
  pmNotXor = 10      'R2_NOTXORPEN   As Long = 10 'zum drüberzeichnen
  pmNop = 11         'R2_NOP         As Long = 11
  pmMergeNotPen = 12 'R2_MERGENOTPEN As Long = 12
  pmCopy = 13        'R2_COPYPEN     As Long = 13
  pmMergePenNot = 14 'R2_MERGEPENNOT As Long = 14
  pmMerge = 15       'R2_MERGEPEN    As Long = 15
  pmWhite = 16       'R2_WHITE       As Long = 16
End Enum
Public Type TPenData
  Handle As HPEN
  Color As TColor
  Width As Long
  Style As TPenStyle
End Type


'TColor #####################################################################
Public Enum TColor
  '[_] '[-&H7FFFFFF& - 1 To &H80FFFFFF&]
'                                      BBGGRR
  clBlack = &H2000000         'Schwarz
  clMaroon = &H2000080        'Kastanienbraun
  clGreen = &H2008200         'Grün
  clOlive = &H2008284         'Olivgrün
  clNavy = &H2840000          'Marineblau
  clPurple = &H2840084        'Purpur
  clTeal = &H2848200          'Blaugrün
  clGray = &H2848284          'Grau
  clSilver = &H2C6C3C6        'Silber
  clRed = &H20000FF           'Rot
  clLime = &H200FF00         'Gelbgrün
  clYellow = &H200FFFF        'Gelb
      
  clBlue = &H2FF0000          'Blau
  clFuchsia = &H2FF00FF       'Lila
  clAqua = &H2FFFF00          'Hellblau eigentlich türkis
  clWhite = &H2FFFFFF         'Weiß
  clMoneyGreen = &H2C6DFC6    'Minzgrün
  clSkyBlue = &H2F7CBA5       'Himmelblau
  clCream = &H2F7FBFF         'Creme
  clMedGray = &H2A4A0A0       'Mittelgrau
  clDkGray = &H2808080        'Dunkelgrau
  clLtGray = &H2C0C0C0        'Hellgrau
  clNone = &H1FFFFFF          'Weiß in Windows 9x, Schwarz in NT.

  clSystemColor = &H80000000
  clScrollBar = (clSystemColor Or COLOR_SCROLLBAR)            'Aktuelle Farbe der Bildlaufleiste.
  clBackground = (clSystemColor Or COLOR_BACKGROUND)          'Aktuelle Hintergrundfarbe des Windows-Desktops
  clActiveCaption = (clSystemColor Or COLOR_ACTIVECAPTION)       'Aktuelle Farbe der Titelleiste des aktiven Fensters.
  clInactiveCaption = (clSystemColor Or COLOR_INACTIVECAPTION)   'Aktuelle Farbe der Titelleiste der inaktiven Fenster.
  clMenu = (clSystemColor Or COLOR_MENU)                      'Aktuelle Hintergrundfarbe der Menüs.
  clWindow = (clSystemColor Or COLOR_WINDOW)                  'Aktuelle Hintergrundfarbe der Fenster.
  clWindowFrame = (clSystemColor Or COLOR_WINDOWFRAME)        'Aktuelle Farbe des Fensterrahmens.
  clMenuText = (clSystemColor Or COLOR_MENUTEXT)              'Aktuelle Textfarbe der Menüs.
  clWindowText = (clSystemColor Or COLOR_WINDOWTEXT)          'Aktuelle Textfarbe der Fenster.
  clCaptionText = (clSystemColor Or COLOR_CAPTIONTEXT)        'Aktuelle Textfarbe der Titelleiste des aktiven Fensters.
  clActiveBorder = (clSystemColor Or COLOR_ACTIVEBORDER)      'Aktuelle Rahmenfarbe des aktiven Fensters.
  clInactiveBorder = (clSystemColor Or COLOR_INACTIVEBORDER)     'Aktuelle Rahmenfarbe der inaktiven Fenster.
  clAppWorkSpace = (clSystemColor Or COLOR_APPWORKSPACE)      'Aktuelle Farbe des Arbeitsbereichs der Anwendung.
  clHighlight = (clSystemColor Or COLOR_HIGHLIGHT)            'Aktuelle Hintergrundfarbe des markierten Textes.
  clHighlightText = (clSystemColor Or COLOR_HIGHLIGHTTEXT)       'Aktuelle Farbe des markierten Textes.
  clBtnFace = (clSystemColor Or COLOR_BTNFACE)                'Aktuelle Farbe einer Schaltfläche.
  clBtnShadow = (clSystemColor Or COLOR_BTNSHADOW)            'Aktuelle Schattenfarbe einer Schaltfläche.
  clGrayText = (clSystemColor Or COLOR_GRAYTEXT)              'Aktuelle Farbe für abgedunktelten Text.
  clBtnText = (clSystemColor Or COLOR_BTNTEXT)                'Aktuelle Textfarbe der Schaltflächen.
  clInactiveCaptionText = (clSystemColor Or COLOR_INACTIVECAPTIONTEXT)       'Aktuelle Textfarbe der Titelleiste der inaktiven Fenster.
  clBtnHighlight = (clSystemColor Or COLOR_BTNHIGHLIGHT)      'Aktuelle Hervorhebungsfarbe der Schaltflächen.
  cl3DDkShadow = (clSystemColor Or COLOR_3DDKSHADOW)          'Nur Windows 95 und NT 4.0: Dunkler Schatten für dreidimensionale Elemente.
  cl3DLight = (clSystemColor Or COLOR_3DLIGHT)                'Nur Windows 95 und NT 4.0: Helle Farbe für dreidimensionale Elemente (für Kanten, die zur Lichtquelle zeigen).
  clInfoText = (clSystemColor Or COLOR_INFOTEXT)              'Nur Windows 95 und NT 4.0: Textfarbe für Kurzhinweise.
  clInfoBk = (clSystemColor Or COLOR_INFOBK)                  'Nur Windows 95 und NT 4.0: Hintergrundfarbe für Kurzhinweise.
  clGradientActiveCaption = (clSystemColor Or COLOR_GRADIENTACTIVECAPTION)       'Windows 98 oder Windows 2000: Rechte Farbe im Farbverlauf der Titelleiste eines aktiven Fensters. clActiveCaption gibt die Farbe der linken Seite an.
  clGradientInactiveCaption = (clSystemColor Or COLOR_GRADIENTINACTIVECAPTION)    'Windows 98 oder Windows 2000: Rechte Farbe im Farbverlauf der Titelleiste eines inaktiven Fensters. clInactiveCaption gibt die Farbe der linken Seite an.
  clDefault = &H20000000             'Die Standardfarbe des Steuerelements, dem die Farbe zugewiesen wird.
End Enum

'Public StockPen As HPEN
'Public StockBrush As HBRUSH
'Public StockFont As HFONT


