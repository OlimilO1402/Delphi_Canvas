Attribute VB_Name = "ModEnumType"
Option Explicit
'StockObjects
Public StockPen As Long 'HPEN 'Gegenseitige Abhängigkeit von Modulen wenn StockPen as HPEN
Public StockBrush As Long 'HBRUSH
Public StockFont As Long 'HFONT


'  { Color Types }
Public Const CTLCOLOR_MSGBOX    As Long = 0
Public Const CTLCOLOR_EDIT      As Long = 1
Public Const CTLCOLOR_LISTBOX   As Long = 2
Public Const CTLCOLOR_BTN       As Long = 3
Public Const CTLCOLOR_DLG       As Long = 4
Public Const CTLCOLOR_SCROLLBAR As Long = 5
Public Const CTLCOLOR_STATIC    As Long = 6
Public Const CTLCOLOR_MAX       As Long = 7

Public Const COLOR_SCROLLBAR           As Long = 0
Public Const COLOR_BACKGROUND          As Long = 1
Public Const COLOR_ACTIVECAPTION       As Long = 2
Public Const COLOR_INACTIVECAPTION     As Long = 3
Public Const COLOR_MENU                As Long = 4
Public Const COLOR_WINDOW              As Long = 5
Public Const COLOR_WINDOWFRAME         As Long = 6
Public Const COLOR_MENUTEXT            As Long = 7
Public Const COLOR_WINDOWTEXT          As Long = 8
Public Const COLOR_CAPTIONTEXT         As Long = 9
Public Const COLOR_ACTIVEBORDER        As Long = 10
Public Const COLOR_INACTIVEBORDER      As Long = 11
Public Const COLOR_APPWORKSPACE        As Long = 12
Public Const COLOR_HIGHLIGHT           As Long = 13
Public Const COLOR_HIGHLIGHTTEXT       As Long = 14
Public Const COLOR_BTNFACE             As Long = 15
Public Const COLOR_BTNSHADOW           As Long = &H10
Public Const COLOR_GRAYTEXT            As Long = 17
Public Const COLOR_BTNTEXT             As Long = 18
Public Const COLOR_INACTIVECAPTIONTEXT As Long = 19
Public Const COLOR_BTNHIGHLIGHT        As Long = 20
Public Const COLOR_3DDKSHADOW          As Long = 21
Public Const COLOR_3DLIGHT             As Long = 22
Public Const COLOR_INFOTEXT            As Long = 23
Public Const COLOR_INFOBK              As Long = 24

Public Const COLOR_HOTLIGHT            As Long = 26
Public Const COLOR_GRADIENTACTIVECAPTION As Long = 27
Public Const COLOR_GRADIENTINACTIVECAPTION As Long = 28
Public Const COLOR_MENUHILIGHT         As Long = 29
Public Const COLOR_MENUBAR             As Long = 30

Public Const COLOR_ENDCOLORS   As Long = COLOR_MENUBAR
Public Const COLOR_DESKTOP     As Long = COLOR_BACKGROUND
Public Const COLOR_3DFACE      As Long = COLOR_BTNFACE
Public Const COLOR_3DSHADOW    As Long = COLOR_BTNSHADOW
Public Const COLOR_3DHIGHLIGHT As Long = COLOR_BTNHIGHLIGHT
Public Const COLOR_3DHILIGHT   As Long = COLOR_BTNHIGHLIGHT
Public Const COLOR_BTNHILIGHT  As Long = COLOR_BTNHIGHLIGHT

'  { Coordinate Modes }
Public Const ABSOLUTE = 1
Public Const RELATIVE = 2

'  { Stock Logical Objects }
Public Const WHITE_BRUSH         As Long = 0
Public Const LTGRAY_BRUSH        As Long = 1
Public Const GRAY_BRUSH          As Long = 2
Public Const DKGRAY_BRUSH        As Long = 3
Public Const BLACK_BRUSH         As Long = 4
Public Const NULL_BRUSH          As Long = 5
Public Const HOLLOW_BRUSH        As Long = NULL_BRUSH
Public Const WHITE_PEN           As Long = 6
Public Const BLACK_PEN           As Long = 7
Public Const NULL_PEN            As Long = 8
Public Const OEM_FIXED_FONT      As Long = 10
Public Const ANSI_FIXED_FONT     As Long = 11
Public Const ANSI_VAR_FONT       As Long = 12
Public Const SYSTEM_FONT         As Long = 13
Public Const DEVICE_DEFAULT_FONT As Long = 14
Public Const DEFAULT_PALETTE     As Long = 15
Public Const SYSTEM_FIXED_FONT   As Long = &H10
Public Const DEFAULT_GUI_FONT    As Long = 17
Public Const DC_BRUSH            As Long = 18
Public Const DC_PEN              As Long = 19
Public Const STOCK_LAST          As Long = 19

'  {$EXTERNALSYM CLR_INVALID}
Public Const CLR_INVALID  As Long = &HFFFFFFFF

'API-GDI BrushStyle Constanten
Public Const BS_SOLID         As Long = 0
Public Const BS_NULL           As Long = 1
Public Const BS_HOLLOW         As Long = BS_NULL
Public Const BS_HATCHED        As Long = 2
Public Const BS_PATTERN        As Long = 3
Public Const BS_INDEXED        As Long = 4
Public Const BS_DIBPATTERN     As Long = 5
Public Const BS_DIBPATTERNPT   As Long = 6
Public Const BS_PATTERN8X8     As Long = 7
Public Const BS_DIBPATTERN8X8 As Long = 8
Public Const BS_MONOPATTERN    As Long = 9

'API-GDI Hatch Style Constanten
Public Const HS_HORIZONTAL As Long = 0  '{ ----- }
Public Const HS_VERTICAL   As Long = 1  '{ ||||| }
Public Const HS_FDIAGONAL  As Long = 2  '{ ///// }
Public Const HS_BDIAGONAL  As Long = 3  '{ \\\\\ }
Public Const HS_CROSS      As Long = 4  '{ +++++ }
Public Const HS_DIAGCROSS  As Long = 5  '{ xxxxx }

Public Const ETO_GRAYED         As Long = 1
Public Const ETO_OPAQUE         As Long = 2
Public Const ETO_CLIPPED        As Long = 4
Public Const ETO_GLYPH_INDEX    As Long = &H10
Public Const ETO_RTLREADING     As Long = &H80
Public Const ETO_NUMERICSLOCAL  As Long = &H400
Public Const ETO_NUMERICSLATIN  As Long = &H800
Public Const ETO_IGNORELANGUAGE As Long = &H1000
Public Const ETO_PDY            As Long = &H2000

'  {$EXTERNALSYM DEFAULT_PITCH}
Public Const DEFAULT_PITCH      As Long = 0
Public Const FIXED_PITCH        As Long = 1
Public Const VARIABLE_PITCH     As Long = 2
Public Const MONO_FONT          As Long = 8

'API-GDI PenStyle Constanten
Public Const PS_COSMETIC As Long = &H0
Public Const PS_STYLE_MASK As Long = &HF&
Public Const PS_GEOMETRIC As Long = &H10000
Public Const PS_TYPE_MASK As Long = &HF0000

Public Const PS_SOLID       As Long = 0  '{ _______ }
Public Const PS_DASH        As Long = 1  '{ ------- }
Public Const PS_DOT         As Long = 2  '{ ....... }
Public Const PS_DASHDOT     As Long = 3  '{ _._._._ }
Public Const PS_DASHDOTDOT  As Long = 4  '{ _.._.._ }
Public Const PS_NULL        As Long = 5
Public Const PS_INSIDEFRAME As Long = 6
Public Const PS_USERSTYLE   As Long = 7
Public Const PS_ALTERNATE   As Long = 8

Public Const PS_ENDCAP_ROUND As Long = &H0
Public Const PS_ENDCAP_SQUARE As Long = &H100
Public Const PS_ENDCAP_FLAT As Long = &H200
Public Const PS_ENDCAP_MASK As Long = &HF00&
'Public Enum EEndCapStyle
'  esRound = PS_GEOMETRIC Or PS_ENDCAP_ROUND
'  esSquare = PS_GEOMETRIC Or PS_ENDCAP_SQUARE
'  esFlat = PS_GEOMETRIC Or PS_ENDCAP_FLAT
'  esMask = PS_GEOMETRIC Or PS_ENDCAP_MASK
'End Enum

Public Const PS_JOIN_ROUND As Long = &H0
Public Const PS_JOIN_BEVEL As Long = &H1000
Public Const PS_JOIN_MITER As Long = &H2000
Public Const PS_JOIN_MASK As Long = &HF000&
Public Enum EJoinStyle
  [_]
End Enum

'  { Binary raster ops }
Public Const R2_BLACK       As Long = 1     '{  0   }
Public Const R2_NOTMERGEPEN As Long = 2     '{ DPon }
Public Const R2_MASKNOTPEN  As Long = 3     '{ DPna }
Public Const R2_NOTCOPYPEN  As Long = 4     '{ PN   }
Public Const R2_MASKPENNOT  As Long = 5     '{ PDna }
Public Const R2_NOT         As Long = 6     '{ Dn   }
Public Const R2_XORPEN      As Long = 7     '{ DPx  }
Public Const R2_NOTMASKPEN  As Long = 8     '{ DPan }
Public Const R2_MASKPEN     As Long = 9     '{ DPa  }
Public Const R2_NOTXORPEN   As Long = 10    '{ DPxn }
Public Const R2_NOP         As Long = 11    '{ D    }
Public Const R2_MERGENOTPEN As Long = 12    '{ DPno }
Public Const R2_COPYPEN     As Long = 13    '{ P    }
Public Const R2_MERGEPENNOT As Long = 14    '{ PDno }
Public Const R2_MERGEPEN    As Long = 15    '{ DPo  }
Public Const R2_WHITE       As Long = 16 '&H10  '{  1   }
Public Const R2_LAST        As Long = 16 '&H10  '
  
'{ ExtFloodFill style flags }
Public Const FLOODFILLBORDER  As Long = 0
Public Const FLOODFILLSURFACE As Long = 1
'
'sind jetzt im enum TColor Klassenmodul Graphics
'                                      BBGGRR
'Public Const clBlack      As Long = &H2000000  'Schwarz
'Public Const clMaroon     As Long = &H2000080  'Kastanienbraun
'Public Const clGreen      As Long = &H2008200  'Grün
'Public Const clOlive      As Long = &H2008284  'Olivgrün
'Public Const clNavy       As Long = &H2840000  'Marineblau
'Public Const clPurple     As Long = &H2840084  'Purpur
'Public Const clTeal       As Long = &H2848200  'Blaugrün
'Public Const clGray       As Long = &H2848284  'Grau
'Public Const clSilver     As Long = &H2C6C3C6  'Silber
'Public Const clRed        As Long = &H20000FF  'Rot
'Public Const clLime       As Long = &H200FF00  'Gelbgrün
'Public Const clYellow     As Long = &H200FFFF  'Gelb
'
'Public Const clBlue       As Long = &H2FF0000  'Blau
'Public Const clFuchsia    As Long = &H2FF00FF  'Lila
'Public Const clAqua       As Long = &H2FFFF00  'Hellblau eigentlich türkis
'Public Const clWhite      As Long = &H2FFFFFF  'Weiß
'Public Const clMoneyGreen As Long = &H2C6DFC6  'Minzgrün
'Public Const clSkyBlue    As Long = &H2F7CBA5  'Himmelblau
'Public Const clCream      As Long = &H2F7FBFF  'Creme
'Public Const clMedGray    As Long = &H2A4A0A0  'Mittelgrau
'Public Const clDkGray     As Long = &H2808080  'Dunkelgrau
'Public Const clLtGray     As Long = &H2C0C0C0  'Hellgrau
'Public Const clNone       As Long = &H1FFFFFF  'Weiß in Windows 9x, Schwarz in NT.
'
'Public Const clSystemColor     As Long = &H80000000
'Public Const clScrollBar       As Long = (clSystemColor Or COLOR_SCROLLBAR)    'Aktuelle Farbe der Bildlaufleiste.
'Public Const clBackground      As Long = (clSystemColor Or COLOR_BACKGROUND)   'Aktuelle Hintergrundfarbe des Windows-Desktops
'Public Const clActiveCaption   As Long = (clSystemColor Or COLOR_ACTIVECAPTION)   'Aktuelle Farbe der Titelleiste des aktiven Fensters.
'Public Const clInactiveCaption As Long = (clSystemColor Or COLOR_INACTIVECAPTION) 'Aktuelle Farbe der Titelleiste der inaktiven Fenster.
'Public Const clMenu            As Long = (clSystemColor Or COLOR_MENU)         'Aktuelle Hintergrundfarbe der Menüs.
'Public Const clWindow          As Long = (clSystemColor Or COLOR_WINDOW)       'Aktuelle Hintergrundfarbe der Fenster.
'Public Const clWindowFrame     As Long = (clSystemColor Or COLOR_WINDOWFRAME)  'Aktuelle Farbe des Fensterrahmens.
'Public Const clMenuText        As Long = (clSystemColor Or COLOR_MENUTEXT)     'Aktuelle Textfarbe der Menüs.
'Public Const clWindowText      As Long = (clSystemColor Or COLOR_WINDOWTEXT)   'Aktuelle Textfarbe der Fenster.
'Public Const clCaptionText     As Long = (clSystemColor Or COLOR_CAPTIONTEXT)  'Aktuelle Textfarbe der Titelleiste des aktiven Fensters.
'Public Const clActiveBorder    As Long = (clSystemColor Or COLOR_ACTIVEBORDER) 'Aktuelle Rahmenfarbe des aktiven Fensters.
'Public Const clInactiveBorder  As Long = (clSystemColor Or COLOR_INACTIVEBORDER)  'Aktuelle Rahmenfarbe der inaktiven Fenster.
'Public Const clAppWorkSpace    As Long = (clSystemColor Or COLOR_APPWORKSPACE) 'Aktuelle Farbe des Arbeitsbereichs der Anwendung.
'Public Const clHighlight       As Long = (clSystemColor Or COLOR_HIGHLIGHT)    'Aktuelle Hintergrundfarbe des markierten Textes.
'Public Const clHighlightText   As Long = (clSystemColor Or COLOR_HIGHLIGHTTEXT)   'Aktuelle Farbe des markierten Textes.
'Public Const clBtnFace         As Long = (clSystemColor Or COLOR_BTNFACE)      'Aktuelle Farbe einer Schaltfläche.
'Public Const clBtnShadow       As Long = (clSystemColor Or COLOR_BTNSHADOW)    'Aktuelle Schattenfarbe einer Schaltfläche.
'Public Const clGrayText        As Long = (clSystemColor Or COLOR_GRAYTEXT)     'Aktuelle Farbe für abgedunktelten Text.
'Public Const clBtnText         As Long = (clSystemColor Or COLOR_BTNTEXT)      'Aktuelle Textfarbe der Schaltflächen.
'Public Const clInactiveCaptionText     As Long = (clSystemColor Or COLOR_INACTIVECAPTIONTEXT) 'Aktuelle Textfarbe der Titelleiste der inaktiven Fenster.
'Public Const clBtnHighlight    As Long = (clSystemColor Or COLOR_BTNHIGHLIGHT) 'Aktuelle Hervorhebungsfarbe der Schaltflächen.
'Public Const cl3DDkShadow      As Long = (clSystemColor Or COLOR_3DDKSHADOW)   'Nur Windows 95 und NT 4.0: Dunkler Schatten für dreidimensionale Elemente.
'Public Const cl3DLight         As Long = (clSystemColor Or COLOR_3DLIGHT)      'Nur Windows 95 und NT 4.0: Helle Farbe für dreidimensionale Elemente (für Kanten, die zur Lichtquelle zeigen).
'Public Const clInfoText        As Long = (clSystemColor Or COLOR_INFOTEXT)     'Nur Windows 95 und NT 4.0: Textfarbe für Kurzhinweise.
'Public Const clInfoBk          As Long = (clSystemColor Or COLOR_INFOBK)       'Nur Windows 95 und NT 4.0: Hintergrundfarbe für Kurzhinweise.
'Public Const clGradientActiveCaption   As Long = (clSystemColor Or COLOR_GRADIENTACTIVECAPTION)   'Windows 98 oder Windows 2000: Rechte Farbe im Farbverlauf der Titelleiste eines aktiven Fensters. clActiveCaption gibt die Farbe der linken Seite an.
'Public Const clGradientInactiveCaption As Long = (clSystemColor Or COLOR_GRADIENTINACTIVECAPTION)  'Windows 98 oder Windows 2000: Rechte Farbe im Farbverlauf der Titelleiste eines inaktiven Fensters. clInactiveCaption gibt die Farbe der linken Seite an.
'Public Const clDefault   As Long = &H20000000         'Die Standardfarbe des Steuerelements, dem die Farbe zugewiesen wird.

'  { Ternary raster operations }
Public Const SRCCOPY     As Long = &HCC0020  '{ dest = source                    }
Public Const SRCPAINT    As Long = &HEE0086  '{ dest = source OR dest            }
Public Const SRCAND      As Long = &H8800C6  '{ dest = source AND dest           }
Public Const SRCINVERT   As Long = &H660046  '{ dest = source XOR dest           }
Public Const SRCERASE    As Long = &H440328  '{ dest = source AND (NOT dest )    }
Public Const NOTSRCCOPY  As Long = &H330008  '{ dest = (NOT source)              }
Public Const NOTSRCERASE As Long = &H1100A6  '{ dest = (NOT src) AND (NOT dest)  }
Public Const MERGECOPY   As Long = &HC000CA  '{ dest = (source AND pattern)      }
Public Const MERGEPAINT  As Long = &HBB0226  '{ dest = (NOT source) OR dest      }
Public Const PATCOPY     As Long = &HF00021  '{ dest = pattern                   }
Public Const PATPAINT    As Long = &HFB0A09  '{ dest = DPSnoo                    }
Public Const PATINVERT   As Long = &H5A0049  '{ dest = pattern XOR dest          }
Public Const DSTINVERT   As Long = &H550009  '{ dest = (NOT dest)                }
Public Const BLACKNESS   As Long = &H42      '{ dest = BLACK                     }
Public Const WHITENESS   As Long = &HFF0062  '{ dest = WHITE                     }

'  { tmPitchAndFamily flags }
Public Const TMPF_FIXED_PITCH = 1
Public Const TMPF_VECTOR = 2
Public Const TMPF_TRUETYPE = 4
Public Const TMPF_DEVICE = 8
'für den BRush
Public Const TRANSPARENT As Long = 1
Public Const OPAQUE      As Long = 2
'für Font
Public Const OUT_DEFAULT_PRECIS        As Long = 0
Public Const OUT_STRING_PRECIS         As Long = 1
Public Const OUT_CHARACTER_PRECIS      As Long = 2
Public Const OUT_STROKE_PRECIS         As Long = 3
Public Const OUT_TT_PRECIS             As Long = 4
Public Const OUT_DEVICE_PRECIS         As Long = 5
Public Const OUT_RASTER_PRECIS         As Long = 6
Public Const OUT_TT_ONLY_PRECIS        As Long = 7
Public Const OUT_OUTLINE_PRECIS        As Long = 8
Public Const OUT_SCREEN_OUTLINE_PRECIS As Long = 9

Public Const CLIP_DEFAULT_PRECIS As Long = 0
Public Const CLIP_CHARACTER_PRECIS As Long = 1
Public Const CLIP_STROKE_PRECIS As Long = 2
Public Const CLIP_MASK As Long = 15
Public Const CLIP_LH_ANGLES As Long = 16   '(1 shl 4)
Public Const CLIP_TT_ALWAYS As Long = 32   '(2 shl 4)
Public Const CLIP_EMBEDDED As Long = 128 '(8 shl 4)

'Public Type TPoint
'  X As Long
'  Y As Long
'End Type
'Public Type Rect
'    Left As Long
'    Top As Long
'    Right As Long
'    Bottom As Long
'End Type
'
'Public Type Size
'  cX As Long
'  cY As Long
'End Type
'Public Type POLYTEXT
'    X As Long
'    Y As Long
'    n As Long
'    lpStr As String
'    uiFlags As Long
'    rcl As Rect
'    pdx As Long
'End Type
'
'Public Type LOGPEN
'  lopnStyle As Long
'  lopnWidth As TPoint
'  lopnColor As Long
'End Type
'
'Public Type LOGBRUSH
'  lbStyle As Long
'  lbColor As Long
'  lbHatch As Long
'End Type

' für  Fonts:

Public Const LF_FACESIZE As Long = 32
Public Const LF_FULLFACESIZE As Long = 64

Public Type LOGFONT
  lfHeight                     As Long
  lfWidth                      As Long
  lfEscapement                 As Long
  lfOrientation                As Long
  lfWeight                     As Long
  lfItalic                     As Byte
  lfUnderline                  As Byte
  lfStrikeOut                  As Byte
  lfCharSet                    As Byte
  lfOutPrecision               As Byte
  lfClipPrecision              As Byte
  lfQuality                    As Byte
  lfPitchAndFamily             As Byte
  lfFaceName As String * LF_FACESIZE
End Type
'nur für GetAllFonts
Public Type tLOGFONT
  lfHeight                     As Long
  lfWidth                      As Long
  lfEscapement                 As Long
  lfOrientation                As Long
  lfWeight                     As Long
  lfItalic                     As Byte
  lfUnderline                  As Byte
  lfStrikeOut                  As Byte
  lfCharSet                    As Byte
  lfOutPrecision               As Byte
  lfClipPrecision              As Byte
  lfQuality                    As Byte
  lfPitchAndFamily             As Byte
  lfFaceName(LF_FACESIZE)      As Byte  ' As String * LF_FACESIZE
End Type

Public Type ENUMLOGFONTEX
    elfLogFont As tLOGFONT
    elfFullName(LF_FULLFACESIZE) As Byte '  As String * LF_FULLFACESIZE '
    elfStyle(LF_FACESIZE) As Byte  ' As String * LF_FACESIZE '
    elfScript(LF_FACESIZE) As Byte ' As String * LF_FACESIZE '
End Type

Public Const DEFAULT_QUALITY    As Long = 0
Public Const DRAFT_QUALITY      As Long = 1
Public Const PROOF_QUALITY      As Long = 2
Public Const NONANTIALIASED_QUALITY As Long = 3
Public Const ANTIALIASED_QUALITY As Long = 4


Public Const ANSI_CHARSET        As Long = 0
Public Const DEFAULT_CHARSET     As Long = 1
Public Const SYMBOL_CHARSET      As Long = 2
Public Const MAC_CHARSET         As Long = 77
Public Const SHIFTJIS_CHARSET    As Long = &H80 '128
Public Const HANGEUL_CHARSET     As Long = 129
Public Const JOHAB_CHARSET       As Long = 130
Public Const GB2312_CHARSET      As Long = 134
Public Const CHINESEBIG5_CHARSET As Long = 136
Public Const GREEK_CHARSET       As Long = 161
Public Const TURKISH_CHARSET     As Long = 162
Public Const VIETNAMESE_CHARSET  As Long = 163
Public Const HEBREW_CHARSET      As Long = 177
Public Const ARABIC_CHARSET      As Long = 178
Public Const BALTIC_CHARSET      As Long = 186
Public Const RUSSIAN_CHARSET     As Long = 204
Public Const THAI_CHARSET        As Long = 222
Public Const EASTEUROPE_CHARSET  As Long = 238
Public Const OEM_CHARSET         As Long = 255

Public Const FS_LATIN1          As Long = 1
Public Const FS_LATIN2          As Long = 2
Public Const FS_CYRILLIC        As Long = 4
Public Const FS_GREEK           As Long = 8
Public Const FS_TURKISH         As Long = &H10 '16
Public Const FS_HEBREW          As Long = &H20 '32
Public Const FS_ARABIC          As Long = &H40 '64
Public Const FS_BALTIC          As Long = &H80 '128
Public Const FS_VIETNAMESE      As Long = &H100 '256
Public Const FS_THAI            As Long = &H10000 '65536
Public Const FS_JISJAPAN        As Long = &H20000 '131072
Public Const FS_CHINESESIMP     As Long = &H40000 '262144
Public Const FS_WANSUNG         As Long = &H80000 '524288
Public Const FS_CHINESETRAD     As Long = &H100000 '1048576
Public Const FS_JOHAB           As Long = &H200000
Public Const FS_SYMBOL          As Long = &H80000000 'DWORD($80000000)

'  { Font Families }
Public Const FF_DONTCARE   As Long = 0  '(0 shl 4) '  { Don't care or don't know. }
Public Const FF_ROMAN      As Long = 16 '(1 shl 4) '  { Variable stroke width, serifed. }
                                                   '  { Times Roman, Century Schoolbook, etc. }
Public Const FF_SWISS      As Long = 32 '(2 shl 4) '  { Variable stroke width, sans-serifed. }
                                                   '  { Helvetica, Swiss, etc. }
Public Const FF_MODERN     As Long = 48 '(3 shl 4) '  { Constant stroke width, serifed or sans-serifed. }
                                                   '  { Pica, Elite, Courier, etc. }
Public Const FF_SCRIPT     As Long = 64 '(4 shl 4) '  { Cursive, etc. }
Public Const FF_DECORATIVE As Long = 80 '(5 shl 4) '  { Old English, etc. }

'  { Font Weights }
Public Const FW_DONTCARE   As Long = 0
Public Const FW_THIN       As Long = 100
Public Const FW_EXTRALIGHT As Long = 200
Public Const FW_LIGHT      As Long = 300
Public Const FW_NORMAL     As Long = 400
Public Const FW_MEDIUM     As Long = 500
Public Const FW_SEMIBOLD   As Long = 600
Public Const FW_BOLD       As Long = 700
Public Const FW_EXTRABOLD  As Long = 800
Public Const FW_HEAVY      As Long = 900
Public Const FW_ULTRALIGHT As Long = FW_EXTRALIGHT
Public Const FW_REGULAR    As Long = FW_NORMAL
Public Const FW_DEMIBOLD   As Long = FW_SEMIBOLD
Public Const FW_ULTRABOLD  As Long = FW_EXTRABOLD
Public Const FW_BLACK      As Long = FW_HEAVY

Public Const PANOSE_COUNT              As Long = 10
Public Const PAN_FAMILYTYPE_INDEX      As Long = 0
Public Const PAN_SERIFSTYLE_INDEX      As Long = 1
Public Const PAN_WEIGHT_INDEX          As Long = 2
Public Const PAN_PROPORTION_INDEX      As Long = 3
Public Const PAN_CONTRAST_INDEX        As Long = 4
Public Const PAN_STROKEVARIATION_INDEX As Long = 5
Public Const PAN_ARMSTYLE_INDEX        As Long = 6
Public Const PAN_LETTERFORM_INDEX      As Long = 7
Public Const PAN_MIDLINE_INDEX         As Long = 8
Public Const PAN_XHEIGHT_INDEX         As Long = 9

Public Const PAN_CULTURE_LATIN         As Long = 0

'  { Device Parameters for GetDeviceCaps() }
Public Const DRIVERVERSION   As Long = 0 ';     { Device driver version                     }
Public Const TECHNOLOGY      As Long = 2 ';     { Device classification                     }
Public Const HORZSIZE        As Long = 4 ';     { Horizontal size in millimeters            }
Public Const VERTSIZE        As Long = 6 ';     { Vertical size in millimeters              }
Public Const HORZRES         As Long = 8 ';     { Horizontal width in pixels                }
Public Const VERTRES         As Long = 10 ';    { Vertical height in pixels                 }
Public Const BITSPIXEL       As Long = 12 ';    { Number of bits per pixel                  }
Public Const PLANES          As Long = 14 ';    { Number of planes                          }
Public Const NUMBRUSHES      As Long = &H10 ';  { Number of brushes the device has          }
Public Const NUMPENS         As Long = 18 ';    { Number of pens the device has             }
Public Const NUMMARKERS      As Long = 20 ';    { Number of markers the device has          }
Public Const NUMFONTS        As Long = 22 ';    { Number of fonts the device has            }
Public Const NUMCOLORS       As Long = 24 ';    { Number of colors the device supports      }
Public Const PDEVICESIZE     As Long = 26 ';    { Size required for device descriptor       }
Public Const CURVECAPS       As Long = 28 ';    { Curve capabilities                        }
Public Const LINECAPS        As Long = 30 ';    { Line capabilities                         }
Public Const POLYGONALCAPS   As Long = &H20 '32   { Polygonal capabilities                    }
Public Const TEXTCAPS        As Long = 34 ';    { Text capabilities                         }
Public Const CLIPCAPS        As Long = 36 ';    { Clipping capabilities                     }
Public Const RASTERCAPS      As Long = 38 ';    { Bitblt capabilities                       }
Public Const ASPECTX         As Long = 40 ';    { Length of the X leg                       }
Public Const ASPECTY         As Long = 42 ';    { Length of the Y leg                       }
Public Const ASPECTXY        As Long = 44 ';    { Length of the hypotenuse                  }
Public Const SHADEBLENDCAPS  As Long = 45 ';    { Shading and Blending caps                 }
  
Public Const LOGPIXELSX      As Long = 88 ';    { Logical pixelsinch in X =96 on most comp.  }
Public Const LOGPIXELSY      As Long = 90 ';    { Logical pixelsinch in Y =96 on most comp.  }

Public Const SIZEPALETTE     As Long = 104 ';   { Number of entries in physical palette     }
Public Const NUMRESERVED     As Long = 106 ';   { Number of reserved entries in palette     }
Public Const COLORRES        As Long = 108 ';   { Actual color resolution                   }

'  { Printing related DeviceCaps. These replace the appropriate Escapes }
Public Const PHYSICALWIDTH   As Long = 110 ';     { Physical Width in device units            }
Public Const PHYSICALHEIGHT  As Long = 111 ';     { Physical Height in device units           }
Public Const PHYSICALOFFSETX As Long = 112 ';     { Physical Printable Area x margin          }
Public Const PHYSICALOFFSETY As Long = 113 ';     { Physical Printable Area y margin          }
Public Const SCALINGFACTORX  As Long = 114 ';     { Scaling factor x                          }
Public Const SCALINGFACTORY  As Long = 115 ';     { Scaling factor y                          }

'  { Display driver specific}
Public Const VREFRESH       As Long = 116 ';     { Current vertical refresh rate of the display device (for displays only) in Hz}
Public Const DESKTOPVERTRES As Long = 117 ';     { Horizontal width of entire desktop in pixels                                  }
Public Const DESKTOPHORZRES As Long = 118 ';     { Vertical height of entire desktop in pixels                                  }
Public Const BLTALIGNMENT   As Long = 119 ';     { Preferred blt alignment                  }



