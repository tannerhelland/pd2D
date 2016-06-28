Attribute VB_Name = "GDI_Plus"
'***************************************************************************
'GDI+ Interface
'Copyright 2012-2016 by Tanner Helland
'Created: 1/September/12
'Last updated: 26/June/16
'Last update: add more integer-specific rendering functions
'
'This interface provides a means for interacting with various GDI+ features.  GDI+ was originally used as a fallback for image loading
' and saving if the FreeImage DLL was not found, but over time it has become more and more integrated into PD.  As of version 6.0, GDI+
' is used for a number of specialized tasks, including viewport rendering of 32bpp images, regional blur of selection masks, antialiased
' lines and circles on various dialogs, and more.
'
'Note that - by design - some enums in this class differ subtly from the actual GDI+ enums.  This is a deliberate decision
' to make certain enums play more nicely with other imaging libraries and/or features.  PD handles translation between the
' correct enums as necessary.
'
'These routines are adapted from the work of a number of other talented VB programmers.  Since GDI+ is not well-documented
' for VB users, I first pieced this module together from the following pieces of code:
' Avery P's initial GDI+ deconstruction: http://planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=37541&lngWId=1
' Carles P.V.'s iBMP implementation: http://planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=42376&lngWId=1
' Robert Rayment's PaintRR implementation: http://planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=66991&lngWId=1
' Many thanks to these individuals for their outstanding work on graphics in VB.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'As of 2016, this module is undergoing massive reorganization.  Enums, constants, and functions that have been migrated
' to the new (clean) format are placed in this top section.

Public Enum GP_Result
    GP_OK = 0
    GP_GenericError = 1
    GP_InvalidParameter = 2
    GP_OutOfMemory = 3
    GP_ObjectBusy = 4
    GP_InsufficientBuffer = 5
    GP_NotImplemented = 6
    GP_Win32Error = 7
    GP_WrongState = 8
    GP_Aborted = 9
    GP_FileNotFound = 10
    GP_ValueOverflow = 11
    GP_AccessDenied = 12
    GP_UnknownImageFormat = 13
    GP_FontFamilyNotFound = 14
    GP_FontStyleNotFound = 15
    GP_NotTrueTypeFont = 16
    GP_UnsupportedGDIPlusVersion = 17
    GP_GDIPlusNotInitialized = 18
    GP_PropertyNotFound = 19
    GP_PropertyNotSupported = 20
End Enum

#If False Then
    Private Const GP_OK = 0, GP_GenericError = 1, GP_InvalidParameter = 2, GP_OutOfMemory = 3, GP_ObjectBusy = 4, GP_InsufficientBuffer = 5, GP_NotImplemented = 6, GP_Win32Error = 7, GP_WrongState = 8, GP_Aborted = 9, GP_FileNotFound = 10, GP_ValueOverflow = 11, GP_AccessDenied = 12, GP_UnknownImageFormat = 13
    Private Const GP_FontFamilyNotFound = 14, GP_FontStyleNotFound = 15, GP_NotTrueTypeFont = 16, GP_UnsupportedGDIPlusVersion = 17, GP_GDIPlusNotInitialized = 18, GP_PropertyNotFound = 19, GP_PropertyNotSupported = 20
#End If

Private Type GDIPlusStartupInput
    GDIPlusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Private Enum GP_DebugEventLevel
    GP_DebugEventLevelFatal = 0
    GP_DebugEventLevelWarning = 1
End Enum

#If False Then
    Private Const GP_DebugEventLevelFatal = 0, GP_DebugEventLevelWarning = 1
#End If

'Drawing-related enums

Public Enum GP_QualityMode      'Note that many other settings just wrap these default Quality Mode values
    GP_QM_Invalid = -1
    GP_QM_Default = 0
    GP_QM_Low = 1
    GP_QM_High = 2
End Enum

#If False Then
    Private Const GP_QM_Invalid = -1, GP_QM_Default = 0, GP_QM_Low = 1, GP_QM_High = 2
#End If

Public Enum GP_BrushType        'IMPORTANT NOTE!  This enum is *not* the same as PD's internal 2D brush modes!
    GP_BT_SolidColor = 0
    GP_BT_HatchFill = 1
    GP_BT_TextureFill = 2
    GP_BT_PathGradient = 3
    GP_BT_LinearGradient = 4
End Enum

#If False Then
    Private Const GP_BT_SolidColor = 0, GP_BT_HatchFill = 1, GP_BT_TextureFill = 2, GP_BT_PathGradient = 3, GP_BT_LinearGradient = 4
#End If

'Coloar adjustments are handled internally, at present, so we don't need to expose them to other objects
Private Enum GP_ColorAdjustType
    GP_CAT_Default = 0
    GP_CAT_Bitmap = 1
    GP_CAT_Brush = 2
    GP_CAT_Pen = 3
    GP_CAT_Text = 4
    GP_CAT_Count = 5
    GP_CAT_Any = 6
End Enum

#If False Then
    Private Const GP_CAT_Default = 0, GP_CAT_Bitmap = 1, GP_CAT_Brush = 2, GP_CAT_Pen = 3, GP_CAT_Text = 4, GP_CAT_Count = 5, GP_CAT_Any = 6
#End If

Private Enum GP_ColorMatrixFlags
    GP_CMF_Default = 0
    GP_CMF_SkipGrays = 1
    GP_CMF_AltGray = 2
End Enum

#If False Then
    Private Const GP_CMF_Default = 0, GP_CMF_SkipGrays = 1, GP_CMF_AltGray = 2
#End If

Public Enum GP_CombineMode
    GP_CM_Replace = 0
    GP_CM_Intersect = 1
    GP_CM_Union = 2
    GP_CM_Xor = 3
    GP_CM_Exclude = 4
    GP_CM_Complement = 5
End Enum

#If False Then
    Private Const GP_CM_Replace = 0, GP_CM_Intersect = 1, GP_CM_Union = 2, GP_CM_Xor = 3, GP_CM_Exclude = 4, GP_CM_Complement = 5
#End If

'Compositing mode is the closest GDI+ comes to offering "blend modes".  The default mode alpha-blends the source
' with the destination; "copy" mode overwrites the destination completely.
Public Enum GP_CompositingMode
    GP_CM_SourceOver = 0
    GP_CM_SourceCopy = 1
End Enum

#If False Then
    Private Const GP_CM_SourceOver = 0, GP_CM_SourceCopy = 1
#End If

'Alpha compositing qualities, which affects how GDI+ blends pixels.  Use with caution, as gamma-corrected blending
' yields non-inutitive results!
Public Enum GP_CompositingQuality
    GP_CQ_Invalid = GP_QM_Invalid
    GP_CQ_Default = GP_QM_Default
    GP_CQ_HighSpeed = GP_QM_Low
    GP_CQ_HighQuality = GP_QM_High
    GP_CQ_GammaCorrected = 3&
    GP_CQ_AssumeLinear = 4&
End Enum

#If False Then
    Private Const GP_CQ_Invalid = GP_QM_Invalid, GP_CQ_Default = GP_QM_Default, GP_CQ_HighSpeed = GP_QM_Low, GP_CQ_HighQuality = GP_QM_High, GP_CQ_GammaCorrected = 3&, GP_CQ_AssumeLinear = 4&
#End If

Public Enum GP_DashCap
    GP_DC_Flat = 0
    GP_DC_Square = 0     'This is not a typo; it's supplied as a convenience enum to match supported GP_LineCap values (which differentiate between flat and square, as they should)
    GP_DC_Round = 2
    GP_DC_Triangle = 3
End Enum

#If False Then
    Private Const GP_DC_Flat = 0, GP_DC_Square = 0, GP_DC_Round = 2, GP_DC_Triangle = 3
#End If

Public Enum GP_DashStyle
    GP_DS_Solid = 0&
    GP_DS_Dash = 1&
    GP_DS_Dot = 2&
    GP_DS_DashDot = 3&
    GP_DS_DashDotDot = 4&
    GP_DS_Custom = 5&
End Enum

#If False Then
    Private Const GP_DS_Solid = 0&, GP_DS_Dash = 1&, GP_DS_Dot = 2&, GP_DS_DashDot = 3&, GP_DS_DashDotDot = 4&, GP_DS_Custom = 5&
#End If

Public Enum GP_FillMode
    GP_FM_Alternate = 0&
    GP_FM_Winding = 1&
End Enum

#If False Then
    Private Const GP_FM_Alternate = 0&, GP_FM_Winding = 1&
#End If

Public Enum GP_InterpolationMode
    GP_IM_Invalid = GP_QM_Invalid
    GP_IM_Default = GP_QM_Default
    GP_IM_LowQuality = GP_QM_Low
    GP_IM_HighQuality = GP_QM_High
    GP_IM_Bilinear = 3
    GP_IM_Bicubic = 4
    GP_IM_NearestNeighbor = 5
    GP_IM_HighQualityBilinear = 6
    GP_IM_HighQualityBicubic = 7
End Enum

#If False Then
    Private Const GP_IM_Invalid = GP_QM_Invalid, GP_IM_Default = GP_QM_Default, GP_IM_LowQuality = GP_QM_Low, GP_IM_HighQuality = GP_QM_High, GP_IM_Bilinear = 3, GP_IM_Bicubic = 4, GP_IM_NearestNeighbor = 5, GP_IM_HighQualityBilinear = 6, GP_IM_HighQualityBicubic = 7
#End If

Public Enum GP_LineCap
    GP_LC_Flat = 0&
    GP_LC_Square = 1&
    GP_LC_Round = 2&
    GP_LC_Triangle = 3&
    GP_LC_NoAnchor = &H10
    GP_LC_SquareAnchor = &H11
    GP_LC_RoundAnchor = &H12
    GP_LC_DiamondAnchor = &H13
    GP_LC_ArrowAnchor = &H14
    GP_LC_Custom = &HFF
End Enum

#If False Then
    Private Const GP_LC_Flat = 0, GP_LC_Square = 1, GP_LC_Round = 2, GP_LC_Triangle = 3, GP_LC_NoAnchor = &H10, GP_LC_SquareAnchor = &H11, GP_LC_RoundAnchor = &H12, GP_LC_DiamondAnchor = &H13, GP_LC_ArrowAnchor = &H14, GP_LC_Custom = &HFF
#End If

Public Enum GP_LineJoin
    GP_LJ_Miter = 0&
    GP_LJ_Bevel = 1&
    GP_LJ_Round = 2&
End Enum

#If False Then
    Private Const GP_LJ_Miter = 0&, GP_LJ_Bevel = 1&, GP_LJ_Round = 2&
#End If

Public Enum GP_MatrixOrder
    GP_MO_Prepend = 0&
    GP_MO_Append = 1&
End Enum

#If False Then
    Private Const GP_MO_Prepend = 0&, GP_MO_Append = 1&
#End If

Public Enum GP_PatternStyle
    GP_PS_Horizontal = 0
    GP_PS_Vertical = 1
    GP_PS_ForwardDiagonal = 2
    GP_PS_BackwardDiagonal = 3
    GP_PS_Cross = 4
    GP_PS_DiagonalCross = 5
    GP_PS_05Percent = 6
    GP_PS_10Percent = 7
    GP_PS_20Percent = 8
    GP_PS_25Percent = 9
    GP_PS_30Percent = 10
    GP_PS_40Percent = 11
    GP_PS_50Percent = 12
    GP_PS_60Percent = 13
    GP_PS_70Percent = 14
    GP_PS_75Percent = 15
    GP_PS_80Percent = 16
    GP_PS_90Percent = 17
    GP_PS_LightDownwardDiagonal = 18
    GP_PS_LightUpwardDiagonal = 19
    GP_PS_DarkDownwardDiagonal = 20
    GP_PS_DarkUpwardDiagonal = 21
    GP_PS_WideDownwardDiagonal = 22
    GP_PS_WideUpwardDiagonal = 23
    GP_PS_LightVertical = 24
    GP_PS_LightHorizontal = 25
    GP_PS_NarrowVertical = 26
    GP_PS_NarrowHorizontal = 27
    GP_PS_DarkVertical = 28
    GP_PS_DarkHorizontal = 29
    GP_PS_DashedDownwardDiagonal = 30
    GP_PS_DashedUpwardDiagonal = 31
    GP_PS_DashedHorizontal = 32
    GP_PS_DashedVertical = 33
    GP_PS_SmallConfetti = 34
    GP_PS_LargeConfetti = 35
    GP_PS_ZigZag = 36
    GP_PS_Wave = 37
    GP_PS_DiagonalBrick = 38
    GP_PS_HorizontalBrick = 39
    GP_PS_Weave = 40
    GP_PS_Plaid = 41
    GP_PS_Divot = 42
    GP_PS_DottedGrid = 43
    GP_PS_DottedDiamond = 44
    GP_PS_Shingle = 45
    GP_PS_Trellis = 46
    GP_PS_Sphere = 47
    GP_PS_SmallGrid = 48
    GP_PS_SmallCheckerBoard = 49
    GP_PS_LargeCheckerBoard = 50
    GP_PS_OutlinedDiamond = 51
    GP_PS_SolidDiamond = 52
End Enum

#If False Then
    Private Const GP_PS_Horizontal = 0, GP_PS_Vertical = 1, GP_PS_ForwardDiagonal = 2, GP_PS_BackwardDiagonal = 3, GP_PS_Cross = 4, GP_PS_DiagonalCross = 5, GP_PS_05Percent = 6, GP_PS_10Percent = 7, GP_PS_20Percent = 8, GP_PS_25Percent = 9, GP_PS_30Percent = 10, GP_PS_40Percent = 11, GP_PS_50Percent = 12, GP_PS_60Percent = 13, GP_PS_70Percent = 14, GP_PS_75Percent = 15, GP_PS_80Percent = 16, GP_PS_90Percent = 17, GP_PS_LightDownwardDiagonal = 18, GP_PS_LightUpwardDiagonal = 19, GP_PS_DarkDownwardDiagonal = 20, GP_PS_DarkUpwardDiagonal = 21, GP_PS_WideDownwardDiagonal = 22, GP_PS_WideUpwardDiagonal = 23, GP_PS_LightVertical = 24, GP_PS_LightHorizontal = 25
    Private Const GP_PS_NarrowVertical = 26, GP_PS_NarrowHorizontal = 27, GP_PS_DarkVertical = 28, GP_PS_DarkHorizontal = 29, GP_PS_DashedDownwardDiagonal = 30, GP_PS_DashedUpwardDiagonal = 31, GP_PS_DashedHorizontal = 32, GP_PS_DashedVertical = 33, GP_PS_SmallConfetti = 34, GP_PS_LargeConfetti = 35, GP_PS_ZigZag = 36, GP_PS_Wave = 37, GP_PS_DiagonalBrick = 38, GP_PS_HorizontalBrick = 39, GP_PS_Weave = 40, GP_PS_Plaid = 41, GP_PS_Divot = 42, GP_PS_DottedGrid = 43, GP_PS_DottedDiamond = 44, GP_PS_Shingle = 45, GP_PS_Trellis = 46, GP_PS_Sphere = 47, GP_PS_SmallGrid = 48, GP_PS_SmallCheckerBoard = 49, GP_PS_LargeCheckerBoard = 50
    Private Const GP_PS_OutlinedDiamond = 51, GP_PS_SolidDiamond = 52
#End If

Public Enum GP_PenAlignment
    GP_PA_Center = 0&
    GP_PA_Inset = 1&
End Enum

#If False Then
    Private Const GP_PA_Center = 0&, GP_PA_Inset = 1&
#End If

'PixelOffsetMode controls how GDI+ calculates positioning.  Normally, each a pixel is treated as a unit square that covers
' the area between [0, 0] and [1, 1].  However, for point-based objects like paths, GDI+ can treat coordinates as if they
' are centered over [0.5, 0.5] offsets within each pixel.  This typically yields prettier path renders, at some consequence
' to rendering performance.  (See http://drilian.com/2008/11/25/understanding-half-pixel-and-half-texel-offsets/)
Public Enum GP_PixelOffsetMode
    GP_POM_Invalid = GP_QM_Invalid
    GP_POM_Default = GP_QM_Default
    GP_POM_HighSpeed = GP_QM_Low
    GP_POM_HighQuality = GP_QM_High
    GP_POM_None = 3&
    GP_POM_Half = 4&
End Enum

#If False Then
    Private Const GP_POM_Invalid = QualityModeInvalid, GP_POM_Default = QualityModeDefault, GP_POM_HighSpeed = QualityModeLow, GP_POM_HighQuality = QualityModeHigh, GP_POM_None = 3, GP_POM_Half = 4
#End If

Public Enum GP_SmoothingMode
    GP_SM_Invalid = GP_QM_Invalid
    GP_SM_Default = GP_QM_Default
    GP_SM_HighSpeed = GP_QM_Low
    GP_SM_HighQuality = GP_QM_High
    GP_SM_None = 3&
    GP_SM_AntiAlias = 4&
End Enum

#If False Then
    Private Const GP_SM_Invalid = GP_QM_Invalid, GP_SM_Default = GP_QM_Default, GP_SM_HighSpeed = GP_QM_Low, GP_SM_HighQuality = GP_QM_High, GP_SM_None = 3, GP_SM_AntiAlias = 4
#End If

Public Enum GP_Unit
    GP_U_World = 0&
    GP_U_Display = 1&
    GP_U_Pixel = 2&
    GP_U_Point = 3&
    GP_U_Inch = 4&
    GP_U_Document = 5&
    GP_U_Millimeter = 6&
End Enum

#If False Then
    Private Const GP_U_World = 0, GP_U_Display = 1, GP_U_Pixel = 2, GP_U_Point = 3, GP_U_Inch = 4, GP_U_Document = 5, GP_U_Millimeter = 6
#End If

Public Enum GP_WrapMode
    GP_WM_Tile = 0
    GP_WM_TileFlipX = 1
    GP_WM_TileFlipY = 2
    GP_WM_TileFlipXY = 3
    GP_WM_Clamp = 4
End Enum

#If False Then
    Private Const GP_WM_Tile = 0, GP_WM_TileFlipX = 1, GP_WM_TileFlipY = 2, GP_WM_TileFlipXY = 3, GP_WM_Clamp = 4
#End If

Private Const PixelFormat32bppPARGB = &HE200B

'GDI interop is made easier by declaring a few GDI-specific structs
Private Type BITMAPINFOHEADER
    Size As Long
    Width As Long
    Height As Long
    Planes As Integer
    BitCount As Integer
    Compression As Long
    ImageSize As Long
    xPelsPerMeter As Long
    yPelsPerMeter As Long
    Colorused As Long
    ColorImportant As Long
End Type

Private Type BITMAPINFO
    Header As BITMAPINFOHEADER
    Colors(0 To 255) As RGBQUAD
End Type

'This (stupid) type is used so we can take advantage of LSet when performing some conversions
Private Type tmpLong
    lngResult As Long
End Type

'Core GDI+ functions:
Private Declare Function GdiplusStartup Lib "gdiplus" (ByRef gdipToken As Long, ByRef startupStruct As GDIPlusStartupInput, Optional ByVal OutputBuffer As Long = 0&) As GP_Result
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal gdipToken As Long) As GP_Result

'Object creation/destruction/property functions
Private Declare Function GdipAddPathRectangle Lib "gdiplus" (ByVal hPath As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal rectWidth As Single, ByVal rectHeight As Single) As GP_Result
Private Declare Function GdipAddPathEllipse Lib "gdiplus" (ByVal hPath As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal rectWidth As Single, ByVal rectHeight As Single) As GP_Result
Private Declare Function GdipAddPathLine Lib "gdiplus" (ByVal hPath As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As GP_Result
Private Declare Function GdipAddPathCurve2 Lib "gdiplus" (ByVal hPath As Long, ByVal ptrToFloatArray As Long, ByVal numOfPoints As Long, ByVal curveTension As Single) As GP_Result
Private Declare Function GdipAddPathClosedCurve2 Lib "gdiplus" (ByVal hPath As Long, ByVal ptrToFloatArray As Long, ByVal numOfPoints As Long, ByVal curveTension As Single) As GP_Result
Private Declare Function GdipAddPathBezier Lib "gdiplus" (ByVal hPath As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal x3 As Single, ByVal y3 As Single, ByVal x4 As Single, ByVal y4 As Single) As GP_Result
Private Declare Function GdipAddPathLine2 Lib "gdiplus" (ByVal hPath As Long, ByVal ptrToFloatArray As Long, ByVal numOfPoints As Long) As GP_Result
Private Declare Function GdipAddPathPolygon Lib "gdiplus" (ByVal hPath As Long, ByVal ptrToFloatArray As Long, ByVal numOfPoints As Long) As GP_Result
Private Declare Function GdipAddPathArc Lib "gdiplus" (ByVal hPath As Long, ByVal x As Single, ByVal y As Single, ByVal arcWidth As Single, ByVal arcHeight As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As GP_Result
Private Declare Function GdipAddPathPath Lib "gdiplus" (ByVal hPath As Long, ByVal pathToAdd As Long, ByVal connectToPreviousPoint As Long) As GP_Result

Private Declare Function GdipCloneMatrix Lib "gdiplus" (ByVal srcMatrix As Long, ByRef dstMatrix As Long) As GP_Result
Private Declare Function GdipClonePath Lib "gdiplus" (ByVal srcPath As Long, ByRef dstPath As Long) As GP_Result
Private Declare Function GdipCloneRegion Lib "gdiplus" (ByVal srcRegion As Long, ByRef dstRegion As Long) As GP_Result
Private Declare Function GdipClosePathFigure Lib "gdiplus" (ByVal hPath As Long) As GP_Result

Private Declare Function GdipCombineRegionRect Lib "gdiplus" (ByVal hRegion As Long, ByRef srcRectF As RECTF, ByVal dstCombineMode As GP_CombineMode) As GP_Result
Private Declare Function GdipCombineRegionRectI Lib "gdiplus" (ByVal hRegion As Long, ByRef srcRectL As RECTL, ByVal dstCombineMode As GP_CombineMode) As GP_Result
Private Declare Function GdipCombineRegionRegion Lib "gdiplus" (ByVal dstRegion As Long, ByVal srcRegion As Long, ByVal dstCombineMode As GP_CombineMode) As GP_Result
Private Declare Function GdipCombineRegionPath Lib "gdiplus" (ByVal dstRegion As Long, ByVal srcPath As Long, ByVal dstCombineMode As GP_CombineMode) As GP_Result

Private Declare Function GdipCreateBitmapFromGdiDib Lib "gdiplus" (ByRef origGDIBitmapInfo As BITMAPINFO, ByRef srcBitmapData As Any, ByRef dstGdipBitmap As Long) As GP_Result
Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal bmpWidth As Long, ByVal bmpHeight As Long, ByVal bmpStride As Long, ByVal bmpPixelFormat As Long, ByRef Scan0 As Any, ByRef dstGdipBitmap As Long) As GP_Result
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, ByRef dstGraphics As Long) As GP_Result
Private Declare Function GdipCreateHatchBrush Lib "gdiplus" (ByVal bHatchStyle As GP_PatternStyle, ByVal bForeColor As Long, ByVal bBackColor As Long, ByRef dstBrush As Long) As GP_Result
Private Declare Function GdipCreateImageAttributes Lib "gdiplus" (ByRef dstImageAttributes As Long) As GP_Result
Private Declare Function GdipCreateLineBrush Lib "gdiplus" (ByRef firstPoint As POINTFLOAT, ByRef secondPoint As POINTFLOAT, ByVal firstRGBA As Long, ByVal secondRGBA As Long, ByVal brushWrapMode As GP_WrapMode, ByRef dstBrush As Long) As GP_Result
Private Declare Function GdipCreateLineBrushFromRectWithAngle Lib "gdiplus" (ByRef srcRect As RECTF, ByVal firstRGBA As Long, ByVal secondRGBA As Long, ByVal gradAngle As Single, ByVal isAngleScalable As Long, ByVal gradientWrapMode As GP_WrapMode, ByRef dstLineGradientBrush As Long) As GP_Result
Private Declare Function GdipCreateMatrix Lib "gdiplus" (ByRef dstMatrix As Long) As GP_Result
Private Declare Function GdipCreatePath Lib "gdiplus" (ByVal pathFillMode As GP_FillMode, ByRef dstPath As Long) As GP_Result
Private Declare Function GdipCreatePathGradientFromPath Lib "gdiplus" (ByVal ptrToSrcPath As Long, ByRef dstPathGradientBrush As Long) As GP_Result
Private Declare Function GdipCreatePen1 Lib "gdiplus" (ByVal srcColor As Long, ByVal srcWidth As Single, ByVal srcUnit As GP_Unit, ByRef dstPen As Long) As GP_Result
Private Declare Function GdipCreatePenFromBrush Lib "gdiplus" Alias "GdipCreatePen2" (ByVal srcBrush As Long, ByVal penWidth As Single, ByVal srcUnit As GP_Unit, ByRef dstPen As Long) As GP_Result
Private Declare Function GdipCreateRegion Lib "gdiplus" (ByRef dstRegion As Long) As GP_Result
Private Declare Function GdipCreateRegionPath Lib "gdiplus" (ByVal hPath As Long, ByRef hRegion As Long) As GP_Result
Private Declare Function GdipCreateRegionRect Lib "gdiplus" (ByRef srcRect As RECTF, ByRef hRegion As Long) As GP_Result
Private Declare Function GdipCreateRegionRgnData Lib "gdiplus" (ByVal ptrToRegionData As Long, ByVal sizeOfRegionData As Long, ByRef dstRegion As Long) As GP_Result
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal srcColor As Long, ByRef dstBrush As Long) As GP_Result
Private Declare Function GdipCreateTexture Lib "gdiplus" (ByVal hImage As Long, ByVal textureWrapMode As GP_WrapMode, ByRef dstTexture As Long) As GP_Result

Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal hBrush As Long) As GP_Result
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As GP_Result
Private Declare Function GdipDeleteMatrix Lib "gdiplus" (ByVal hMatrix As Long) As GP_Result
Private Declare Function GdipDeletePath Lib "gdiplus" (ByVal hPath As Long) As GP_Result
Private Declare Function GdipDeletePen Lib "gdiplus" (ByVal hPen As Long) As GP_Result
Private Declare Function GdipDeleteRegion Lib "gdiplus" (ByVal hRegion As Long) As GP_Result
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal hImage As Long) As GP_Result
Private Declare Function GdipDisposeImageAttributes Lib "gdiplus" (ByVal hImageAttributes As Long) As GP_Result

Private Declare Function GdipDrawArc Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal x As Single, ByVal y As Single, ByVal nWidth As Single, ByVal nHeight As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As GP_Result
Private Declare Function GdipDrawArcI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal startAngle As Long, ByVal sweepAngle As Long) As GP_Result
Private Declare Function GdipDrawClosedCurve2 Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal ptrToPointFloats As Long, ByVal numOfPoints As Long, ByVal curveTension As Single) As GP_Result
Private Declare Function GdipDrawClosedCurve2I Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal ptrToPointLongs As Long, ByVal numOfPoints As Long, ByVal curveTension As Single) As GP_Result
Private Declare Function GdipDrawCurve2 Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal ptrToPointFloats As Long, ByVal numOfPoints As Long, ByVal curveTension As Single) As GP_Result
Private Declare Function GdipDrawCurve2I Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal ptrToPointLongs As Long, ByVal numOfPoints As Long, ByVal curveTension As Single) As GP_Result
Private Declare Function GdipDrawEllipse Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal x As Single, ByVal y As Single, ByVal nWidth As Single, ByVal nHeight As Single) As GP_Result
Private Declare Function GdipDrawEllipseI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long) As GP_Result
Private Declare Function GdipDrawImage Lib "gdiplus" (ByVal dstGraphics As Long, ByVal srcImage As Long, ByVal x As Single, ByVal y As Single) As GP_Result
Private Declare Function GdipDrawImageI Lib "gdiplus" (ByVal dstGraphics As Long, ByVal srcImage As Long, ByVal x As Long, ByVal y As Long) As GP_Result
Private Declare Function GdipDrawImageRect Lib "gdiplus" (ByVal dstGraphics As Long, ByVal srcImage As Long, ByVal x As Single, ByVal y As Single, ByVal dstWidth As Single, ByVal dstHeight As Single) As GP_Result
Private Declare Function GdipDrawImageRectI Lib "gdiplus" (ByVal dstGraphics As Long, ByVal srcImage As Long, ByVal x As Long, ByVal y As Long, ByVal dstWidth As Long, ByVal dstHeight As Long) As GP_Result
Private Declare Function GdipDrawImageRectRect Lib "gdiplus" (ByVal dstGraphics As Long, ByVal srcImage As Long, ByVal dstX As Single, ByVal dstY As Single, ByVal dstWidth As Single, ByVal dstHeight As Single, ByVal srcX As Single, ByVal srcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, ByVal srcUnit As GP_Unit, Optional ByVal newImgAttributes As Long = 0, Optional ByVal progCallbackFunction As Long = 0, Optional ByVal progCallbackData As Long = 0) As GP_Result
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal dstGraphics As Long, ByVal srcImage As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal srcUnit As GP_Unit, Optional ByVal newImgAttributes As Long = 0, Optional ByVal progCallbackFunction As Long = 0, Optional ByVal progCallbackData As Long = 0) As GP_Result
Private Declare Function GdipDrawImagePointsRect Lib "gdiplus" (ByVal dstGraphics As Long, ByVal srcImage As Long, ByVal ptrToPointFloats As Long, ByVal dstPtCount As Long, ByVal srcX As Single, ByVal srcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, ByVal srcUnit As GP_Unit, Optional ByVal newImgAttributes As Long = 0, Optional ByVal progCallbackFunction As Long = 0, Optional ByVal progCallbackData As Long = 0) As GP_Result
Private Declare Function GdipDrawImagePointsRectI Lib "gdiplus" (ByVal dstGraphics As Long, ByVal srcImage As Long, ByVal ptrToPointInts As Long, ByVal dstPtCount As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal srcUnit As GP_Unit, Optional ByVal newImgAttributes As Long = 0, Optional ByVal progCallbackFunction As Long = 0, Optional ByVal progCallbackData As Long = 0) As GP_Result
Private Declare Function GdipDrawLine Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As GP_Result
Private Declare Function GdipDrawLineI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As GP_Result
Private Declare Function GdipDrawLines Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal ptrToPointFloats As Long, ByVal numOfPoints As Long) As GP_Result
Private Declare Function GdipDrawLinesI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal ptrToPointLongs As Long, ByVal numOfPoints As Long) As GP_Result
Private Declare Function GdipDrawPath Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal hPath As Long) As GP_Result
Private Declare Function GdipDrawPolygon Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal ptrToPointFloats As Long, ByVal numOfPoints As Long) As GP_Result
Private Declare Function GdipDrawPolygonI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal ptrToPointLongs As Long, ByVal numOfPoints As Long) As GP_Result
Private Declare Function GdipDrawRectangle Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal x As Single, ByVal y As Single, ByVal nWidth As Single, ByVal nHeight As Single) As GP_Result
Private Declare Function GdipDrawRectangleI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long) As GP_Result

Private Declare Function GdipFillClosedCurve2 Lib "gdiplus" (ByVal hGraphics As Long, ByVal hBrush As Long, ByVal ptrToPointFloats As Long, ByVal numOfPoints As Long, ByVal curveTension As Single, ByVal fillMode As GP_FillMode) As GP_Result
Private Declare Function GdipFillClosedCurve2I Lib "gdiplus" (ByVal hGraphics As Long, ByVal hBrush As Long, ByVal ptrToPointLongs As Long, ByVal numOfPoints As Long, ByVal curveTension As Single, ByVal fillMode As GP_FillMode) As GP_Result
Private Declare Function GdipFillEllipse Lib "gdiplus" (ByVal hGraphics As Long, ByVal hBrush As Long, ByVal x As Single, ByVal y As Single, ByVal nWidth As Single, ByVal nHeight As Single) As GP_Result
Private Declare Function GdipFillEllipseI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hBrush As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long) As GP_Result
Private Declare Function GdipFillPath Lib "gdiplus" (ByVal hGraphics As Long, ByVal hBrush As Long, ByVal hPath As Long) As GP_Result
Private Declare Function GdipFillPolygon Lib "gdiplus" (ByVal hGraphics As Long, ByVal hBrush As Long, ByVal ptrToPointFloats As Long, ByVal numOfPoints As Long, ByVal fillMode As GP_FillMode) As GP_Result
Private Declare Function GdipFillPolygonI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hBrush As Long, ByVal ptrToPointLongs As Long, ByVal numOfPoints As Long, ByVal fillMode As GP_FillMode) As GP_Result
Private Declare Function GdipFillRectangle Lib "gdiplus" (ByVal hGraphics As Long, ByVal hBrush As Long, ByVal x As Single, ByVal y As Single, ByVal nWidth As Single, ByVal nHeight As Single) As GP_Result
Private Declare Function GdipFillRectangleI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hBrush As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long) As GP_Result

Private Declare Function GdipGetClip Lib "gdiplus" (ByVal hGraphics As Long, ByRef dstRegion As Long) As GP_Result
Private Declare Function GdipGetCompositingQuality Lib "gdiplus" (ByVal hGraphics As Long, ByRef dstCompositingQuality As GP_CompositingQuality) As GP_Result
Private Declare Function GdipGetInterpolationMode Lib "gdiplus" (ByVal hGraphics As Long, ByRef dstInterpolationMode As GP_InterpolationMode) As GP_Result
Private Declare Function GdipGetPathFillMode Lib "gdiplus" (ByVal hPath As Long, ByRef dstFillRule As GP_FillMode) As GP_Result
Private Declare Function GdipGetPathWorldBounds Lib "gdiplus" (ByVal hPath As Long, ByRef dstBounds As RECTF, ByVal tmpTransformMatrix As Long, ByVal tmpPenHandle As Long) As GP_Result
Private Declare Function GdipGetPathWorldBoundsI Lib "gdiplus" (ByVal hPath As Long, ByRef dstBounds As RECTL, ByVal tmpTransformMatrix As Long, ByVal tmpPenHandle As Long) As GP_Result
Private Declare Function GdipGetPenColor Lib "gdiplus" (ByVal hPen As Long, ByRef dstPARGBColor As Long) As GP_Result
Private Declare Function GdipGetPenDashCap Lib "gdiplus" Alias "GdipGetPenDashCap197819" (ByVal hPen As Long, ByRef dstCap As GP_DashCap) As GP_Result
Private Declare Function GdipGetPenDashStyle Lib "gdiplus" (ByVal hPen As Long, ByRef dstDashStyle As GP_DashStyle) As GP_Result
Private Declare Function GdipGetPenEndCap Lib "gdiplus" (ByVal hPen As Long, ByRef dstLineCap As GP_LineCap) As GP_Result
Private Declare Function GdipGetPenStartCap Lib "gdiplus" (ByVal hPen As Long, ByRef dstLineCap As GP_LineCap) As GP_Result
Private Declare Function GdipGetPenLineJoin Lib "gdiplus" (ByVal hPen As Long, ByRef dstLineJoin As GP_LineJoin) As GP_Result
Private Declare Function GdipGetPenMiterLimit Lib "gdiplus" (ByVal hPen As Long, ByRef dstMiterLimit As Single) As GP_Result
Private Declare Function GdipGetPenMode Lib "gdiplus" (ByVal hPen As Long, ByRef dstPenMode As GP_PenAlignment) As GP_Result
Private Declare Function GdipGetPenWidth Lib "gdiplus" (ByVal hPen As Long, ByRef dstWidth As Single) As GP_Result
Private Declare Function GdipGetPixelOffsetMode Lib "gdiplus" (ByVal hGraphics As Long, ByRef dstMode As GP_PixelOffsetMode) As GP_Result
Private Declare Function GdipGetRegionBounds Lib "gdiplus" (ByVal hRegion As Long, ByVal hGraphics As Long, ByRef dstRectF As RECTF) As GP_Result
Private Declare Function GdipGetRegionBoundsI Lib "gdiplus" (ByVal hRegion As Long, ByVal hGraphics As Long, ByRef dstRectL As RECTL) As GP_Result
Private Declare Function GdipGetRenderingOrigin Lib "gdiplus" (ByVal hGraphics As Long, ByRef dstX As Long, ByRef dstY As Long) As GP_Result
Private Declare Function GdipGetSmoothingMode Lib "gdiplus" (ByVal hGraphics As Long, ByRef dstMode As GP_SmoothingMode) As GP_Result
Private Declare Function GdipGetSolidFillColor Lib "gdiplus" (ByVal hBrush As Long, ByRef dstColor As Long) As GP_Result
Private Declare Function GdipGetTextureWrapMode Lib "gdiplus" (ByVal hBrush As Long, ByRef dstWrapMode As GP_WrapMode) As GP_Result

Private Declare Function GdipInvertMatrix Lib "gdiplus" (ByVal hMatrix As Long) As GP_Result

Private Declare Function GdipIsEmptyRegion Lib "gdiplus" (ByVal srcRegion As Long, ByVal srcGraphics As Long, ByRef dstResult As Long) As GP_Result
Private Declare Function GdipIsEqualRegion Lib "gdiplus" (ByVal srcRegion1 As Long, ByVal srcRegion2 As Long, ByVal srcGraphics As Long, ByRef dstResult As Long) As GP_Result
Private Declare Function GdipIsInfiniteRegion Lib "gdiplus" (ByVal srcRegion As Long, ByVal srcGraphics As Long, ByRef dstResult As Long) As GP_Result
Private Declare Function GdipIsMatrixInvertible Lib "gdiplus" (ByVal hMatrix As Long, ByRef dstResult As Long) As Long
Private Declare Function GdipIsOutlineVisiblePathPoint Lib "gdiplus" (ByVal hPath As Long, ByVal x As Single, ByVal y As Single, ByVal hPen As Long, ByVal hGraphicsOptional As Long, ByRef dstResult As Long) As GP_Result
Private Declare Function GdipIsOutlineVisiblePathPointI Lib "gdiplus" (ByVal hPath As Long, ByVal x As Long, ByVal y As Long, ByVal hPen As Long, ByVal hGraphicsOptional As Long, ByRef dstResult As Long) As GP_Result
Private Declare Function GdipIsVisiblePathPoint Lib "gdiplus" (ByVal hPath As Long, ByVal x As Single, ByVal y As Single, ByVal hGraphicsOptional As Long, ByRef dstResult As Long) As GP_Result
Private Declare Function GdipIsVisiblePathPointI Lib "gdiplus" (ByVal hPath As Long, ByVal x As Long, ByVal y As Long, ByVal hGraphicsOptional As Long, ByRef dstResult As Long) As GP_Result

Private Declare Function GdipResetClip Lib "gdiplus" (ByVal hGraphics As Long) As GP_Result
Private Declare Function GdipResetPath Lib "gdiplus" (ByVal hPath As Long) As GP_Result

Private Declare Function GdipRotateMatrix Lib "gdiplus" (ByVal hMatrix As Long, ByVal rotateAngle As Single, ByVal mOrder As GP_MatrixOrder) As GP_Result
Private Declare Function GdipScaleMatrix Lib "gdiplus" (ByVal hMatrix As Long, ByVal scaleX As Single, ByVal scaleY As Single, ByVal mOrder As GP_MatrixOrder) As GP_Result

Private Declare Function GdipSetClipRect Lib "gdiplus" (ByVal hGraphics As Long, ByVal x As Single, ByVal y As Single, ByVal nWidth As Single, ByVal nHeight As Single, ByVal useCombineMode As GP_CombineMode) As GP_Result
Private Declare Function GdipSetClipRegion Lib "gdiplus" (ByVal hGraphics As Long, ByVal hRegion As Long, ByVal useCombineMode As GP_CombineMode) As GP_Result
Private Declare Function GdipSetCompositingMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal newCompositingMode As GP_CompositingMode) As GP_Result
Private Declare Function GdipSetCompositingQuality Lib "gdiplus" (ByVal hGraphics As Long, ByVal newCompositingQuality As GP_CompositingQuality) As GP_Result
Private Declare Function GdipSetEmpty Lib "gdiplus" (ByVal hRegion As Long) As GP_Result
Private Declare Function GdipSetImageAttributesWrapMode Lib "gdiplus" (ByVal hImageAttributes As Long, ByVal newWrapMode As GP_WrapMode, ByVal argbOfClampMode As Long, ByVal bClampMustBeZero As Long) As GP_Result
Private Declare Function GdipSetImageAttributesColorMatrix Lib "gdiplus" (ByVal hImageAttributes As Long, ByVal typeOfAdjustment As GP_ColorAdjustType, ByVal enableSeparateAdjustmentFlag As Long, ByVal ptrToColorMatrix As Long, ByVal ptrToGrayscaleMatrix As Long, ByVal extraColorMatrixFlags As GP_ColorMatrixFlags) As GP_Result
Private Declare Function GdipSetInfinite Lib "gdiplus" (ByVal hRegion As Long) As GP_Result
Private Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal newInterpolationMode As GP_InterpolationMode) As GP_Result
Private Declare Function GdipSetLinePresetBlend Lib "gdiplus" (ByVal hBrush As Long, ByVal ptrToFirstColor As Long, ByVal ptrToFirstPosition As Long, ByVal numOfPoints As Long) As GP_Result
Private Declare Function GdipSetPathGradientCenterPoint Lib "gdiplus" (ByVal hBrush As Long, ByRef newCenterPoints As POINTFLOAT) As GP_Result
Private Declare Function GdipSetPathGradientPresetBlend Lib "gdiplus" (ByVal hBrush As Long, ByVal ptrToFirstColor As Long, ByVal ptrToFirstPosition As Long, ByVal numOfPoints As Long) As GP_Result
Private Declare Function GdipSetPathGradientWrapMode Lib "gdiplus" (ByVal hBrush As Long, ByVal newWrapMode As GP_WrapMode) As GP_Result
Private Declare Function GdipSetPathFillMode Lib "gdiplus" (ByVal hPath As Long, ByVal pathFillMode As GP_FillMode) As GP_Result
Private Declare Function GdipSetPenColor Lib "gdiplus" (ByVal hPen As Long, ByVal pARGBColor As Long) As GP_Result
Private Declare Function GdipSetPenDashCap Lib "gdiplus" Alias "GdipSetPenDashCap197819" (ByVal hPen As Long, ByVal newCap As GP_DashCap) As GP_Result
Private Declare Function GdipSetPenDashStyle Lib "gdiplus" (ByVal hPen As Long, ByVal newDashStyle As GP_DashStyle) As GP_Result
Private Declare Function GdipSetPenEndCap Lib "gdiplus" (ByVal hPen As Long, ByVal endCap As GP_LineCap) As GP_Result
Private Declare Function GdipSetPenLineCap Lib "gdiplus" Alias "GdipSetPenLineCap197819" (ByVal hPen As Long, ByVal startCap As GP_LineCap, ByVal endCap As GP_LineCap, ByVal dashCap As GP_DashCap) As GP_Result
Private Declare Function GdipSetPenLineJoin Lib "gdiplus" (ByVal hPen As Long, ByVal newLineJoin As GP_LineJoin) As GP_Result
Private Declare Function GdipSetPenMiterLimit Lib "gdiplus" (ByVal hPen As Long, ByVal newMiterLimit As Single) As GP_Result
Private Declare Function GdipSetPenMode Lib "gdiplus" (ByVal hPen As Long, ByVal penMode As GP_PenAlignment) As GP_Result
Private Declare Function GdipSetPenStartCap Lib "gdiplus" (ByVal hPen As Long, ByVal startCap As GP_LineCap) As GP_Result
Private Declare Function GdipSetPenWidth Lib "gdiplus" (ByVal hPen As Long, ByVal penWidth As Single) As GP_Result
Private Declare Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal newMode As GP_PixelOffsetMode) As GP_Result
Private Declare Function GdipSetRenderingOrigin Lib "gdiplus" (ByVal hGraphics As Long, ByVal x As Long, ByVal y As Long) As GP_Result
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal newMode As GP_SmoothingMode) As GP_Result
Private Declare Function GdipSetSolidFillColor Lib "gdiplus" (ByVal hBrush As Long, ByVal newColor As Long) As GP_Result
Private Declare Function GdipSetTextureWrapMode Lib "gdiplus" (ByVal hBrush As Long, ByVal newWrapMode As GP_WrapMode) As GP_Result

Private Declare Function GdipShearMatrix Lib "gdiplus" (ByVal hMatrix As Long, ByVal shearX As Single, ByVal shearY As Single, ByVal mOrder As GP_MatrixOrder) As GP_Result
Private Declare Function GdipStartPathFigure Lib "gdiplus" (ByVal hPath As Long) As GP_Result

Private Declare Function GdipTranslateMatrix Lib "gdiplus" (ByVal hMatrix As Long, ByVal offsetX As Single, ByVal offsetY As Single, ByVal mOrder As GP_MatrixOrder) As GP_Result
Private Declare Function GdipTransformMatrixPoints Lib "gdiplus" (ByVal hMatrix As Long, ByVal ptrToFirstPointF As Long, ByVal numOfPoints As Long) As GP_Result
Private Declare Function GdipTransformPath Lib "gdiplus" (ByVal hPath As Long, ByVal hMatrix As Long) As GP_Result

Private Declare Function GdipWidenPath Lib "gdiplus" (ByVal hPath As Long, ByVal hPen As Long, ByVal hTransformMatrix As Long, ByVal allowableError As Single) As GP_Result
Private Declare Function GdipWindingModeOutline Lib "gdiplus" (ByVal hPath As Long, ByVal hTransformationMatrix As Long, ByVal allowableError As Single) As GP_Result

'Non-GDI+ helper functions:
Private Declare Function CopyMemory_Strict Lib "kernel32" Alias "RtlMoveMemory" (ByVal ptrDst As Long, ByVal ptrSrc As Long, ByVal numOfBytes As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32" (ByVal oColor As OLE_COLOR, ByVal hPalette As Long, ByRef cColorRef As Long) As Long

'Internally cached values:

'Startup values
Private m_GDIPlusToken As Long, m_GDIPlus11Available As Boolean

'Some GDI+ functions require world transformation data.  This dummy graphics container is used to host any such transformations.
' It is created when GDI+ is initialized, and destroyed when GDI+ is released.  To be a good citizen, please undo any world transforms
' before a function releases.  This ensures that subsequent functions are not messed up.
Private m_TransformDIB As pd2DDIB, m_TransformGraphics As Long

'To modify opacity in GDI+, an image attributes matrix is used.  Rather than recreating one every time an alpha operation is required,
' we simply create a default identity matrix at initialization, then re-use it as necessary.
Private m_AttributesMatrix() As Single

'At start-up, this function is called to determine whether or not we have GDI+ available on this machine.
Public Function GDIP_StartEngine(Optional ByVal hookDebugProc As Boolean = False) As Boolean
    
    'Prep a generic GDI+ startup interface
    Dim gdiCheck As GDIPlusStartupInput
    With gdiCheck
        .GDIPlusVersion = 1&
        
        'Hypothetically you could set a callback function here, but I haven't tested this thoroughly, so use with caution!
        'If hookDebugProc Then
        '    .DebugEventCallback = FakeProcPtr(AddressOf GDIP_Debug_Proc)
        'Else
            .DebugEventCallback = 0&
        'End If
        
        .SuppressBackgroundThread = 0&
        .SuppressExternalCodecs = 0&
    End With
    
    'Retrieve a GDI+ token for this session
    GDIP_StartEngine = CBool(GdiplusStartup(m_GDIPlusToken, gdiCheck, 0&) = GP_OK)
    If GDIP_StartEngine Then
        
        'As a convenience, create a dummy graphics container.  This is useful for various GDI+ functions that require world
        ' transformation data.
        Set m_TransformDIB = New pd2DDIB
        m_TransformDIB.CreateBlank 8, 8, 32, 0, 0
        GdipCreateFromHDC m_TransformDIB.GetDIBDC, m_TransformGraphics
        
        'Note that these dummy objects are released when GDI+ terminates.
        
        'Next, create a default identity matrix for image attributes.
        ReDim m_AttributesMatrix(0 To 4, 0 To 4) As Single
        m_AttributesMatrix(0, 0) = 1#
        m_AttributesMatrix(1, 1) = 1#
        m_AttributesMatrix(2, 2) = 1#
        m_AttributesMatrix(3, 3) = 1#
        m_AttributesMatrix(4, 4) = 1#
        
        'Next, check to see if v1.1 is available.  This allows for advanced fx work.
        Dim hMod As Long, strGDIPName As String
        strGDIPName = "gdiplus.dll"
        hMod = LoadLibrary(StrPtr(strGDIPName))
        If (hMod <> 0) Then
            Dim testAddress As Long
            testAddress = GetProcAddress(hMod, "GdipDrawImageFX")
            m_GDIPlus11Available = CBool(testAddress <> 0)
            FreeLibrary hMod
        Else
            m_GDIPlus11Available = False
        End If
        
    Else
        m_GDIPlus11Available = False
    End If

End Function

'At shutdown, this function must be called to release our GDI+ instance
Public Function GDIP_StopEngine() As Boolean

    'Release any dummy containers we have created
    GdipDeleteGraphics m_TransformGraphics
    Set m_TransformDIB = Nothing
    
    'Release GDI+ using the same token we received at startup time
    GDIP_StopEngine = CBool(GdiplusShutdown(m_GDIPlusToken) = GP_OK)
    
End Function

'Want to know if GDI+ v1.1 is available?  Use this wrapper.
Public Function IsGDIPlusV11Available() As Boolean
    IsGDIPlusV11Available = m_GDIPlus11Available
End Function

Private Function FakeProcPtr(ByVal AddressOfResult As Long) As Long
    FakeProcPtr = AddressOfResult
End Function

'At GDI+ startup, the caller can request that we provide a debug proc for GDI+ to call on warnings and errors.
' This is that proc.
'
'NOTE: this feature is currently disabled due to lack of testing.
Private Function GDIP_Debug_Proc(ByVal deLevel As GP_DebugEventLevel, ByVal ptrChar As Long) As Long
    
    'Pull the GDI+ message into a local string
    'Dim cUnicode As pdUnicode
    'Set cUnicode = New pdUnicode
    
    Dim debugString As String
    'debugString = cUnicode.ConvertCharPointerToVBString(ptrChar, False)
    debugString = "Unknown GDI+ error was passed to the GDIPlus debug procedure."
    
    If (deLevel = GP_DebugEventLevelWarning) Then
        Debug.Print "GDI+ WARNING: " & debugString
    ElseIf (deLevel = GP_DebugEventLevelFatal) Then
        Debug.Print "GDI+ ERROR: " & debugString
    Else
        Debug.Print "GDI+ UNKNOWN: " & debugString
    End If
    
End Function

Private Function InternalGDIPlusError(Optional ByVal errName As String = vbNullString, Optional ByVal errDescription As String = vbNullString, Optional ByVal errNumber As GP_Result = GP_OK)
        
    'If the caller passes an error number but no error name, attempt to automatically populate
    ' it based on the error number.
    If ((Len(errName) = 0) And (errNumber <> GP_OK)) Then
        
        Select Case errNumber
            Case GP_GenericError
                errName = "Generic Error"
            Case GP_InvalidParameter
                errName = "Invalid parameter"
            Case GP_OutOfMemory
                errName = "Out of memory"
            Case GP_ObjectBusy
                errName = "Object busy"
            Case GP_InsufficientBuffer
                errName = "Insufficient buffer size"
            Case GP_NotImplemented
                errName = "Feature is not implemented"
            Case GP_Win32Error
                errName = "Win32 error"
            Case GP_WrongState
                errName = "Wrong state"
            Case GP_Aborted
                errName = "Operation aborted"
            Case GP_FileNotFound
                errName = "File not found"
            Case GP_ValueOverflow
                errName = "Value too large (overflow)"
            Case GP_AccessDenied
                errName = "Access denied"
            Case GP_UnknownImageFormat
                errName = "Image format was not recognized"
            Case GP_FontFamilyNotFound
                errName = "Font family not found"
            Case GP_FontStyleNotFound
                errName = "Font style not found"
            Case GP_NotTrueTypeFont
                errName = "Font is not TrueType (only TT fonts are supported)"
            Case GP_UnsupportedGDIPlusVersion
                errName = "GDI+ version is not supported"
            Case GP_GDIPlusNotInitialized
                errName = "GDI+ was not initialized correctly"
            Case GP_PropertyNotFound
                errName = "Property missing"
            Case GP_PropertyNotSupported
                errName = "Property not supported"
            Case Else
                errName = "Undefined error (number doesn't match known returns)"
        End Select
        
    End If
    
    Dim tmpString As String
    tmpString = "WARNING!  Internal GDI+ error #" & errNumber & ", """ & errName & """"
    If (Len(errDescription) <> 0) Then tmpString = tmpString & ": " & errDescription
    Debug.Print tmpString
    
End Function

'GDI+ requires RGBQUAD colors with alpha in the 4th byte.  This function returns an RGBQUAD (long-type) from a standard RGB()
' long and supplied alpha.  It's not a very efficient conversion, but I need it so infrequently that I don't really care.
Public Function FillQuadWithVBRGB(ByVal vbRGB As Long, ByVal alphaValue As Byte) As Long
    
    'The vbRGB constant may be an OLE color constant; if that happens, we want to convert it to a normal RGB quad.
    vbRGB = TranslateColor(vbRGB)
    
    Dim dstQuad As RGBQUAD
    dstQuad.Red = Drawing2D.ExtractRed(vbRGB)
    dstQuad.Green = Drawing2D.ExtractGreen(vbRGB)
    dstQuad.Blue = Drawing2D.ExtractBlue(vbRGB)
    dstQuad.Alpha = alphaValue
    
    Dim placeHolder As tmpLong
    LSet placeHolder = dstQuad
    
    FillQuadWithVBRGB = placeHolder.lngResult
    
End Function

'Given a long-type pARGB value returned from GDI+, retrieve just the opacity value on the scale [0, 100]
Public Function GetOpacityFromPARGB(ByVal pARGB As Long) As Single
    Dim srcQuad As RGBQUAD
    CopyMemory_Strict VarPtr(srcQuad), VarPtr(pARGB), 4&
    GetOpacityFromPARGB = CSng(srcQuad.Alpha) * CSng(100# / 255#)
End Function

'Given a long-type pARGB value returned from GDI+, retrieve just the RGB component in combined vbRGB format
Public Function GetColorFromPARGB(ByVal pARGB As Long) As Long
    
    Dim srcQuad As RGBQUAD
    CopyMemory_Strict VarPtr(srcQuad), VarPtr(pARGB), 4&
    
    If (srcQuad.Alpha = 255) Then
        GetColorFromPARGB = RGB(srcQuad.Red, srcQuad.Green, srcQuad.Blue)
    Else
    
        Dim tmpSingle As Single
        tmpSingle = CSng(srcQuad.Alpha) / 255
        
        If (tmpSingle <> 0) Then
            Dim tmpRed As Long, tmpGreen As Long, tmpBlue As Long
            tmpRed = CSng(srcQuad.Red) / tmpSingle
            tmpGreen = CSng(srcQuad.Green) / tmpSingle
            tmpBlue = CSng(srcQuad.Blue) / tmpSingle
            GetColorFromPARGB = RGB(tmpRed, tmpGreen, tmpBlue)
        Else
            GetColorFromPARGB = 0
        End If
        
    End If
    
End Function

'Translate an OLE color to an RGB Long.  Note that the API function returns -1 on failure; if this happens, we return white.
Private Function TranslateColor(ByVal colorRef As Long) As Long
    If OleTranslateColor(colorRef, 0, TranslateColor) Then TranslateColor = vbWhite
End Function

Public Function GetGDIPlusSolidBrushHandle(ByVal brushColor As Long, Optional ByVal brushOpacity As Byte = 255) As Long
    GdipCreateSolidFill FillQuadWithVBRGB(brushColor, brushOpacity), GetGDIPlusSolidBrushHandle
End Function

Public Function GetGDIPlusPatternBrushHandle(ByVal brushPattern As GP_PatternStyle, ByVal bFirstColor As Long, ByVal bFirstColorOpacity As Byte, ByVal bSecondColor As Long, ByVal bSecondColorOpacity As Byte) As Long
    GdipCreateHatchBrush brushPattern, FillQuadWithVBRGB(bFirstColor, bFirstColorOpacity), FillQuadWithVBRGB(bSecondColor, bSecondColorOpacity), GetGDIPlusPatternBrushHandle
End Function

Public Function GetGDIPlusLinearBrushHandle(ByRef srcRect As RECTF, ByVal firstRGBA As Long, ByVal secondRGBA As Long, ByVal gradAngle As Single, ByVal isAngleScalable As Long, ByVal gradientWrapMode As PD_2D_WrapMode) As Long
    GdipCreateLineBrushFromRectWithAngle srcRect, firstRGBA, secondRGBA, gradAngle, isAngleScalable, gradientWrapMode, GetGDIPlusLinearBrushHandle
End Function

Public Function OverrideGDIPlusLinearGradient(ByVal hBrush As Long, ByVal ptrToFirstColor As Long, ByVal ptrToFirstPosition As Long, ByVal numOfPoints As Long) As Boolean
    OverrideGDIPlusLinearGradient = CBool(GdipSetLinePresetBlend(hBrush, ptrToFirstColor, ptrToFirstPosition, numOfPoints) = GP_OK)
End Function

Public Function GetGDIPlusPathBrushHandle(ByVal hGraphicsPath As Long) As Long
    GdipCreatePathGradientFromPath hGraphicsPath, GetGDIPlusPathBrushHandle
End Function

Public Function SetGDIPlusPathBrushCenter(ByVal hBrush As Long, ByVal centerX As Single, ByVal centerY As Single) As Long
    Dim centerPoint As POINTFLOAT
    centerPoint.x = centerX
    centerPoint.y = centerY
    GdipSetPathGradientCenterPoint hBrush, centerPoint
End Function

Public Function SetGDIPlusPathBrushWrap(ByVal hBrush As Long, ByVal newWrapMode As GP_WrapMode) As Boolean
    SetGDIPlusPathBrushWrap = CBool(GdipSetPathGradientWrapMode(hBrush, newWrapMode) = GP_OK)
End Function

Public Function OverrideGDIPlusPathGradient(ByVal hBrush As Long, ByVal ptrToFirstColor As Long, ByVal ptrToFirstPosition As Long, ByVal numOfPoints As Long) As Boolean
    OverrideGDIPlusPathGradient = CBool(GdipSetPathGradientPresetBlend(hBrush, ptrToFirstColor, ptrToFirstPosition, numOfPoints) = GP_OK)
End Function

'Simpler shorthand function for obtaining a GDI+ bitmap handle from a pd2DDIB object.  Note that 24/32bpp cases have to be
' handled separately because GDI+ is unpredictable at automatically detecting color depth with 32-bpp DIBs.  (This behavior
' is forgivable, given GDI's unreliable handling of alpha bytes.)
Public Function GetGdipBitmapHandleFromDIB(ByRef dstBitmapHandle As Long, ByRef srcDIB As pd2DDIB) As Boolean
    
    If (srcDIB Is Nothing) Then Exit Function
    
    If (srcDIB.GetDIBColorDepth = 32) Then
        GetGdipBitmapHandleFromDIB = CBool(GdipCreateBitmapFromScan0(srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, srcDIB.GetDIBWidth * 4, PixelFormat32bppPARGB, ByVal srcDIB.GetDIBPointer, dstBitmapHandle) = GP_OK)
    Else
    
        'Use GdipCreateBitmapFromGdiDib for 24bpp DIBs
        Dim imgHeader As BITMAPINFO
        With imgHeader.Header
            .Size = Len(imgHeader.Header)
            .Planes = 1
            .BitCount = srcDIB.GetDIBColorDepth
            .Width = srcDIB.GetDIBWidth
            .Height = -srcDIB.GetDIBHeight
        End With
        GetGdipBitmapHandleFromDIB = CBool(GdipCreateBitmapFromGdiDib(imgHeader, ByVal srcDIB.GetDIBPointer, dstBitmapHandle) = GP_OK)
        
    End If

End Function

'Retrieving a bitmap from a DC is a messy and performance-intensive process.  Avoid it if at all possible.
Public Function GetGdipBitmapHandleFromDC(ByVal srcDC As Long) As Long

End Function

'Because of the way GDI+ texture brushes work, it is significantly easier to initialize one from a full DIB object
' (which *always* guarantees bitmap bits will be available) vs a GDI+ Graphics object, which is more like a DC in
' that it could be a non-bitmap, or dimensionless, or other weird criteria.
Public Function GetGDIPlusTextureBrush(ByRef srcDIB As pd2DDIB, Optional ByVal brushWrapMode As GP_WrapMode = GP_WM_Tile) As Long
    Dim srcBitmap As Long, tmpReturn As GP_Result
    GetGdipBitmapHandleFromDIB srcBitmap, srcDIB
    tmpReturn = GdipCreateTexture(srcBitmap, brushWrapMode, GetGDIPlusTextureBrush)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError , , tmpReturn
    tmpReturn = GdipDisposeImage(srcBitmap)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError , , tmpReturn
End Function

'Retrieve a persistent handle to a GDI+-format graphics container.  Optionally, a smoothing mode can be specified so that it does
' not have to be repeatedly specified by a caller function.  (GDI+ sets smoothing mode by graphics container, not by function call.)
Public Function GetGDIPlusGraphicsFromDC(ByVal srcDC As Long, Optional ByVal graphicsAntialiasing As GP_SmoothingMode = GP_SM_None, Optional ByVal graphicsPixelOffsetMode As GP_PixelOffsetMode = GP_POM_None) As Long
    Dim hGraphics As Long
    If (GdipCreateFromHDC(srcDC, hGraphics) = GP_OK) Then
        SetGDIPlusGraphicsProperty hGraphics, P2_SurfaceAntialiasing, graphicsAntialiasing
        SetGDIPlusGraphicsProperty hGraphics, P2_SurfacePixelOffset, graphicsPixelOffsetMode
        GetGDIPlusGraphicsFromDC = hGraphics
    Else
        GetGDIPlusGraphicsFromDC = 0
    End If
End Function

'Shorthand function for quickly creating a new GDI+ pen.  This can be useful if many drawing operations are going to be applied with the same pen.
' (Note that a single parameter is used to set both pen and dash endcaps; if you want these to differ, you must call the separate
' SetPenDashCap function, below.)
Public Function GetGDIPlusPenHandle(ByVal penColor As Long, Optional ByVal penOpacity As Long = 255&, Optional ByVal penWidth As Single = 1#, Optional ByVal penLineCap As GP_LineCap = GP_LC_Flat, Optional ByVal penLineJoin As GP_LineJoin = GP_LJ_Miter, Optional ByVal penDashMode As GP_DashStyle = GP_DS_Solid, Optional ByVal penMiterLimit As Single = 3#, Optional ByVal penAlignment As GP_PenAlignment = GP_PA_Center) As Long

    'Create the base pen
    Dim hPen As Long
    GdipCreatePen1 FillQuadWithVBRGB(penColor, penOpacity), penWidth, GP_U_Pixel, hPen
    
    If (hPen <> 0) Then
        
        GdipSetPenLineCap hPen, penLineCap, penLineCap, 0&
        GdipSetPenLineJoin hPen, penLineJoin
        
        If (penDashMode <> GP_DS_Solid) Then
            
            GdipSetPenDashStyle hPen, penDashMode
            
            'Mirror the line cap across the dashes as well
            If (penLineCap = GP_LC_ArrowAnchor) Or (penLineCap = GP_LC_DiamondAnchor) Then
                GdipSetPenDashCap hPen, GP_DC_Triangle
            ElseIf (penLineCap = GP_LC_Round) Or (penLineCap = GP_LC_RoundAnchor) Then
                GdipSetPenDashCap hPen, GP_DC_Round
            Else
                GdipSetPenDashCap hPen, GP_DC_Flat
            End If
            
        End If
        
        'To avoid major miter errors, we default to 3.0 for a miter limit.  (GDI+ defaults to 10, which can easily cause artifacts.)
        GdipSetPenMiterLimit hPen, penMiterLimit
        
        'Finally, if a non-standard alignment was specified, apply it last
        If (penAlignment <> GP_PA_Center) Then GdipSetPenMode hPen, penAlignment
        
    End If
    
    GetGDIPlusPenHandle = hPen

End Function

Public Function GetGDIPlusPenFromBrush(ByVal hBrush As Long, ByVal penWidth As Single, Optional ByVal penUnit As GP_Unit = GP_U_Pixel) As Long
    GdipCreatePenFromBrush hBrush, penWidth, penUnit, GetGDIPlusPenFromBrush
End Function

Public Function GetGDIPlusRegionHandle() As Long
    GdipCreateRegion GetGDIPlusRegionHandle
End Function

Public Function ReleaseGDIPlusBrush(ByVal srcHandle As Long) As Boolean
    If (srcHandle <> 0) Then
        ReleaseGDIPlusBrush = CBool(GdipDeleteBrush(srcHandle) = GP_OK)
    Else
        ReleaseGDIPlusBrush = True
    End If
End Function

Public Function ReleaseGDIPlusGraphics(ByVal srcHandle As Long) As Boolean
    If (srcHandle <> 0) Then
        ReleaseGDIPlusGraphics = CBool(GdipDeleteGraphics(srcHandle) = GP_OK)
    Else
        ReleaseGDIPlusGraphics = True
    End If
End Function

Public Function ReleaseGDIPlusImage(ByVal srcHandle As Long) As Boolean
    If (srcHandle <> 0) Then
        ReleaseGDIPlusImage = CBool(GdipDisposeImage(srcHandle) = GP_OK)
    Else
        ReleaseGDIPlusImage = True
    End If
End Function

Public Function ReleaseGDIPlusPen(ByVal srcHandle As Long) As Boolean
    If (srcHandle <> 0) Then
        ReleaseGDIPlusPen = CBool(GdipDeletePen(srcHandle) = GP_OK)
    Else
        ReleaseGDIPlusPen = True
    End If
End Function

Public Function ReleaseGDIPlusRegion(ByVal srcHandle As Long) As Boolean
    If (srcHandle <> 0) Then
        ReleaseGDIPlusRegion = CBool(GdipDeleteRegion(srcHandle) = GP_OK)
    Else
        ReleaseGDIPlusRegion = True
    End If
End Function

'NOTE!  ALL OPACITY SETTINGS are treated as singles on the range [0, 100], *not* as bytes on the range [0, 255].
'NOTE!  When getting or setting brush settings, you need to make sure the current brush type matches.  For example: if your
'       brush handle points to a solid brush, getting/setting its pattern style is meaningless.  You need to set the
'       relevant brush mode PRIOR to getting/setting other settings.
'NOTE!  Some brush settings cannot be set or retrieved.  For example, GDI+ does not allow you to change hatch style, color,
'       or opacity after brush creation.  You must create a new brush from scratch.  If you use the pd2DBrush class instead
'       of interfacing with these functions directly, nuances like this are handled automatically.
Public Function GetGDIPlusBrushProperty(ByVal hBrush As Long, ByVal propID As PD_2D_BRUSH_SETTINGS) As Variant
    
    If (hBrush <> 0) Then
        
        Dim gResult As GP_Result
        Dim tmpLong As Long, tmpSingle As Single
        
        Select Case propID
            
            'GDI+ does provide a function for this, but their enums differ from ours (by design).
            ' As such, you cannot set brush mode with this function; use the pd2DBrush class, instead.
            Case P2_BrushMode
                GetGDIPlusBrushProperty = 0&
                
           Case P2_BrushColor
                gResult = GdipGetSolidFillColor(hBrush, tmpLong)
                GetGDIPlusBrushProperty = GetColorFromPARGB(tmpLong)
                
            Case P2_BrushOpacity
                gResult = GdipGetSolidFillColor(hBrush, tmpLong)
                GetGDIPlusBrushProperty = GetOpacityFromPARGB(tmpLong)
                
            'Not directly supported by GDI+; use the pd2DBrush class to handle this
            Case P2_BrushPatternStyle
                GetGDIPlusBrushProperty = 0&
                
            'Not directly supported by GDI+; use the pd2DBrush class to handle this
            Case P2_BrushPattern1Color
                GetGDIPlusBrushProperty = 0&
                
            'Not directly supported by GDI+; use the pd2DBrush class to handle this
            Case P2_BrushPattern1Opacity
                GetGDIPlusBrushProperty = 0#
                
            'Not directly supported by GDI+; use the pd2DBrush class to handle this
            Case P2_BrushPattern2Color
                GetGDIPlusBrushProperty = 0&
                
            'Not directly supported by GDI+; use the pd2DBrush class to handle this
            Case P2_BrushPattern2Opacity
                GetGDIPlusBrushProperty = 0#
                
            'Not directly supported by GDI+; use the pd2DBrush class to handle this
            Case P2_BrushGradientAllSettings
                GetGDIPlusBrushProperty = vbNullString
                
            'Not directly supported by GDI+; use the pd2DBrush class to handle this
            Case P2_BrushGradientShape
                GetGDIPlusBrushProperty = P2_GS_Linear
                
            'Not directly supported by GDI+; use the pd2DBrush class to handle this
            Case P2_BrushGradientAngle
                GetGDIPlusBrushProperty = 0#
            
            'Not directly supported by GDI+; use the pd2DBrush class to handle this
            Case P2_BrushGradientWrapMode
                GetGDIPlusBrushProperty = P2_WM_TileFlipXY
            
            'Not directly supported by GDI+; use the pd2DBrush class to handle this
            Case P2_BrushGradientNodes
                GetGDIPlusBrushProperty = vbNullString
                
            Case P2_BrushTextureWrapMode
                gResult = GdipGetTextureWrapMode(hBrush, tmpLong)
                GetGDIPlusBrushProperty = tmpLong
                
        End Select
        
        If (gResult <> GP_OK) Then
            InternalGDIPlusError "GetGDIPlusBrushProperty Error", "Bad GP_RESULT value", gResult
        End If
    
    Else
        InternalGDIPlusError "GetGDIPlusBrushProperty Error", "Null brush handle"
    End If
    
End Function

Public Function SetGDIPlusBrushProperty(ByVal hBrush As Long, ByVal propID As PD_2D_BRUSH_SETTINGS, ByVal newSetting As Variant) As Boolean
    
    If (hBrush <> 0) Then
        
        Dim tmpColor As Long, tmpOpacity As Single
        
        Select Case propID
            
            'Not directly supported by GDI+; use the pd2DBrush class to handle this
            Case P2_BrushMode
                SetGDIPlusBrushProperty = False
                
            Case P2_BrushColor
                tmpOpacity = GetGDIPlusBrushProperty(hBrush, P2_BrushOpacity)
                SetGDIPlusBrushProperty = CBool(GdipSetSolidFillColor(hBrush, FillQuadWithVBRGB(CLng(newSetting), tmpOpacity * 2.55)) = GP_OK)
                
            Case P2_BrushOpacity
                tmpColor = GetGDIPlusBrushProperty(hBrush, P2_BrushColor)
                SetGDIPlusBrushProperty = CBool(GdipSetSolidFillColor(hBrush, FillQuadWithVBRGB(tmpColor, CSng(newSetting) * 2.55)) = GP_OK)
                
            'Not directly supported by GDI+; use the pd2DBrush class to handle this
            Case P2_BrushPatternStyle
                SetGDIPlusBrushProperty = False
                
            'Not directly supported by GDI+; use the pd2DBrush class to handle this
            Case P2_BrushPattern1Color
                SetGDIPlusBrushProperty = False
                
            'Not directly supported by GDI+; use the pd2DBrush class to handle this
            Case P2_BrushPattern1Opacity
                SetGDIPlusBrushProperty = False
                
            'Not directly supported by GDI+; use the pd2DBrush class to handle this
            Case P2_BrushPattern2Color
                SetGDIPlusBrushProperty = False
                
            'Not directly supported by GDI+; use the pd2DBrush class to handle this
            Case P2_BrushPattern2Opacity
                SetGDIPlusBrushProperty = False
                
            'Not directly supported by GDI+; use the pd2DBrush class to handle this
            Case P2_BrushGradientAllSettings
                SetGDIPlusBrushProperty = False
            
            'Not directly supported by GDI+; use the pd2DBrush class to handle this
            Case P2_BrushGradientShape
                SetGDIPlusBrushProperty = False
                
            'Not directly supported by GDI+; use the pd2DBrush class to handle this
            Case P2_BrushGradientAngle
                SetGDIPlusBrushProperty = False
            
            'Not directly supported by GDI+; use the pd2DBrush class to handle this
            Case P2_BrushGradientWrapMode
                SetGDIPlusBrushProperty = False
            
            'Not directly supported by GDI+; use the pd2DBrush class to handle this
            Case P2_BrushGradientNodes
                SetGDIPlusBrushProperty = False
                
            Case P2_BrushTextureWrapMode
                SetGDIPlusBrushProperty = CBool(GdipSetTextureWrapMode(hBrush, CLng(newSetting)) = GP_OK)
                
        End Select
    
    Else
        InternalGDIPlusError "SetGDIPlusBrushProperty Error", "Null brush handle"
    End If
    
End Function

Public Function GetGDIPlusGraphicsProperty(ByVal hGraphics As Long, ByVal propID As PD_2D_SURFACE_SETTINGS) As Variant
    
    If (hGraphics <> 0) Then
        
        Dim gResult As GP_Result
        Dim tmpLong1 As Long, tmpLong2 As Long
        
        Select Case propID
            
            Case P2_SurfaceAntialiasing
                gResult = GdipGetSmoothingMode(hGraphics, tmpLong1)
                GetGDIPlusGraphicsProperty = tmpLong1
                
            Case P2_SurfacePixelOffset
                gResult = GdipGetPixelOffsetMode(hGraphics, tmpLong1)
                GetGDIPlusGraphicsProperty = tmpLong1
                
            Case P2_SurfaceRenderingOriginX
                gResult = GdipGetRenderingOrigin(hGraphics, tmpLong1, tmpLong2)
                GetGDIPlusGraphicsProperty = tmpLong1
            
            Case P2_SurfaceRenderingOriginY
                gResult = GdipGetRenderingOrigin(hGraphics, tmpLong1, tmpLong2)
                GetGDIPlusGraphicsProperty = tmpLong2
                
            Case P2_SurfaceBlendUsingSRGBGamma
                gResult = GdipGetCompositingQuality(hGraphics, tmpLong1)
                GetGDIPlusGraphicsProperty = tmpLong1
                
            Case P2_SurfaceResizeQuality
                gResult = GdipGetInterpolationMode(hGraphics, tmpLong1)
                GetGDIPlusGraphicsProperty = tmpLong1
                
        End Select
        
        If (gResult <> GP_OK) Then
            InternalGDIPlusError "GetGDIPlusGraphicsProperty Error", "Bad GP_RESULT value", gResult
        End If
    
    Else
        InternalGDIPlusError "GetGDIPlusGraphicsProperty Error", "Null graphics handle"
    End If
    
End Function

Public Function SetGDIPlusGraphicsProperty(ByVal hGraphics As Long, ByVal propID As PD_2D_SURFACE_SETTINGS, ByVal newSetting As Variant) As Boolean
    
    If (hGraphics <> 0) Then
        
        Select Case propID
            
            Case P2_SurfaceAntialiasing
                SetGDIPlusGraphicsProperty = CBool(GdipSetSmoothingMode(hGraphics, CLng(newSetting)) = GP_OK)
                
            Case P2_SurfacePixelOffset
                SetGDIPlusGraphicsProperty = CBool(GdipSetPixelOffsetMode(hGraphics, CLng(newSetting)) = GP_OK)
            
            Case P2_SurfaceRenderingOriginX
                SetGDIPlusGraphicsProperty = CBool(GdipSetRenderingOrigin(hGraphics, CLng(newSetting), GetGDIPlusGraphicsProperty(hGraphics, P2_SurfaceRenderingOriginY)) = GP_OK)
            
            Case P2_SurfaceRenderingOriginY
                SetGDIPlusGraphicsProperty = CBool(GdipSetRenderingOrigin(hGraphics, GetGDIPlusGraphicsProperty(hGraphics, P2_SurfaceRenderingOriginX), CLng(newSetting)) = GP_OK)
                
            Case P2_SurfaceBlendUsingSRGBGamma
                SetGDIPlusGraphicsProperty = CBool(GdipSetCompositingQuality(hGraphics, CLng(newSetting)) = GP_OK)
                
            Case P2_SurfaceResizeQuality
                SetGDIPlusGraphicsProperty = CBool(GdipSetInterpolationMode(hGraphics, CLng(newSetting)) = GP_OK)
            
        End Select
    
    Else
        InternalGDIPlusError "GetGDIPlusGraphicsProperty Error", "Null graphics handle"
    End If
    
End Function

'NOTE!  PEN OPACITY setting is treated as a single on the range [0, 100], *not* as a byte on the range [0, 255]
Public Function GetGDIPlusPenProperty(ByVal hPen As Long, ByVal propID As PD_2D_PEN_SETTINGS) As Variant
    
    If (hPen <> 0) Then
        
        Dim gResult As GP_Result
        Dim tmpLong As Long, tmpSingle As Single
        
        Select Case propID
            
            Case P2_PenStyle
                gResult = GdipGetPenDashStyle(hPen, tmpLong)
                GetGDIPlusPenProperty = tmpLong
            
            Case P2_PenColor
                gResult = GdipGetPenColor(hPen, tmpLong)
                GetGDIPlusPenProperty = GetColorFromPARGB(tmpLong)
                
            Case P2_PenOpacity
                gResult = GdipGetPenColor(hPen, tmpLong)
                GetGDIPlusPenProperty = GetOpacityFromPARGB(tmpLong)
                
            Case P2_PenWidth
                gResult = GdipGetPenWidth(hPen, tmpSingle)
                GetGDIPlusPenProperty = tmpSingle
                
            Case P2_PenLineJoin
                gResult = GdipGetPenLineJoin(hPen, tmpLong)
                GetGDIPlusPenProperty = tmpLong
                
            Case P2_PenLineCap
                gResult = GdipGetPenStartCap(hPen, tmpLong)
                GetGDIPlusPenProperty = tmpLong
                
            Case P2_PenDashCap
                gResult = GdipGetPenDashCap(hPen, tmpLong)
                GetGDIPlusPenProperty = tmpLong
                
            Case P2_PenMiterLimit
                gResult = GdipGetPenMiterLimit(hPen, tmpSingle)
                GetGDIPlusPenProperty = tmpSingle
                
            Case P2_PenAlignment
                gResult = GdipGetPenMode(hPen, tmpLong)
                GetGDIPlusPenProperty = tmpLong
                
            Case P2_PenStartCap
                gResult = GdipGetPenStartCap(hPen, tmpLong)
                GetGDIPlusPenProperty = tmpLong
            
            Case P2_PenEndCap
                gResult = GdipGetPenEndCap(hPen, tmpLong)
                GetGDIPlusPenProperty = tmpLong
                
        End Select
        
        If (gResult <> GP_OK) Then
            InternalGDIPlusError "GetGDIPlusPenProperty Error", "Bad GP_RESULT value", gResult
        End If
    
    Else
        InternalGDIPlusError "GetGDIPlusPenProperty Error", "Null pen handle"
    End If
    
End Function

'NOTE!  PEN OPACITY setting is treated as a single on the range [0, 100], *not* as a byte on the range [0, 255]
Public Function SetGDIPlusPenProperty(ByVal hPen As Long, ByVal propID As PD_2D_PEN_SETTINGS, ByVal newSetting As Variant) As Boolean
    
    If (hPen <> 0) Then
        
        Dim tmpColor As Long, tmpOpacity As Single, tmpLong As Long
        
        Select Case propID
            
            Case P2_PenStyle
                SetGDIPlusPenProperty = CBool(GdipSetPenDashStyle(hPen, CLng(newSetting)) = GP_OK)
                
            Case P2_PenColor
                tmpOpacity = GetGDIPlusPenProperty(hPen, P2_PenOpacity)
                SetGDIPlusPenProperty = CBool(GdipSetPenColor(hPen, FillQuadWithVBRGB(CLng(newSetting), tmpOpacity * 2.55)) = GP_OK)
                
            Case P2_PenOpacity
                tmpColor = GetGDIPlusPenProperty(hPen, P2_PenColor)
                SetGDIPlusPenProperty = CBool(GdipSetPenColor(hPen, FillQuadWithVBRGB(tmpColor, CSng(newSetting) * 2.55)) = GP_OK)
                
            Case P2_PenWidth
                SetGDIPlusPenProperty = CBool(GdipSetPenDashStyle(hPen, CSng(newSetting)) = GP_OK)
                
            Case P2_PenLineJoin
                SetGDIPlusPenProperty = CBool(GdipSetPenLineJoin(hPen, CLng(newSetting)) = GP_OK)
                
            Case P2_PenLineCap
                tmpLong = GetGDIPlusPenProperty(hPen, P2_PenDashCap)
                SetGDIPlusPenProperty = CBool(GdipSetPenLineCap(hPen, CLng(newSetting), CLng(newSetting), tmpLong) = GP_OK)
                
            Case P2_PenDashCap
                SetGDIPlusPenProperty = CBool(GdipSetPenDashCap(hPen, CLng(newSetting)) = GP_OK)
                
            Case P2_PenMiterLimit
                SetGDIPlusPenProperty = CBool(GdipSetPenMiterLimit(hPen, CSng(newSetting)) = GP_OK)
                
            Case P2_PenAlignment
                SetGDIPlusPenProperty = CBool(GdipSetPenMode(hPen, CLng(newSetting)) = GP_OK)
            
            Case P2_PenStartCap
                SetGDIPlusPenProperty = CBool(GdipSetPenStartCap(hPen, CLng(newSetting)) = GP_OK)
            
            Case P2_PenEndCap
                SetGDIPlusPenProperty = CBool(GdipSetPenEndCap(hPen, CLng(newSetting)) = GP_OK)
                
        End Select
    
    Else
        InternalGDIPlusError "SetGDIPlusPenProperty Error", "Null pen handle"
    End If
    
End Function

'All generic draw and fill functions follow

'GDI+ arcs use bounding boxes to describe their placement.  As such, we manually convert the incoming centerX/Y and radius values
' to bounding box coordinates.
Public Function GDIPlus_DrawArcF(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal centerX As Single, ByVal centerY As Single, ByVal arcRadius As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As Boolean
    GDIPlus_DrawArcF = CBool(GdipDrawArc(dstGraphics, srcPen, centerX - arcRadius, centerY - arcRadius, arcRadius * 2, arcRadius * 2, startAngle, sweepAngle) = GP_OK)
End Function

Public Function GDIPlus_DrawArcI(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal centerX As Long, ByVal centerY As Long, ByVal arcRadius As Long, ByVal startAngle As Long, ByVal sweepAngle As Long) As Boolean
    GDIPlus_DrawArcI = CBool(GdipDrawArcI(dstGraphics, srcPen, centerX - arcRadius, centerY - arcRadius, arcRadius * 2, arcRadius * 2, startAngle, sweepAngle) = GP_OK)
End Function

Public Function GDIPlus_DrawClosedCurveF(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal ptrToPtFArray As Long, ByVal numOfPoints As Long, Optional ByVal curveTension As Single = 0.5) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawClosedCurve2(dstGraphics, srcPen, ptrToPtFArray, numOfPoints, curveTension)
    GDIPlus_DrawClosedCurveF = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_DrawClosedCurveI(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal ptrToPtLArray As Long, ByVal numOfPoints As Long, Optional ByVal curveTension As Single = 0.5) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawClosedCurve2I(dstGraphics, srcPen, ptrToPtLArray, numOfPoints, curveTension)
    GDIPlus_DrawClosedCurveI = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_DrawCurveF(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal ptrToPtFArray As Long, ByVal numOfPoints As Long, Optional ByVal curveTension As Single = 0.5) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawCurve2(dstGraphics, srcPen, ptrToPtFArray, numOfPoints, curveTension)
    GDIPlus_DrawCurveF = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_DrawCurveI(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal ptrToPtLArray As Long, ByVal numOfPoints As Long, Optional ByVal curveTension As Single = 0.5) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawCurve2I(dstGraphics, srcPen, ptrToPtLArray, numOfPoints, curveTension)
    GDIPlus_DrawCurveI = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_DrawImageF(ByVal dstGraphics As Long, ByVal srcImage As Long, ByVal dstX As Single, ByVal dstY As Single) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawImage(dstGraphics, srcImage, dstX, dstY)
    GDIPlus_DrawImageF = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_DrawImageI(ByVal dstGraphics As Long, ByVal srcImage As Long, ByVal dstX As Long, ByVal dstY As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawImageI(dstGraphics, srcImage, dstX, dstY)
    GDIPlus_DrawImageI = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_DrawImageRectF(ByVal dstGraphics As Long, ByVal srcImage As Long, ByVal dstX As Single, ByVal dstY As Single, ByVal dstWidth As Single, ByVal dstHeight As Single) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawImageRect(dstGraphics, srcImage, dstX, dstY, dstWidth, dstHeight)
    GDIPlus_DrawImageRectF = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_DrawImageRectI(ByVal dstGraphics As Long, ByVal srcImage As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawImageRectI(dstGraphics, srcImage, dstX, dstY, dstWidth, dstHeight)
    GDIPlus_DrawImageRectI = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_DrawImageRectRectF(ByVal dstGraphics As Long, ByVal srcImage As Long, ByVal dstX As Single, ByVal dstY As Single, ByVal dstWidth As Single, ByVal dstHeight As Single, ByVal srcX As Single, ByVal srcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, Optional ByVal opacityModifier As Single = 1#) As Boolean
    
    'Modified opacity requires us to create a temporary image attributes object
    Dim imgAttributesHandle As Long
    If (opacityModifier <> 1#) Then
        GdipCreateImageAttributes imgAttributesHandle
        m_AttributesMatrix(3, 3) = opacityModifier
        GdipSetImageAttributesColorMatrix imgAttributesHandle, GP_CAT_Bitmap, 1&, VarPtr(m_AttributesMatrix(0, 0)), 0&, GP_CMF_Default
    End If
    
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawImageRectRect(dstGraphics, srcImage, dstX, dstY, dstWidth, dstHeight, srcX, srcY, srcWidth, srcHeight, GP_U_Pixel, imgAttributesHandle, 0&, 0&)
    GDIPlus_DrawImageRectRectF = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
    
    'As necessary, release our image attributes object, and reset the alpha value of the master identity matrix
    If (opacityModifier <> 1#) Then
        GdipDisposeImageAttributes imgAttributesHandle
        m_AttributesMatrix(3, 3) = 1#
    End If
        
End Function

Public Function GDIPlus_DrawImageRectRectI(ByVal dstGraphics As Long, ByVal srcImage As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, Optional ByVal opacityModifier As Single = 1#) As Boolean
    
    'Modified opacity requires us to create a temporary image attributes object
    Dim imgAttributesHandle As Long
    If (opacityModifier <> 1#) Then
        GdipCreateImageAttributes imgAttributesHandle
        m_AttributesMatrix(3, 3) = opacityModifier
        GdipSetImageAttributesColorMatrix imgAttributesHandle, GP_CAT_Bitmap, 1&, VarPtr(m_AttributesMatrix(0, 0)), 0&, GP_CMF_Default
    End If
    
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawImageRectRectI(dstGraphics, srcImage, dstX, dstY, dstWidth, dstHeight, srcX, srcY, srcWidth, srcHeight, GP_U_Pixel, imgAttributesHandle, 0&, 0&)
    GDIPlus_DrawImageRectRectI = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
    
    'As necessary, release our image attributes object, and reset the alpha value of the master identity matrix
    If (opacityModifier <> 1#) Then
        GdipDisposeImageAttributes imgAttributesHandle
        m_AttributesMatrix(3, 3) = 1#
    End If
        
End Function

Public Function GDIPlus_DrawLineF(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As Boolean
    GDIPlus_DrawLineF = CBool(GdipDrawLine(dstGraphics, srcPen, x1, y1, x2, y2) = GP_OK)
End Function

Public Function GDIPlus_DrawLineI(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Boolean
    GDIPlus_DrawLineI = CBool(GdipDrawLineI(dstGraphics, srcPen, x1, y1, x2, y2) = GP_OK)
End Function

Public Function GDIPlus_DrawLinesF(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal ptrToPtFArray As Long, ByVal numOfPoints As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawLines(dstGraphics, srcPen, ptrToPtFArray, numOfPoints)
    GDIPlus_DrawLinesF = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_DrawLinesI(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal ptrToPtLArray As Long, ByVal numOfPoints As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawLinesI(dstGraphics, srcPen, ptrToPtLArray, numOfPoints)
    GDIPlus_DrawLinesI = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_DrawPath(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal srcPath As Long) As Boolean
    GDIPlus_DrawPath = CBool(GdipDrawPath(dstGraphics, srcPen, srcPath) = GP_OK)
End Function

Public Function GDIPlus_DrawPolygonF(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal ptrToPtFArray As Long, ByVal numOfPoints As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawPolygon(dstGraphics, srcPen, ptrToPtFArray, numOfPoints)
    GDIPlus_DrawPolygonF = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_DrawPolygonI(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal ptrToPtLArray As Long, ByVal numOfPoints As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawPolygonI(dstGraphics, srcPen, ptrToPtLArray, numOfPoints)
    GDIPlus_DrawPolygonI = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_DrawRectF(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal rectLeft As Single, ByVal rectTop As Single, ByVal rectWidth As Single, ByVal rectHeight As Single) As Boolean
    GDIPlus_DrawRectF = CBool(GdipDrawRectangle(dstGraphics, srcPen, rectLeft, rectTop, rectWidth, rectHeight) = GP_OK)
End Function

Public Function GDIPlus_DrawRectI(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal rectLeft As Long, ByVal rectTop As Long, ByVal rectWidth As Long, ByVal rectHeight As Long) As Boolean
    GDIPlus_DrawRectI = CBool(GdipDrawRectangleI(dstGraphics, srcPen, rectLeft, rectTop, rectWidth, rectHeight) = GP_OK)
End Function

Public Function GDIPlus_DrawEllipseF(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal ellipseLeft As Single, ByVal ellipseTop As Single, ByVal ellipseWidth As Single, ByVal ellipseHeight As Single) As Boolean
    GDIPlus_DrawEllipseF = CBool(GdipDrawEllipse(dstGraphics, srcPen, ellipseLeft, ellipseTop, ellipseWidth, ellipseHeight) = GP_OK)
End Function

Public Function GDIPlus_DrawEllipseI(ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal ellipseLeft As Long, ByVal ellipseTop As Long, ByVal ellipseWidth As Long, ByVal ellipseHeight As Long) As Boolean
    GDIPlus_DrawEllipseI = CBool(GdipDrawEllipseI(dstGraphics, srcPen, ellipseLeft, ellipseTop, ellipseWidth, ellipseHeight) = GP_OK)
End Function

Public Function GDIPlus_FillClosedCurveF(ByVal dstGraphics As Long, ByVal srcBrush As Long, ByVal ptrToPtFArray As Long, ByVal numOfPoints As Long, Optional ByVal curveTension As Single = 0.5, Optional ByVal fillMode As GP_FillMode = GP_FM_Winding) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipFillClosedCurve2(dstGraphics, srcBrush, ptrToPtFArray, numOfPoints, curveTension, fillMode)
    GDIPlus_FillClosedCurveF = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_FillClosedCurveI(ByVal dstGraphics As Long, ByVal srcBrush As Long, ByVal ptrToPtLArray As Long, ByVal numOfPoints As Long, Optional ByVal curveTension As Single = 0.5, Optional ByVal fillMode As GP_FillMode = GP_FM_Winding) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipFillClosedCurve2I(dstGraphics, srcBrush, ptrToPtLArray, numOfPoints, curveTension, fillMode)
    GDIPlus_FillClosedCurveI = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_FillPath(ByVal dstGraphics As Long, ByVal srcBrush As Long, ByVal srcPath As Long) As Boolean
    GDIPlus_FillPath = CBool(GdipFillPath(dstGraphics, srcBrush, srcPath) = GP_OK)
End Function

Public Function GDIPlus_FillEllipseF(ByVal dstGraphics As Long, ByVal srcBrush As Long, ByVal ellipseLeft As Single, ByVal ellipseTop As Single, ByVal ellipseWidth As Single, ByVal ellipseHeight As Single) As Boolean
    GDIPlus_FillEllipseF = CBool(GdipFillEllipse(dstGraphics, srcBrush, ellipseLeft, ellipseTop, ellipseWidth, ellipseHeight) = GP_OK)
End Function

Public Function GDIPlus_FillEllipseI(ByVal dstGraphics As Long, ByVal srcBrush As Long, ByVal ellipseLeft As Long, ByVal ellipseTop As Long, ByVal ellipseWidth As Long, ByVal ellipseHeight As Long) As Boolean
    GDIPlus_FillEllipseI = CBool(GdipFillEllipseI(dstGraphics, srcBrush, ellipseLeft, ellipseTop, ellipseWidth, ellipseHeight) = GP_OK)
End Function

Public Function GDIPlus_FillPolygonF(ByVal dstGraphics As Long, ByVal srcBrush As Long, ByVal ptrToPtFArray As Long, ByVal numOfPoints As Long, Optional ByVal fillMode As GP_FillMode = GP_FM_Winding) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipFillPolygon(dstGraphics, srcBrush, ptrToPtFArray, numOfPoints, fillMode)
    GDIPlus_FillPolygonF = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_FillPolygonI(ByVal dstGraphics As Long, ByVal srcBrush As Long, ByVal ptrToPtLArray As Long, ByVal numOfPoints As Long, Optional ByVal fillMode As GP_FillMode = GP_FM_Winding) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipFillPolygonI(dstGraphics, srcBrush, ptrToPtLArray, numOfPoints, fillMode)
    GDIPlus_FillPolygonI = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_FillRectF(ByVal dstGraphics As Long, ByVal srcBrush As Long, ByVal rectLeft As Single, ByVal rectTop As Single, ByVal rectWidth As Single, ByVal rectHeight As Single) As Boolean
    GDIPlus_FillRectF = CBool(GdipFillRectangle(dstGraphics, srcBrush, rectLeft, rectTop, rectWidth, rectHeight) = GP_OK)
End Function

Public Function GDIPlus_FillRectI(ByVal dstGraphics As Long, ByVal srcBrush As Long, ByVal rectLeft As Long, ByVal rectTop As Long, ByVal rectWidth As Long, ByVal rectHeight As Long) As Boolean
    GDIPlus_FillRectI = CBool(GdipFillRectangleI(dstGraphics, srcBrush, rectLeft, rectTop, rectWidth, rectHeight) = GP_OK)
End Function

Public Function GDIPlus_GraphicsGetClipRegion(ByVal srcGraphics As Long) As Long
    Dim tmpReturn As GP_Result
    tmpReturn = GdipGetClip(srcGraphics, GDIPlus_GraphicsGetClipRegion)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_GraphicsResetClipRegion(ByVal dstGraphics As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipResetClip(dstGraphics)
    GDIPlus_GraphicsResetClipRegion = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_GraphicsSetClipRect(ByVal dstGraphics As Long, ByVal clipX As Single, ByVal clipY As Single, ByVal clipWidth As Single, ByVal clipHeight As Single, Optional ByVal useCombineMode As GP_CombineMode = GP_CM_Replace) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipSetClipRect(dstGraphics, clipX, clipY, clipWidth, clipHeight, useCombineMode)
    GDIPlus_GraphicsSetClipRect = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_GraphicsSetClipRegion(ByVal dstGraphics As Long, ByVal srcRegion As Long, Optional ByVal useCombineMode As GP_CombineMode = GP_CM_Replace) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipSetClipRegion(dstGraphics, srcRegion, useCombineMode)
    GDIPlus_GraphicsSetClipRegion = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_GraphicsSetCompositingMode(ByVal dstGraphics As Long, Optional ByVal newCompositeMode As GP_CompositingMode = GP_CM_SourceOver) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipSetCompositingMode(dstGraphics, newCompositeMode)
    GDIPlus_GraphicsSetCompositingMode = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_MatrixCreate() As Long
    Dim tmpReturn As GP_Result
    tmpReturn = GdipCreateMatrix(GDIPlus_MatrixCreate)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_MatrixClone(ByVal srcMatrix As Long) As Long
    Dim tmpReturn As GP_Result
    tmpReturn = GdipCloneMatrix(srcMatrix, GDIPlus_MatrixClone)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_MatrixDelete(ByVal hMatrix As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDeleteMatrix(hMatrix)
    GDIPlus_MatrixDelete = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_MatrixInvert(ByVal hMatrix As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipInvertMatrix(hMatrix)
    GDIPlus_MatrixInvert = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_MatrixIsInvertible(ByVal hMatrix As Long) As Boolean
    Dim tmpResult As Long, tmpReturn As GP_Result
    tmpReturn = GdipIsMatrixInvertible(hMatrix, tmpResult)
    GDIPlus_MatrixIsInvertible = CBool(tmpResult <> 0)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_MatrixRotate(ByVal hMatrix As Long, ByVal rotateAngle As Single, Optional ByVal operationOrder As GP_MatrixOrder = GP_MO_Append) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipRotateMatrix(hMatrix, rotateAngle, operationOrder)
    GDIPlus_MatrixRotate = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_MatrixScale(ByVal hMatrix As Long, ByVal scaleX As Single, ByVal scaleY As Single, Optional ByVal operationOrder As GP_MatrixOrder = GP_MO_Append) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipScaleMatrix(hMatrix, scaleX, scaleY, operationOrder)
    GDIPlus_MatrixScale = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_MatrixShear(ByVal hMatrix As Long, ByVal shearX As Single, ByVal shearY As Single, Optional ByVal operationOrder As GP_MatrixOrder = GP_MO_Append) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipShearMatrix(hMatrix, shearX, shearY, operationOrder)
    GDIPlus_MatrixShear = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_MatrixTransformListOfPoints(ByVal hMatrix As Long, ByVal ptrToFirstPointF As Long, ByVal numOfPoints As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipTransformMatrixPoints(hMatrix, ptrToFirstPointF, numOfPoints)
    GDIPlus_MatrixTransformListOfPoints = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_MatrixTranslate(ByVal hMatrix As Long, ByVal offsetX As Single, ByVal offsetY As Single, Optional ByVal operationOrder As GP_MatrixOrder = GP_MO_Append) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipTranslateMatrix(hMatrix, offsetX, offsetY, operationOrder)
    GDIPlus_MatrixTranslate = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathAddArc(ByVal hPath As Long, ByVal x As Single, ByVal y As Single, ByVal arcWidth As Single, ByVal arcHeight As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipAddPathArc(hPath, x, y, arcWidth, arcHeight, startAngle, sweepAngle)
    GDIPlus_PathAddArc = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathAddBezier(ByVal hPath As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal x3 As Single, ByVal y3 As Single, ByVal x4 As Single, ByVal y4 As Single) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipAddPathBezier(hPath, x1, y1, x2, y2, x3, y3, x4, y4)
    GDIPlus_PathAddBezier = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathAddClosedCurve(ByVal hPath As Long, ByVal ptrToFloatArray As Long, ByVal numOfPoints As Long, Optional ByVal curveTension As Single = 0.5) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipAddPathClosedCurve2(hPath, ptrToFloatArray, numOfPoints, curveTension)
    GDIPlus_PathAddClosedCurve = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathAddCurve(ByVal hPath As Long, ByVal ptrToFloatArray As Long, ByVal numOfPoints As Long, ByVal curveTension As Single) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipAddPathCurve2(hPath, ptrToFloatArray, numOfPoints, curveTension)
    GDIPlus_PathAddCurve = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathAddEllipse(ByVal hPath As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal rectWidth As Single, ByVal rectHeight As Single) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipAddPathEllipse(hPath, x1, y1, rectWidth, rectHeight)
    GDIPlus_PathAddEllipse = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathAddLine(ByVal hPath As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipAddPathLine(hPath, x1, y1, x2, y2)
    GDIPlus_PathAddLine = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathAddLines(ByVal hPath As Long, ByVal ptrToFloatArray As Long, ByVal numOfPoints As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipAddPathLine2(hPath, ptrToFloatArray, numOfPoints)
    GDIPlus_PathAddLines = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathAddPath(ByVal hPath As Long, ByVal pathToAdd As Long, ByVal connectToPreviousPoint As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipAddPathPath(hPath, pathToAdd, connectToPreviousPoint)
    GDIPlus_PathAddPath = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathAddPolygon(ByVal hPath As Long, ByVal ptrToFloatArray As Long, ByVal numOfPoints As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipAddPathPolygon(hPath, ptrToFloatArray, numOfPoints)
    GDIPlus_PathAddPolygon = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathAddRectangle(ByVal hPath As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal rectWidth As Single, ByVal rectHeight As Single) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipAddPathRectangle(hPath, x1, y1, rectWidth, rectHeight)
    GDIPlus_PathAddRectangle = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathClone(ByVal srcPath As Long) As Long
    Dim tmpReturn As GP_Result
    tmpReturn = GdipClonePath(srcPath, GDIPlus_PathClone)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathCloseFigure(ByVal hPath As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipClosePathFigure(hPath)
    GDIPlus_PathCloseFigure = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathCreate(Optional ByVal initFillRule As GP_FillMode = GP_FM_Alternate) As Long
    Dim tmpReturn As GP_Result
    tmpReturn = GdipCreatePath(initFillRule, GDIPlus_PathCreate)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathDelete(ByVal hPath As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDeletePath(hPath)
    GDIPlus_PathDelete = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathDoesPointTouchOutlineF(ByVal hPath As Long, ByVal srcX As Single, ByVal srcY As Single, ByVal hPen As Long) As Boolean
    Dim tmpReturn As GP_Result, tmpResult As Long
    tmpReturn = GdipIsOutlineVisiblePathPoint(hPath, srcX, srcY, hPen, 0&, tmpResult)
    GDIPlus_PathDoesPointTouchOutlineF = CBool(tmpResult <> 0)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathDoesPointTouchOutlineL(ByVal hPath As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal hPen As Long) As Boolean
    Dim tmpReturn As GP_Result, tmpResult As Long
    tmpReturn = GdipIsOutlineVisiblePathPointI(hPath, srcX, srcY, hPen, 0&, tmpResult)
    GDIPlus_PathDoesPointTouchOutlineL = CBool(tmpResult <> 0)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathGetFillRule(ByVal hPath As Long) As GP_FillMode
    Dim tmpReturn As GP_Result
    tmpReturn = GdipGetPathFillMode(hPath, GDIPlus_PathGetFillRule)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathGetPathBoundsF(ByVal hPath As Long, Optional ByVal hTransform As Long = 0, Optional ByVal hPen As Long = 0) As RECTF
    Dim tmpReturn As GP_Result
    tmpReturn = GdipGetPathWorldBounds(hPath, GDIPlus_PathGetPathBoundsF, hTransform, hPen)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathGetPathBoundsL(ByVal hPath As Long, Optional ByVal hTransform As Long = 0, Optional ByVal hPen As Long = 0) As RECTL
    Dim tmpReturn As GP_Result
    tmpReturn = GdipGetPathWorldBoundsI(hPath, GDIPlus_PathGetPathBoundsL, hTransform, hPen)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathIsPointInsideF(ByVal hPath As Long, ByVal srcX As Single, ByVal srcY As Single) As Boolean
    Dim tmpReturn As GP_Result, tmpResult As Long
    tmpReturn = GdipIsVisiblePathPoint(hPath, srcX, srcY, 0&, tmpResult)
    GDIPlus_PathIsPointInsideF = CBool(tmpResult <> 0)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathIsPointInsideL(ByVal hPath As Long, ByVal srcX As Long, ByVal srcY As Long) As Boolean
    Dim tmpReturn As GP_Result, tmpResult As Long
    tmpReturn = GdipIsVisiblePathPointI(hPath, srcX, srcY, 0&, tmpResult)
    GDIPlus_PathIsPointInsideL = CBool(tmpResult <> 0)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathReset(ByVal hPath As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipResetPath(hPath)
    GDIPlus_PathReset = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathSetFillRule(ByVal hPath As Long, ByVal newFillRule As GP_FillMode) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipSetPathFillMode(hPath, newFillRule)
    GDIPlus_PathSetFillRule = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathStartFigure(ByVal hPath As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipStartPathFigure(hPath)
    GDIPlus_PathStartFigure = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathTransform(ByVal hPath As Long, ByVal hTransformMatrix As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipTransformPath(hPath, hTransformMatrix)
    GDIPlus_PathTransform = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathWiden(ByVal hPath As Long, ByVal hPen As Long, ByVal hTransformMatrix As Long, ByVal allowableError As Single) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipWidenPath(hPath, hPen, hTransformMatrix, allowableError)
    GDIPlus_PathWiden = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_PathWindingModeOutline(ByVal hPath As Long, ByVal hTransformMatrix As Long, ByVal allowableError As Single) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipWindingModeOutline(hPath, hTransformMatrix, allowableError)
    GDIPlus_PathWindingModeOutline = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_RegionAddRectF(ByVal dstRegion As Long, ByRef srcRectF As RECTF, Optional ByVal useCombineMode As GP_CombineMode = GP_CM_Replace) As Boolean
    GDIPlus_RegionAddRectF = CBool(GdipCombineRegionRect(dstRegion, srcRectF, useCombineMode) = GP_OK)
End Function

Public Function GDIPlus_RegionAddRectL(ByVal dstRegion As Long, ByRef srcRectL As RECTL, Optional ByVal useCombineMode As GP_CombineMode = GP_CM_Replace) As Boolean
    GDIPlus_RegionAddRectL = CBool(GdipCombineRegionRectI(dstRegion, srcRectL, useCombineMode) = GP_OK)
End Function

Public Function GDIPlus_RegionAddRegion(ByVal dstRegion As Long, ByVal srcRegion As Long, Optional ByVal useCombineMode As GP_CombineMode = GP_CM_Replace) As Boolean
    GDIPlus_RegionAddRegion = CBool(GdipCombineRegionRegion(dstRegion, srcRegion, useCombineMode) = GP_OK)
End Function

Public Function GDIPlus_RegionAddPath(ByVal dstRegion As Long, ByVal srcPath As Long, Optional ByVal useCombineMode As GP_CombineMode = GP_CM_Replace) As Boolean
    GDIPlus_RegionAddPath = CBool(GdipCombineRegionPath(dstRegion, srcPath, useCombineMode) = GP_OK)
End Function

Public Function GDIPlus_RegionClone(ByVal srcRegion As Long, ByRef dstRegion As Long) As Boolean
    Dim tmpReturn As Long
    tmpReturn = GdipCloneRegion(srcRegion, dstRegion)
    GDIPlus_RegionClone = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_RegionGetClipRectF(ByVal srcRegion As Long) As RECTF
    Dim tmpReturn As Long
    tmpReturn = GdipGetRegionBounds(srcRegion, m_TransformGraphics, GDIPlus_RegionGetClipRectF)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_RegionGetClipRectI(ByVal srcRegion As Long) As RECTL
    Dim tmpReturn As Long
    tmpReturn = GdipGetRegionBoundsI(srcRegion, m_TransformGraphics, GDIPlus_RegionGetClipRectI)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_RegionIsInfinite(ByVal srcRegion As Long) As Boolean
    Dim tmpResult As Long, tmpReturn As GP_Result
    tmpReturn = GdipIsInfiniteRegion(srcRegion, m_TransformGraphics, tmpResult)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
    GDIPlus_RegionIsInfinite = CBool(tmpResult <> 0)
End Function

Public Function GDIPlus_RegionIsEmpty(ByVal srcRegion As Long) As Boolean
    Dim tmpResult As Long, tmpReturn As GP_Result
    tmpReturn = GdipIsEmptyRegion(srcRegion, m_TransformGraphics, tmpResult)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
    GDIPlus_RegionIsEmpty = CBool(tmpResult <> 0)
End Function

Public Function GDIPlus_RegionsAreEqual(ByVal srcRegion1 As Long, ByVal srcRegion2 As Long) As Boolean
    Dim tmpResult As Long, tmpReturn As GP_Result
    tmpReturn = GdipIsEqualRegion(srcRegion1, srcRegion2, m_TransformGraphics, tmpResult)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
    GDIPlus_RegionsAreEqual = CBool(tmpResult <> 0)
End Function

Public Function GDIPlus_RegionSetEmpty(ByVal dstRegion As Long) As Boolean
    GDIPlus_RegionSetEmpty = CBool(GdipSetEmpty(dstRegion) = GP_OK)
End Function

Public Function GDIPlus_RegionSetInfinite(ByVal dstRegion As Long) As Boolean
    GDIPlus_RegionSetInfinite = CBool(GdipSetInfinite(dstRegion) = GP_OK)
End Function

