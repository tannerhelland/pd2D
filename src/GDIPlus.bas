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

Public Enum GP_BitmapLockMode
    GP_BLM_Read = &H1
    GP_BLM_Write = &H2
    GP_BLM_UserInputBuf = &H4
End Enum

#If False Then
    Private Const GP_BLM_Read = &H1, GP_BLM_Write = &H2, GP_BLM_UserInputBuf = &H4
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

Public Enum GP_EncoderValueType
    GP_EVT_Byte = 1
    GP_EVT_ASCII = 2
    GP_EVT_Short = 3
    GP_EVT_Long = 4
    GP_EVT_Rational = 5
    GP_EVT_LongRange = 6
    GP_EVT_Undefined = 7
    GP_EVT_RationalRange = 8
    GP_EVT_Pointer = 9
End Enum

#If False Then
    Private Const GP_EVT_Byte = 1, GP_EVT_ASCII = 2, GP_EVT_Short = 3, GP_EVT_Long = 4, GP_EVT_Rational = 5, GP_EVT_LongRange = 6, GP_EVT_Undefined = 7, GP_EVT_RationalRange = 8, GP_EVT_Pointer = 9
#End If

Public Enum GP_EncoderValue
    GP_EV_ColorTypeCMYK = 0
    GP_EV_ColorTypeYCCK = 1
    GP_EV_CompressionLZW = 2
    GP_EV_CompressionCCITT3 = 3
    GP_EV_CompressionCCITT4 = 4
    GP_EV_CompressionRle = 5
    GP_EV_CompressionNone = 6
    GP_EV_ScanMethodInterlaced = 7
    GP_EV_ScanMethodNonInterlaced = 8
    GP_EV_VersionGif87 = 9
    GP_EV_VersionGif89 = 10
    GP_EV_RenderProgressive = 11
    GP_EV_RenderNonProgressive = 12
    GP_EV_TransformRotate90 = 13
    GP_EV_TransformRotate180 = 14
    GP_EV_TransformRotate270 = 15
    GP_EV_TransformFlipHorizontal = 16
    GP_EV_TransformFlipVertical = 17
    GP_EV_MultiFrame = 18
    GP_EV_LastFrame = 19
    GP_EV_Flush = 20
    GP_EV_FrameDimensionTime = 21
    GP_EV_FrameDimensionResolution = 22
    GP_EV_FrameDimensionPage = 23
    GP_EV_ColorTypeGray = 24
    GP_EV_ColorTypeRGB = 25
End Enum

#If False Then
    Private Const GP_EV_ColorTypeCMYK = 0, GP_EV_ColorTypeYCCK = 1, GP_EV_CompressionLZW = 2, GP_EV_CompressionCCITT3 = 3, GP_EV_CompressionCCITT4 = 4, GP_EV_CompressionRle = 5, GP_EV_CompressionNone = 6, GP_EV_ScanMethodInterlaced = 7, GP_EV_ScanMethodNonInterlaced = 8, GP_EV_VersionGif87 = 9, GP_EV_VersionGif89 = 10
    Private Const GP_EV_RenderProgressive = 11, GP_EV_RenderNonProgressive = 12, GP_EV_TransformRotate90 = 13, GP_EV_TransformRotate180 = 14, GP_EV_TransformRotate270 = 15, GP_EV_TransformFlipHorizontal = 16, GP_EV_TransformFlipVertical = 17, GP_EV_MultiFrame = 18, GP_EV_LastFrame = 19, GP_EV_Flush = 20
    Private Const GP_EV_FrameDimensionTime = 21, GP_EV_FrameDimensionResolution = 22, GP_EV_FrameDimensionPage = 23, GP_EV_ColorTypeGray = 24, GP_EV_ColorTypeRGB = 25
#End If

Public Enum GP_FillMode
    GP_FM_Alternate = 0&
    GP_FM_Winding = 1&
End Enum

#If False Then
    Private Const GP_FM_Alternate = 0&, GP_FM_Winding = 1&
#End If

Public Enum GP_ImageType
    GP_IT_Unknown = 0
    GP_IT_Bitmap = 1
    GP_IT_Metafile = 2
End Enum

#If False Then
    Private Const GP_IT_Unknown = 0, GP_IT_Bitmap = 1, GP_IT_Metafile = 2
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

'EMFs can be converted between various formats.  GDI+ prefers "EMF+", which supports GDI+ primitives as well
Public Enum GP_MetafileType
   GP_MT_Invalid = 0
   GP_MT_Wmf = 1
   GP_MT_WmfPlaceable = 2
   GP_MT_Emf = 3              'Old-style EMF consisting only of GDI commands
   GP_MT_EmfPlus = 4          'New-style EMF+ consisting only of GDI+ commands
   GP_MT_EmfDual = 5          'New-style EMF+ with GDI fallbacks for legacy rendering
End Enum

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

'GDI+ pixel format IDs use a bitfield system:
' [0, 7] = format index
' [8, 15] = pixel size (in bits)
' [16, 23] = flags
' [24, 31] = reserved (current unused)

'Note also that pixel format is *not* 100% reliable.  Behavior differs between OSes, even for the "same"
' major GDI+ version.  (See http://stackoverflow.com/questions/5065371/how-to-identify-cmyk-images-in-asp-net-using-c-sharp)
Public Enum GP_PixelFormat
    GP_PF_Indexed = &H10000         'Image uses a palette to define colors
    GP_PF_GDI = &H20000             'Is a format supported by GDI
    GP_PF_Alpha = &H40000           'Alpha channel present
    GP_PF_PreMultAlpha = &H80000    'Alpha is premultiplied (not always correct; manual verification should be used)
    GP_PF_HDR = &H100000            'High bit-depth colors are in use (e.g. 48-bpp or 64-bpp; behavior is unpredictable on XP)
    GP_PF_Canonical = &H200000      'Canonical formats: 32bppARGB, 32bppPARGB, 64bppARGB, 64bppPARGB
    
    GP_PF_32bppCMYK = &H200F        'CMYK is never returned on XP or Vista; ImageFlags can be checked as a failsafe
                                    ' (Conversely, ImageFlags are unreliable on Win 7 - this is the shit we deal with
                                    '  as Windows developers!)
    
    GP_PF_1bppIndexed = &H30101
    GP_PF_4bppIndexed = &H30402
    GP_PF_8bppIndexed = &H30803
    GP_PF_16bppGreyscale = &H101004
    GP_PF_16bppRGB555 = &H21005
    GP_PF_16bppRGB565 = &H21006
    GP_PF_16bppARGB1555 = &H61007
    GP_PF_24bppRGB = &H21808
    GP_PF_32bppRGB = &H22009
    GP_PF_32bppARGB = &H26200A
    GP_PF_32bppPARGB = &HE200B
    GP_PF_48bppRGB = &H10300C
    GP_PF_64bppARGB = &H34400D
    GP_PF_64bppPARGB = &H1C400E
End Enum

#If False Then
    Private Const GP_PF_Indexed = &H10000, GP_PF_GDI = &H20000, GP_PF_Alpha = &H40000, GP_PF_PreMultAlpha = &H80000, GP_PF_HDR = &H100000, GP_PF_Canonical = &H200000, GP_PF_32bppCMYK = &H200F
    Private Const GP_PF_1bppIndexed = &H30101, GP_PF_4bppIndexed = &H30402, GP_PF_8bppIndexed = &H30803, GP_PF_16bppGreyscale = &H101004, GP_PF_16bppRGB555 = &H21005, GP_PF_16bppRGB565 = &H21006
    Private Const GP_PF_16bppARGB1555 = &H61007, GP_PF_24bppRGB = &H21808, GP_PF_32bppRGB = &H22009, GP_PF_32bppARGB = &H26200A, GP_PF_32bppPARGB = &HE200B, GP_PF_48bppRGB = &H10300C, GP_PF_64bppARGB = &H34400D, GP_PF_64bppPARGB = &H1C400E
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

'Property tags describe image metadata.  Metadata is very complicated to read and/or write, because tags are encoded
' in a variety of ways.  Refer to https://msdn.microsoft.com/en-us/library/ms534416(v=vs.85).aspx for details.
' pd2D uses these sparingly; do not expect it to perform full metadata preservation.
Public Enum GP_PropertyTag
    GP_PT_Artist = &H13B
    GP_PT_BitsPerSample = &H102
    GP_PT_CellHeight = &H109
    GP_PT_CellWidth = &H108
    GP_PT_ChrominanceTable = &H5091
    GP_PT_ColorMap = &H140
    GP_PT_ColorTransferFunction = &H501A
    GP_PT_Compression = &H103
    GP_PT_Copyright = &H8298
    GP_PT_DateTime = &H132
    GP_PT_DocumentName = &H10D
    GP_PT_DotRange = &H150
    GP_PT_EquipMake = &H10F
    GP_PT_EquipModel = &H110
    GP_PT_ExifAperture = &H9202
    GP_PT_ExifBrightness = &H9203
    GP_PT_ExifCfaPattern = &HA302
    GP_PT_ExifColorSpace = &HA001
    GP_PT_ExifCompBPP = &H9102
    GP_PT_ExifCompConfig = &H9101
    GP_PT_ExifDTDigitized = &H9004
    GP_PT_ExifDTDigSS = &H9292
    GP_PT_ExifDTOrig = &H9003
    GP_PT_ExifDTOrigSS = &H9291
    GP_PT_ExifDTSubsec = &H9290
    GP_PT_ExifExposureBias = &H9204
    GP_PT_ExifExposureIndex = &HA215
    GP_PT_ExifExposureProg = &H8822
    GP_PT_ExifExposureTime = &H829A
    GP_PT_ExifFileSource = &HA300
    GP_PT_ExifFlash = &H9209
    GP_PT_ExifFlashEnergy = &HA20B
    GP_PT_ExifFNumber = &H829D
    GP_PT_ExifFocalLength = &H920A
    GP_PT_ExifFocalResUnit = &HA210
    GP_PT_ExifFocalXRes = &HA20E
    GP_PT_ExifFocalYRes = &HA20F
    GP_PT_ExifFPXVer = &HA000
    GP_PT_ExifIFD = &H8769
    GP_PT_ExifInterop = &HA005
    GP_PT_ExifISOSpeed = &H8827
    GP_PT_ExifLightSource = &H9208
    GP_PT_ExifMakerNote = &H927C
    GP_PT_ExifMaxAperture = &H9205
    GP_PT_ExifMeteringMode = &H9207
    GP_PT_ExifOECF = &H8828
    GP_PT_ExifPixXDim = &HA002
    GP_PT_ExifPixYDim = &HA003
    GP_PT_ExifRelatedWav = &HA004
    GP_PT_ExifSceneType = &HA301
    GP_PT_ExifSensingMethod = &HA217
    GP_PT_ExifShutterSpeed = &H9201
    GP_PT_ExifSpatialFR = &HA20C
    GP_PT_ExifSpectralSense = &H8824
    GP_PT_ExifSubjectDist = &H9206
    GP_PT_ExifSubjectLoc = &HA214
    GP_PT_ExifUserComment = &H9286
    GP_PT_ExifVer = &H9000
    GP_PT_ExtraSamples = &H152
    GP_PT_FillOrder = &H10A
    GP_PT_FrameDelay = &H5100
    GP_PT_FreeByteCounts = &H121
    GP_PT_FreeOffset = &H120
    GP_PT_Gamma = &H301
    GP_PT_GlobalPalette = &H5102
    GP_PT_GpsAltitude = &H6
    GP_PT_GpsAltitudeRef = &H5
    GP_PT_GpsDestBear = &H18
    GP_PT_GpsDestBearRef = &H17
    GP_PT_GpsDestDist = &H1A
    GP_PT_GpsDestDistRef = &H19
    GP_PT_GpsDestLat = &H14
    GP_PT_GpsDestLatRef = &H13
    GP_PT_GpsDestLong = &H16
    GP_PT_GpsDestLongRef = &H15
    GP_PT_GpsGpsDop = &HB
    GP_PT_GpsGpsMeasureMode = &HA
    GP_PT_GpsGpsSatellites = &H8
    GP_PT_GpsGpsStatus = &H9
    GP_PT_GpsGpsTime = &H7
    GP_PT_GpsIFD = &H8825
    GP_PT_GpsImgDir = &H11
    GP_PT_GpsImgDirRef = &H10
    GP_PT_GpsLatitude = &H2
    GP_PT_GpsLatitudeRef = &H1
    GP_PT_GpsLongitude = &H4
    GP_PT_GpsLongitudeRef = &H3
    GP_PT_GpsMapDatum = &H12
    GP_PT_GpsSpeed = &HD
    GP_PT_GpsSpeedRef = &HC
    GP_PT_GpsTrack = &HF
    GP_PT_GpsTrackRef = &HE
    GP_PT_GpsVer = &H0
    GP_PT_GrayResponseCurve = &H123
    GP_PT_GrayResponseUnit = &H122
    GP_PT_GridSize = &H5011
    GP_PT_HalftoneDegree = &H500C
    GP_PT_HalftoneHints = &H141
    GP_PT_HalftoneLPI = &H500A
    GP_PT_HalftoneLPIUnit = &H500B
    GP_PT_HalftoneMisc = &H500E
    GP_PT_HalftoneScreen = &H500F
    GP_PT_HalftoneShape = &H500D
    GP_PT_HostComputer = &H13C
    GP_PT_ICCProfile = &H8773
    GP_PT_ICCProfileDescriptor = &H302
    GP_PT_ImageDescription = &H10E
    GP_PT_ImageHeight = &H101
    GP_PT_ImageTitle = &H320
    GP_PT_ImageWidth = &H100
    GP_PT_IndexBackground = &H5103
    GP_PT_IndexTransparent = &H5104
    GP_PT_InkNames = &H14D
    GP_PT_InkSet = &H14C
    GP_PT_JPEGACTables = &H209
    GP_PT_JPEGDCTables = &H208
    GP_PT_JPEGInterFormat = &H201
    GP_PT_JPEGInterLength = &H202
    GP_PT_JPEGLosslessPredictors = &H205
    GP_PT_JPEGPointTransforms = &H206
    GP_PT_JPEGProc = &H200
    GP_PT_JPEGQTables = &H207
    GP_PT_JPEGQuality = &H5010
    GP_PT_JPEGRestartInterval = &H203
    GP_PT_LoopCount = &H5101
    GP_PT_LuminanceTable = &H5090
    GP_PT_MaxSampleValue = &H119
    GP_PT_MinSampleValue = &H118
    GP_PT_NewSubfileType = &HFE
    GP_PT_NumberOfInks = &H14E
    GP_PT_Orientation = &H112
    GP_PT_PageName = &H11D
    GP_PT_PageNumber = &H129
    GP_PT_PaletteHistogram = &H5113
    GP_PT_PhotometricInterp = &H106
    GP_PT_PixelPerUnitX = &H5111
    GP_PT_PixelPerUnitY = &H5112
    GP_PT_PixelUnit = &H5110
    GP_PT_PlanarConfig = &H11C
    GP_PT_Predictor = &H13D
    GP_PT_PrimaryChromaticities = &H13F
    GP_PT_PrintFlags = &H5005
    GP_PT_PrintFlagsBleedWidth = &H5008
    GP_PT_PrintFlagsBleedWidthScale = &H5009
    GP_PT_PrintFlagsCrop = &H5007
    GP_PT_PrintFlagsVersion = &H5006
    GP_PT_REFBlackWhite = &H214
    GP_PT_ResolutionUnit = &H128
    GP_PT_ResolutionXLengthUnit = &H5003
    GP_PT_ResolutionXUnit = &H5001
    GP_PT_ResolutionYLengthUnit = &H5004
    GP_PT_ResolutionYUnit = &H5002
    GP_PT_RowsPerStrip = &H116
    GP_PT_SampleFormat = &H153
    GP_PT_SamplesPerPixel = &H115
    GP_PT_SMaxSampleValue = &H155
    GP_PT_SMinSampleValue = &H154
    GP_PT_SoftwareUsed = &H131
    GP_PT_SRGBRenderingIntent = &H303
    GP_PT_StripBytesCount = &H117
    GP_PT_StripOffsets = &H111
    GP_PT_SubfileType = &HFF
    GP_PT_T4Option = &H124
    GP_PT_T6Option = &H125
    GP_PT_TargetPrinter = &H151
    GP_PT_ThreshHolding = &H107
    GP_PT_ThumbnailArtist = &H5034
    GP_PT_ThumbnailBitsPerSample = &H5022
    GP_PT_ThumbnailColorDepth = &H5015
    GP_PT_ThumbnailCompressedSize = &H5019
    GP_PT_ThumbnailCompression = &H5023
    GP_PT_ThumbnailCopyRight = &H503B
    GP_PT_ThumbnailData = &H501B
    GP_PT_ThumbnailDateTime = &H5033
    GP_PT_ThumbnailEquipMake = &H5026
    GP_PT_ThumbnailEquipModel = &H5027
    GP_PT_ThumbnailFormat = &H5012
    GP_PT_ThumbnailHeight = &H5014
    GP_PT_ThumbnailImageDescription = &H5025
    GP_PT_ThumbnailImageHeight = &H5021
    GP_PT_ThumbnailImageWidth = &H5020
    GP_PT_ThumbnailOrientation = &H5029
    GP_PT_ThumbnailPhotometricInterp = &H5024
    GP_PT_ThumbnailPlanarConfig = &H502F
    GP_PT_ThumbnailPlanes = &H5016
    GP_PT_ThumbnailPrimaryChromaticities = &H5036
    GP_PT_ThumbnailRawBytes = &H5017
    GP_PT_ThumbnailRefBlackWhite = &H503A
    GP_PT_ThumbnailResolutionUnit = &H5030
    GP_PT_ThumbnailResolutionX = &H502D
    GP_PT_ThumbnailResolutionY = &H502E
    GP_PT_ThumbnailRowsPerStrip = &H502B
    GP_PT_ThumbnailSamplesPerPixel = &H502A
    GP_PT_ThumbnailSize = &H5018
    GP_PT_ThumbnailSoftwareUsed = &H5032
    GP_PT_ThumbnailStripBytesCount = &H502C
    GP_PT_ThumbnailStripOffsets = &H5028
    GP_PT_ThumbnailTransferFunction = &H5031
    GP_PT_ThumbnailWhitePoint = &H5035
    GP_PT_ThumbnailWidth = &H5013
    GP_PT_ThumbnailYCbCrCoefficients = &H5037
    GP_PT_ThumbnailYCbCrPositioning = &H5039
    GP_PT_ThumbnailYCbCrSubsampling = &H5038
    GP_PT_TileByteCounts = &H145
    GP_PT_TileLength = &H143
    GP_PT_TileOffset = &H144
    GP_PT_TileWidth = &H142
    GP_PT_TransferFunction = &H12D
    GP_PT_TransferRange = &H156
    GP_PT_WhitePoint = &H13E
    GP_PT_XPosition = &H11E
    GP_PT_XResolution = &H11A
    GP_PT_YCbCrCoefficients = &H211
    GP_PT_YCbCrPositioning = &H213
    GP_PT_YCbCrSubsampling = &H212
    GP_PT_YPosition = &H11F
    GP_PT_YResolution = &H11B
End Enum

#If False Then
    Private Const GP_PT_Artist = &H13B, GP_PT_BitsPerSample = &H102, GP_PT_CellHeight = &H109, GP_PT_CellWidth = &H108, GP_PT_ChrominanceTable = &H5091, GP_PT_ColorMap = &H140, GP_PT_ColorTransferFunction = &H501A, GP_PT_Compression = &H103, GP_PT_Copyright = &H8298, GP_PT_DateTime = &H132, GP_PT_DocumentName = &H10D, GP_PT_DotRange = &H150, GP_PT_EquipMake = &H10F, GP_PT_EquipModel = &H110, GP_PT_ExifAperture = &H9202, GP_PT_ExifBrightness = &H9203, GP_PT_ExifCfaPattern = &HA302, GP_PT_ExifColorSpace = &HA001
    Private Const GP_PT_ExifCompBPP = &H9102, GP_PT_ExifCompConfig = &H9101, GP_PT_ExifDTDigitized = &H9004, GP_PT_ExifDTDigSS = &H9292, GP_PT_ExifDTOrig = &H9003, GP_PT_ExifDTOrigSS = &H9291, GP_PT_ExifDTSubsec = &H9290, GP_PT_ExifExposureBias = &H9204, GP_PT_ExifExposureIndex = &HA215, GP_PT_ExifExposureProg = &H8822, GP_PT_ExifExposureTime = &H829A, GP_PT_ExifFileSource = &HA300, GP_PT_ExifFlash = &H9209, GP_PT_ExifFlashEnergy = &HA20B, GP_PT_ExifFNumber = &H829D, GP_PT_ExifFocalLength = &H920A
    Private Const GP_PT_ExifFocalResUnit = &HA210, GP_PT_ExifFocalXRes = &HA20E, GP_PT_ExifFocalYRes = &HA20F, GP_PT_ExifFPXVer = &HA000, GP_PT_ExifIFD = &H8769, GP_PT_ExifInterop = &HA005, GP_PT_ExifISOSpeed = &H8827, GP_PT_ExifLightSource = &H9208, GP_PT_ExifMakerNote = &H927C, GP_PT_ExifMaxAperture = &H9205, GP_PT_ExifMeteringMode = &H9207, GP_PT_ExifOECF = &H8828, GP_PT_ExifPixXDim = &HA002, GP_PT_ExifPixYDim = &HA003, GP_PT_ExifRelatedWav = &HA004, GP_PT_ExifSceneType = &HA301
    Private Const GP_PT_ExifSensingMethod = &HA217, GP_PT_ExifShutterSpeed = &H9201, GP_PT_ExifSpatialFR = &HA20C, GP_PT_ExifSpectralSense = &H8824, GP_PT_ExifSubjectDist = &H9206, GP_PT_ExifSubjectLoc = &HA214, GP_PT_ExifUserComment = &H9286, GP_PT_ExifVer = &H9000, GP_PT_ExtraSamples = &H152, GP_PT_FillOrder = &H10A, GP_PT_FrameDelay = &H5100, GP_PT_FreeByteCounts = &H121, GP_PT_FreeOffset = &H120, GP_PT_Gamma = &H301, GP_PT_GlobalPalette = &H5102, GP_PT_GpsAltitude = &H6
    Private Const GP_PT_GpsAltitudeRef = &H5, GP_PT_GpsDestBear = &H18, GP_PT_GpsDestBearRef = &H17, GP_PT_GpsDestDist = &H1A, GP_PT_GpsDestDistRef = &H19, GP_PT_GpsDestLat = &H14, GP_PT_GpsDestLatRef = &H13, GP_PT_GpsDestLong = &H16, GP_PT_GpsDestLongRef = &H15, GP_PT_GpsGpsDop = &HB, GP_PT_GpsGpsMeasureMode = &HA, GP_PT_GpsGpsSatellites = &H8, GP_PT_GpsGpsStatus = &H9, GP_PT_GpsGpsTime = &H7, GP_PT_GpsIFD = &H8825, GP_PT_GpsImgDir = &H11, GP_PT_GpsImgDirRef = &H10, GP_PT_GpsLatitude = &H2
    Private Const GP_PT_GpsLatitudeRef = &H1, GP_PT_GpsLongitude = &H4, GP_PT_GpsLongitudeRef = &H3, GP_PT_GpsMapDatum = &H12, GP_PT_GpsSpeed = &HD, GP_PT_GpsSpeedRef = &HC, GP_PT_GpsTrack = &HF, GP_PT_GpsTrackRef = &HE, GP_PT_GpsVer = &H0, GP_PT_GrayResponseCurve = &H123, GP_PT_GrayResponseUnit = &H122, GP_PT_GridSize = &H5011, GP_PT_HalftoneDegree = &H500C, GP_PT_HalftoneHints = &H141, GP_PT_HalftoneLPI = &H500A, GP_PT_HalftoneLPIUnit = &H500B, GP_PT_HalftoneMisc = &H500E, GP_PT_HalftoneScreen = &H500F
    Private Const GP_PT_HalftoneShape = &H500D, GP_PT_HostComputer = &H13C, GP_PT_ICCProfile = &H8773, GP_PT_ICCProfileDescriptor = &H302, GP_PT_ImageDescription = &H10E, GP_PT_ImageHeight = &H101, GP_PT_ImageTitle = &H320, GP_PT_ImageWidth = &H100, GP_PT_IndexBackground = &H5103, GP_PT_IndexTransparent = &H5104, GP_PT_InkNames = &H14D, GP_PT_InkSet = &H14C, GP_PT_JPEGACTables = &H209, GP_PT_JPEGDCTables = &H208, GP_PT_JPEGInterFormat = &H201, GP_PT_JPEGInterLength = &H202, GP_PT_JPEGLosslessPredictors = &H205
    Private Const GP_PT_JPEGPointTransforms = &H206, GP_PT_JPEGProc = &H200, GP_PT_JPEGQTables = &H207, GP_PT_JPEGQuality = &H5010, GP_PT_JPEGRestartInterval = &H203, GP_PT_LoopCount = &H5101, GP_PT_LuminanceTable = &H5090, GP_PT_MaxSampleValue = &H119, GP_PT_MinSampleValue = &H118, GP_PT_NewSubfileType = &HFE, GP_PT_NumberOfInks = &H14E, GP_PT_Orientation = &H112, GP_PT_PageName = &H11D, GP_PT_PageNumber = &H129, GP_PT_PaletteHistogram = &H5113, GP_PT_PhotometricInterp = &H106, GP_PT_PixelPerUnitX = &H5111
    Private Const GP_PT_PixelPerUnitY = &H5112, GP_PT_PixelUnit = &H5110, GP_PT_PlanarConfig = &H11C, GP_PT_Predictor = &H13D, GP_PT_PrimaryChromaticities = &H13F, GP_PT_PrintFlags = &H5005, GP_PT_PrintFlagsBleedWidth = &H5008, GP_PT_PrintFlagsBleedWidthScale = &H5009, GP_PT_PrintFlagsCrop = &H5007, GP_PT_PrintFlagsVersion = &H5006, GP_PT_REFBlackWhite = &H214, GP_PT_ResolutionUnit = &H128, GP_PT_ResolutionXLengthUnit = &H5003, GP_PT_ResolutionXUnit = &H5001, GP_PT_ResolutionYLengthUnit = &H5004
    Private Const GP_PT_ResolutionYUnit = &H5002, GP_PT_RowsPerStrip = &H116, GP_PT_SampleFormat = &H153, GP_PT_SamplesPerPixel = &H115, GP_PT_SMaxSampleValue = &H155, GP_PT_SMinSampleValue = &H154, GP_PT_SoftwareUsed = &H131, GP_PT_SRGBRenderingIntent = &H303, GP_PT_StripBytesCount = &H117, GP_PT_StripOffsets = &H111, GP_PT_SubfileType = &HFF, GP_PT_T4Option = &H124, GP_PT_T6Option = &H125, GP_PT_TargetPrinter = &H151, GP_PT_ThreshHolding = &H107, GP_PT_ThumbnailArtist = &H5034, GP_PT_ThumbnailBitsPerSample = &H5022
    Private Const GP_PT_ThumbnailColorDepth = &H5015, GP_PT_ThumbnailCompressedSize = &H5019, GP_PT_ThumbnailCompression = &H5023, GP_PT_ThumbnailCopyRight = &H503B, GP_PT_ThumbnailData = &H501B, GP_PT_ThumbnailDateTime = &H5033, GP_PT_ThumbnailEquipMake = &H5026, GP_PT_ThumbnailEquipModel = &H5027, GP_PT_ThumbnailFormat = &H5012, GP_PT_ThumbnailHeight = &H5014, GP_PT_ThumbnailImageDescription = &H5025, GP_PT_ThumbnailImageHeight = &H5021, GP_PT_ThumbnailImageWidth = &H5020, GP_PT_ThumbnailOrientation = &H5029, GP_PT_ThumbnailPhotometricInterp = &H5024
    Private Const GP_PT_ThumbnailPlanarConfig = &H502F, GP_PT_ThumbnailPlanes = &H5016, GP_PT_ThumbnailPrimaryChromaticities = &H5036, GP_PT_ThumbnailRawBytes = &H5017, GP_PT_ThumbnailRefBlackWhite = &H503A, GP_PT_ThumbnailResolutionUnit = &H5030, GP_PT_ThumbnailResolutionX = &H502D, GP_PT_ThumbnailResolutionY = &H502E, GP_PT_ThumbnailRowsPerStrip = &H502B, GP_PT_ThumbnailSamplesPerPixel = &H502A, GP_PT_ThumbnailSize = &H5018, GP_PT_ThumbnailSoftwareUsed = &H5032, GP_PT_ThumbnailStripBytesCount = &H502C, GP_PT_ThumbnailStripOffsets = &H5028
    Private Const GP_PT_ThumbnailTransferFunction = &H5031, GP_PT_ThumbnailWhitePoint = &H5035, GP_PT_ThumbnailWidth = &H5013, GP_PT_ThumbnailYCbCrCoefficients = &H5037, GP_PT_ThumbnailYCbCrPositioning = &H5039, GP_PT_ThumbnailYCbCrSubsampling = &H5038, GP_PT_TileByteCounts = &H145, GP_PT_TileLength = &H143, GP_PT_TileOffset = &H144, GP_PT_TileWidth = &H142, GP_PT_TransferFunction = &H12D, GP_PT_TransferRange = &H156, GP_PT_WhitePoint = &H13E, GP_PT_XPosition = &H11E, GP_PT_XResolution = &H11A, GP_PT_YCbCrCoefficients = &H211
    Private Const GP_PT_YCbCrPositioning = &H213, GP_PT_YCbCrSubsampling = &H212, GP_PT_YPosition = &H11F, GP_PT_YResolution = &H11B
#End If

Public Enum GP_RotateFlip
    GP_RF_NoneFlipNone = 0
    GP_RF_90FlipNone = 1
    GP_RF_180FlipNone = 2
    GP_RF_270FlipNone = 3
    GP_RF_NoneFlipX = 4
    GP_RF_90FlipX = 5
    GP_RF_180FlipX = 6
    GP_RF_270FlipX = 7
    GP_RF_NoneFlipY = GP_RF_180FlipX
    GP_RF_90FlipY = GP_RF_270FlipX
    GP_RF_180FlipY = GP_RF_NoneFlipX
    GP_RF_270FlipY = GP_RF_90FlipX
    GP_RF_NoneFlipXY = GP_RF_180FlipNone
    GP_RF_90FlipXY = GP_RF_270FlipNone
    GP_RF_180FlipXY = GP_RF_NoneFlipNone
    GP_RF_270FlipXY = GP_RF_90FlipNone
End Enum

#If False Then
    Private Const GP_RF_NoneFlipNone = 0, GP_RF_90FlipNone = 1, GP_RF_180FlipNone = 2, GP_RF_270FlipNone = 3, GP_RF_NoneFlipX = 4, GP_RF_90FlipX = 5, GP_RF_180FlipX = 6, GP_RF_270FlipX = 7, GP_RF_NoneFlipY = GP_RF_180FlipX
    Private Const GP_RF_90FlipY = GP_RF_270FlipX, GP_RF_180FlipY = GP_RF_NoneFlipX, GP_RF_270FlipY = GP_RF_90FlipX, GP_RF_NoneFlipXY = GP_RF_180FlipNone, GP_RF_90FlipXY = GP_RF_270FlipNone, GP_RF_180FlipXY = GP_RF_NoneFlipNone, GP_RF_270FlipXY = GP_RF_90FlipNone
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

'GDI+ uses a modified bitmap struct when performing things like raster format conversions
Public Type GP_BitmapData
    BD_Width As Long
    BD_Height As Long
    BD_Stride As Long
    BD_PixelFormat As GP_PixelFormat
    BD_Scan0 As Long
    BD_Reserved As Long
End Type

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

'Exporting images via GDI+ is a big headache.  A number of convoluted structs are required if the user
' wants to custom-set any image properties.
Private Type GP_EncoderParameter
    EP_GUID(0 To 15) As Byte
    EP_NumOfValues As Long
    EP_ValueType As GP_EncoderValueType
    EP_ValuePtr As Long
End Type

Private Type GP_EncoderParameters
    EP_Count As Long
    EP_Parameter As GP_EncoderParameter
End Type

Private Type GP_ImageCodecInfo
    IC_ClassID(0 To 15) As Byte
    IC_FormatID(0 To 15) As Byte
    IC_CodecName As Long
    IC_DllName As Long
    IC_FormatDescription As Long
    IC_FilenameExtension As Long
    IC_MimeType As Long
    IC_Flags As Long
    IC_Version As Long
    IC_SigCount As Long
    IC_SigSize As Long
    IC_SigPattern As Long
    IC_SigMask As Long
End Type

'GDI+ uses GUIDs to define image formats.  VB6 doesn't let us predeclare byte arrays (at least not easily),
' so we save ourselves the trouble and just use string versions.
Private Const GP_FF_GUID_Undefined = "{B96B3CA9-0728-11D3-9D7B-0000F81EF32E}"
Private Const GP_FF_GUID_MemoryBMP = "{B96B3CAA-0728-11D3-9D7B-0000F81EF32E}"
Private Const GP_FF_GUID_BMP = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}"
Private Const GP_FF_GUID_EMF = "{B96B3CAC-0728-11D3-9D7B-0000F81EF32E}"
Private Const GP_FF_GUID_WMF = "{B96B3CAD-0728-11D3-9D7B-0000F81EF32E}"
Private Const GP_FF_GUID_JPEG = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
Private Const GP_FF_GUID_PNG = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
Private Const GP_FF_GUID_GIF = "{B96B3CB0-0728-11D3-9D7B-0000F81EF32E}"
Private Const GP_FF_GUID_TIFF = "{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}"
Private Const GP_FF_GUID_EXIF = "{B96B3CB2-0728-11D3-9D7B-0000F81EF32E}"
Private Const GP_FF_GUID_Icon = "{B96B3CB5-0728-11D3-9D7B-0000F81EF32E}"

'Like image formats, export encoder properties are also defined by GUID.  These values come from the Win 8.1
' version of gdiplusimaging.h.  Note that some are restricted to GDI+ v1.1.
Private Const GP_EP_Compression As String = "{E09D739D-CCD4-44EE-8EBA-3FBF8BE4FC58}"
Private Const GP_EP_ColorDepth As String = "{66087055-AD66-4C7C-9A18-38A2310B8337}"
Private Const GP_EP_ScanMethod As String = "{3A4E2661-3109-4E56-8536-42C156E7DCFA}"
Private Const GP_EP_Version As String = "{24D18C76-814A-41A4-BF53-1C219CCCF797}"
Private Const GP_EP_RenderMethod As String = "{6D42C53A-229A-4825-8BB7-5C99E2B9A8B8}"
Private Const GP_EP_Quality As String = "{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"
Private Const GP_EP_Transformation As String = "{8D0EB2D1-A58E-4EA8-AA14-108074B7B6F9}"
Private Const GP_EP_LuminanceTable As String = "{EDB33BCE-0266-4A77-B904-27216099E717}"
Private Const GP_EP_ChrominanceTable As String = "{F2E455DC-09B3-4316-8260-676ADA32481C}"
Private Const GP_EP_SaveFlag As String = "{292266FC-AC40-47BF-8CFC-A85B89A655DE}"

'REQUIRES GDI+ v1.1 OR LATER!
Private Const GP_EP_ColorSpace As String = "{AE7A62A0-EE2C-49D8-9D07-1BA8A927596E}"
Private Const GP_EP_SaveAsCMYK As String = "{A219BBC9-0A9D-4005-A3EE-3A421B8BB06C}"

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

Private Declare Function GdipBitmapLockBits Lib "gdiplus" (ByVal hImage As Long, ByRef srcRect As RECTL, ByVal lockMode As GP_BitmapLockMode, ByVal dstPixelFormat As GP_PixelFormat, ByRef srcBitmapData As GP_BitmapData) As GP_Result
Private Declare Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal hImage As Long, ByRef srcBitmapData As GP_BitmapData) As GP_Result

Private Declare Function GdipCloneBitmapAreaI Lib "gdiplus" (ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal newPixelFormat As GP_PixelFormat, ByVal hSrcBitmap As Long, ByRef hDstBitmap As Long) As GP_Result
Private Declare Function GdipCloneMatrix Lib "gdiplus" (ByVal srcMatrix As Long, ByRef dstMatrix As Long) As GP_Result
Private Declare Function GdipClonePath Lib "gdiplus" (ByVal srcPath As Long, ByRef dstPath As Long) As GP_Result
Private Declare Function GdipCloneRegion Lib "gdiplus" (ByVal srcRegion As Long, ByRef dstRegion As Long) As GP_Result
Private Declare Function GdipClosePathFigure Lib "gdiplus" (ByVal hPath As Long) As GP_Result

Private Declare Function GdipCombineRegionRect Lib "gdiplus" (ByVal hRegion As Long, ByRef srcRectF As RECTF, ByVal dstCombineMode As GP_CombineMode) As GP_Result
Private Declare Function GdipCombineRegionRectI Lib "gdiplus" (ByVal hRegion As Long, ByRef srcRectL As RECTL, ByVal dstCombineMode As GP_CombineMode) As GP_Result
Private Declare Function GdipCombineRegionRegion Lib "gdiplus" (ByVal dstRegion As Long, ByVal srcRegion As Long, ByVal dstCombineMode As GP_CombineMode) As GP_Result
Private Declare Function GdipCombineRegionPath Lib "gdiplus" (ByVal dstRegion As Long, ByVal srcPath As Long, ByVal dstCombineMode As GP_CombineMode) As GP_Result

'This EMF convert function only works on Vista+!
Private Declare Function GdipConvertToEmfPlus Lib "gdiplus" (ByVal hGraphics As Long, ByVal srcMetafile As Long, ByRef conversionSuccess As Long, ByVal typeOfEMF As GP_MetafileType, ByVal ptrToMetafileDescription As Long, ByRef dstMetafilePtr As Long) As GP_Result

Private Declare Function GdipCreateBitmapFromGdiDib Lib "gdiplus" (ByRef origGDIBitmapInfo As BITMAPINFO, ByRef srcBitmapData As Any, ByRef dstGdipBitmap As Long) As GP_Result
Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal bmpWidth As Long, ByVal bmpHeight As Long, ByVal bmpStride As Long, ByVal bmpPixelFormat As GP_PixelFormat, ByRef Scan0 As Any, ByRef dstGdipBitmap As Long) As GP_Result
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
Private Declare Function GdipGetImageEncoders Lib "gdiplus" (ByVal numOfEncoders As Long, ByVal sizeOfEncoders As Long, ByVal ptrToDstEncoders As Long) As GP_Result
Private Declare Function GdipGetImageEncodersSize Lib "gdiplus" (ByRef numOfEncoders As Long, ByRef sizeOfEncoders As Long) As GP_Result
Private Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal hImage As Long, ByRef dstHeight As Long) As GP_Result
Private Declare Function GdipGetImageHorizontalResolution Lib "gdiplus" (ByVal hImage As Long, ByRef dstHResolution As Single) As GP_Result
Private Declare Function GdipGetImagePixelFormat Lib "gdiplus" (ByVal hImage As Long, ByRef dstPixelFormat As GP_PixelFormat) As GP_Result
Private Declare Function GdipGetImageType Lib "gdiplus" (ByVal srcImage As Long, ByRef dstImageType As GP_ImageType) As GP_Result
Private Declare Function GdipGetImageVerticalResolution Lib "gdiplus" (ByVal hImage As Long, ByRef dstVResolution As Single) As GP_Result
Private Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal hImage As Long, ByRef dstWidth As Long) As GP_Result
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
Private Declare Function GdipGetPropertyItem Lib "gdiplus" (ByVal hImage As Long, ByVal gpPropertyID As Long, ByVal srcPropertySize As Long, ByVal ptrToDstBuffer As Long) As GP_Result
Private Declare Function GdipGetPropertyItemSize Lib "gdiplus" (ByVal hImage As Long, ByVal gpPropertyID As Long, ByRef dstPropertySize As Long) As GP_Result
Private Declare Function GdipGetRegionBounds Lib "gdiplus" (ByVal hRegion As Long, ByVal hGraphics As Long, ByRef dstRectF As RECTF) As GP_Result
Private Declare Function GdipGetRegionBoundsI Lib "gdiplus" (ByVal hRegion As Long, ByVal hGraphics As Long, ByRef dstRectL As RECTL) As GP_Result
Private Declare Function GdipGetRenderingOrigin Lib "gdiplus" (ByVal hGraphics As Long, ByRef dstX As Long, ByRef dstY As Long) As GP_Result
Private Declare Function GdipGetSmoothingMode Lib "gdiplus" (ByVal hGraphics As Long, ByRef dstMode As GP_SmoothingMode) As GP_Result
Private Declare Function GdipGetSolidFillColor Lib "gdiplus" (ByVal hBrush As Long, ByRef dstColor As Long) As GP_Result
Private Declare Function GdipGetTextureWrapMode Lib "gdiplus" (ByVal hBrush As Long, ByRef dstWrapMode As GP_WrapMode) As GP_Result

Private Declare Function GdipGetImageRawFormat Lib "gdiplus" (ByVal hImage As Long, ByVal ptrToDstGuid As Long) As GP_Result
Private Declare Function GdipImageRotateFlip Lib "gdiplus" (ByVal hImage As Long, ByVal rotateFlipType As GP_RotateFlip) As GP_Result
Private Declare Function GdipInvertMatrix Lib "gdiplus" (ByVal hMatrix As Long) As GP_Result

Private Declare Function GdipIsEmptyRegion Lib "gdiplus" (ByVal srcRegion As Long, ByVal srcGraphics As Long, ByRef dstResult As Long) As GP_Result
Private Declare Function GdipIsEqualRegion Lib "gdiplus" (ByVal srcRegion1 As Long, ByVal srcRegion2 As Long, ByVal srcGraphics As Long, ByRef dstResult As Long) As GP_Result
Private Declare Function GdipIsInfiniteRegion Lib "gdiplus" (ByVal srcRegion As Long, ByVal srcGraphics As Long, ByRef dstResult As Long) As GP_Result
Private Declare Function GdipIsMatrixInvertible Lib "gdiplus" (ByVal hMatrix As Long, ByRef dstResult As Long) As Long
Private Declare Function GdipIsOutlineVisiblePathPoint Lib "gdiplus" (ByVal hPath As Long, ByVal x As Single, ByVal y As Single, ByVal hPen As Long, ByVal hGraphicsOptional As Long, ByRef dstResult As Long) As GP_Result
Private Declare Function GdipIsOutlineVisiblePathPointI Lib "gdiplus" (ByVal hPath As Long, ByVal x As Long, ByVal y As Long, ByVal hPen As Long, ByVal hGraphicsOptional As Long, ByRef dstResult As Long) As GP_Result
Private Declare Function GdipIsVisiblePathPoint Lib "gdiplus" (ByVal hPath As Long, ByVal x As Single, ByVal y As Single, ByVal hGraphicsOptional As Long, ByRef dstResult As Long) As GP_Result
Private Declare Function GdipIsVisiblePathPointI Lib "gdiplus" (ByVal hPath As Long, ByVal x As Long, ByVal y As Long, ByVal hGraphicsOptional As Long, ByRef dstResult As Long) As GP_Result

Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal ptrSrcFilename As Long, ByRef dstGdipImage As Long) As GP_Result
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal srcIStream As Long, ByRef dstGdipImage As Long) As GP_Result

Private Declare Function GdipResetClip Lib "gdiplus" (ByVal hGraphics As Long) As GP_Result
Private Declare Function GdipResetPath Lib "gdiplus" (ByVal hPath As Long) As GP_Result
Private Declare Function GdipRotateMatrix Lib "gdiplus" (ByVal hMatrix As Long, ByVal rotateAngle As Single, ByVal mOrder As GP_MatrixOrder) As GP_Result

Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal hImage As Long, ByVal ptrToFilename As Long, ByVal ptrToEncoderGUID As Long, ByVal ptrToEncoderParams As Long) As GP_Result
Private Declare Function GdipSaveImageToStream Lib "gdiplus" (ByVal hImage As Long, ByVal dstIStream As Long, ByVal ptrToEncoderGUID As Long, ByVal ptrToEncoderParams As Long) As GP_Result
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
Private Declare Function CLSIDFromString Lib "ole32" (ByVal ptrToGuidString As Long, ByVal ptrToByteArray As Long) As Long
Private Declare Function CopyMemory_Strict Lib "kernel32" Alias "RtlMoveMemory" (ByVal ptrDst As Long, ByVal ptrSrc As Long, ByVal numOfBytes As Long) As Long
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ByVal ptrToDstStream As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetHGlobalFromStream Lib "ole32" (ByVal srcIStream As Long, ByRef dstHGlobal As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal ptrToString As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal ptrToString As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32" (ByVal oColor As OLE_COLOR, ByVal hPalette As Long, ByRef cColorRef As Long) As Long
Private Declare Function PutMem4 Lib "msvbvm60" (ByVal ptrDst As Long, ByVal newValue As Long) As Long
Private Declare Function StringFromCLSID Lib "ole32" (ByVal ptrToGuid As Long, ByRef ptrToDstString As Long) As Long
Private Declare Function SysAllocString Lib "oleaut32" (ByVal srcWCharPtr As Long) As Long
Private Declare Function SysAllocStringByteLen Lib "oleaut32" (ByVal srcAnsiPtr As Long, ByVal srcLength As Long) As String

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
    If (errNumber <> 0) Then
        tmpString = "WARNING!  Internal GDI+ error #" & errNumber & ", """ & errName & """"
    Else
        tmpString = "WARNING!  GDI+ module error, """ & errName & """"
    End If
    
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
        GetGdipBitmapHandleFromDIB = CBool(GdipCreateBitmapFromScan0(srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, srcDIB.GetDIBWidth * 4, GP_PF_32bppPARGB, ByVal srcDIB.GetDIBPointer, dstBitmapHandle) = GP_OK)
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

Public Function GDIPlus_DrawImagePointsRectF(ByVal dstGraphics As Long, ByVal srcImage As Long, ByRef dstPlgPoints() As POINTFLOAT, ByVal srcX As Single, ByVal srcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, Optional ByVal opacityModifier As Single = 1#) As Boolean
    
    'Modified opacity requires us to create a temporary image attributes object
    Dim imgAttributesHandle As Long
    If (opacityModifier <> 1#) Then
        GdipCreateImageAttributes imgAttributesHandle
        m_AttributesMatrix(3, 3) = opacityModifier
        GdipSetImageAttributesColorMatrix imgAttributesHandle, GP_CAT_Bitmap, 1&, VarPtr(m_AttributesMatrix(0, 0)), 0&, GP_CMF_Default
    End If
    
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawImagePointsRect(dstGraphics, srcImage, VarPtr(dstPlgPoints(0)), 3, srcX, srcY, srcWidth, srcHeight, GP_U_Pixel, imgAttributesHandle, 0&, 0&)
    GDIPlus_DrawImagePointsRectF = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
    
    'As necessary, release our image attributes object, and reset the alpha value of the master identity matrix
    If (opacityModifier <> 1#) Then
        GdipDisposeImageAttributes imgAttributesHandle
        m_AttributesMatrix(3, 3) = 1#
    End If
        
End Function

Public Function GDIPlus_DrawImagePointsRectI(ByVal dstGraphics As Long, ByVal srcImage As Long, ByRef dstPlgPoints() As POINTLONG, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, Optional ByVal opacityModifier As Single = 1#) As Boolean
    
    'Modified opacity requires us to create a temporary image attributes object
    Dim imgAttributesHandle As Long
    If (opacityModifier <> 1#) Then
        GdipCreateImageAttributes imgAttributesHandle
        m_AttributesMatrix(3, 3) = opacityModifier
        GdipSetImageAttributesColorMatrix imgAttributesHandle, GP_CAT_Bitmap, 1&, VarPtr(m_AttributesMatrix(0, 0)), 0&, GP_CMF_Default
    End If
    
    Dim tmpReturn As GP_Result
    tmpReturn = GdipDrawImagePointsRectI(dstGraphics, srcImage, VarPtr(dstPlgPoints(0)), 3, srcX, srcY, srcWidth, srcHeight, GP_U_Pixel, imgAttributesHandle, 0&, 0&)
    GDIPlus_DrawImagePointsRectI = CBool(tmpReturn = GP_OK)
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

'WARNING!  If a graphics object has never specified a clipping region, the default region is infinite.
' For reasons unknown, GDI+ is finicky about returning such a region; it often reports "Object Busy" for no
' apparent reason.  I'm not sure of a good workaround.
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

'Note that this function creates an image from an array containing a valid image file (e.g. not an array with
' bare RGB values).  This is helpful for interop with other software, or if you prefer to roll your own filesystem code.
Public Function GDIPlus_ImageCreateFromArray(ByRef srcArray() As Byte, Optional ByRef isImageMetafile As Boolean = False) As Long
    
    'GDI+ requires a stream object for import, so we're going to wrap a temporary stream around the source array.
    Dim tmpStream As Long
    
    Dim tmpHMem As Long
    Const GMEM_MOVEABLE As Long = &H2&
    tmpHMem = GlobalAlloc(GMEM_MOVEABLE, UBound(srcArray) - LBound(srcArray) + 1)
    If (tmpHMem <> 0) Then
        
        Dim tmpLockMem As Long
        tmpLockMem = GlobalLock(tmpHMem)
        If (tmpLockMem <> 0) Then
            CopyMemory_Strict tmpLockMem, VarPtr(srcArray(LBound(srcArray))), UBound(srcArray) - LBound(srcArray) + 1
            GlobalUnlock tmpHMem
            CreateStreamOnHGlobal tmpHMem, 1&, VarPtr(tmpStream)
        End If
        
    End If
    
    If (tmpStream <> 0) Then
    
        Dim tmpReturn As GP_Result
        tmpReturn = GdipLoadImageFromStream(tmpStream, GDIPlus_ImageCreateFromArray)
        If (tmpReturn = GP_OK) Then
            Dim imgType As GP_ImageType
            GdipGetImageType GDIPlus_ImageCreateFromArray, imgType
            isImageMetafile = CBool(imgType = GP_IT_Metafile)
        Else
            InternalGDIPlusError vbNullString, vbNullString, tmpReturn
        End If
        
    Else
        InternalGDIPlusError "IStream failure", "GDIPlus_ImageCreateFromArray() failed to wrap an IStream around the source array; load aborted."
    End If
    
End Function

Public Function GDIPlus_ImageCreateFromFile(ByVal srcFilename As String, Optional ByRef isImageMetafile As Boolean = False) As Long
    Dim tmpReturn As GP_Result
    tmpReturn = GdipLoadImageFromFile(StrPtr(srcFilename), GDIPlus_ImageCreateFromFile)
    If (tmpReturn = GP_OK) Then
        Dim imgType As GP_ImageType
        GdipGetImageType GDIPlus_ImageCreateFromFile, imgType
        isImageMetafile = CBool(imgType = GP_IT_Metafile)
    Else
        InternalGDIPlusError vbNullString, vbNullString, tmpReturn
    End If
End Function

'This function only works on bitmaps (never metafiles!), and the source image *must* already be in 32-bpp format.
Public Function GDIPlus_ImageForcePremultipliedAlpha(ByVal hImage As Long, ByVal imgWidth As Long, ByVal imgHeight As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipCloneBitmapAreaI(0, 0, imgWidth, imgHeight, GP_PF_32bppPARGB, hImage, hImage)
    GDIPlus_ImageForcePremultipliedAlpha = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

'RANDOM FACT! GdipGetImageDimension works fine on bitmaps.  On metafiles, it returns bizarre values that may be
' astronomically large.  I assume that metafile dimensions are not necessarily returned in pixels (though pixels
' are the default for bitmaps...?).  Anyway, to avoid this problem, we only use GdipGetImageWidth/Height, which
' always return "correct" pixel values.
Public Function GDIPlus_ImageGetDimensions(ByVal hImage As Long, ByRef dstWidth As Long, ByRef dstHeight As Long) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipGetImageWidth(hImage, dstWidth)
    If (tmpReturn = GP_OK) Then
        tmpReturn = GdipGetImageHeight(hImage, dstHeight)
        GDIPlus_ImageGetDimensions = CBool(tmpReturn = GP_OK)
    Else
        GDIPlus_ImageGetDimensions = False
    End If
End Function

Public Function GDIPlus_ImageGetFileFormat(ByVal hImage As Long) As PD_2D_FileFormatImport
    GDIPlus_ImageGetFileFormat = GetPd2dFileFormatFromGUID(GDIPlus_ImageGetFileFormatGUID(hImage))
End Function

Public Function GDIPlus_ImageGetFileFormatGUID(ByVal hImage As Long) As String
    
    Dim tmpReturn As GP_Result
    
    'Start by retrieving the raw bytes of the GUID
    Dim guidBytes() As Byte
    ReDim guidBytes(0 To 15) As Byte
    tmpReturn = GdipGetImageRawFormat(hImage, VarPtr(guidBytes(0)))
    
    If (tmpReturn = GP_OK) Then
    
        'Byte array comparisons against predefined constants are messy in VB, so retrieve a string instead
        Dim imgStringPointer As Long
        If (StringFromCLSID(VarPtr(guidBytes(0)), imgStringPointer) = 0) Then
            Dim strLength As Long
            strLength = lstrlenW(imgStringPointer)
            If (strLength <> 0) Then
                GDIPlus_ImageGetFileFormatGUID = String$(strLength, 48)
                CopyMemory_Strict StrPtr(GDIPlus_ImageGetFileFormatGUID), imgStringPointer, strLength * 2
            End If
        Else
            InternalGDIPlusError "Failed to convert CLSID to string", "GDIPlus_ImageGetFileFormatGUID failed"
        End If
        
    Else
        GDIPlus_ImageGetFileFormatGUID = GP_FF_GUID_Undefined
        InternalGDIPlusError vbNullString, vbNullString, tmpReturn
    End If
    
End Function

'Given a GDI+ GUID format identifier, return a Long-type pd2D file format identifier
Private Function GetPd2dFileFormatFromGUID(ByRef srcGUID As String) As PD_2D_FileFormatImport
    Select Case srcGUID
        Case GP_FF_GUID_BMP, GP_FF_GUID_MemoryBMP
            GetPd2dFileFormatFromGUID = P2_FFI_BMP
        Case GP_FF_GUID_EMF
            GetPd2dFileFormatFromGUID = P2_FFI_EMF
        Case GP_FF_GUID_WMF
            GetPd2dFileFormatFromGUID = P2_FFI_WMF
        Case GP_FF_GUID_JPEG
            GetPd2dFileFormatFromGUID = P2_FFI_JPEG
        Case GP_FF_GUID_PNG
            GetPd2dFileFormatFromGUID = P2_FFI_PNG
        Case GP_FF_GUID_GIF
            GetPd2dFileFormatFromGUID = P2_FFI_GIF
        Case GP_FF_GUID_TIFF
            GetPd2dFileFormatFromGUID = P2_FFI_TIFF
        Case GP_FF_GUID_Icon
            GetPd2dFileFormatFromGUID = P2_FFI_ICO
        Case Else
            GetPd2dFileFormatFromGUID = P2_FFI_Undefined
    End Select
End Function

'Given a pd2D file format, return a matching GDI+ GUID format identifier (as a string; you'll need to manually
' convert this to a byte array, FYI!)
Private Function GetGUIDFromPd2dFileFormat(ByVal srcFileFormat As PD_2D_FileFormatImport) As String
    Select Case srcFileFormat
        Case P2_FFI_BMP
            GetGUIDFromPd2dFileFormat = GP_FF_GUID_BMP
        Case P2_FFI_EMF
            GetGUIDFromPd2dFileFormat = GP_FF_GUID_EMF
        Case P2_FFI_WMF
            GetGUIDFromPd2dFileFormat = GP_FF_GUID_WMF
        Case P2_FFI_JPEG
            GetGUIDFromPd2dFileFormat = GP_FF_GUID_JPEG
        Case P2_FFI_PNG
            GetGUIDFromPd2dFileFormat = GP_FF_GUID_PNG
        Case P2_FFI_GIF
            GetGUIDFromPd2dFileFormat = GP_FF_GUID_GIF
        Case P2_FFI_TIFF
            GetGUIDFromPd2dFileFormat = GP_FF_GUID_TIFF
        Case P2_FFI_ICO
            GetGUIDFromPd2dFileFormat = GP_FF_GUID_Icon
        Case Else
            GetGUIDFromPd2dFileFormat = vbNullString
    End Select
End Function

Public Function GDIPlus_ImageGetPixelFormat(ByVal hImage As Long) As GP_PixelFormat
    Dim tmpReturn As GP_Result
    tmpReturn = GdipGetImagePixelFormat(hImage, GDIPlus_ImageGetPixelFormat)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

'It's important to check the return value of this function; it will be FALSE if the image does not
' contain/provide the requested property.  Also note that all properties are returned as byte arrays.
' It is up to the caller to make sense of this return, presumably using the MSDN guide at
' https://msdn.microsoft.com/en-us/library/ms534416(v=vs.85).aspx
Public Function GDIPlus_ImageGetProperty(ByVal hImage As Long, ByVal gpPropertyID As GP_PropertyTag, ByRef dstBuffer() As Byte) As Boolean
    
    Dim tmpReturn As GP_Result, propSize As Long
    tmpReturn = GdipGetPropertyItemSize(hImage, gpPropertyID, propSize)
    If (tmpReturn = GP_OK) Then
    
        If (propSize > 0) Then
            ReDim dstBuffer(0 To propSize - 1) As Byte
            tmpReturn = GdipGetPropertyItem(hImage, gpPropertyID, propSize, VarPtr(dstBuffer(0)))
            If (tmpReturn = GP_OK) Then
                GDIPlus_ImageGetProperty = True
            Else
                InternalGDIPlusError vbNullString, vbNullString, tmpReturn
                GDIPlus_ImageGetProperty = False
            End If
        Else
            GDIPlus_ImageGetProperty = False
        End If
        
    Else
        'NOTE: it's totally okay for an image to not have a given property.  This is not a meaningful error,
        ' so we do not report it.
        'InternalGDIPlusError vbNullString, vbNullString, tmpReturn
        GDIPlus_ImageGetProperty = False
    End If
    
End Function

Public Function GDIPlus_ImageGetResolution(ByVal hImage As Long, ByRef dstHResolution As Single, ByRef dstVResolution As Single) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipGetImageHorizontalResolution(hImage, dstHResolution)
    If (tmpReturn = GP_OK) Then
        tmpReturn = GdipGetImageVerticalResolution(hImage, dstVResolution)
        GDIPlus_ImageGetResolution = CBool(tmpReturn = GP_OK)
    Else
        GDIPlus_ImageGetResolution = False
    End If
End Function

Public Function GDIPlus_ImageLockBits(ByVal hImage As Long, ByRef srcRect As RECTL, ByRef srcCopyData As GP_BitmapData, ByVal lockFlags As GP_BitmapLockMode, ByVal dstPixelFormat As GP_PixelFormat) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipBitmapLockBits(hImage, srcRect, lockFlags, dstPixelFormat, srcCopyData)
    GDIPlus_ImageLockBits = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

Public Function GDIPlus_ImageRotateFlip(ByVal hImage As Long, ByVal typeOfRotateFlip As GP_RotateFlip) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipImageRotateFlip(hImage, typeOfRotateFlip)
    GDIPlus_ImageRotateFlip = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

'Save a surface to a VB byte array.  The destination array *must* be dynamic, and does not need to be dimensionsed.
' (It will be auto-dimensioned correctly by thsi function.)
' As with saving to file, note that the only export property currently supported is JPEG quality; other properties
' are set automatically by GDI+.
Public Function GDIPlus_ImageSaveToArray(ByVal hImage As Long, ByRef dstArray() As Byte, Optional ByVal dstFileFormat As PD_2D_FileFormatExport = P2_FFE_PNG, Optional ByVal jpegQuality As Long = 85) As Boolean
        
    On Error GoTo GDIPlusSaveError
    
    'GDI+ uses GUIDs to define image export encoders; retrieve the relevant encoder GUID now
    Dim exporterGUID(0 To 15) As Byte
    If GetEncoderGUIDForPd2dFormat(dstFileFormat, VarPtr(exporterGUID(0))) Then
    
        'Like export format, GDI+ also uses GUIDs to define export properties.  If multiple encoder parameters
        ' are in use, these need to be merged into sequential order (because GDI+ only takes a pointer).
        ' pd2D does not currently cover this use-case; it always assumes there are only 0 or 1 parameters in use.
        ' To use multiple parameters, you would need copy the first GP_EncoderParameters entry into the
        ' fullEncoderParams() array, like normal, but with the Count value set to the number of parameters.
        ' Then, you would need to copy subsequent parameters into place *after* it.  (But *only* the parameters,
        ' not additional "Count" values.)
        '
        'Look at PhotoDemon's source code for an example of how to do this.
        Dim paramsInUse As Boolean: paramsInUse = False
        Dim tmpEncoderParams As GP_EncoderParameters, tmpConstString As String
        Dim fullEncoderParams() As Byte
        
        If (dstFileFormat = P2_FFE_JPEG) Then
            
            paramsInUse = True
            
            tmpEncoderParams.EP_Count = 1
            With tmpEncoderParams.EP_Parameter
                .EP_NumOfValues = 1
                .EP_ValueType = GP_EVT_Long
                tmpConstString = GP_EP_Quality
                CLSIDFromString StrPtr(tmpConstString), VarPtr(.EP_GUID(0))
                .EP_ValuePtr = VarPtr(jpegQuality)
            End With
            
        End If
        
        'Prep an IStream to receive the export.  Note that we deliberately mark the stream as "free on release",
        ' which spares us from manually releasing the stream's contents.  (They will be auto-freed when the tmpStream
        ' object goes out of scope.)
        Dim tmpStream As Long
        CreateStreamOnHGlobal 0&, 1&, VarPtr(tmpStream)
        
        'Perform the export
        Dim tmpReturn As GP_Result
        If paramsInUse Then
            tmpReturn = GdipSaveImageToStream(hImage, tmpStream, VarPtr(exporterGUID(0)), VarPtr(tmpEncoderParams))
        Else
            tmpReturn = GdipSaveImageToStream(hImage, tmpStream, VarPtr(exporterGUID(0)), 0&)
        End If
        
        If (tmpReturn = GP_OK) Then
        
            'We now need to copy the contents of the stream into a VB array
            Dim tmpHMem As Long, hMemSize As Long
            If (GetHGlobalFromStream(tmpStream, tmpHMem) = 0) Then
                hMemSize = GlobalSize(tmpHMem)
                If (hMemSize <> 0) Then
                
                    Dim lockedMem As Long
                    lockedMem = GlobalLock(tmpHMem)
                    If (lockedMem <> 0) Then
                        ReDim dstArray(0 To hMemSize - 1) As Byte
                        CopyMemory_Strict VarPtr(dstArray(0)), lockedMem, hMemSize
                        GlobalUnlock lockedMem
                        GDIPlus_ImageSaveToArray = True
                    End If
                    
                End If
            End If
            
        Else
            GDIPlus_ImageSaveToArray = False
            InternalGDIPlusError "Image was not saved", "GDIPlus_ImageSaveToArray() failed; additional details follow"
            InternalGDIPlusError vbNullString, vbNullString, tmpReturn
        End If
    
    Else
        InternalGDIPlusError "Image was not saved", "GDIPlus_ImageSaveToArray() failed; no encoder found for that image format"
    End If
    
    Exit Function
    
GDIPlusSaveError:
    InternalGDIPlusError "Image was not saved", "A VB error occurred inside GDIPlus_ImageSaveToFile: " & Err.Description
    GDIPlus_ImageSaveToArray = False
End Function

'Save a surface to file.  The only property currently supported is JPEG quality; other properties are set automatically by GDI+.
Public Function GDIPlus_ImageSaveToFile(ByVal hImage As Long, ByVal dstFilename As String, Optional ByVal dstFileFormat As PD_2D_FileFormatExport = P2_FFE_PNG, Optional ByVal jpegQuality As Long = 85) As Boolean
        
    On Error GoTo GDIPlusSaveError
    
    'GDI+ uses GUIDs to define image export encoders; retrieve the relevant encoder GUID now
    Dim exporterGUID(0 To 15) As Byte
    If GetEncoderGUIDForPd2dFormat(dstFileFormat, VarPtr(exporterGUID(0))) Then
    
        'Like export format, GDI+ also uses GUIDs to define export properties.  If multiple encoder parameters
        ' are in use, these need to be merged into sequential order (because GDI+ only takes a pointer).
        ' pd2D does not currently cover this use-case; it always assumes there are only 0 or 1 parameters in use.
        ' To use multiple parameters, you would need copy the first GP_EncoderParameters entry into the
        ' fullEncoderParams() array, like normal, but with the Count value set to the number of parameters.
        ' Then, you would need to copy subsequent parameters into place *after* it.  (But *only* the parameters,
        ' not additional "Count" values.)
        '
        'Look at PhotoDemon's source code for an example of how to do this.
        Dim paramsInUse As Boolean: paramsInUse = False
        Dim tmpEncoderParams As GP_EncoderParameters, tmpConstString As String
        Dim fullEncoderParams() As Byte
        
        If (dstFileFormat = P2_FFE_JPEG) Then
            
            paramsInUse = True
            
            tmpEncoderParams.EP_Count = 1
            With tmpEncoderParams.EP_Parameter
                .EP_NumOfValues = 1
                .EP_ValueType = GP_EVT_Long
                tmpConstString = GP_EP_Quality
                CLSIDFromString StrPtr(tmpConstString), VarPtr(.EP_GUID(0))
                .EP_ValuePtr = VarPtr(jpegQuality)
            End With
            
        End If
        
        'Perform the export and return
        Dim tmpReturn As GP_Result
        If paramsInUse Then
            tmpReturn = GdipSaveImageToFile(hImage, StrPtr(dstFilename), VarPtr(exporterGUID(0)), VarPtr(tmpEncoderParams))
        Else
            tmpReturn = GdipSaveImageToFile(hImage, StrPtr(dstFilename), VarPtr(exporterGUID(0)), 0&)
        End If
        
        If (tmpReturn = GP_OK) Then
            GDIPlus_ImageSaveToFile = True
        Else
            GDIPlus_ImageSaveToFile = False
            InternalGDIPlusError "Image was not saved", "GDIPlus_ImageSaveToFile() failed to save " & dstFilename & "; additional details follow"
            InternalGDIPlusError vbNullString, vbNullString, tmpReturn
        End If
    
    Else
        InternalGDIPlusError "Image was not saved", "GDIPlus_ImageSaveToFile() failed to save " & dstFilename & "; no encoder found for that image format"
    End If
    
    Exit Function
    
GDIPlusSaveError:
    InternalGDIPlusError "Image was not saved", "A VB error occurred inside GDIPlus_ImageSaveToFile: " & Err.Description
    GDIPlus_ImageSaveToFile = False
End Function

'When exporting images, we need to find the unique GUID for a given exporter.  Matching via mimetype is a
' straightforward way to do this, and is the recommended solution from MSDN (see https://msdn.microsoft.com/en-us/library/ms533843(v=vs.85).aspx)
Private Function GetEncoderGUIDForPd2dFormat(ByVal srcFormat As PD_2D_FileFormatExport, ByVal ptrToDstGuid As Long) As Boolean
    
    GetEncoderGUIDForPd2dFormat = False
    
    'Generate a matching mimetype for the given format
    Dim srcMimetype As String
    Select Case srcFormat
        Case P2_FFE_BMP
            srcMimetype = "image/bmp"
        Case P2_FFE_GIF
            srcMimetype = "image/gif"
        Case P2_FFE_JPEG
            srcMimetype = "image/jpeg"
        Case P2_FFE_PNG
            srcMimetype = "image/png"
        Case P2_FFE_TIFF
            srcMimetype = "image/tiff"
        Case Else
            srcMimetype = vbNullString
    End Select
    
    If (Len(srcMimetype) <> 0) Then
        
        'Start by retrieving the number of encoders, and the size of the full encoder list
        Dim numOfEncoders As Long, sizeOfEncoders As Long
        If (GdipGetImageEncodersSize(numOfEncoders, sizeOfEncoders) = GP_OK) Then
            If (numOfEncoders > 0) And (sizeOfEncoders > 0) Then
            
                Dim encoderBuffer() As Byte
                Dim tmpCodec As GP_ImageCodecInfo
                
                'Hypothetically, we could probably pull the encoder list directly into a GP_ImageCodecInfo() array,
                ' but I haven't tested to see if the byte values of the encoder sizes are exact.  To avoid any problems,
                ' let's just dump the return into a byte array, then parse out what we need as we go.
                ReDim encoderBuffer(0 To sizeOfEncoders - 1) As Byte
                If (GdipGetImageEncoders(numOfEncoders, sizeOfEncoders, VarPtr(encoderBuffer(0))) = GP_OK) Then
                
                    'Iterate through the encoder list, searching for a match
                    Dim i As Long, strLength As Long, tmpMimeType As String
                    For i = 0 To numOfEncoders - 1
                    
                        'Extract this codec
                        CopyMemory_Strict VarPtr(tmpCodec), VarPtr(encoderBuffer(0)) + LenB(tmpCodec) * i, LenB(tmpCodec)
                        
                        'Extract the codec's mimetype
                        strLength = lstrlenW(tmpCodec.IC_MimeType)
                        If (strLength <> 0) Then
                            tmpMimeType = String$(strLength, 0&)
                            CopyMemory_Strict StrPtr(tmpMimeType), tmpCodec.IC_MimeType, strLength * 2
                            
                            'If we find a match, copy the encoder GUID and exit
                            If (StrComp(srcMimetype, tmpMimeType, vbBinaryCompare) = 0) Then
                                GetEncoderGUIDForPd2dFormat = True
                                CopyMemory_Strict ptrToDstGuid, VarPtr(tmpCodec.IC_ClassID(0)), 16&
                                Exit For
                            End If
                        End If
                        
                    Next i
                
                End If
                
            End If
        End If
        
    End If

End Function

Public Function GDIPlus_ImageUnlockBits(ByVal hImage As Long, ByRef srcCopyData As GP_BitmapData) As Boolean
    Dim tmpReturn As GP_Result
    tmpReturn = GdipBitmapUnlockBits(hImage, srcCopyData)
    GDIPlus_ImageUnlockBits = CBool(tmpReturn = GP_OK)
    If (tmpReturn <> GP_OK) Then InternalGDIPlusError vbNullString, vbNullString, tmpReturn
End Function

'Convert an EMF or WMF to the new EMF+ format.  Note that this is done in-memory, and the source file is not touched.
' Conversion allows us to render the metafile with antialiasing, alpha bytes, and more.
' REQUIRES GDI+ v1.1 (Win 7 or later only; conditionally available on Vista if explicitly requested via manifest)
'
'If successful, this function will generate a new handle.  It *must* be freed separately from the old handle!
Public Function GDIPlus_ImageUpgradeMetafile(ByVal hImage As Long, ByVal srcGraphicsForConvertSettings As Long, ByRef dstNewMetafile As Long) As Boolean
    
    dstNewMetafile = 0
    
    Dim tmpReturn As GP_Result
    tmpReturn = GdipConvertToEmfPlus(srcGraphicsForConvertSettings, hImage, ByVal 0&, GP_MT_EmfDual, 0&, dstNewMetafile)
    
    If (tmpReturn = GP_OK) Then
        GDIPlus_ImageUpgradeMetafile = True
    Else
        GDIPlus_ImageUpgradeMetafile = False
        InternalGDIPlusError vbNullString, vbNullString, tmpReturn
    End If
    
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

