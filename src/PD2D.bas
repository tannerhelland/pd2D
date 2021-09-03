Attribute VB_Name = "PD2D"
'***************************************************************************
'PhotoDemon 2D Painting class (interface for using pd2DBrush and pd2DPen on pd2DSurface objects)
'Copyright 2012-2021 by Tanner Helland
'Created: 01/September/12
'Last updated: 03/July/17
'Last update: kill off pd2DPainter class; migrate all commands to this module, instead
'
'All source code in this file is licensed under a modified BSD license. This means you may use the code in your own
' projects IF you provide attribution. For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'This master debug-mode flag modifies behavior in various pd2D objects (for example, some objects
' will track create/destroy behavior to make it easier to track down leaks).  I do *not* recommend
' enabling it in production builds as it has perf repercussions.
Public Const PD2D_DEBUG_MODE As Boolean = False

'If possible (e.g. painting without stretching), this painter class will drop back to bare AlphaBlend calls
' for image rendering.  This provides a meaningful performance improvement over GDI+ draw calls.
Private Declare Function AlphaBlend Lib "gdi32" Alias "GdiAlphaBlend" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long, ByVal blendFunct As Long) As Long

'In the future, this module may support multiple different rendering backends.  At present, however, only GDI+ is used.
Public Enum PD_2D_RENDERING_BACKEND
    P2_DefaultBackend = 0
    P2_GDIPlusBackend = 1
End Enum

#If False Then
    Private Const P2_DefaultBackend = 0, P2_GDIPlusBackend = 1
#End If

'When wrapping a DC, a surface needs to know the size of the object being painted on.  If an hWnd is supplied alongside
' the DC, we'll use that to auto-detect dimensions; otherwise, the caller needs to provide them.  (If the size is
' unknown, we'll use the size of the bitmap currently selected into the DC, but that's *not* reliable - so don't use it
' unless you know what you're doing!)
'
'This enum is only used internally.
Public Enum PD_2D_SIZE_DETECTION
    P2_SizeUnknown = 0
    P2_SizeFromHWnd = 1
    P2_SizeFromCaller = 2
End Enum

#If False Then
    Private Const P2_SizeUnknown = 0, P2_SizeFromHWnd = 1, P2_SizeFromCaller = 2
#End If

'The whole point of PD2D is to avoid backend-specific parameters.  As such, we necessarily wrap a number of
' GDI+ enums with our own P2-prefixed enums.  This seems redundant (and it is), but this is what makes it possible
' to support backends with different capabilities.
'
'As such, all PD2D classes operate on the enums defined in this class.  Where appropriate, they internally
' remap these values to backend-specific ones.

Public Enum PD_2D_Antialiasing
    P2_AA_None = 0&
    P2_AA_HighQuality = 1&
End Enum

#If False Then
    Private Const P2_AA_None = 0&, P2_AA_HighQuality = 1&
#End If

Public Enum PD_2D_BrushMode
    P2_BM_Solid = 0
    P2_BM_Pattern = 1
    P2_BM_Gradient = 2
    P2_BM_Texture = 3
End Enum

#If False Then
    Private Const P2_BM_Solid = 0, P2_BM_Pattern = 1, P2_BM_Gradient = 2, P2_BM_Texture = 3
#End If

Public Enum PD_2D_CombineMode
    P2_CM_Replace = 0
    P2_CM_Intersect = 1
    P2_CM_Union = 2
    P2_CM_Xor = 3
    P2_CM_Exclude = 4
    P2_CM_Complement = 5
End Enum

#If False Then
    Private Const P2_CM_Replace = 0, P2_CM_Intersect = 1, P2_CM_Union = 2, P2_CM_Xor = 3, P2_CM_Exclude = 4, P2_CM_Complement = 5
#End If

Public Enum PD_2D_CompositeMode
    P2_CM_Blend = 0
    P2_CM_Overwrite = 1
End Enum

#If False Then
    Private Const P2_CM_Blend = 0, P2_CM_Overwrite = 1
#End If

Public Enum PD_2D_DashCap
    P2_DC_Flat = 0
    P2_DC_Square = 1        'NOTE: GDI+ does not support square dash caps - only flat ones - so square simply remaps to flat
    P2_DC_Round = 2
    P2_DC_Triangle = 3
End Enum

#If False Then
    Private Const P2_DC_Flat = 0, P2_DC_Square = 1, P2_DC_Round = 2, P2_DC_Triangle = 3
#End If

Public Enum PD_2D_DashStyle
    P2_DS_Solid = 0&
    P2_DS_Dash = 1&
    P2_DS_Dot = 2&
    P2_DS_DashDot = 3&
    P2_DS_DashDotDot = 4&
    P2_DS_Custom = 5&
End Enum

#If False Then
    Private Const P2_DS_Solid = 0&, P2_DS_Dash = 1&, P2_DS_Dot = 2&, P2_DS_DashDot = 3&, P2_DS_DashDotDot = 4&, P2_DS_Custom = 5&
#End If

Public Enum PD_2D_FileFormatImport
    P2_FFI_Undefined = -1
    P2_FFI_BMP = 0
    P2_FFI_ICO = 1
    P2_FFI_JPEG = 2
    P2_FFI_GIF = 3
    P2_FFI_PNG = 4
    P2_FFI_TIFF = 5
    P2_FFI_WMF = 6
    P2_FFI_EMF = 7
End Enum

#If False Then
    Private Const P2_FFI_Undefined = -1, P2_FFI_BMP = 0, P2_FFI_ICO = 1, P2_FFI_JPEG = 2, P2_FFI_GIF = 3, P2_FFI_PNG = 4, P2_FFI_TIFF = 5, P2_FFI_WMF = 6, P2_FFI_EMF = 7
#End If

Public Enum PD_2D_FileFormatExport
    P2_FFE_BMP = 0
    P2_FFE_GIF = 1
    P2_FFE_JPEG = 2
    P2_FFE_PNG = 3
    P2_FFE_TIFF = 4
End Enum

#If False Then
    Private Const P2_FFE_BMP = 0, P2_FFE_GIF = 1, P2_FFE_JPEG = 2, P2_FFE_PNG = 3, P2_FFE_TIFF = 4
#End If

Public Enum PD_2D_FillRule
    P2_FR_OddEven = 0&
    P2_FR_Winding = 1&
End Enum

#If False Then
    Private Const P2_FR_OddEven = 0&, P2_FR_Winding = 1&
#End If

Public Enum PD_2D_GradientShape
    P2_GS_Linear = 0
    P2_GS_Reflection = 1
    P2_GS_Radial = 2
    P2_GS_Rectangle = 3
    P2_GS_Diamond = 4
End Enum

#If False Then
    Private Const P2_GS_Linear = 0, P2_GS_Reflection = 1, P2_GS_Radial = 2, P2_GS_Rectangle = 3, P2_GS_Diamond = 4
#End If

Public Enum PD_2D_LineCap
    P2_LC_Flat = 0&
    P2_LC_Square = 1&
    P2_LC_Round = 2&
    P2_LC_Triangle = 3&
    P2_LC_FlatAnchor = &H10
    P2_LC_SquareAnchor = &H11
    P2_LC_RoundAnchor = &H12
    P2_LC_DiamondAnchor = &H13
    P2_LC_ArrowAnchor = &H14
    P2_LC_Custom = &HFF
End Enum

#If False Then
    Private Const P2_LC_Flat = 0, P2_LC_Square = 1, P2_LC_Round = 2, P2_LC_Triangle = 3, P2_LC_FlatAnchor = &H10, P2_LC_SquareAnchor = &H11, P2_LC_RoundAnchor = &H12, P2_LC_DiamondAnchor = &H13, P2_LC_ArrowAnchor = &H14, P2_LC_Custom = &HFF
#End If

Public Enum PD_2D_LineJoin
    P2_LJ_Miter = 0&
    P2_LJ_Bevel = 1&
    P2_LJ_Round = 2&
End Enum

#If False Then
    Private Const P2_LJ_Miter = 0&, P2_LJ_Bevel = 1&, P2_LJ_Round = 2&
#End If

Public Enum PD_2D_PatternStyle
    P2_PS_Horizontal = 0
    P2_PS_Vertical = 1
    P2_PS_ForwardDiagonal = 2
    P2_PS_BackwardDiagonal = 3
    P2_PS_Cross = 4
    P2_PS_DiagonalCross = 5
    P2_PS_05Percent = 6
    P2_PS_10Percent = 7
    P2_PS_20Percent = 8
    P2_PS_25Percent = 9
    P2_PS_30Percent = 10
    P2_PS_40Percent = 11
    P2_PS_50Percent = 12
    P2_PS_60Percent = 13
    P2_PS_70Percent = 14
    P2_PS_75Percent = 15
    P2_PS_80Percent = 16
    P2_PS_90Percent = 17
    P2_PS_LightDownwardDiagonal = 18
    P2_PS_LightUpwardDiagonal = 19
    P2_PS_DarkDownwardDiagonal = 20
    P2_PS_DarkUpwardDiagonal = 21
    P2_PS_WideDownwardDiagonal = 22
    P2_PS_WideUpwardDiagonal = 23
    P2_PS_LightVertical = 24
    P2_PS_LightHorizontal = 25
    P2_PS_NarrowVertical = 26
    P2_PS_NarrowHorizontal = 27
    P2_PS_DarkVertical = 28
    P2_PS_DarkHorizontal = 29
    P2_PS_DashedDownwardDiagonal = 30
    P2_PS_DashedUpwardDiagonal = 31
    P2_PS_DashedHorizontal = 32
    P2_PS_DashedVertical = 33
    P2_PS_SmallConfetti = 34
    P2_PS_LargeConfetti = 35
    P2_PS_ZigZag = 36
    P2_PS_Wave = 37
    P2_PS_DiagonalBrick = 38
    P2_PS_HorizontalBrick = 39
    P2_PS_Weave = 40
    P2_PS_Plaid = 41
    P2_PS_Divot = 42
    P2_PS_DottedGrid = 43
    P2_PS_DottedDiamond = 44
    P2_PS_Shingle = 45
    P2_PS_Trellis = 46
    P2_PS_Sphere = 47
    P2_PS_SmallGrid = 48
    P2_PS_SmallCheckerBoard = 49
    P2_PS_LargeCheckerBoard = 50
    P2_PS_OutlinedDiamond = 51
    P2_PS_SolidDiamond = 52
End Enum

#If False Then
    Private Const P2_PS_Horizontal = 0, P2_PS_Vertical = 1, P2_PS_ForwardDiagonal = 2, P2_PS_BackwardDiagonal = 3, P2_PS_Cross = 4, P2_PS_DiagonalCross = 5, P2_PS_05Percent = 6, P2_PS_10Percent = 7, P2_PS_20Percent = 8, P2_PS_25Percent = 9, P2_PS_30Percent = 10, P2_PS_40Percent = 11, P2_PS_50Percent = 12, P2_PS_60Percent = 13, P2_PS_70Percent = 14, P2_PS_75Percent = 15, P2_PS_80Percent = 16, P2_PS_90Percent = 17, P2_PS_LightDownwardDiagonal = 18, P2_PS_LightUpwardDiagonal = 19, P2_PS_DarkDownwardDiagonal = 20, P2_PS_DarkUpwardDiagonal = 21, P2_PS_WideDownwardDiagonal = 22, P2_PS_WideUpwardDiagonal = 23, P2_PS_LightVertical = 24, P2_PS_LightHorizontal = 25
    Private Const P2_PS_NarrowVertical = 26, P2_PS_NarrowHorizontal = 27, P2_PS_DarkVertical = 28, P2_PS_DarkHorizontal = 29, P2_PS_DashedDownwardDiagonal = 30, P2_PS_DashedUpwardDiagonal = 31, P2_PS_DashedHorizontal = 32, P2_PS_DashedVertical = 33, P2_PS_SmallConfetti = 34, P2_PS_LargeConfetti = 35, P2_PS_ZigZag = 36, P2_PS_Wave = 37, P2_PS_DiagonalBrick = 38, P2_PS_HorizontalBrick = 39, P2_PS_Weave = 40, P2_PS_Plaid = 41, P2_PS_Divot = 42, P2_PS_DottedGrid = 43, P2_PS_DottedDiamond = 44, P2_PS_Shingle = 45, P2_PS_Trellis = 46, P2_PS_Sphere = 47, P2_PS_SmallGrid = 48, P2_PS_SmallCheckerBoard = 49, P2_PS_LargeCheckerBoard = 50
    Private Const P2_PS_OutlinedDiamond = 51, P2_PS_SolidDiamond = 52
#End If

Public Enum PD_2D_PixelOffset
    P2_PO_Normal = 0
    P2_PO_Half = 1
End Enum

#If False Then
    Private Const P2_PO_Normal = 0, P2_PO_Half = 1
#End If

Public Enum PD_2D_ResizeQuality
    P2_RQ_Fast = 0
    P2_RQ_Bilinear = 1
    P2_RQ_Bicubic = 2
End Enum

#If False Then
    Private Const P2_RQ_Fast = 0, P2_RQ_Bilinear = 1, P2_RQ_Bicubic = 2
#End If

'Surfaces come in a few different varieties.  Note that some actions may not be available for certain surface types.
Public Enum PD_2D_SurfaceType
    P2_ST_Uninitialized = -1    'The default value of a new surface; the surface is empty, and cannot be painted to
    P2_ST_WrapperOnly = 0       'This surface is just a wrapper around an existing hDC; pdSurface did not create it
    P2_ST_Bitmap = 1            'This surface is a bitmap (raster) surface, created and owned by a pdSurface instance
End Enum

#If False Then
    Private Const P2_ST_WrapperOnly = 0, P2_ST_Bitmap = 1
#End If

Public Enum PD_2D_TransformOrder
    P2_TO_Prepend = 0&
    P2_TO_Append = 1&
End Enum

#If False Then
    Private Const P2_TO_Prepend = 0&, P2_TO_Append = 1&
#End If

Public Enum PD_2D_WrapMode
    P2_WM_Tile = 0
    P2_WM_TileFlipX = 1
    P2_WM_TileFlipY = 2
    P2_WM_TileFlipXY = 3
    
    'IMPORTANT NOTE: clamp wrap mode does not work on all GDI+ calls; for example, it fails miserably
    ' on linear gradients for unknown reasons.  (See https://stackoverflow.com/questions/33225410/why-does-setting-lineargradientbrush-wrapmode-to-clamp-fail-with-argumentexcepti)
    P2_WM_Clamp = 4
End Enum

#If False Then
    Private Const P2_WM_Tile = 0, P2_WM_TileFlipX = 1, P2_WM_TileFlipY = 2, P2_WM_TileFlipXY = 3, P2_WM_Clamp = 4
#End If

'Certain structs are immensely helpful when drawing
Public Type RGBQuad
    Blue As Byte
    Green As Byte
    Red As Byte
    Alpha As Byte
End Type

Public Type PointFloat
    x As Single
    y As Single
End Type

Public Type PointLong
    x As Long
    y As Long
End Type

Public Type RectL
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type RectF
    Left As Single
    Top As Single
    Width As Single
    Height As Single
End Type

'SafeArray types for pointing VB arrays at arbitrary memory locations (in our case, bitmap data)
Public Type SafeArrayBound
    cElements As Long
    lBound   As Long
End Type

Public Type SafeArray2D
    cDims      As Integer
    fFeatures  As Integer
    cbElements As Long
    cLocks     As Long
    pvData     As Long
    Bounds(1)  As SafeArrayBound
End Type

Public Type SafeArray1D
    cDims      As Integer
    fFeatures  As Integer
    cbElements As Long
    cLocks     As Long
    pvData     As Long
    cElements As Long
    lBound   As Long
End Type

'PD's gradient format is straightforward, and it's declared here so functions can easily create their own gradient interfaces.
Public Type GradientPoint
    PointRGB As Long
    PointOpacity As Single
    PointPosition As Single
End Type

'If GDI+ is initialized successfully, this will be set to TRUE
Private m_GDIPlusAvailable As Boolean

'When debug mode is active, live counts of various drawing objects are tracked on a per-backend basis.  This is crucial for
' leak detection - these numbers should always match the number of active class instances.
Private m_BrushCount_GDIPlus As Long, m_PathCount_GDIPlus As Long, m_PenCount_GDIPlus As Long, m_RegionCount_GDIPlus As Long, m_SurfaceCount_GDIPlus As Long, m_TransformCount_GDIPlus As Long

'Some APIs are used *so* frequently throughout PD that we declare them publicly
Public Declare Sub CopyMemoryStrict Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpDst As Long, ByVal lpSrc As Long, ByVal byteLength As Long)
Public Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (ByVal dstPointer As Long, ByVal numOfBytes As Long, ByVal fillValue As Byte)
Public Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByVal dstPointer As Long, ByVal numOfBytes As Long)

Public Declare Function VarPtrArray Lib "msvbvm60" Alias "VarPtr" (ptr() As Any) As Long

'Not all of these functions are used in PD2D, but they are enumerated here for your convenience.
' Uncomment if curious.
'Public Declare Sub GetMem1 Lib "msvbvm60" (ByVal ptrSrc As Long, ByRef dstByte As Byte)
'Public Declare Sub GetMem1_Ptr Lib "msvbvm60" Alias "GetMem1" (ByVal ptrSrc As Long, ByVal ptrDst1 As Long)
'Public Declare Sub GetMem2 Lib "msvbvm60" (ByVal ptrSrc As Long, ByRef dstInteger As Integer)
'Public Declare Sub GetMem2_Ptr Lib "msvbvm60" Alias "GetMem2" (ByVal ptrSrc As Long, ByVal ptrDst2 As Long)
Public Declare Sub GetMem4 Lib "msvbvm60" (ByVal ptrSrc As Long, ByRef dstValue As Long)
Public Declare Sub GetMem4_Ptr Lib "msvbvm60" Alias "GetMem4" (ByVal ptrSrc As Long, ByVal ptrDst4 As Long)
'Public Declare Sub GetMem8_Ptr Lib "msvbvm60" Alias "GetMem8" (ByVal ptrSrc As Long, ByVal ptrDst8 As Long)
'Public Declare Sub PutMem1 Lib "msvbvm60" (ByVal ptrDst As Long, ByVal newValue As Byte)
Public Declare Sub PutMem2 Lib "msvbvm60" (ByVal ptrDst As Long, ByVal newValue As Integer)
Public Declare Sub PutMem4 Lib "msvbvm60" (ByVal ptrDst As Long, ByVal newValue As Long)

Private Declare Function RtlCompareMemory Lib "ntdll" (ByVal ptrSource1 As Long, ByVal ptrSource2 As Long, ByVal Length As Long) As Long

'Helper color functions for moving individual RGB components between RGB() Longs.  Note that these functions only
' return values in the range [0, 255], but declaring them as integers prevents overflow during intermediary math steps.
Public Function ExtractRed(ByVal srcColor As Long) As Integer
    ExtractRed = srcColor And 255
End Function

Public Function ExtractGreen(ByVal srcColor As Long) As Integer
    ExtractGreen = (srcColor \ 256) And 255
End Function

Public Function ExtractBlue(ByVal srcColor As Long) As Integer
    ExtractBlue = (srcColor \ 65536) And 255
End Function

Public Function GetNameOfFileFormat(ByVal srcFormat As PD_2D_FileFormatImport) As String
    Select Case srcFormat
        Case P2_FFI_BMP
            GetNameOfFileFormat = "BMP"
        Case P2_FFI_ICO
            GetNameOfFileFormat = "Icon"
        Case P2_FFI_JPEG
            GetNameOfFileFormat = "JPEG"
        Case P2_FFI_GIF
            GetNameOfFileFormat = "GIF"
        Case P2_FFI_PNG
            GetNameOfFileFormat = "PNG"
        Case P2_FFI_TIFF
            GetNameOfFileFormat = "TIFF"
        Case P2_FFI_WMF
            GetNameOfFileFormat = "WMF"
        Case P2_FFI_EMF
            GetNameOfFileFormat = "EMF"
        Case Else
            GetNameOfFileFormat = "Unknown file format"
    End Select
End Function

Public Function MemCmp(ByVal ptr1 As Long, ByVal ptr2 As Long, ByVal bytesToCompare As Long) As Boolean
    Dim bytesEqual As Long
    bytesEqual = RtlCompareMemory(ptr1, ptr2, bytesToCompare)
    MemCmp = (bytesEqual = bytesToCompare)
End Function

'Shortcut function for creating a new rectangular region with the default rendering backend
Public Function QuickCreateRegionRectangle(ByRef dstRegion As pd2DRegion, ByVal rLeft As Single, ByVal rTop As Single, ByVal rWidth As Single, ByVal rHeight As Single) As Boolean
    If (dstRegion Is Nothing) Then Set dstRegion = New pd2DRegion Else dstRegion.ResetAllProperties
    With dstRegion
        QuickCreateRegionRectangle = .AddRectangleF(rLeft, rTop, rWidth, rHeight, P2_CM_Replace)
    End With
End Function

'Shortcut function for quickly creating a blank surface with the default rendering backend and default rendering settings
Public Function QuickCreateBlankSurface(ByRef dstSurface As pd2DSurface, ByVal surfaceWidth As Long, ByVal surfaceHeight As Long, Optional ByVal surfaceSupportsAlpha As Boolean = True, Optional ByVal enableAntialiasing As Boolean = False, Optional ByVal initialColor As Long = vbWhite, Optional ByVal initialOpacity As Single = 100#) As Boolean
    If (dstSurface Is Nothing) Then Set dstSurface = New pd2DSurface Else dstSurface.ResetAllProperties
    With dstSurface
        If enableAntialiasing Then .SetSurfaceAntialiasing P2_AA_HighQuality Else .SetSurfaceAntialiasing P2_AA_None
        QuickCreateBlankSurface = .CreateBlankSurface(surfaceWidth, surfaceHeight, surfaceSupportsAlpha, initialColor, initialOpacity)
    End With
End Function

'Shortcut function for creating a new surface with the default rendering backend and default rendering settings
Public Function QuickCreateSurfaceFromDC(ByRef dstSurface As pd2DSurface, ByVal srcDC As Long, Optional ByVal enableAntialiasing As Boolean = False, Optional ByVal srcHwnd As Long = 0) As Boolean
    If (dstSurface Is Nothing) Then Set dstSurface = New pd2DSurface Else dstSurface.ResetAllProperties
    With dstSurface
        If enableAntialiasing Then .SetSurfaceAntialiasing P2_AA_HighQuality Else .SetSurfaceAntialiasing P2_AA_None
        QuickCreateSurfaceFromDC = .WrapSurfaceAroundDC(srcDC, srcHwnd)
    End With
End Function

Public Function QuickCreateSurfaceFromDIB(ByRef dstSurface As pd2DSurface, ByVal srcDIB As pd2DDIB, Optional ByVal enableAntialiasing As Boolean = False) As Boolean
    If (dstSurface Is Nothing) Then Set dstSurface = New pd2DSurface Else dstSurface.ResetAllProperties
    With dstSurface
        If enableAntialiasing Then .SetSurfaceAntialiasing P2_AA_HighQuality Else .SetSurfaceAntialiasing P2_AA_None
        QuickCreateSurfaceFromDIB = .WrapSurfaceAroundpd2ddib(srcDIB)
    End With
End Function

Public Function QuickCreateSurfaceFromFile(ByRef dstSurface As pd2DSurface, ByVal srcPath As String) As Boolean
    If (dstSurface Is Nothing) Then Set dstSurface = New pd2DSurface Else dstSurface.ResetAllProperties
    With dstSurface
        QuickCreateSurfaceFromFile = .CreateSurfaceFromFile(srcPath)
    End With
End Function

'Shortcut function for creating a solid brush
Public Function QuickCreateSolidBrush(ByRef dstBrush As pd2DBrush, Optional ByVal brushColor As Long = vbWhite, Optional ByVal brushOpacity As Single = 100#) As Boolean
    If (dstBrush Is Nothing) Then Set dstBrush = New pd2DBrush Else dstBrush.ResetAllProperties
    With dstBrush
        .SetBrushColor brushColor
        .SetBrushOpacity brushOpacity
        QuickCreateSolidBrush = .CreateBrush()
    End With
End Function

'Shortcut function for creating a two-color gradient brush
Public Function QuickCreateTwoColorGradientBrush(ByRef dstBrush As pd2DBrush, ByRef gradientBoundary As RectF, Optional ByVal firstColor As Long = vbBlack, Optional ByVal secondColor As Long = vbWhite, Optional ByVal firstColorOpacity As Single = 100#, Optional ByVal secondColorOpacity As Single = 100#, Optional ByVal gradientShape As PD_2D_GradientShape = P2_GS_Linear, Optional ByVal gradientAngle As Single = 0#) As Boolean
    
    If (dstBrush Is Nothing) Then Set dstBrush = New pd2DBrush Else dstBrush.ResetAllProperties
    
    Dim tmpGradient As pd2DGradient
    Set tmpGradient = New pd2DGradient
    With tmpGradient
        .SetGradientShape gradientShape
        .SetGradientAngle gradientAngle
        .CreateTwoPointGradient firstColor, secondColor, firstColorOpacity, secondColorOpacity
    End With
    
    With dstBrush
        .SetBrushMode P2_BM_Gradient
        .SetBoundaryRect gradientBoundary
        .SetBrushGradientAllSettings tmpGradient.GetGradientAsString
        QuickCreateTwoColorGradientBrush = .CreateBrush()
    End With
    
End Function

'Shortcut function for creating a solid pen
Public Function QuickCreateSolidPen(ByRef dstPen As pd2DPen, Optional ByVal penWidth As Single = 1!, Optional ByVal penColor As Long = vbWhite, Optional ByVal penOpacity As Single = 100!, Optional ByVal penLineJoin As PD_2D_LineJoin = P2_LJ_Miter, Optional ByVal penLineCap As PD_2D_LineCap = P2_LC_Flat) As Boolean
    If (dstPen Is Nothing) Then Set dstPen = New pd2DPen Else dstPen.ResetAllProperties
    With dstPen
        .SetPenWidth penWidth
        .SetPenColor penColor
        .SetPenOpacity penOpacity
        .SetPenLineJoin penLineJoin
        .SetPenLineCap penLineCap
        QuickCreateSolidPen = .CreatePen()
    End With
End Function

'Shortcut function for creating two pens for UI purposes.  This function could use a clearer name, but "UI pens" consist
' of a wide, semi-translucent black pen on bottom, and a thin, less-translucent white pen on top.  This combination
' of pens are perfect for drawing on any arbitrary background of any color or pattern, and ensuring the line will still
' be visible.  (This approach is used in modern software instead of the old "invert" pen approach of past decades.)
'
'If the line is currently being hovered or otherwise interacted with, you can set "useHighlightColor" to TRUE.  This will
' return the top pen in the current highlight color (hard-coded at the top of this module) instead of plain white.
'
'By design, pen width is not settable via this function.  The top pen will always be 1.6 pixels wide (a size required
' to bypass GDI+ subpixel flaws between 1 and 2 pixels) while the bottom pen will always be 3.0 pixels wide.
Public Function QuickCreatePairOfUIPens(ByRef dstPenBase As pd2DPen, ByRef dstPenTop As pd2DPen, Optional ByVal useHighlightColor As Boolean = False, Optional ByVal penLineJoin As PD_2D_LineJoin = P2_LJ_Miter, Optional ByVal penLineCap As PD_2D_LineCap = P2_LC_Flat) As Boolean
    
    If (dstPenBase Is Nothing) Then Set dstPenBase = New pd2DPen Else dstPenBase.ResetAllProperties
    If (dstPenTop Is Nothing) Then Set dstPenTop = New pd2DPen Else dstPenTop.ResetAllProperties
    
    With dstPenBase
        .SetPenWidth 3!
        .SetPenColor vbBlack
        .SetPenOpacity 75!
        .SetPenLineJoin penLineJoin
        .SetPenLineCap penLineCap
        QuickCreatePairOfUIPens = .CreatePen()
    End With
    
    With dstPenTop
        .SetPenWidth 1.6!
        If useHighlightColor Then .SetPenColor vbBlue Else .SetPenColor vbWhite
        .SetPenOpacity 87.5!
        .SetPenLineJoin penLineJoin
        .SetPenLineCap penLineCap
        QuickCreatePairOfUIPens = (QuickCreatePairOfUIPens And .CreatePen())
    End With
    
End Function

'LoadPicture replacement.  All pd2D interactions are handled internally, so you just pass the target object
' and source file path.
'
'The target object needs to have a DC property, or this function will fail.
Public Function QuickLoadPicture(ByRef dstObject As Object, ByVal srcPath As String, Optional ByVal resizeImageToFit As Boolean = True) As Boolean
    
    On Error GoTo LoadPictureFail
    
    Dim srcSurface As pd2DSurface
    If PD2D.QuickCreateSurfaceFromFile(srcSurface, srcPath) Then
        
        Dim dstSurface As pd2DSurface
        If PD2D.QuickCreateSurfaceFromDC(dstSurface, dstObject.hDC, True, dstObject.hWnd) Then
            
            If resizeImageToFit Then
                
                'If the source surface is smaller than the destination surface, center the image to fit
                If ((srcSurface.GetSurfaceWidth < dstSurface.GetSurfaceWidth) And (srcSurface.GetSurfaceHeight < dstSurface.GetSurfaceHeight)) Then
                    QuickLoadPicture = PD2D.DrawSurfaceI(dstSurface, (dstSurface.GetSurfaceWidth - srcSurface.GetSurfaceWidth) \ 2, (dstSurface.GetSurfaceHeight - srcSurface.GetSurfaceHeight) \ 2, srcSurface)
                Else
                
                    'Calculate the correct target size, and use that size when painting.
                    Dim newWidth As Long, newHeight As Long
                    PD2D_Math.ConvertAspectRatio srcSurface.GetSurfaceWidth, srcSurface.GetSurfaceHeight, dstSurface.GetSurfaceWidth, dstSurface.GetSurfaceHeight, newWidth, newHeight
                    
                    dstSurface.SetSurfaceResizeQuality P2_RQ_Bicubic
                    QuickLoadPicture = PD2D.DrawSurfaceResizedI(dstSurface, (dstSurface.GetSurfaceWidth - newWidth) \ 2, (dstSurface.GetSurfaceHeight - newHeight) \ 2, newWidth, newHeight, srcSurface)
                    
                End If
                
            Else
                QuickLoadPicture = PD2D.DrawSurfaceI(dstSurface, 0, 0, srcSurface)
            End If
            
        End If
        
    End If
    
    Exit Function
    
LoadPictureFail:
    InternalError "QuickLoadPicture", Err.Description, Err.Number
    QuickLoadPicture = False
End Function

Public Function IsRenderingEngineActive(Optional ByVal targetBackend As PD_2D_RENDERING_BACKEND = P2_DefaultBackend) As Boolean
    Select Case targetBackend
        Case P2_DefaultBackend, P2_GDIPlusBackend
            IsRenderingEngineActive = m_GDIPlusAvailable
        Case Else
            IsRenderingEngineActive = False
    End Select
End Function

'Start a new rendering backend
Public Function StartRenderingEngine(Optional ByVal targetBackend As PD_2D_RENDERING_BACKEND = P2_DefaultBackend) As Boolean

    Select Case targetBackend
            
        Case P2_DefaultBackend, P2_GDIPlusBackend
            StartRenderingEngine = PD2D_GDIPlus.GDIP_StartEngine(False)
            m_GDIPlusAvailable = StartRenderingEngine
            
        Case Else
            InternalError "StartRenderingEngine", "unknown backend"
    
    End Select

End Function

'Stop a running rendering backend
Public Function StopRenderingEngine(Optional ByVal targetBackend As PD_2D_RENDERING_BACKEND = P2_DefaultBackend) As Boolean
        
    Select Case targetBackend
            
        Case P2_DefaultBackend, P2_GDIPlusBackend
            
            'Prior to release, see if any GDI+ object counts are non-zero; if they are, the caller needs to
            ' be notified, because those resources should be released before the backend disappears.
            If PD2D_DEBUG_MODE Then
                If (m_BrushCount_GDIPlus <> 0) Then InternalError "StopRenderingEngine", "There are still " & m_BrushCount_GDIPlus & " brush(es) active.  Release these objects before shutting down the drawing backend."
                If (m_PathCount_GDIPlus <> 0) Then InternalError "StopRenderingEngine", "There are still " & m_PathCount_GDIPlus & " path(s) active.  Release these objects before shutting down the drawing backend."
                If (m_PenCount_GDIPlus <> 0) Then InternalError "StopRenderingEngine", "There are still " & m_PenCount_GDIPlus & " pen(s) active.  Release these objects before shutting down the drawing backend."
                If (m_RegionCount_GDIPlus <> 0) Then InternalError "StopRenderingEngine", "There are still " & m_RegionCount_GDIPlus & " region(s) active.  Release these objects before shutting down the drawing backend."
                If (m_SurfaceCount_GDIPlus <> 0) Then InternalError "StopRenderingEngine", "There are still " & m_SurfaceCount_GDIPlus & " surface(s) active.  Release these objects before shutting down the drawing backend."
                If (m_TransformCount_GDIPlus <> 0) Then InternalError "StopRenderingEngine", "There are still " & m_TransformCount_GDIPlus & " transform(s) active.  Release these objects before shutting down the drawing backend."
            End If
            
            StopRenderingEngine = PD2D_GDIPlus.GDIP_StopEngine()
            m_GDIPlusAvailable = False
            
        Case Else
            InternalError "StopRenderingEngine", "unknown backend"
    
    End Select
    
End Function

'DEBUG FUNCTIONS FOLLOW.  These functions should not be called directly.  They are invoked by other pd2D class when PD2D_DEBUG_MODE = TRUE.
Public Sub DEBUG_NotifyBrushCountChange(ByVal objectCreated As Boolean)
    If objectCreated Then m_BrushCount_GDIPlus = m_BrushCount_GDIPlus + 1 Else m_BrushCount_GDIPlus = m_BrushCount_GDIPlus - 1
End Sub

Public Sub DEBUG_NotifyPathCountChange(ByVal objectCreated As Boolean)
    If objectCreated Then m_PathCount_GDIPlus = m_PathCount_GDIPlus + 1 Else m_PathCount_GDIPlus = m_PathCount_GDIPlus - 1
End Sub

Public Sub DEBUG_NotifyPenCountChange(ByVal objectCreated As Boolean)
    If objectCreated Then m_PenCount_GDIPlus = m_PenCount_GDIPlus + 1 Else m_PenCount_GDIPlus = m_PenCount_GDIPlus - 1
End Sub

Public Sub DEBUG_NotifyRegionCountChange(ByVal objectCreated As Boolean)
    If objectCreated Then m_RegionCount_GDIPlus = m_RegionCount_GDIPlus + 1 Else m_RegionCount_GDIPlus = m_RegionCount_GDIPlus - 1
End Sub

Public Sub DEBUG_NotifySurfaceCountChange(ByVal objectCreated As Boolean)
    If objectCreated Then m_SurfaceCount_GDIPlus = m_SurfaceCount_GDIPlus + 1 Else m_SurfaceCount_GDIPlus = m_SurfaceCount_GDIPlus - 1
End Sub

Public Sub DEBUG_NotifyTransformCountChange(ByVal objectCreated As Boolean)
    If objectCreated Then m_TransformCount_GDIPlus = m_TransformCount_GDIPlus + 1 Else m_TransformCount_GDIPlus = m_TransformCount_GDIPlus - 1
End Sub

'In a default build, external pd2D classes relay any internal errors to this function.  You may wish to modify those classes
' to raise their own error events, or perhaps handle their errors internally.  (By default, pd2D does *not* halt on errors.)
Public Sub DEBUG_NotifyError(ByRef errClassName As String, ByRef errFunctionName As String, ByRef errDescription As String, Optional ByVal errNum As Long = 0)
    Debug.Print "WARNING!  pd2D error in " & errClassName & "." & errFunctionName & ": " & errDescription
    If (errNum <> 0) Then Debug.Print "  (If it helps, an error number was also reported: #" & errNum & ")"
End Sub

'These functions exist to help with XML serialization.  They use consistent names against
' the current enums, and if the enums ever change order in the future, existing XML strings
' will still produce correct results.
Public Function XML_GetNameOfWrapMode(ByVal srcWrapMode As PD_2D_WrapMode) As String
    Select Case srcWrapMode
        Case P2_WM_Tile
            XML_GetNameOfWrapMode = "tile"
        Case P2_WM_TileFlipX
            XML_GetNameOfWrapMode = "tile-flip-x"
        Case P2_WM_TileFlipY
            XML_GetNameOfWrapMode = "tile-flip-y"
        Case P2_WM_TileFlipXY
            XML_GetNameOfWrapMode = "tile-flip-xy"
        Case P2_WM_Clamp
            XML_GetNameOfWrapMode = "clamp"
        Case Else
            XML_GetNameOfWrapMode = "tile"
    End Select
End Function

Public Function XML_GetWrapModeFromName(ByRef srcName As String) As PD_2D_WrapMode
    Select Case srcName
        Case "tile"
            XML_GetWrapModeFromName = P2_WM_Tile
        Case "tile-flip-x"
            XML_GetWrapModeFromName = P2_WM_TileFlipX
        Case "tile-flip-y"
            XML_GetWrapModeFromName = P2_WM_TileFlipY
        Case "tile-flip-xy"
            XML_GetWrapModeFromName = P2_WM_TileFlipXY
        Case "clamp"
            XML_GetWrapModeFromName = P2_WM_Clamp
        Case Else
            XML_GetWrapModeFromName = P2_WM_Tile
    End Select
End Function

Public Function XML_GetNameOfBrushMode(ByVal srcBrushMode As PD_2D_BrushMode) As String
    Select Case srcBrushMode
        Case P2_BM_Solid
            XML_GetNameOfBrushMode = "solid"
        Case P2_BM_Pattern
            XML_GetNameOfBrushMode = "pattern"
        Case P2_BM_Gradient
            XML_GetNameOfBrushMode = "gradient"
        Case P2_BM_Texture
            XML_GetNameOfBrushMode = "texture"
        Case Else
            XML_GetNameOfBrushMode = "solid"
    End Select
End Function

Public Function XML_GetBrushModeFromName(ByRef srcName As String) As PD_2D_BrushMode
    Select Case srcName
        Case "solid"
            XML_GetBrushModeFromName = P2_BM_Solid
        Case "pattern"
            XML_GetBrushModeFromName = P2_BM_Pattern
        Case "gradient"
            XML_GetBrushModeFromName = P2_BM_Gradient
        Case "texture"
            XML_GetBrushModeFromName = P2_BM_Texture
        Case Else
            XML_GetBrushModeFromName = P2_BM_Solid
    End Select
End Function

Public Function XML_GetNameOfPattern(ByVal srcPattern As PD_2D_PatternStyle) As String
    Select Case srcPattern
        Case P2_PS_Horizontal
            XML_GetNameOfPattern = "x"
        Case P2_PS_Vertical
            XML_GetNameOfPattern = "y"
        Case P2_PS_ForwardDiagonal
            XML_GetNameOfPattern = "forward-dg"
        Case P2_PS_BackwardDiagonal
            XML_GetNameOfPattern = "backward-dg"
        Case P2_PS_Cross
            XML_GetNameOfPattern = "cross"
        Case P2_PS_DiagonalCross
            XML_GetNameOfPattern = "dg-cross"
        Case P2_PS_05Percent
            XML_GetNameOfPattern = "pc-05"
        Case P2_PS_10Percent
            XML_GetNameOfPattern = "pc-10"
        Case P2_PS_20Percent
            XML_GetNameOfPattern = "pc-20"
        Case P2_PS_25Percent
            XML_GetNameOfPattern = "pc-25"
        Case P2_PS_30Percent
            XML_GetNameOfPattern = "pc-30"
        Case P2_PS_40Percent
            XML_GetNameOfPattern = "pc-40"
        Case P2_PS_50Percent
            XML_GetNameOfPattern = "pc-50"
        Case P2_PS_60Percent
            XML_GetNameOfPattern = "pc-60"
        Case P2_PS_70Percent
            XML_GetNameOfPattern = "pc-70"
        Case P2_PS_75Percent
            XML_GetNameOfPattern = "pc-75"
        Case P2_PS_80Percent
            XML_GetNameOfPattern = "pc-80"
        Case P2_PS_90Percent
            XML_GetNameOfPattern = "pc-90"
        Case P2_PS_LightDownwardDiagonal
            XML_GetNameOfPattern = "light-down-dg"
        Case P2_PS_LightUpwardDiagonal
            XML_GetNameOfPattern = "light-up-dg"
        Case P2_PS_DarkDownwardDiagonal
            XML_GetNameOfPattern = "dark-down-dg"
        Case P2_PS_DarkUpwardDiagonal
            XML_GetNameOfPattern = "dark-up-dg"
        Case P2_PS_WideDownwardDiagonal
            XML_GetNameOfPattern = "wide-down-dg"
        Case P2_PS_WideUpwardDiagonal
            XML_GetNameOfPattern = "wide-up-dg"
        Case P2_PS_LightVertical
            XML_GetNameOfPattern = "light-y"
        Case P2_PS_LightHorizontal
            XML_GetNameOfPattern = "light-x"
        Case P2_PS_NarrowVertical
            XML_GetNameOfPattern = "narrow-y"
        Case P2_PS_NarrowHorizontal
            XML_GetNameOfPattern = "narrow-x"
        Case P2_PS_DarkVertical
            XML_GetNameOfPattern = "dark-y"
        Case P2_PS_DarkHorizontal
            XML_GetNameOfPattern = "dark-x"
        Case P2_PS_DashedDownwardDiagonal
            XML_GetNameOfPattern = "dash-down-dg"
        Case P2_PS_DashedUpwardDiagonal
            XML_GetNameOfPattern = "dash-up-dg"
        Case P2_PS_DashedHorizontal
            XML_GetNameOfPattern = "dash-x"
        Case P2_PS_DashedVertical
            XML_GetNameOfPattern = "dash-y"
        Case P2_PS_SmallConfetti
            XML_GetNameOfPattern = "confetti-s"
        Case P2_PS_LargeConfetti
            XML_GetNameOfPattern = "confetti-l"
        Case P2_PS_ZigZag
            XML_GetNameOfPattern = "zigzag"
        Case P2_PS_Wave
            XML_GetNameOfPattern = "wave"
        Case P2_PS_DiagonalBrick
            XML_GetNameOfPattern = "brick-dg"
        Case P2_PS_HorizontalBrick
            XML_GetNameOfPattern = "brick-x"
        Case P2_PS_Weave
            XML_GetNameOfPattern = "weave"
        Case P2_PS_Plaid
            XML_GetNameOfPattern = "plaid"
        Case P2_PS_Divot
            XML_GetNameOfPattern = "divot"
        Case P2_PS_DottedGrid
            XML_GetNameOfPattern = "dot-grid"
        Case P2_PS_DottedDiamond
            XML_GetNameOfPattern = "dot-diamond"
        Case P2_PS_Shingle
            XML_GetNameOfPattern = "shingle"
        Case P2_PS_Trellis
            XML_GetNameOfPattern = "trellis"
        Case P2_PS_Sphere
            XML_GetNameOfPattern = "sphere"
        Case P2_PS_SmallGrid
            XML_GetNameOfPattern = "grid-s"
        Case P2_PS_SmallCheckerBoard
            XML_GetNameOfPattern = "checker-s"
        Case P2_PS_LargeCheckerBoard
            XML_GetNameOfPattern = "checker-l"
        Case P2_PS_OutlinedDiamond
            XML_GetNameOfPattern = "diamond-outline"
        Case P2_PS_SolidDiamond
            XML_GetNameOfPattern = "diamond-solid"
        Case Else
            XML_GetNameOfPattern = "x"
    End Select
End Function

Public Function XML_GetPatternFromName(ByRef srcName As String) As PD_2D_PatternStyle
    Select Case srcName
        Case "x"
            XML_GetPatternFromName = P2_PS_Horizontal
        Case "y"
            XML_GetPatternFromName = P2_PS_Vertical
        Case "forward-dg"
            XML_GetPatternFromName = P2_PS_ForwardDiagonal
        Case "backward-dg"
            XML_GetPatternFromName = P2_PS_BackwardDiagonal
        Case "cross"
            XML_GetPatternFromName = P2_PS_Cross
        Case "dg-cross"
            XML_GetPatternFromName = P2_PS_DiagonalCross
        Case "pc-05"
            XML_GetPatternFromName = P2_PS_05Percent
        Case "pc-10"
            XML_GetPatternFromName = P2_PS_10Percent
        Case "pc-20"
            XML_GetPatternFromName = P2_PS_20Percent
        Case "pc-25"
            XML_GetPatternFromName = P2_PS_25Percent
        Case "pc-30"
            XML_GetPatternFromName = P2_PS_30Percent
        Case "pc-40"
            XML_GetPatternFromName = P2_PS_40Percent
        Case "pc-50"
            XML_GetPatternFromName = P2_PS_50Percent
        Case "pc-60"
            XML_GetPatternFromName = P2_PS_60Percent
        Case "pc-70"
            XML_GetPatternFromName = P2_PS_70Percent
        Case "pc-75"
            XML_GetPatternFromName = P2_PS_75Percent
        Case "pc-80"
            XML_GetPatternFromName = P2_PS_80Percent
        Case "pc-90"
            XML_GetPatternFromName = P2_PS_90Percent
        Case "light-down-dg"
            XML_GetPatternFromName = P2_PS_LightDownwardDiagonal
        Case "light-up-dg"
            XML_GetPatternFromName = P2_PS_LightUpwardDiagonal
        Case "dark-down-dg"
            XML_GetPatternFromName = P2_PS_DarkDownwardDiagonal
        Case "dark-up-dg"
            XML_GetPatternFromName = P2_PS_DarkUpwardDiagonal
        Case "wide-down-dg"
            XML_GetPatternFromName = P2_PS_WideDownwardDiagonal
        Case "wide-up-dg"
            XML_GetPatternFromName = P2_PS_WideUpwardDiagonal
        Case "light-y"
            XML_GetPatternFromName = P2_PS_LightVertical
        Case "light-x"
            XML_GetPatternFromName = P2_PS_LightHorizontal
        Case "narrow-y"
            XML_GetPatternFromName = P2_PS_NarrowVertical
        Case "narrow-x"
            XML_GetPatternFromName = P2_PS_NarrowHorizontal
        Case "dark-y"
            XML_GetPatternFromName = P2_PS_DarkVertical
        Case "dark-x"
            XML_GetPatternFromName = P2_PS_DarkHorizontal
        Case "dash-down-dg"
            XML_GetPatternFromName = P2_PS_DashedDownwardDiagonal
        Case "dash-up-dg"
            XML_GetPatternFromName = P2_PS_DashedUpwardDiagonal
        Case "dash-x"
            XML_GetPatternFromName = P2_PS_DashedHorizontal
        Case "dash-y"
            XML_GetPatternFromName = P2_PS_DashedVertical
        Case "confetti-s"
            XML_GetPatternFromName = P2_PS_SmallConfetti
        Case "confetti-l"
            XML_GetPatternFromName = P2_PS_LargeConfetti
        Case "zigzag"
            XML_GetPatternFromName = P2_PS_ZigZag
        Case "wave"
            XML_GetPatternFromName = P2_PS_Wave
        Case "brick-dg"
            XML_GetPatternFromName = P2_PS_DiagonalBrick
        Case "brick-x"
            XML_GetPatternFromName = P2_PS_HorizontalBrick
        Case "weave"
            XML_GetPatternFromName = P2_PS_Weave
        Case "plaid"
            XML_GetPatternFromName = P2_PS_Plaid
        Case "divot"
            XML_GetPatternFromName = P2_PS_Divot
        Case "dot-grid"
            XML_GetPatternFromName = P2_PS_DottedGrid
        Case "dot-diamond"
            XML_GetPatternFromName = P2_PS_DottedDiamond
        Case "shingle"
            XML_GetPatternFromName = P2_PS_Shingle
        Case "trellis"
            XML_GetPatternFromName = P2_PS_Trellis
        Case "sphere"
            XML_GetPatternFromName = P2_PS_Sphere
        Case "grid-s"
            XML_GetPatternFromName = P2_PS_SmallGrid
        Case "checker-s"
            XML_GetPatternFromName = P2_PS_SmallCheckerBoard
        Case "checker-l"
            XML_GetPatternFromName = P2_PS_LargeCheckerBoard
        Case "diamond-outline"
            XML_GetPatternFromName = P2_PS_OutlinedDiamond
        Case "diamond-solid"
            XML_GetPatternFromName = P2_PS_SolidDiamond
        Case Else
            XML_GetPatternFromName = P2_PS_Horizontal
    End Select
End Function

Public Function XML_GetNameOfGradientShape(ByVal srcShape As PD_2D_GradientShape) As String
    Select Case srcShape
        Case P2_GS_Linear
            XML_GetNameOfGradientShape = "linear"
        Case P2_GS_Reflection
            XML_GetNameOfGradientShape = "reflect"
        Case P2_GS_Radial
            XML_GetNameOfGradientShape = "radial"
        Case P2_GS_Rectangle
            XML_GetNameOfGradientShape = "rectangle"
        Case P2_GS_Diamond
            XML_GetNameOfGradientShape = "diamond"
        Case Else
            XML_GetNameOfGradientShape = "linear"
    End Select
End Function

Public Function XML_GetGradientShapeFromName(ByRef srcName As String) As PD_2D_GradientShape
    Select Case srcName
        Case "linear"
            XML_GetGradientShapeFromName = P2_GS_Linear
        Case "reflect"
            XML_GetGradientShapeFromName = P2_GS_Reflection
        Case "radial"
            XML_GetGradientShapeFromName = P2_GS_Radial
        Case "rectangle"
            XML_GetGradientShapeFromName = P2_GS_Rectangle
        Case "diamond"
            XML_GetGradientShapeFromName = P2_GS_Diamond
        Case Else
            XML_GetGradientShapeFromName = P2_GS_Linear
    End Select
End Function

Public Function XML_GetNameOfLineCap(ByVal srcLineCap As PD_2D_LineCap) As String
    Select Case srcLineCap
        Case P2_LC_Flat
            XML_GetNameOfLineCap = "flat"
        Case P2_LC_Square
            XML_GetNameOfLineCap = "square"
        Case P2_LC_Round
            XML_GetNameOfLineCap = "round"
        Case P2_LC_Triangle
            XML_GetNameOfLineCap = "triangle"
        Case P2_LC_FlatAnchor
            XML_GetNameOfLineCap = "anchor-flat"
        Case P2_LC_SquareAnchor
            XML_GetNameOfLineCap = "anchor-square"
        Case P2_LC_RoundAnchor
            XML_GetNameOfLineCap = "anchor-round"
        Case P2_LC_DiamondAnchor
            XML_GetNameOfLineCap = "anchor-diamond"
        Case P2_LC_ArrowAnchor
            XML_GetNameOfLineCap = "anchor-arrow"
        Case P2_LC_Custom
            XML_GetNameOfLineCap = "custom"
        Case Else
            XML_GetNameOfLineCap = "flat"
    End Select
End Function

Public Function XML_GetLineCapFromName(ByRef srcName As String) As PD_2D_LineCap
    Select Case srcName
        Case "flat"
            XML_GetLineCapFromName = P2_LC_Flat
        Case "square"
            XML_GetLineCapFromName = P2_LC_Square
        Case "round"
            XML_GetLineCapFromName = P2_LC_Round
        Case "triangle"
            XML_GetLineCapFromName = P2_LC_Triangle
        Case "anchor-flat"
            XML_GetLineCapFromName = P2_LC_FlatAnchor
        Case "anchor-square"
            XML_GetLineCapFromName = P2_LC_SquareAnchor
        Case "anchor-round"
            XML_GetLineCapFromName = P2_LC_RoundAnchor
        Case "anchor-diamond"
            XML_GetLineCapFromName = P2_LC_DiamondAnchor
        Case "anchor-arrow"
            XML_GetLineCapFromName = P2_LC_ArrowAnchor
        Case "custom"
            XML_GetLineCapFromName = P2_LC_Custom
        Case Else
            XML_GetLineCapFromName = P2_LC_Flat
    End Select
End Function

Public Function XML_GetNameOfDashCap(ByVal srcDashCap As PD_2D_DashCap) As String
    Select Case srcDashCap
        Case P2_DC_Flat
            XML_GetNameOfDashCap = "flat"
        Case P2_DC_Square
            XML_GetNameOfDashCap = "square"
        Case P2_DC_Round
            XML_GetNameOfDashCap = "round"
        Case P2_DC_Triangle
            XML_GetNameOfDashCap = "triangle"
        Case Else
            XML_GetNameOfDashCap = "flat"
    End Select
End Function

Public Function XML_GetDashCapFromName(ByRef srcName As String) As PD_2D_DashCap
    Select Case srcName
        Case "flat"
            XML_GetDashCapFromName = P2_DC_Flat
        Case "square"
            XML_GetDashCapFromName = P2_DC_Square
        Case "round"
            XML_GetDashCapFromName = P2_DC_Round
        Case "triangle"
            XML_GetDashCapFromName = P2_DC_Triangle
        Case Else
            XML_GetDashCapFromName = P2_DC_Flat
    End Select
End Function

Public Function XML_GetNameOfLineJoin(ByVal srcLineJoin As PD_2D_LineJoin) As String
    Select Case srcLineJoin
        Case P2_LJ_Miter
            XML_GetNameOfLineJoin = "miter"
        Case P2_LJ_Bevel
            XML_GetNameOfLineJoin = "bevel"
        Case P2_LJ_Round
            XML_GetNameOfLineJoin = "round"
        Case Else
            XML_GetNameOfLineJoin = "miter"
    End Select
End Function

Public Function XML_GetLineJoinFromName(ByRef srcName As String) As PD_2D_LineJoin
    Select Case srcName
        Case "miter"
            XML_GetLineJoinFromName = P2_LJ_Miter
        Case "bevel"
            XML_GetLineJoinFromName = P2_LJ_Bevel
        Case "round"
            XML_GetLineJoinFromName = P2_LJ_Round
        Case Else
            XML_GetLineJoinFromName = P2_LJ_Miter
    End Select
End Function

Public Function XML_GetNameOfDashStyle(ByVal srcPenStyle As PD_2D_DashStyle) As String
    Select Case srcPenStyle
        Case P2_DS_Solid
            XML_GetNameOfDashStyle = "solid"
        Case P2_DS_Dash
            XML_GetNameOfDashStyle = "dash"
        Case P2_DS_Dot
            XML_GetNameOfDashStyle = "dot"
        Case P2_DS_DashDot
            XML_GetNameOfDashStyle = "dash-dot"
        Case P2_DS_DashDotDot
            XML_GetNameOfDashStyle = "dash-dot-dot"
        Case P2_DS_Custom
            XML_GetNameOfDashStyle = "custom"
        Case Else
            XML_GetNameOfDashStyle = "solid"
    End Select
End Function

Public Function XML_GetDashStyleFromName(ByRef srcName As String) As PD_2D_DashStyle
    Select Case srcName
        Case "solid"
            XML_GetDashStyleFromName = P2_DS_Solid
        Case "dash"
            XML_GetDashStyleFromName = P2_DS_Dash
        Case "dot"
            XML_GetDashStyleFromName = P2_DS_Dot
        Case "dash-dot"
            XML_GetDashStyleFromName = P2_DS_DashDot
        Case "dash-dot-dot"
            XML_GetDashStyleFromName = P2_DS_DashDotDot
        Case "custom"
            XML_GetDashStyleFromName = P2_DS_Custom
        Case Else
            XML_GetDashStyleFromName = P2_DS_Solid
    End Select
End Function

'Copy functions.  Copying one surface onto another surface does *not* perform any blending.  It performs a wholesale
' replacement of the destination bytes with the source bytes.
Public Function CopySurfaceI(ByRef dstSurface As pd2DSurface, ByVal dstX As Long, ByVal dstY As Long, ByRef srcSurface As pd2DSurface) As Boolean
    Dim srcWidth As Long, srcHeight As Long
    srcWidth = srcSurface.GetSurfaceWidth
    srcHeight = srcSurface.GetSurfaceHeight
    CopySurfaceI = PD2D_GDI.BitBltWrapper(dstSurface.GetSurfaceDC, dstX, dstY, srcWidth, srcHeight, srcSurface.GetSurfaceDC, 0, 0, vbSrcCopy)
End Function

'Whenever floating-point coordinates are used, we must use GDI+ for rendering.  This is always slower than PD2D_GDI.
Public Function CopySurfaceF(ByRef dstSurface As pd2DSurface, ByVal dstX As Single, ByVal dstY As Single, ByRef srcSurface As pd2DSurface) As Boolean
    PD2D_GDIPlus.GDIPlus_GraphicsSetCompositingMode dstSurface.GetHandle, GP_CM_SourceCopy
    CopySurfaceF = PD2D_GDIPlus.GDIPlus_DrawImageF(dstSurface.GetHandle, srcSurface.GetGdipImageHandle, dstX, dstY)
    PD2D_GDIPlus.GDIPlus_GraphicsSetCompositingMode dstSurface.GetHandle, GP_CM_SourceOver
End Function

'These crop functions are identical to the ones above, except they allow the user to control source width/height instead of
' inferring it automatically.
Public Function CopySurfaceCroppedI(ByRef dstSurface As pd2DSurface, ByVal dstX As Long, ByVal dstY As Long, ByVal cropWidth As Long, ByVal cropHeight As Long, ByRef srcSurface As pd2DSurface) As Boolean
    CopySurfaceCroppedI = PD2D_GDI.BitBltWrapper(dstSurface.GetSurfaceDC, dstX, dstY, cropWidth, cropHeight, srcSurface.GetSurfaceDC, 0, 0, vbSrcCopy)
End Function

Public Function CopySurfaceCroppedF(ByRef dstSurface As pd2DSurface, ByVal dstX As Single, ByVal dstY As Single, ByVal cropWidth As Long, ByVal cropHeight As Long, ByRef srcSurface As pd2DSurface) As Boolean
    PD2D_GDIPlus.GDIPlus_GraphicsSetCompositingMode dstSurface.GetHandle, GP_CM_SourceCopy
    CopySurfaceCroppedF = PD2D_GDIPlus.GDIPlus_DrawImageRectF(dstSurface.GetHandle, srcSurface.GetGdipImageHandle, dstX, dstY, cropWidth, cropHeight)
    PD2D_GDIPlus.GDIPlus_GraphicsSetCompositingMode dstSurface.GetHandle, GP_CM_SourceOver
End Function

'You might think we could just wrap StretchBlt here, but StretchBlt is inconsistent in its handling of alpha channels.
' GDI+ is actually pretty comparable speed-wise in nearest-neighbor mode, so this isn't a huge penalty.
Public Function CopySurfaceResizedI(ByRef dstSurface As pd2DSurface, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByRef srcSurface As pd2DSurface) As Boolean
    PD2D_GDIPlus.GDIPlus_GraphicsSetCompositingMode dstSurface.GetHandle, GP_CM_SourceCopy
    CopySurfaceResizedI = PD2D_GDIPlus.GDIPlus_DrawImageRectI(dstSurface.GetHandle, srcSurface.GetGdipImageHandle, dstX, dstY, dstWidth, dstHeight)
    PD2D_GDIPlus.GDIPlus_GraphicsSetCompositingMode dstSurface.GetHandle, GP_CM_SourceOver
End Function

Public Function CopySurfaceResizedF(ByRef dstSurface As pd2DSurface, ByVal dstX As Single, ByVal dstY As Single, ByVal dstWidth As Single, ByVal dstHeight As Single, ByRef srcSurface As pd2DSurface) As Boolean
    PD2D_GDIPlus.GDIPlus_GraphicsSetCompositingMode dstSurface.GetHandle, GP_CM_SourceCopy
    CopySurfaceResizedF = PD2D_GDIPlus.GDIPlus_DrawImageRectF(dstSurface.GetHandle, srcSurface.GetGdipImageHandle, dstX, dstY, dstWidth, dstHeight)
    PD2D_GDIPlus.GDIPlus_GraphicsSetCompositingMode dstSurface.GetHandle, GP_CM_SourceOver
End Function

Public Function CopySurfaceResizedCroppedF(ByRef dstSurface As pd2DSurface, ByVal dstX As Single, ByVal dstY As Single, ByVal dstWidth As Single, ByVal dstHeight As Single, ByRef srcSurface As pd2DSurface, ByVal srcX As Single, ByVal srcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single) As Boolean
    PD2D_GDIPlus.GDIPlus_GraphicsSetCompositingMode dstSurface.GetHandle, GP_CM_SourceCopy
    CopySurfaceResizedCroppedF = PD2D_GDIPlus.GDIPlus_DrawImageRectRectF(dstSurface.GetHandle, srcSurface.GetGdipImageHandle, dstX, dstY, dstWidth, dstHeight, srcX, srcY, srcWidth, srcHeight)
    PD2D_GDIPlus.GDIPlus_GraphicsSetCompositingMode dstSurface.GetHandle, GP_CM_SourceOver
End Function

Public Function CopySurfaceResizedCroppedI(ByRef dstSurface As pd2DSurface, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByRef srcSurface As pd2DSurface, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long) As Boolean
    PD2D_GDIPlus.GDIPlus_GraphicsSetCompositingMode dstSurface.GetHandle, GP_CM_SourceCopy
    CopySurfaceResizedCroppedI = PD2D_GDIPlus.GDIPlus_DrawImageRectRectI(dstSurface.GetHandle, srcSurface.GetGdipImageHandle, dstX, dstY, dstWidth, dstHeight, srcX, srcY, srcWidth, srcHeight)
    PD2D_GDIPlus.GDIPlus_GraphicsSetCompositingMode dstSurface.GetHandle, GP_CM_SourceOver
End Function

'Draw functions.  Given a target pd2dSurface object and a source pd2dPen, apply the pen to the surface in said shape.
Public Function DrawArcF(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal centerX As Single, ByVal centerY As Single, ByVal arcRadius As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As Boolean
    DrawArcF = PD2D_GDIPlus.GDIPlus_DrawArcF(dstSurface.GetHandle, srcPen.GetHandle, centerX, centerY, arcRadius, startAngle, sweepAngle)
End Function

Public Function DrawArcI(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal centerX As Long, ByVal centerY As Long, ByVal arcRadius As Long, ByVal startAngle As Long, ByVal sweepAngle As Long) As Boolean
    DrawArcI = PD2D_GDIPlus.GDIPlus_DrawArcI(dstSurface.GetHandle, srcPen.GetHandle, centerX, centerY, arcRadius, startAngle, sweepAngle)
End Function

Public Function DrawCircleF(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal centerX As Single, ByVal centerY As Single, ByVal circleRadius As Single) As Boolean
    DrawCircleF = DrawEllipseF(dstSurface, srcPen, centerX - circleRadius, centerY - circleRadius, circleRadius * 2, circleRadius * 2)
End Function

Public Function DrawCircleI(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal centerX As Long, ByVal centerY As Long, ByVal circleRadius As Long) As Boolean
    DrawCircleI = DrawEllipseI(dstSurface, srcPen, centerX - circleRadius, centerY - circleRadius, circleRadius * 2, circleRadius * 2)
End Function

Public Function DrawEllipseF(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal ellipseLeft As Single, ByVal ellipseTop As Single, ByVal ellipseWidth As Single, ByVal ellipseHeight As Single) As Boolean
    DrawEllipseF = PD2D_GDIPlus.GDIPlus_DrawEllipseF(dstSurface.GetHandle, srcPen.GetHandle, ellipseLeft, ellipseTop, ellipseWidth, ellipseHeight)
End Function

Public Function DrawEllipseF_AbsoluteCoords(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal ellipseLeft As Single, ByVal ellipseTop As Single, ByVal ellipseRight As Single, ByVal ellipseBottom As Single) As Boolean
    DrawEllipseF_AbsoluteCoords = PD2D.DrawEllipseF(dstSurface, srcPen, ellipseLeft, ellipseTop, ellipseRight - ellipseLeft, ellipseBottom - ellipseTop)
End Function

Public Function DrawEllipseF_FromRectF(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByRef srcRect As RectF) As Boolean
    DrawEllipseF_FromRectF = PD2D.DrawEllipseF(dstSurface, srcPen, srcRect.Left, srcRect.Top, srcRect.Width, srcRect.Height)
End Function

Public Function DrawEllipseI(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal ellipseLeft As Long, ByVal ellipseTop As Long, ByVal ellipseWidth As Long, ByVal ellipseHeight As Long) As Boolean
    DrawEllipseI = PD2D_GDIPlus.GDIPlus_DrawEllipseI(dstSurface.GetHandle, srcPen.GetHandle, ellipseLeft, ellipseTop, ellipseWidth, ellipseHeight)
End Function

Public Function DrawEllipseI_AbsoluteCoords(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal ellipseLeft As Long, ByVal ellipseTop As Long, ByVal ellipseRight As Long, ByVal ellipseBottom As Long) As Boolean
    DrawEllipseI_AbsoluteCoords = PD2D.DrawEllipseI(dstSurface, srcPen, ellipseLeft, ellipseTop, ellipseRight - ellipseLeft, ellipseBottom - ellipseTop)
End Function

Public Function DrawEllipseI_FromRectL(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByRef srcRect As RectL) As Boolean
    DrawEllipseI_FromRectL = PD2D.DrawEllipseI(dstSurface, srcPen, srcRect.Left, srcRect.Top, srcRect.Right - srcRect.Left, srcRect.Bottom - srcRect.Top)
End Function

'Drawing entire surfaces onto each other is significantly more convoluted than drawing shapes, because GDI+ Graphics
' objects do not support direct access to bits (which is actually forgivable, because a Graphics object may not have
' a raster object selected into it).  Instead, we must generate - on-the-fly - a GDI+ Image object as our
' "source image".  The surface class helps with this.
'
'Also, note that wherever possible we try to bypass GDI+ and just use GDI, which is totally sufficient for 24-bpp
' targets and/or integer-only coordinates.
Public Function DrawSurfaceI(ByRef dstSurface As pd2DSurface, ByVal dstX As Long, ByVal dstY As Long, ByRef srcSurface As pd2DSurface, Optional ByVal customOpacity As Single = 100#) As Boolean
    
    'Because this function doesn't require stretching, we can drop back to AlphaBlend for improved performance.
    ' (This is only possible because pd2D operates in the premultiplied alpha space; if it didn't, we'd be forced
    ' to use slower GDI+ calls everywhere.)
    Dim srcWidth As Long, srcHeight As Long
    srcWidth = srcSurface.GetSurfaceWidth
    srcHeight = srcSurface.GetSurfaceHeight
    
    If (srcSurface.GetSurfaceAlphaSupport Or (customOpacity <> 100)) Then
        DrawSurfaceI = AlphaBlendWrapper(dstSurface.GetSurfaceDC, dstX, dstY, srcWidth, srcHeight, srcSurface.GetSurfaceDC, 0, 0, srcWidth, srcHeight, srcSurface.GetSurfaceAlphaSupport, customOpacity * 2.55)
    Else
        DrawSurfaceI = PD2D_GDI.BitBltWrapper(dstSurface.GetSurfaceDC, dstX, dstY, srcWidth, srcHeight, srcSurface.GetSurfaceDC, 0, 0, vbSrcCopy)
    End If
    
End Function

Private Function AlphaBlendWrapper(ByVal hDstDC As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal hSrcDC As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, Optional ByVal srcIs32bpp As Boolean = True, Optional ByVal blendOpacity As Long = 255) As Boolean

    Dim abParams As Long
    
    'Use the image's current alpha channel, and blend it with the supplied customAlpha value
    If srcIs32bpp Then
        abParams = blendOpacity * &H10000 Or &H1000000
        
        'If the source is a pdDIB object, we could actually test for premultiplication here (and in fact,
        ' pdDIB provides its own AlphaBlend wrapper that handles this for us).
    
    'Ignore alpha channel, and only use the supplied customAlpha value.
    Else
        
        ' (My memory is fuzzy after so many years, but I seem to recall old versions of Windows sometimes failing
        '  to AlphaBlend if the alpha value was exactly 255 - as a failsafe, let's use 254 as necessary.
        '  TODO: test this on XP, Win 7, Win 10 to confirm behavior.)
        If (blendOpacity = 255) Then blendOpacity = 254
        abParams = (blendOpacity * &H10000)
    End If
    
    AlphaBlend hDstDC, dstX, dstY, dstWidth, dstHeight, hSrcDC, srcX, srcY, srcWidth, srcHeight, abParams
    
End Function

'Whenever floating-point coordinates are used, we must use GDI+ for rendering.  This is always slower than PD2D_GDI.
Public Function DrawSurfaceF(ByRef dstSurface As pd2DSurface, ByVal dstX As Single, ByVal dstY As Single, ByRef srcSurface As pd2DSurface, Optional ByVal customOpacity As Single = 100#) As Boolean
    
    Dim srcWidth As Long, srcHeight As Long
    srcWidth = srcSurface.GetSurfaceWidth
    srcHeight = srcSurface.GetSurfaceHeight
    
    'Custom opacity requires a totally different (and far more complicated) GDI+ function
    If (customOpacity <> 100#) Then
        DrawSurfaceF = PD2D_GDIPlus.GDIPlus_DrawImageRectRectF(dstSurface.GetHandle, srcSurface.GetGdipImageHandle, dstX, dstY, srcWidth, srcHeight, 0#, 0#, srcWidth, srcHeight, customOpacity * 0.01)
    Else
        DrawSurfaceF = PD2D_GDIPlus.GDIPlus_DrawImageF(dstSurface.GetHandle, srcSurface.GetGdipImageHandle, dstX, dstY)
    End If
    
End Function

Public Function DrawSurfaceCroppedI(ByRef dstSurface As pd2DSurface, ByVal dstX As Long, ByVal dstY As Long, ByVal cropWidth As Long, ByVal cropHeight As Long, ByRef srcSurface As pd2DSurface, ByVal srcX As Long, ByVal srcY As Long, Optional ByVal customOpacity As Single = 100#) As Boolean
    
    'Because this function doesn't require stretching, we can drop back to AlphaBlend for improved performance.
    ' (This is only possible because pd2D operates in the premultiplied alpha space; if it didn't, we'd be forced
    ' to use slower GDI+ calls everywhere.)
    If (srcSurface.GetSurfaceAlphaSupport Or (customOpacity <> 100)) Then
        DrawSurfaceCroppedI = AlphaBlendWrapper(dstSurface.GetSurfaceDC, dstX, dstY, cropWidth, cropHeight, srcSurface.GetSurfaceDC, srcX, srcY, cropWidth, cropHeight, srcSurface.GetSurfaceAlphaSupport, customOpacity * 2.55)
    Else
        DrawSurfaceCroppedI = PD2D_GDI.BitBltWrapper(dstSurface.GetSurfaceDC, dstX, dstY, cropWidth, cropHeight, srcSurface.GetSurfaceDC, srcX, srcY, vbSrcCopy)
    End If
    
End Function

Public Function DrawSurfaceCroppedF(ByRef dstSurface As pd2DSurface, ByVal dstX As Single, ByVal dstY As Single, ByVal cropWidth As Single, ByVal cropHeight As Single, ByRef srcSurface As pd2DSurface, ByVal srcX As Single, ByVal srcY As Single, Optional ByVal customOpacity As Single = 100#) As Boolean
    DrawSurfaceCroppedF = PD2D_GDIPlus.GDIPlus_DrawImageRectRectF(dstSurface.GetHandle, srcSurface.GetGdipImageHandle, dstX, dstY, cropWidth, cropHeight, srcX, srcY, cropWidth, cropHeight, customOpacity * 0.01)
End Function

Public Function DrawSurfaceResizedI(ByRef dstSurface As pd2DSurface, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByRef srcSurface As pd2DSurface, Optional ByVal customOpacity As Single = 100#) As Boolean
    
    Dim srcWidth As Long, srcHeight As Long
    srcWidth = srcSurface.GetSurfaceWidth
    srcHeight = srcSurface.GetSurfaceHeight
    
    If (customOpacity <> 100#) Then
        DrawSurfaceResizedI = PD2D_GDIPlus.GDIPlus_DrawImageRectRectI(dstSurface.GetHandle, srcSurface.GetGdipImageHandle, dstX, dstY, dstWidth, dstHeight, 0#, 0#, srcWidth, srcHeight, customOpacity * 0.01)
    Else
        DrawSurfaceResizedI = PD2D_GDIPlus.GDIPlus_DrawImageRectI(dstSurface.GetHandle, srcSurface.GetGdipImageHandle, dstX, dstY, dstWidth, dstHeight)
    End If
    
End Function

'Whenever floating-point coordinates are used, we must use GDI+ for rendering.  This is always slower than PD2D_GDI.
Public Function DrawSurfaceResizedF(ByRef dstSurface As pd2DSurface, ByVal dstX As Single, ByVal dstY As Single, ByVal dstWidth As Single, ByVal dstHeight As Single, ByRef srcSurface As pd2DSurface, Optional ByVal customOpacity As Single = 100#) As Boolean
    
    Dim srcWidth As Long, srcHeight As Long
    srcWidth = srcSurface.GetSurfaceWidth
    srcHeight = srcSurface.GetSurfaceHeight
    
    If (customOpacity <> 100#) Then
        DrawSurfaceResizedF = PD2D_GDIPlus.GDIPlus_DrawImageRectRectF(dstSurface.GetHandle, srcSurface.GetGdipImageHandle, dstX, dstY, dstWidth, dstHeight, 0#, 0#, srcWidth, srcHeight, customOpacity * 0.01)
    Else
        DrawSurfaceResizedF = PD2D_GDIPlus.GDIPlus_DrawImageRectF(dstSurface.GetHandle, srcSurface.GetGdipImageHandle, dstX, dstY, dstWidth, dstHeight)
    End If
    
End Function

Public Function DrawSurfaceResizedCroppedF(ByRef dstSurface As pd2DSurface, ByVal dstX As Single, ByVal dstY As Single, ByVal dstWidth As Single, ByVal dstHeight As Single, ByRef srcSurface As pd2DSurface, ByVal srcX As Single, ByVal srcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, Optional ByVal customOpacity As Single = 100#) As Boolean
    DrawSurfaceResizedCroppedF = PD2D_GDIPlus.GDIPlus_DrawImageRectRectF(dstSurface.GetHandle, srcSurface.GetGdipImageHandle, dstX, dstY, dstWidth, dstHeight, srcX, srcY, srcWidth, srcHeight, customOpacity * 0.01)
End Function

Public Function DrawSurfaceResizedCroppedI(ByRef dstSurface As pd2DSurface, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByRef srcSurface As pd2DSurface, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, Optional ByVal customOpacity As Single = 100#) As Boolean
    DrawSurfaceResizedCroppedI = PD2D_GDIPlus.GDIPlus_DrawImageRectRectI(dstSurface.GetHandle, srcSurface.GetGdipImageHandle, dstX, dstY, dstWidth, dstHeight, srcX, srcY, srcWidth, srcHeight, customOpacity * 0.01)
End Function

Public Function DrawSurfaceRotatedF(ByRef dstSurface As pd2DSurface, ByVal dstCenterX As Single, ByVal dstCenterY As Single, ByVal rotateAngle As Single, ByRef srcSurface As pd2DSurface, ByVal srcX As Single, ByVal srcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, Optional ByVal customOpacity As Single = 100#) As Boolean
    
    'Create a transform that describes the rotation
    Dim cTransform As pd2DTransform: Set cTransform = New pd2DTransform
    cTransform.ApplyTranslation -1 * (srcX + (srcX + srcWidth)) / 2, -1 * (srcY + (srcY + srcHeight)) / 2
    cTransform.ApplyRotation rotateAngle
    cTransform.ApplyTranslation dstCenterX, dstCenterY
    
    'Translate the corner points of the image to match.  (Note that the order of points is important; GDI+ requires points
    ' in top-left, top-right, bottom-left order, with the fourth point being optional.)
    Dim imgCorners() As PointFloat
    ReDim imgCorners(0 To 3) As PointFloat
    imgCorners(0).x = srcX
    imgCorners(0).y = srcY
    imgCorners(1).x = srcX + srcWidth
    imgCorners(1).y = srcY
    imgCorners(2).x = srcX
    imgCorners(2).y = srcY + srcHeight
    
    cTransform.ApplyTransformToPointFs VarPtr(imgCorners(0)), 3
    
    'Draw the image, using the new corner points as the destination!
    DrawSurfaceRotatedF = PD2D_GDIPlus.GDIPlus_DrawImagePointsRectF(dstSurface.GetHandle, srcSurface.GetGdipImageHandle, imgCorners, srcX, srcY, srcWidth, srcHeight, customOpacity * 0.01)
    
End Function

Public Function DrawSurfaceTransformedF(ByRef dstSurface As pd2DSurface, ByRef srcSurface As pd2DSurface, ByRef srcTransform As pd2DTransform, ByVal srcX As Single, ByVal srcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, Optional ByVal customOpacity As Single = 100#) As Boolean
    
    'Translate the corner points of the image to match.  (Note that the order of points is important; GDI+ requires points
    ' in top-left, top-right, bottom-left order, with the fourth point being optional.)
    Dim imgCorners() As PointFloat
    ReDim imgCorners(0 To 3) As PointFloat
    imgCorners(0).x = srcX
    imgCorners(0).y = srcY
    imgCorners(1).x = srcX + srcWidth - 1
    imgCorners(1).y = srcY
    imgCorners(2).x = srcX
    imgCorners(2).y = srcY + srcHeight - 1
    
    srcTransform.ApplyTransformToPointFs VarPtr(imgCorners(0)), 3
    
    'Draw the image, using the new corner points as the destination!
    DrawSurfaceTransformedF = PD2D_GDIPlus.GDIPlus_DrawImagePointsRectF(dstSurface.GetHandle, srcSurface.GetGdipImageHandle, imgCorners, srcX, srcY, srcWidth, srcHeight, customOpacity * 0.01)
    
End Function

Public Function DrawLineF(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As Boolean
    DrawLineF = PD2D_GDIPlus.GDIPlus_DrawLineF(dstSurface.GetHandle, srcPen.GetHandle, x1, y1, x2, y2)
    If (Not DrawLineF) Then InternalError "DrawLineF", "GDI+ failure"
End Function

Public Function DrawLineF_FromPtF(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByRef srcPoint1 As PointFloat, ByRef srcPoint2 As PointFloat) As Boolean
    DrawLineF_FromPtF = PD2D_GDIPlus.GDIPlus_DrawLineF(dstSurface.GetHandle, srcPen.GetHandle, srcPoint1.x, srcPoint1.y, srcPoint2.x, srcPoint2.y)
End Function

Public Function DrawLinesF_FromPtF(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal numOfPoints As Long, ByVal ptrToPtFArray As Long, Optional ByVal useCurveAlgorithm As Boolean = False, Optional ByVal curvatureTension As Single = 0.5) As Boolean
    If useCurveAlgorithm Then
        DrawLinesF_FromPtF = PD2D_GDIPlus.GDIPlus_DrawCurveF(dstSurface.GetHandle, srcPen.GetHandle, ptrToPtFArray, numOfPoints, curvatureTension)
    Else
        DrawLinesF_FromPtF = PD2D_GDIPlus.GDIPlus_DrawLinesF(dstSurface.GetHandle, srcPen.GetHandle, ptrToPtFArray, numOfPoints)
    End If
End Function

Public Function DrawLineI(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Boolean
    DrawLineI = PD2D_GDIPlus.GDIPlus_DrawLineI(dstSurface.GetHandle, srcPen.GetHandle, x1, y1, x2, y2)
End Function

Public Function DrawLineI_FromPtL(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByRef srcPoint1 As PointLong, ByRef srcPoint2 As PointLong) As Boolean
    DrawLineI_FromPtL = PD2D_GDIPlus.GDIPlus_DrawLineI(dstSurface.GetHandle, srcPen.GetHandle, srcPoint1.x, srcPoint1.y, srcPoint2.x, srcPoint2.y)
End Function

Public Function DrawLinesI_FromPtL(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal numOfPoints As Long, ByVal ptrToPtLArray As Long, Optional ByVal useCurveAlgorithm As Boolean = False, Optional ByVal curvatureTension As Single = 0.5) As Boolean
    If useCurveAlgorithm Then
        DrawLinesI_FromPtL = PD2D_GDIPlus.GDIPlus_DrawCurveI(dstSurface.GetHandle, srcPen.GetHandle, ptrToPtLArray, numOfPoints, curvatureTension)
    Else
        DrawLinesI_FromPtL = PD2D_GDIPlus.GDIPlus_DrawLinesI(dstSurface.GetHandle, srcPen.GetHandle, ptrToPtLArray, numOfPoints)
    End If
End Function

Public Function DrawPath(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByRef srcPath As pd2DPath) As Boolean
    DrawPath = PD2D_GDIPlus.GDIPlus_DrawPath(dstSurface.GetHandle, srcPen.GetHandle, srcPath.GetHandle)
End Function

'Helper function; the source path is silently cloned and transformed, leaving the original path untouched
Public Function DrawPath_Transformed(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByRef srcPath As pd2DPath, ByRef srcTransform As pd2DTransform) As Boolean
    Dim tmpPath As pd2DPath: Set tmpPath = New pd2DPath
    tmpPath.CloneExistingPath srcPath
    tmpPath.ApplyTransformation srcTransform
    DrawPath_Transformed = PD2D_GDIPlus.GDIPlus_DrawPath(dstSurface.GetHandle, srcPen.GetHandle, tmpPath.GetHandle)
End Function

Public Function DrawPolygonF(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal numOfPoints As Long, ByVal ptrToPtFArray As Long, Optional ByVal useCurveAlgorithm As Boolean = False, Optional ByVal curvatureTension As Single = 0.5) As Boolean
    If useCurveAlgorithm Then
        DrawPolygonF = PD2D_GDIPlus.GDIPlus_DrawClosedCurveF(dstSurface.GetHandle, srcPen.GetHandle, ptrToPtFArray, numOfPoints, curvatureTension)
    Else
        DrawPolygonF = PD2D_GDIPlus.GDIPlus_DrawPolygonF(dstSurface.GetHandle, srcPen.GetHandle, ptrToPtFArray, numOfPoints)
    End If
End Function

Public Function DrawPolygonI(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal numOfPoints As Long, ByVal ptrToPtLArray As Long, Optional ByVal useCurveAlgorithm As Boolean = False, Optional ByVal curvatureTension As Single = 0.5) As Boolean
    If useCurveAlgorithm Then
        DrawPolygonI = PD2D_GDIPlus.GDIPlus_DrawClosedCurveI(dstSurface.GetHandle, srcPen.GetHandle, ptrToPtLArray, numOfPoints, curvatureTension)
    Else
        DrawPolygonI = PD2D_GDIPlus.GDIPlus_DrawPolygonI(dstSurface.GetHandle, srcPen.GetHandle, ptrToPtLArray, numOfPoints)
    End If
End Function

Public Function DrawRectangleF(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal rectLeft As Single, ByVal rectTop As Single, ByVal rectWidth As Single, ByVal rectHeight As Single) As Boolean
    DrawRectangleF = PD2D_GDIPlus.GDIPlus_DrawRectF(dstSurface.GetHandle, srcPen.GetHandle, rectLeft, rectTop, rectWidth, rectHeight)
End Function

Public Function DrawRectangleF_AbsoluteCoords(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal rectLeft As Single, ByVal rectTop As Single, ByVal rectRight As Single, ByVal rectBottom As Single) As Boolean
    DrawRectangleF_AbsoluteCoords = PD2D.DrawRectangleF(dstSurface, srcPen, rectLeft, rectTop, rectRight - rectLeft, rectBottom - rectTop)
End Function

Public Function DrawRectangleF_FromRectF(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByRef srcRect As RectF) As Boolean
    DrawRectangleF_FromRectF = PD2D.DrawRectangleF(dstSurface, srcPen, srcRect.Left, srcRect.Top, srcRect.Width, srcRect.Height)
End Function

Public Function DrawRectangleI(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal rectLeft As Long, ByVal rectTop As Long, ByVal rectWidth As Long, ByVal rectHeight As Long) As Boolean
    DrawRectangleI = PD2D_GDIPlus.GDIPlus_DrawRectI(dstSurface.GetHandle, srcPen.GetHandle, rectLeft, rectTop, rectWidth, rectHeight)
End Function

Public Function DrawRectangleI_AbsoluteCoords(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByVal rectLeft As Long, ByVal rectTop As Long, ByVal rectRight As Long, ByVal rectBottom As Long) As Boolean
    DrawRectangleI_AbsoluteCoords = PD2D.DrawRectangleI(dstSurface, srcPen, rectLeft, rectTop, rectRight - rectLeft, rectBottom - rectTop)
End Function

Public Function DrawRectangleI_FromRectL(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByRef srcRect As RectL) As Boolean
    DrawRectangleI_FromRectL = PD2D.DrawRectangleI(dstSurface, srcPen, srcRect.Left, srcRect.Top, srcRect.Right - srcRect.Left, srcRect.Bottom - srcRect.Top)
End Function

Public Function DrawRoundRectangleF_FromRectF(ByRef dstSurface As pd2DSurface, ByRef srcPen As pd2DPen, ByRef srcRect As RectF, ByVal cornerRadius As Single) As Boolean
        
    'GDI+ has no internal rounded rect function, so we need to manually construct our own path.
    Dim tmpPath As pd2DPath
    Set tmpPath = New pd2DPath
    tmpPath.AddRoundedRectangle_RectF srcRect, cornerRadius
    
    DrawRoundRectangleF_FromRectF = PD2D_GDIPlus.GDIPlus_DrawPath(dstSurface.GetHandle, srcPen.GetHandle, tmpPath.GetHandle)
    
End Function

'Fill functions.  Given a target pd2dSurface and a source pd2dBrush, apply the brush to the surface in said shape.

Public Function FillCircleF(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByVal centerX As Single, ByVal centerY As Single, ByVal circleRadius As Single) As Boolean
    FillCircleF = FillEllipseF(dstSurface, srcBrush, centerX - circleRadius, centerY - circleRadius, circleRadius * 2, circleRadius * 2)
End Function

Public Function FillCircleI(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByVal centerX As Long, ByVal centerY As Long, ByVal circleRadius As Long) As Boolean
    FillCircleI = FillEllipseI(dstSurface, srcBrush, centerX - circleRadius, centerY - circleRadius, circleRadius * 2, circleRadius * 2)
End Function

Public Function FillEllipseF(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByVal ellipseLeft As Single, ByVal ellipseTop As Single, ByVal ellipseWidth As Single, ByVal ellipseHeight As Single) As Boolean
    FillEllipseF = PD2D_GDIPlus.GDIPlus_FillEllipseF(dstSurface.GetHandle, srcBrush.GetHandle, ellipseLeft, ellipseTop, ellipseWidth, ellipseHeight)
End Function

Public Function FillEllipseF_AbsoluteCoords(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByVal ellipseLeft As Single, ByVal ellipseTop As Single, ByVal ellipseRight As Single, ByVal ellipseBottom As Single) As Boolean
    FillEllipseF_AbsoluteCoords = PD2D.FillEllipseF(dstSurface, srcBrush, ellipseLeft, ellipseTop, ellipseRight - ellipseLeft, ellipseBottom - ellipseTop)
End Function

Public Function FillEllipseF_FromRectF(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByRef srcRect As RectF) As Boolean
    FillEllipseF_FromRectF = PD2D.FillEllipseF(dstSurface, srcBrush, srcRect.Left, srcRect.Top, srcRect.Width, srcRect.Height)
End Function

Public Function FillEllipseI(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByVal ellipseLeft As Long, ByVal ellipseTop As Long, ByVal ellipseWidth As Long, ByVal ellipseHeight As Long) As Boolean
    FillEllipseI = PD2D_GDIPlus.GDIPlus_FillEllipseI(dstSurface.GetHandle, srcBrush.GetHandle, ellipseLeft, ellipseTop, ellipseWidth, ellipseHeight)
End Function

Public Function FillEllipseI_AbsoluteCoords(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByVal ellipseLeft As Long, ByVal ellipseTop As Long, ByVal ellipseRight As Long, ByVal ellipseBottom As Long) As Boolean
    FillEllipseI_AbsoluteCoords = PD2D.FillEllipseI(dstSurface, srcBrush, ellipseLeft, ellipseTop, ellipseRight - ellipseLeft, ellipseBottom - ellipseTop)
End Function

Public Function FillEllipseI_FromRectL(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByRef srcRect As RectL) As Boolean
    FillEllipseI_FromRectL = PD2D.FillEllipseI(dstSurface, srcBrush, srcRect.Left, srcRect.Top, srcRect.Right - srcRect.Left, srcRect.Bottom - srcRect.Top)
End Function

Public Function FillPath(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByRef srcPath As pd2DPath) As Boolean
    FillPath = PD2D_GDIPlus.GDIPlus_FillPath(dstSurface.GetHandle, srcBrush.GetHandle, srcPath.GetHandle)
End Function

'Helper function; the source path is silently cloned and transformed, leaving the original path untouched
Public Function FillPath_Transformed(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByRef srcPath As pd2DPath, ByRef srcTransform As pd2DTransform) As Boolean
    Dim tmpPath As pd2DPath: Set tmpPath = New pd2DPath
    tmpPath.CloneExistingPath srcPath
    tmpPath.ApplyTransformation srcTransform
    FillPath_Transformed = PD2D_GDIPlus.GDIPlus_FillPath(dstSurface.GetHandle, srcBrush.GetHandle, tmpPath.GetHandle)
End Function

Public Function FillPolygonF_FromPtF(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByVal numOfPoints As Long, ByVal ptrToPtFArray As Long, Optional ByVal useCurveAlgorithm As Boolean = False, Optional ByVal curvatureTension As Single = 0.5, Optional ByVal fillMode As PD_2D_FillRule = P2_FR_Winding) As Boolean
    If useCurveAlgorithm Then
        FillPolygonF_FromPtF = PD2D_GDIPlus.GDIPlus_FillClosedCurveF(dstSurface.GetHandle, srcBrush.GetHandle, ptrToPtFArray, numOfPoints, curvatureTension, fillMode)
    Else
        FillPolygonF_FromPtF = PD2D_GDIPlus.GDIPlus_FillPolygonF(dstSurface.GetHandle, srcBrush.GetHandle, ptrToPtFArray, numOfPoints, fillMode)
    End If
End Function

Public Function FillRectangleF(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByVal rectLeft As Single, ByVal rectTop As Single, ByVal rectWidth As Single, ByVal rectHeight As Single) As Boolean
    FillRectangleF = PD2D_GDIPlus.GDIPlus_FillRectF(dstSurface.GetHandle, srcBrush.GetHandle, rectLeft, rectTop, rectWidth, rectHeight)
End Function

Public Function FillRectangleF_AbsoluteCoords(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByVal rectLeft As Single, ByVal rectTop As Single, ByVal rectRight As Single, ByVal rectBottom As Single) As Boolean
    FillRectangleF_AbsoluteCoords = PD2D.FillRectangleF(dstSurface, srcBrush, rectLeft, rectTop, rectRight - rectLeft, rectBottom - rectTop)
End Function

Public Function FillRectangleF_FromRectF(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByRef srcRect As RectF) As Boolean
    FillRectangleF_FromRectF = PD2D.FillRectangleF(dstSurface, srcBrush, srcRect.Left, srcRect.Top, srcRect.Width, srcRect.Height)
End Function

Public Function FillRectangleI(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByVal rectLeft As Long, ByVal rectTop As Long, ByVal rectWidth As Long, ByVal rectHeight As Long) As Boolean
    FillRectangleI = PD2D_GDIPlus.GDIPlus_FillRectI(dstSurface.GetHandle, srcBrush.GetHandle, rectLeft, rectTop, rectWidth, rectHeight)
End Function

Public Function FillRectangleI_AbsoluteCoords(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByVal rectLeft As Long, ByVal rectTop As Long, ByVal rectRight As Long, ByVal rectBottom As Long) As Boolean
    FillRectangleI_AbsoluteCoords = PD2D.FillRectangleI(dstSurface, srcBrush, rectLeft, rectTop, rectRight - rectLeft, rectBottom - rectTop)
End Function

Public Function FillRectangleI_FromRectF(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByRef srcRect As RectF) As Boolean
    FillRectangleI_FromRectF = PD2D.FillRectangleI(dstSurface, srcBrush, Int(srcRect.Left), Int(srcRect.Top), Int(PD2D_Math.Frac(srcRect.Left) + srcRect.Width + 0.5), Int(PD2D_Math.Frac(srcRect.Top) + srcRect.Height + 0.5))
End Function

Public Function FillRectangleI_FromRectL(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByRef srcRect As RectL) As Boolean
    FillRectangleI_FromRectL = PD2D.FillRectangleI(dstSurface, srcBrush, srcRect.Left, srcRect.Top, srcRect.Right - srcRect.Left, srcRect.Bottom - srcRect.Top)
End Function

Public Function FillRoundRectangleF(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByVal x As Single, ByVal y As Single, ByVal rectWidth As Single, ByVal rectHeight As Single, ByVal cornerRadius As Single) As Boolean
        
    'GDI+ has no internal rounded rect function, so we need to manually construct our own path.
    Dim tmpPath As pd2DPath
    Set tmpPath = New pd2DPath
    tmpPath.AddRoundedRectangle_Relative x, y, rectWidth, rectHeight, cornerRadius
    
    FillRoundRectangleF = PD2D_GDIPlus.GDIPlus_FillPath(dstSurface.GetHandle, srcBrush.GetHandle, tmpPath.GetHandle)
    
End Function

Public Function FillRoundRectangleF_FromRectF(ByRef dstSurface As pd2DSurface, ByRef srcBrush As pd2DBrush, ByRef srcRect As RectF, ByVal cornerRadius As Single) As Boolean
        
    'GDI+ has no internal rounded rect function, so we need to manually construct our own path.
    Dim tmpPath As pd2DPath
    Set tmpPath = New pd2DPath
    tmpPath.AddRoundedRectangle_RectF srcRect, cornerRadius
    
    FillRoundRectangleF_FromRectF = PD2D_GDIPlus.GDIPlus_FillPath(dstSurface.GetHandle, srcBrush.GetHandle, tmpPath.GetHandle)
    
End Function

'All pd2D classes report errors using an internal function similar to this one.
' Feel free to modify this function to better fit your project
' (for example, maybe you prefer to raise an actual error event).
'
'Note that by default, pd2D build simply dumps all error information to the Immediate window.
Private Sub InternalError(ByRef errFunction As String, ByRef errDescription As String, Optional ByVal errNum As Long = 0)
    PD2D.DEBUG_NotifyError "PD2D", errFunction, errDescription, errNum
End Sub
