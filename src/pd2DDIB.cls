VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pd2DDIB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon DIB Wrapper class (formerly known as "pdLayer")
'Copyright 2012-2016 by Tanner Helland
'Created: 29/August/12
'Last updated: 12/June/16
'Last update: clean-up some esoteric functions by moving them to the DIB_Support module.
'
'This class is the core of PhotoDemon.  Any image interactions are handled through this lightweight Windows DIB wrapper, and as you
' might expect, a number of other capabilities are built into the class.
'
'Note that anything you can do with an hDC property you can do with this class - simply use the getDIBDC function to
' return the DIB's hDC, then do with it what you please. All functions are heavily commented and should be self-explanatory.
'
'SPECIAL THANKS
'In building this class, I utilized a number of other DIB classes for reference and testing. Special thanks to:
' Carles PV's iBMP project: http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=42376&lngWId=1
' Steve McMahon's DIB/SafeArray analysis: http://www.vbaccelerator.com/home/VB/Code/vbMedia/DIB_Sections/True_Colour_DIBSection/article.asp
'
'All source code in this file is licensed under a modified BSD license. This means you may use the code in your own
' projects IF you provide attribution. For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'DIB Types
'In PD, GDI bitmaps are declared inside the GDI module; to use this as a standalone class, uncomment this type dec
'Private Type GDI_Bitmap
'    Type As Long
'    Width As Long
'    Height As Long
'    WidthBytes As Long
'    Planes As Integer
'    BitsPerPixel As Integer
'    Bits As Long
'End Type

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

Private Type BITMAPFILEHEADER
    Type As Integer
    Size As Long
    Reserved1 As Integer
    Reserved2 As Integer
    OffBits As Long
End Type

'Rectangle type for use with SetRect/FillRect API calls
Private Type RECT
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

'Drawing API functions
Private Declare Function BitBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal hSrcDC As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal rastOp As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDestDC As Long, ByVal nStretchMode As Long) As Long
Private Const STRETCHBLT_COLORONCOLOR As Long = 3
Private Const STRETCHBLT_HALFTONE As Long = 4

'SafeArray API functions
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal byteLength As Long)
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (ByVal dstPointer As Long, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Function VarPtrArray Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long

'DIB API functions
' (Note that these are currently declared in FastDrawing as well)
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, lpBitsInfo As BITMAPINFOHEADER, ByVal wUsage As Long, lpBits As Long, ByVal Handle As Long, ByVal dw As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectW" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long

'Object API functions
Private Const OBJ_BITMAP As Long = 7
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetObjectType Lib "gdi32" (ByVal hgdiobj As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

'Rectangle objects to be used with brushes
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long

'Brush creation
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

'Convert a system color (such as "button face" or "inactive window") to a literal RGB value
Private Declare Function OleTranslateColor Lib "olepro32" (ByVal oColor As OLE_COLOR, ByVal HPALETTE As Long, ByRef cColorRef As Long) As Long

'AlphaBlend API call
Private Declare Function AlphaBlend Lib "msimg32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long, ByVal blendFunct As Long) As Boolean

'Variables related to the DIB
Private m_dibDC As Long                 'hDC for this DIB
Private m_dibHandle As Long             'Actual DIB handle for this DIB
Private m_dibHandleOriginal As Long     'Original handle when this DIB is first created (we must store this so we can properly clean up the DIB when we're finished)
Private m_dibBits As Long               'Pointer to the actual DIB pixel bits
Private m_dibHeader As BITMAPINFOHEADER 'Persistent DIB header; this will be used with WAPI to create the DIB initially

'The DIB's width and height
Private m_dibWidth As Long, m_dibHeight As Long

'The DIB's array width (m_dibWidth * 4 for 32-bit, varies for 24-bit due to DWORD-alignment)
Private m_dibStride As Long

'The DIB's color depth (should only ever be 24 or 32)
Private m_dibColorDepth As Long

'DIBs created from loaded image files contain resolution data.  This data isn't important on a per-DIB basis, but the parent
' pdImage object will copy the resolution data from its first-loaded child DIB (if it has a resolution)
Private m_XResolution As Double, m_YResolution As Double, m_dibDPI As Double

'DIBs created from loaded image files will have their original color depth stored here.  Note that PD only works in 24/32 bpp
' mode at present, but this value may contain other bit-depths (e.g. 8bpp if the source data was a GIF).
Private m_OriginalColorDepth As Long

'If the alpha is currently premultiplied, this will be set to TRUE
Private m_IsAlphaPremultiplied As Boolean

'Get/set alpha premultiplication.
' IMPORTANT NOTE!  To make it explicitly clear that modifying this property DOES NOT ACTUALLY MODIFY THE IMAGE, the Set instruction is
' labeled differently.  It is only meant to be used by DIB creation functions, where the premultiplication state is explicitly known prior
' to writing DIB bits.  The counterpart setAlphaPremultiplication function (which is found further down in this file) will actually
' modify image bits as necessary to create the desired premultiplication state.
Public Function GetAlphaPremultiplication() As Boolean
    GetAlphaPremultiplication = m_IsAlphaPremultiplied
End Function

Public Sub SetInitialAlphaPremultiplicationState(ByVal newState As Boolean)
    m_IsAlphaPremultiplied = newState
End Sub

'Get/set original color depth.  Note that this is set after tone-mapping and other actions have been applied
Public Function GetOriginalColorDepth() As Long
    GetOriginalColorDepth = m_OriginalColorDepth
End Function

Public Sub SetOriginalColorDepth(ByVal origColorDepth As Long)
    m_OriginalColorDepth = origColorDepth
End Sub

'Set the DPI of this DIB.  This is only relevant if this DIB has been created directly from an image file.
Public Sub SetDPI(ByVal xRes As Double, ByVal yRes As Double, Optional ByVal sourceIsFreeImage As Boolean = False)
    
    'Many image types do not store resolution information; default to 96 in this case
    If xRes = 0 Then xRes = 96
    If yRes = 0 Then yRes = 96
    
    'FreeImage is pretty damn stupid when it comes to DPI.  If no DPI information is found, it will return "72" by default.
    ' If the source of this call is FreeImage, we want to replace 72 with a more modern default of 96.  Obviously this creates
    ' a problem when images have an actual resolution of 72 DPI, as there is no way to tell if FreeImage set that by default,
    ' or if it is the actual DPI value of the file.  If this occurs, we can use ExifTool to retrieve an actual value as
    ' stored in the file.
    If sourceIsFreeImage Then
        If xRes = 72 Then xRes = 96
        If yRes = 72 Then yRes = 96
    End If
    
    m_XResolution = xRes
    m_YResolution = yRes
    
    'It is extremely rare for x/y resolution to differ, but just in case, calculate an average resolution as well
    m_dibDPI = (xRes + yRes) \ 2

End Sub

'Even though we store separate x and y DPI, PD only deals in a single per-DIB DPI value
Public Function GetDPI() As Double
    GetDPI = m_dibDPI
End Function

'Given a pixel coordinate, return an RGBA quad for that coordinate.
' Returns: TRUE if the pixel lies inside DIB boundaries; FALSE otherwise.  Make sure to check this before using the RGBQUAD.
Friend Function GetPixelRGBQuad(ByVal x As Long, ByVal y As Long, ByRef dstQuad As RGBQUAD) As Boolean

    'Before doing anything else, check to see if the x/y postition lies inside the DIB
    If (x >= 0) And (x < m_dibWidth) And (y >= 0) And (y < m_dibHeight) Then
        
        'The point lies inside the DIB, which means we need to figure out the color at this position
        GetPixelRGBQuad = True
        
        Dim tmpData() As Byte, tSA As SAFEARRAY2D
        Me.WrapArrayAroundDIB tmpData, tSA
        
        Dim quickX As Long
        quickX = x * (m_dibColorDepth \ 8)
        
        'Failsafe bounds check
        If ((quickX + 2) < m_dibStride) Then
        
            With dstQuad
                .Blue = tmpData(quickX, y)
                .Green = tmpData(quickX + 1, y)
                .Red = tmpData(quickX + 2, y)
                If m_dibColorDepth = 32 Then .Alpha = tmpData(quickX + 3, y)
            End With
            
        End If
        
        Me.UnwrapArrayFromDIB tmpData
        
    'This coordinate does not lie inside the layer.
    Else
        GetPixelRGBQuad = False
    End If

End Function

'Convenience functions for wrapping an array around this DIB's bits.  You *must* call the Unwrap function prior to the array
' falling out of scope, or VB will crash.
Friend Sub WrapArrayAroundDIB(ByRef srcArray() As Byte, ByRef srcSafeArray As SAFEARRAY2D)
    PrepInternalSafeArray srcSafeArray
    CopyMemory ByVal VarPtrArray(srcArray()), VarPtr(srcSafeArray), 4&
End Sub

Public Sub UnwrapArrayFromDIB(ByRef srcArray() As Byte)
    CopyMemory ByVal VarPtrArray(srcArray), 0&, 4&
End Sub

'Force all alpha bytes in the DIB to some preset value.  If the value is *not* 255, please remember to premultiply
' alpha accordingly, if the DIB will be rendered to the screen.
Public Function ForceNewAlpha(ByVal newAlpha As Byte) As Boolean
    
    ForceNewAlpha = False
    
    'Make sure this DIB is 32bpp. If it isn't, running this function is pointless.
    If (m_dibColorDepth = 32) Then

        'Make sure this DIB isn't empty
        If (Me.GetDIBDC <> 0) And (m_dibWidth <> 0) And (m_dibHeight <> 0) Then
            
            'Loop through the image and force each alpha value to the user's specified value
            Dim iData() As Byte, tmpSA As SAFEARRAY2D
            Me.WrapArrayAroundDIB iData, tmpSA
            
            Dim x As Long, y As Long, finalX As Long
            finalX = (m_dibWidth - 1) * 4
                
            'Loop through the image, checking alphas as we go
            For y = 0 To m_dibHeight - 1
            For x = 0 To finalX Step 4
                iData(x + 3, y) = newAlpha
            Next x
            Next y
            
            Me.UnwrapArrayFromDIB iData
            ForceNewAlpha = True
            
        Else
            Debug.Print "WARNING!  The target DIB is empty (DC or Width or Height = 0).  pd2DDIB.ForceNewAlpha failed."
        End If
        
    Else
        Debug.Print "WARNING!  You cannot call pd2DDIB.ForceNewAlpha on a 24-bpp DIB!"
    End If
    
End Function

'Translate an OLE color to an RGB Long
Private Function TranslateColor(ByVal colorRef As Long) As Long
    'OleTranslateColor returns -1 if it fails; if that happens, default to white
    If OleTranslateColor(colorRef, 0, TranslateColor) Then
        TranslateColor = RGB(255, 255, 255)
    End If
End Function

Public Sub FreeFromDC()
    If (m_dibDC <> 0) Then
        SelectObject m_dibDC, m_dibHandleOriginal
        GDI.FreeMemoryDC m_dibDC
        m_dibDC = 0
    End If
End Sub

'Return this DIB's color depth
Public Function GetDIBColorDepth() As Long
    GetDIBColorDepth = m_dibColorDepth
End Function

'Return this DIB's array width
Public Function GetDIBStride() As Long
    GetDIBStride = m_dibStride
End Function

'Return this DIB's width
Public Function GetDIBWidth() As Long
    GetDIBWidth = m_dibWidth
End Function

'Return this DIB's height
Public Function GetDIBHeight() As Long
    GetDIBHeight = m_dibHeight
End Function

'Return whether or not this DIB has image data associated with it
Public Function HasImage() As Boolean
    HasImage = (m_dibHandle <> 0)
End Function

'Return whether the DIB is top-down (negative height) or bottom-up (positive height)
Public Function IsDIBTopDown() As Boolean
    IsDIBTopDown = CBool(m_dibHeader.Height < 0)
End Function

'Return this DIB's hDC.  DIBs receive a DC by default, at creation-time.  However, the caller can remove the DIB from its DC by
' calling the FreeFromDC() sub.  If you subsequently request a DC for the DIB, one will be auto-created for you.
Public Function GetDIBDC() As Long
    If (m_dibDC <> 0) Then
        GetDIBDC = m_dibDC
    Else
        If (m_dibHandle <> 0) Then
            m_dibDC = GDI.GetMemoryDC()
            m_dibHandleOriginal = SelectObject(m_dibDC, m_dibHandle)
            GetDIBDC = m_dibDC
        Else
            GetDIBDC = 0
        End If
    End If
End Function

'Return a pointer to this DIB's actual DIB
Public Function GetDIBHandle() As Long
    GetDIBHandle = m_dibHandle
End Function

'Return a pointer to this DIB's pixel data (no header)
Public Function GetDIBPointer() As Long
    GetDIBPointer = m_dibBits
End Function

'Return a pointer to an individual scanline.  For performance reasons, this function does not perform bounds checking -
' so use it carefully!
Public Function GetDIBScanline(ByVal targetScanline As Long) As Long
    GetDIBScanline = m_dibBits + (targetScanline * m_dibStride)
End Function

'Return a pointer to this DIB's header
Public Function GetDIBHeader() As Long
    GetDIBHeader = VarPtr(m_dibHeader)
End Function

Friend Sub CopyDIBHeader(ByRef dstHeader As BITMAPINFOHEADER)
    dstHeader = m_dibHeader
End Sub

'Quickly create a new DIB from an arbitrary DC.  No stretching is applied, by design; StretchBlt is unpredictable with
' 32-bpp data, so any stretching must be applied *after* creation (when PD has full control over the image data).
Public Function CreateFromDC(ByVal srcDC As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, Optional ByVal srcColorDepth As Long = 32, Optional ByVal isSourcePremultiplied As Boolean = False) As Boolean
    
    CreateFromDC = False
    
    'Make sure the DIB we're passed isn't empty
    If (srcDC <> 0) Then
        
        'If we are already at the same size and bit-depth as the source image, we can use our existing DIB and DC as-is.
        If (srcWidth = Me.GetDIBWidth) And (srcHeight = Me.GetDIBHeight) And (srcColorDepth = Me.GetDIBColorDepth) Then
            BitBlt Me.GetDIBDC, 0, 0, srcWidth, srcHeight, srcDC, srcX, srcY, vbSrcCopy
            If (srcColorDepth = 32) Then m_IsAlphaPremultiplied = isSourcePremultiplied Else m_IsAlphaPremultiplied = False
            CreateFromDC = True
        Else
        
            'Create a new, blank DIB the same size as the source DIB
            If CreateBlank(srcWidth, srcHeight, srcColorDepth, 0, 255) Then
                BitBlt Me.GetDIBDC, 0, 0, srcWidth, srcHeight, srcDC, srcX, srcY, vbSrcCopy
                If (srcColorDepth = 32) Then m_IsAlphaPremultiplied = isSourcePremultiplied Else m_IsAlphaPremultiplied = False
                CreateFromDC = True
            End If
            
        End If
        
    End If
    
End Function

'Make a copy of an existing DIB
Public Function CreateFromExistingDIB(ByRef srcDIB As pd2DDIB) As Boolean
    
    CreateFromExistingDIB = False
    
    'Make sure the DIB we're passed isn't empty
    If (srcDIB.GetDIBDC <> 0) Then
            
        'If we are already at the same size and bit-depth as the source image, we can use our existing DIB and DC as-is.
        If (srcDIB.GetDIBWidth = m_dibWidth) And (srcDIB.GetDIBHeight = m_dibHeight) And (srcDIB.GetDIBColorDepth = m_dibColorDepth) Then
        
            'Copy the image data without modification
            BitBlt Me.GetDIBDC, 0, 0, m_dibWidth, m_dibHeight, srcDIB.GetDIBDC, 0, 0, vbSrcCopy
            
            'Copy the source DIB's header as well
            srcDIB.CopyDIBHeader m_dibHeader
            
            'Reset our alpha premultiplication flag to match
            m_IsAlphaPremultiplied = srcDIB.GetAlphaPremultiplication
            
            CreateFromExistingDIB = True
            
        Else
        
            'Create a new, blank DIB the same size as the source DIB
            If CreateBlank(srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, srcDIB.GetDIBColorDepth) Then
                
                'Copy the image data without modification
                BitBlt Me.GetDIBDC, 0, 0, m_dibWidth, m_dibHeight, srcDIB.GetDIBDC, 0, 0, vbSrcCopy
                
                'Reset our alpha premultiplication flag to match
                m_IsAlphaPremultiplied = srcDIB.GetAlphaPremultiplication
                
                CreateFromExistingDIB = True
                
            End If
            
        End If
    
    End If
    
End Function

'Convert this DIB to 24bpp mode
Public Function ConvertTo24bpp(Optional ByVal newBackColor As Long = vbWhite) As Boolean
    
    ConvertTo24bpp = False
    
    'Make sure this DIB is 32bpp. If it isn't, running this function pointless.
    If (m_dibColorDepth = 32) Then

        'Make sure this DIB isn't empty
        If (Me.GetDIBDC <> 0) And (m_dibWidth <> 0) And (m_dibHeight <> 0) Then
    
            'Create a temporary DIB to hold a copy of this DIB's data (because it's about to get deleted)
            Dim tmpDIB As pd2DDIB
            Set tmpDIB = New pd2DDIB
            tmpDIB.CreateFromExistingDIB Me
            
            'Composite the temporary DIB against a white background, per convention.
            tmpDIB.CompositeBackgroundColor Drawing2D.ExtractRed(newBackColor), Drawing2D.ExtractGreen(newBackColor), Drawing2D.ExtractBlue(newBackColor)
            
            'Now erase our own DIB
            Me.EraseDIB
            
            'Create a new DIB that's exactly the same size as the old one
            If Me.CreateBlank(tmpDIB.GetDIBWidth, tmpDIB.GetDIBHeight, 24) Then
                BitBlt m_dibDC, 0, 0, m_dibWidth, m_dibHeight, tmpDIB.GetDIBDC, 0, 0, vbSrcCopy
                ConvertTo24bpp = True
            End If
    
        End If
    
    'If the DIB is already 24-bpp, return TRUE as there's nothing left for us to do
    Else
        ConvertTo24bpp = True
    End If
    
End Function

'Convert this DIB to 32bpp mode
Public Function ConvertTo32bpp(Optional ByVal newTransparency As Byte = 255) As Boolean

    'Make sure this DIB isn't empty
    If (Me.GetDIBDC <> 0) And (m_dibWidth <> 0) And (m_dibHeight <> 0) Then

        Dim iData() As Byte, tmpSA As SAFEARRAY2D
        Dim x As Long, y As Long, quickX As Long
        
        Dim newTransFloat As Double
        newTransFloat = newTransparency / 255

        If (Me.GetDIBColorDepth = 32) Then
            
            'If we are already 32bpp, with premultiplied alpha, unpremultiply the temporary DIB now
            If Me.GetAlphaPremultiplication Then Me.SetAlphaPremultiplication False
            
            'Finally, we need to loop through the image and set all alpha values to 255. Otherwise the image will
            ' be completely transparent (and we don't want that!)
            PrepInternalSafeArray tmpSA
            CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
            
            'Loop through the image, setting the alpha of each pixel to 255 (opaque)
            For y = 0 To m_dibHeight - 1
            For x = 0 To m_dibWidth - 1
                quickX = x * 4
                iData(quickX + 3, y) = iData(quickX + 3, y) * newTransFloat
            Next x
            Next y

            'With our alpha channel complete, point iData() away from the DIB and deallocate it
            CopyMemory ByVal VarPtrArray(iData), 0&, 4
            
            'Per PD-convention, re-apply premultiplication now
            Me.SetAlphaPremultiplication True
            
            'Exit
            ConvertTo32bpp = True
            Exit Function
            
        Else
        
            'Create a temporary DIB to hold a copy of this DIB's data (because it's about to get deleted)
            Dim tmpDIB As pd2DDIB
            Set tmpDIB = New pd2DDIB
            tmpDIB.CreateFromExistingDIB Me
            
            'Now erase our own DIB
            EraseDIB
    
            'Create a new DIB that's exactly the same size as the old one
            If Me.CreateBlank(tmpDIB.GetDIBWidth, tmpDIB.GetDIBHeight, 32) Then
                
                'Copy the image data from the temporary DIB without modification
                BitBlt m_dibDC, 0, 0, m_dibWidth, m_dibHeight, tmpDIB.GetDIBDC, 0, 0, vbSrcCopy
                
                'Finally, we need to loop through the image and set all alpha values to 255. Otherwise the image will
                ' be completely transparent (and we don't want that!)
                PrepInternalSafeArray tmpSA
                CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
                
                'Loop through the image, setting the alpha of each pixel to 255 (opaque)
                For y = 0 To m_dibHeight - 1
                For x = 0 To m_dibWidth - 1
                    iData(x * 4 + 3, y) = newTransparency
                Next x
                Next y
                
                'With our alpha channel complete, point iData() away from the DIB and deallocate it
                CopyMemory ByVal VarPtrArray(iData), 0&, 4
                
                'If the alpha value is not 255, apply premultiplication now
                If (newTransparency <> 255) Then Me.SetAlphaPremultiplication True
                
                'Exit
                ConvertTo32bpp = True
                Exit Function
                
            End If
            
        End If
        
    End If
    
    
    'If we made it to this line, something went horribly wrong
    ConvertTo32bpp = False

End Function

'Give this DIB a picture from a standard VB picture object
Public Function CreateFromPicture(ByRef srcPicture As StdPicture, Optional forceWhiteBackground As Boolean = False) As Boolean

    'Make sure the picture we're passed isn't empty
    If (Not srcPicture Is Nothing) Then
    
        'Make sure the picture is actually a picture
        If (GetObjectType(srcPicture) = OBJ_BITMAP) Then
        
            'Select the picture's attributes into a bitmap object
            Dim tmpBitmap As GDI_Bitmap
            GetObject srcPicture.Handle, Len(tmpBitmap), tmpBitmap
            
            'Use that bitmap object to create a new, blank DIB of the same size
            Dim targetColorDepth As Long
            If (tmpBitmap.BitsPerPixel = 32) Then targetColorDepth = 32 Else targetColorDepth = 24
            If CreateBlank(tmpBitmap.Width, tmpBitmap.Height, targetColorDepth, , 255) Then
            
                'Create a new DC
                Dim tmpDC As Long
                tmpDC = GDI.GetMemoryDC()
                
                'If successful, select the object into that DC
                If (tmpDC <> 0) Then
                
                    'Temporary holder for the object selection
                    Dim oldBitmap As Long
                    oldBitmap = SelectObject(tmpDC, srcPicture.Handle)
                    
                    'Use BitBlt to copy the pixel data to this DIB
                    BitBlt m_dibDC, 0, 0, m_dibWidth, m_dibHeight, tmpDC, 0, 0, vbSrcCopy
                    
                    'Now that we have the pixel data, erase all temporary objects
                    SelectObject tmpDC, oldBitmap
                    GDI.FreeMemoryDC tmpDC
                    
                    'Finally, if the copied image contains an alpha channel (icons, PNGs, etc), it will be set against a
                    ' black background. We typically want the background to be white, so perform a composite if requested.
                    If forceWhiteBackground And (m_dibColorDepth = 32) Then CompositeBackgroundColor
                    
                    'Success!
                    CreateFromPicture = True
                    Exit Function
                    
                End If
            
                CreateFromPicture = False
            
            End If
            
            CreateFromPicture = False
        
        End If
        
        CreateFromPicture = False
    
    End If
    
    CreateFromPicture = False

End Function

'Create a blank DIB. If no colorDepth is specified, it will default to 24bpp (16 million colors, no alpha-channel).
' If the DIB is 32bpp, an optional alpha parameter can be set for the DIB.
Public Function CreateBlank(ByVal iWidth As Long, ByVal iHeight As Long, Optional ByVal colorDepth As Long = 24, Optional ByVal BackColor As Long = vbWhite, Optional ByVal dibAlpha As Long = 0) As Boolean
    
    On Error GoTo CouldNotCreateDIB
    
    'Erase any existing DIB data
    EraseDIB
    
    'PhotoDemon only supports 24 and 32 BPP at present
    If (colorDepth <> 32) And (colorDepth <> 24) Then colorDepth = 24
        
    'The back color may or may not be a system color, so translate it just in case
    If (BackColor <> vbWhite) And (BackColor <> vbBlack) Then BackColor = TranslateColor(BackColor)
    
    'Force the DIB to have a size of at least 1x1
    If (iWidth <= 0) Then iWidth = 1
    If (iHeight <= 0) Then iHeight = 1
    
    'Remember the requested color depth, width, and height
    m_dibColorDepth = colorDepth
    m_dibWidth = iWidth
    m_dibHeight = iHeight
    
    'Prepare the required header
    With m_dibHeader
        .Size = Len(m_dibHeader)
        .Planes = 1
        .BitCount = colorDepth
        .Width = iWidth
        .Height = -iHeight
        'As always, this value needs to be a multiple of four; with 32bpp that's automatic, with 24bpp it is not
        If colorDepth = 32 Then
            m_dibStride = 4 * iWidth
        Else
            m_dibStride = (iWidth * 3 + 3) And &HFFFFFFFC
        End If
        .ImageSize = m_dibStride * iHeight
    End With
    
    'If we don't already have a compatible DC, create one now
    If (Me.GetDIBDC = 0) Then
        m_dibDC = GDI.GetMemoryDC()
        
    'If we already have a DC, empty it
    Else
        
        If (m_dibHandle <> 0) Then
            SelectObject m_dibDC, m_dibHandleOriginal
            DeleteObject m_dibHandle
        End If
        
    End If
    
    If (m_dibDC <> 0) Then
    
        'Activate color management for this DC.
        ' NOTE!  Due to rejiggering of PD's color management engine, this is no longer necessary.  Color management is only applied
        '         at image load-time (to convert into sRGB), and when drawing to the screen - but when drawing to the screen, all
        '         management is handled at the destination.  For performance reasons, we no longer color-manage DIBs themselves,
        '         which allows us to move RGB/A data between them much more efficiently.
        'turnOnColorManagementForDC m_dibDC
        
        'Create a DIB
        m_dibHandle = CreateDIBSection(m_dibDC, m_dibHeader, 0, m_dibBits, 0, 0)
        
        'If successful, select the newly created dib into our DC
        If (m_dibHandle <> 0) Then
            
            'Inside debugging mode, increment a global DIB counter.
            #If DEBUGMODE = 1 Then
                If (m_dibHandle <> 0) Then g_DIBsCreated = g_DIBsCreated + 1
            #End If
            
            'We will later use m_dibHandleOriginal to clear up the memory associated with this DIB
            m_dibHandleOriginal = SelectObject(m_dibDC, m_dibHandle)
        
            'If the DIB is 24bpp, apply the background color now
            If (colorDepth = 24) Then
                
                Dim dibRect As RECT
                SetRect dibRect, 0, 0, iWidth, iHeight
            
                Dim hBrush As Long
                hBrush = CreateSolidBrush(BackColor)
                FillRect m_dibDC, dibRect, hBrush
                DeleteObject hBrush
                
                'Also, reset alpha premultiplication to false
                m_IsAlphaPremultiplied = False
                
            Else
            
                'Finally, set the backColor and/or transparency
                Dim tmpSurface As Long, tmpBrush As Long
                tmpSurface = GDI_Plus.GetGDIPlusGraphicsFromDC(m_dibDC)
                tmpBrush = GDI_Plus.GetGDIPlusSolidBrushHandle(BackColor, dibAlpha)
                GDI_Plus.GDIPlus_FillRectI tmpSurface, tmpBrush, 0, 0, iWidth, iHeight
                GDI_Plus.ReleaseGDIPlusBrush tmpBrush
                GDI_Plus.ReleaseGDIPlusGraphics tmpSurface
                
                'Because alpha premultiplication is unknown, we'll assume it is FALSE
                m_IsAlphaPremultiplied = False
                
            End If
            
        'If DIB creation failed, clear out the work we've done so far
        Else
            EraseDIB
        End If
        
    End If
    
CouldNotCreateDIB:
    
    'Return success contingent on whether we have a DIB pointer or not
    CreateBlank = CBool(m_dibHandle <> 0)
    
End Function

'This will effectively reset everything related to this DIB, including all image data. Use cautiously!
Public Sub EraseDIB(Optional ByVal alsoReleaseDC As Boolean = False)

    'If we have image data, clear it out
    If (m_dibHandle <> 0) Then
        
        If (m_dibDC <> 0) Then SelectObject m_dibDC, m_dibHandleOriginal
        DeleteObject m_dibHandle
        
        'Inside debug mode, decrement a global DIB counter
        #If DEBUGMODE = 1 Then
            g_DIBsDestroyed = g_DIBsDestroyed + 1
        #End If
    
    End If
    
    If alsoReleaseDC And (m_dibDC <> 0) Then
        GDI.FreeMemoryDC m_dibDC
        m_dibDC = 0
    End If
        
    'Reset all associated DIB section variables
    m_dibHandleOriginal = 0
    m_dibHandle = 0
    m_dibBits = 0

    'Reset DIB size
    m_dibWidth = 0
    m_dibHeight = 0
    
    'Reset alpha premultiplication
    m_IsAlphaPremultiplied = False
    
End Sub

'Maintain the DIB's dimensions, but reset all bytes to some value (typically 0).  FillMemory is used for maximum performance.
Public Sub ResetDIB(Optional ByVal fillValue As Byte = 0)

    'Retrieve a pointer to the DIB array and the array's current size
    Dim tmpDIBPointer As Long, tmpDibSize As Long
    RetrieveDIBPointerAndSize tmpDIBPointer, tmpDibSize
    
    'Erase everything
    If (tmpDIBPointer <> 0) And (tmpDibSize > 0) Then FillMemory tmpDIBPointer, tmpDibSize, fillValue
    
End Sub

'INITIALIZE class
Private Sub Class_Initialize()

    'Reset all associated DIB section variables
    m_dibDC = 0
    m_dibHandle = 0
    m_dibHandleOriginal = 0
    m_dibBits = 0
    
    'Reset DIB size
    m_dibWidth = 0
    m_dibHeight = 0
    
End Sub

'TERMINATE class
Private Sub Class_Terminate()
    EraseDIB True
End Sub

'As a helper to external functions, this sub can be used to fill a destination array with a copy the image's current data.
' Note that this is raw image bytes - no header whatsoever - so it's assumed the caller knows how to deal with the data.
Public Sub CopyImageBytesIntoStream(ByRef dstArray() As Byte)

    'Retrieve a pointer to the current DIB bits (in VB format).  Technically we could just use the m_dibBits pointer
    ' and calculate bounds manually, but this provides a foolproof way to do the same thing using existing code.
    Dim iData() As Byte
    Dim tmpSA As SAFEARRAY2D
    PrepInternalSafeArray tmpSA
    CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
    
    'Using that temporary as our source, convert the image to the supplied 1D array.
    Dim aSize As Long
    aSize = (UBound(iData, 1) + 1) * (UBound(iData, 2) + 1) - 1
    ReDim dstArray(0 To aSize) As Byte
    CopyMemory dstArray(0), iData(0, 0), aSize
    
    CopyMemory ByVal VarPtrArray(iData), 0&, 4

End Sub

'Shortcut to copyImageBytesIntoStream; this function gives a pointer to DIB byte 0, and the size of the DIB's contents.  This can be used
' to mimic byte array behavior, without the overhead of actually initializing a VB array and copying the DIB into it.
Public Sub RetrieveDIBPointerAndSize(ByRef dibPointer As Long, ByRef dibSize As Long)
    dibPointer = m_dibBits
    dibSize = m_dibStride * m_dibHeight
End Sub

'The counterpart to copyImageBytesIntoStream, above.  This function will blindly copy an array of bytes over the current
' image's data.  For it to work, the DIB must have been correctly created at the right size and color depth BEFORE CALLING
' THIS FUNCTION.  Otherwise, you will get errors.
Public Sub CopyStreamOverImageArray(ByRef srcArray() As Byte)
    CopyMemory ByVal m_dibBits, srcArray(0), UBound(srcArray) + 1
End Sub

'Sometimes this class needs to access its own DIB bits. Here's how.
Private Sub PrepInternalSafeArray(ByRef dstSafeArray As SAFEARRAY2D)
    With dstSafeArray
        .cbElements = 1
        .cDims = 2
        .cLocks = 1
        .Bounds(0).lBound = 0
        .Bounds(0).cElements = m_dibHeight
        .Bounds(1).lBound = 0
        .Bounds(1).cElements = m_dibStride
        .pvData = m_dibBits
    End With
End Sub

'2D array access is slow, so PD tries to use 1D arrays when possible.  This function initializes a base SafeArray, but the caller
' will need to manually update it to point at a desired line.
Private Sub PrepInternalSafeArray1D(ByRef dstSafeArray As SAFEARRAY1D)
    With dstSafeArray
        .cbElements = 1
        .cDims = 1
        .lBound = 0
        .cElements = m_dibStride
        .pvData = m_dibBits
    End With
End Sub

'Sometimes this class needs to access the DIB bits of other DIBs. Here's how.
Private Sub PrepExternalSafeArray(ByRef srcDIB As pd2DDIB, ByRef dstSafeArray As SAFEARRAY2D)
    With dstSafeArray
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lBound = 0
        .Bounds(0).cElements = srcDIB.GetDIBHeight
        .Bounds(1).lBound = 0
        .Bounds(1).cElements = srcDIB.GetDIBStride
        .pvData = srcDIB.GetDIBPointer
    End With
End Sub

'Pre-composite an image with an alpha-channel against a background color. Until PhotoDemon is capable of rendering transparent
' images itself, this is necessary to give transparent images a white background.
Public Sub CompositeBackgroundColor(Optional ByVal newR As Byte = 255, Optional ByVal newG As Byte = 255, Optional ByVal newB As Byte = 255)

    'This is only useful for images with alpha channels. Exit if no alpha channel is present.
    If (m_dibColorDepth <> 32) Then Exit Sub
    
'    #If DEBUGMODE = 1 Then
'        Debug.Print "Compositing DIB of size " & m_dibWidth & "x" & m_dibHeight & ".  Alpha premultiplication = " & m_IsAlphaPremultiplied & "."
'    #End If
    
    'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
    ' (Because only 32-bpp images are supported, we know each pixel takes up 4-bytes: one each for RGBA.)
    Dim finalX As Long, finalY As Long
    finalX = (m_dibWidth - 1) * 4
    finalY = (m_dibHeight - 1)
    
    '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access), so we cheat and use a
    ' 1D array, which we reset between scanlines.
    Dim scanlineSize As Long, dibPointer As Long
    scanlineSize = Me.GetDIBStride
    dibPointer = Me.GetDIBPointer
    
    Dim dibPixels() As Byte, dibSA As SAFEARRAY1D
    PrepInternalSafeArray1D dibSA
    
    Dim x As Long, y As Long
    Dim checkAlpha As Byte, tmpAlpha As Double
    
    'Because our alpha values are pre-multiplied, we can composite them against the background color via use of a look-up table.
    Dim rLookup(0 To 255) As Byte, gLookUp(0 To 255) As Byte, bLookup(0 To 255) As Byte
    
    'Populate a unique lookup table for each color, based on each possible alpha value (0 to 255)
    For x = 0 To 255
        tmpAlpha = 1 - (x / 255)
        rLookup(x) = Int(newR * tmpAlpha)
        gLookUp(x) = Int(newG * tmpAlpha)
        bLookup(x) = Int(newB * tmpAlpha)
    Next x
    
    'Loop through the image, compositing as we go
    For y = 0 To finalY
        
        'Point our 1D pixel array at the proper scanline
        dibSA.pvData = dibPointer + scanlineSize * y
        CopyMemory ByVal VarPtrArray(dibPixels()), VarPtr(dibSA), 4
        
    For x = 0 To finalX Step 4
        
        'Access the alpha data for this pixel
        checkAlpha = dibPixels(x + 3)
        
        'Ignore opaque pixels
        If (checkAlpha <> 255) Then
            
            'Handle transparent pixels specially (this improves performance)
            If (checkAlpha = 0) Then
                dibPixels(x) = newB
                dibPixels(x + 1) = newG
                dibPixels(x + 2) = newR
            Else
            
                'Use that alpha value to blend the current colors with the newly requested color
                If m_IsAlphaPremultiplied Then
                    dibPixels(x) = dibPixels(x) + bLookup(checkAlpha)
                    dibPixels(x + 1) = dibPixels(x + 1) + gLookUp(checkAlpha)
                    dibPixels(x + 2) = dibPixels(x + 2) + rLookup(checkAlpha)
                Else
                
                    'Convert the alpha value to a floating-point variable
                    tmpAlpha = checkAlpha / 255
                
                    'Use that alpha value to blend the current colors with the newly requested color
                    dibPixels(x) = Blend2Colors(dibPixels(x), newB, tmpAlpha)
                    dibPixels(x + 1) = Blend2Colors(dibPixels(x + 1), newG, tmpAlpha)
                    dibPixels(x + 2) = Blend2Colors(dibPixels(x + 2), newR, tmpAlpha)
                    
                End If
                
            End If
            
            dibPixels(x + 3) = 255
            
        End If
        
    Next x
    Next y
    
    'With our alpha channel complete, point dibPixels() away from the DIB
    CopyMemory ByVal VarPtrArray(dibPixels), 0&, 4

End Sub

'Blend byte1 w/ byte2 based on mixRatio. mixRatio is expected to be a value between 0 and 1.
Private Function Blend2Colors(ByVal Color1 As Long, ByVal Color2 As Long, ByRef mixRatio As Double) As Byte
    Blend2Colors = mixRatio * (Color1 - Color2) + Color2
End Function

'Quick and dirty function to reverse scanlines in the DIB
Public Sub ReverseScanlines(Optional ByVal updateTopDownTracker As Boolean = True)
    
    'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
    ' (Because only 32-bpp images are supported, we know each pixel takes up 4-bytes: one each for RGBA.)
    Dim finalY As Long
    finalY = (m_dibHeight - 1)
    
    '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access), so we cheat and use a
    ' 1D array, which we reset between scanlines.
    Dim scanlineSize As Long, dibPointer As Long
    scanlineSize = Me.GetDIBStride
    dibPointer = Me.GetDIBPointer
    
    Dim tmpScanLine() As Byte
    ReDim tmpScanLine(0 To scanlineSize - 1) As Byte
    
    Dim srcPointer As Long, dstPointer As Long
    
    'Loop through the image, converting alpha as we go
    Dim y As Long
    For y = 0 To (finalY \ 2)
        
        'Copy the source line to a temporary array
        dstPointer = VarPtr(tmpScanLine(0))
        srcPointer = dibPointer + scanlineSize * (finalY - y)
        CopyMemory ByVal dstPointer, ByVal srcPointer, scanlineSize
        
        'Copy the destination line over the source line
        dstPointer = srcPointer
        srcPointer = dibPointer + scanlineSize * y
        CopyMemory ByVal dstPointer, ByVal srcPointer, scanlineSize
        
        'Copy the temporary array into the destination line
        dstPointer = srcPointer
        srcPointer = VarPtr(tmpScanLine(0))
        CopyMemory ByVal dstPointer, ByVal srcPointer, scanlineSize
        
    Next y
    
    If updateTopDownTracker Then m_dibHeader.Height = m_dibHeader.Height * -1

End Sub

'Copy the alpha values from another pd2DDIB object to this one. This is useful when the alpha channel for this DIB must be lost
' during a transformation (typically something involving FreeImage or GDI+), and a temporary DIB was made to preserve the alpha
' data. If the image is currently 24bpp, this function will first convert it to 32bpp before copying the alpha data.
'
'IMPORTANT NOTE: this function does not deal with premultiplied alpha whatsoever.  It simply copies alpha data as-is.
' You must manually set the alpha premultiplication property yourself, and you must also make sure that any number of
' un/premultiply steps are applied to both DIBs to keep the alpha values properly in sync.

'IMPORTANT NOTE: to keep this function fast, NO ARRAY BOUNDS CHECKING IS DONE. Make sure that the passed pd2DDIB object is
' THE SAME SIZE (or larger) than this DIB, or you will experience critical errors.
Public Sub CopyAlphaFromExistingDIB(ByRef srcDIB As pd2DDIB)

    'This is only useful for images with alpha channels. Exit if no alpha channel is present.
    If (m_dibColorDepth <> 32) Then ConvertTo32bpp

    'Prepare a SafeArray that points to our own DIB data
    Dim iData() As Byte
    Dim tmpSA As SAFEARRAY2D
    PrepInternalSafeArray tmpSA
    CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
    
    'Now make a second array that points to the source DIB data
    Dim aData() As Byte
    Dim srcSA As SAFEARRAY2D
    PrepExternalSafeArray srcDIB, srcSA
    CopyMemory ByVal VarPtrArray(aData()), VarPtr(srcSA), 4
    
    Dim x As Long, y As Long, quickX As Long
    
    'Loop through the image, copying alpha values as we go
    For x = 0 To m_dibWidth - 1
        quickX = x * 4
    For y = 0 To m_dibHeight - 1
        
        'Access the alpha data for this pixel
        iData(quickX + 3, y) = aData(quickX + 3, y)
        
    Next y
    Next x
    
    CopyMemory ByVal VarPtrArray(iData), 0&, 4
    CopyMemory ByVal VarPtrArray(aData), 0&, 4
    
End Sub

'DIBs created from GDI+ contain pre-multiplied alpha values. These are great for high-performance blting, but terrible for
' photo editing. This routine can be used to either apply or remove premultiplied alpha from an image.
Public Sub SetAlphaPremultiplication(Optional ByVal applyPremultiplication As Boolean = False, Optional ByVal ignoreEmbeddedValue As Boolean = False)

    'This function doesn't matter if the image isn't 32bpp
    If (m_dibColorDepth <> 32) Then Exit Sub
    
    'If alpha premultiplication already matches the requested state, exit now
    If (m_IsAlphaPremultiplied = applyPremultiplication) And (Not ignoreEmbeddedValue) Then
        Debug.Print "WARNING! Alpha premultiplication = " & applyPremultiplication & " is pointless, as image is already in that state.  (" & m_dibWidth & "x" & m_dibHeight & "x" & m_dibColorDepth & ")  Abandoning request."
        Exit Sub
    End If
    
    'To simplify processing in the inner loop, we process entire scanlines at once, stepping through byte values in pixel increments.
    ' (Because only 32-bpp images are supported, we know each pixel takes up 4-bytes: one each for RGBA.)
    Dim finalX As Long, finalY As Long
    finalX = (m_dibWidth - 1) * 4
    finalY = (m_dibHeight - 1)
    
    'Premultiplication requires a lot of int/float conversions.  To speed things up, we'll use a persistent look-up table
    ' for converting single bytes on the range [0, 255] to 4-byte floats on the range [0, 1].
    Dim intToFloat() As Single
    ReDim intToFloat(0 To 255) As Single
    Dim i As Long
    For i = 0 To 255
        If applyPremultiplication Then
            intToFloat(i) = i / 255
        Else
            If i <> 0 Then intToFloat(i) = 255 / i
        End If
    Next i
    
    '2D array access is slow (because VB must apply multiplication "behind-the-scenes" on each access), so we cheat and use a
    ' 1D array, which we reset between scanlines.
    Dim scanlineSize As Long, dibPointer As Long
    scanlineSize = Me.GetDIBStride
    dibPointer = Me.GetDIBPointer
    
    Dim dibPixels() As Byte, dibSA As SAFEARRAY1D
    PrepInternalSafeArray1D dibSA
    
    Dim x As Long, y As Long, quickX As Long
    Dim r As Long, g As Long, b As Long
    Dim tmpAlpha As Byte, tmpAlphaModifier As Double
    
    'Loop through the image, converting alpha as we go
    For y = 0 To finalY
        
        'Point our 1D pixel array at the proper scanline
        dibSA.pvData = dibPointer + scanlineSize * y
        CopyMemory ByVal VarPtrArray(dibPixels()), VarPtr(dibSA), 4
        
    For x = 0 To finalX Step 4
        
        'Retrieve alpha for the current pixel
        tmpAlpha = dibPixels(x + 3)
        
        'Branch according to applying or removing premultiplication
        If applyPremultiplication Then
        
            'When applying premultiplication, we can ignore fully opaque pixels
            If tmpAlpha <> 255 Then
            
                'We can shortcut the calculation of full transparent pixels (they are made black)
                If tmpAlpha = 0 Then
                    dibPixels(x) = 0
                    dibPixels(x + 1) = 0
                    dibPixels(x + 2) = 0
                Else
            
                    b = dibPixels(x)
                    g = dibPixels(x + 1)
                    r = dibPixels(x + 2)
                    
                    tmpAlphaModifier = intToFloat(tmpAlpha)
                    
                    'Remove premultiplied values by redistributing the colors based on this pixel's alpha value
                    r = (r * tmpAlphaModifier)
                    g = (g * tmpAlphaModifier)
                    b = (b * tmpAlphaModifier)
                    
                    dibPixels(x) = b
                    dibPixels(x + 1) = g
                    dibPixels(x + 2) = r
                    
                End If
            
            End If
        
        Else
            
            'When removing premultiplication, we can ignore fully opaque and fully transparent pixels.
            ' (Note that VB doesn't short-circuit AND statements, so we manually nest the IFs.)
            If (tmpAlpha <> 255) Then
                If (tmpAlpha <> 0) Then
                
                    b = dibPixels(x)
                    g = dibPixels(x + 1)
                    r = dibPixels(x + 2)
                    
                    tmpAlphaModifier = intToFloat(tmpAlpha)
                    
                    'Remove premultiplied values by redistributing the colors based on this pixel's alpha value
                    r = (r * tmpAlphaModifier)
                    g = (g * tmpAlphaModifier)
                    b = (b * tmpAlphaModifier)
                    
                    'Unfortunately, OOB checks are necessary for malformed DIBs
                    If (r > 255) Then r = 255
                    If (g > 255) Then g = 255
                    If (b > 255) Then b = 255
                                    
                    dibPixels(x) = b
                    dibPixels(x + 1) = g
                    dibPixels(x + 2) = r
                    
                End If
            End If
        
        End If
        
    Next x
    Next y
    
    'With our alpha channel complete, point dibPixels() away from the DIB
    CopyMemory ByVal VarPtrArray(dibPixels), 0&, 4
    
    'Mark the new premultiplication state
    m_IsAlphaPremultiplied = applyPremultiplication
    
End Sub

'Note that this function is simply a wrapper to the AlphaBlend API call, meaning it expects premultiplied image data
Public Sub AlphaBlendToDC(ByVal dstDC As Long, Optional ByVal customAlpha As Long = 255, Optional ByVal dstX As Long = 0, Optional ByVal dstY As Long = 0, Optional ByVal newWidth As Long = 0, Optional ByVal newHeight As Long = 0)
    
    Dim bfParams As Long
    If (newWidth = 0) Then newWidth = m_dibWidth
    If (newHeight = 0) Then newHeight = m_dibHeight
    
    If (m_dibColorDepth = 32) Then
        
        'Use the image's current alpha channel, and blend it with the supplied customAlpha value
        bfParams = customAlpha * &H10000 Or &H1000000
        
        'Also, raise a warning if premultiplication is not set
        If (Not m_IsAlphaPremultiplied) Then
            Debug.Print "WARNING!  Premultiplied alpha state unmarked, but alphaBlendToDC being called (" & m_dibWidth & ", " & m_dibHeight & ")!"
        End If
    
    Else
        'Ignore alpha channel, and only use the supplied customAlpha value
        bfParams = (customAlpha * &H10000)
    End If
        
    AlphaBlend dstDC, dstX, dstY, newWidth, newHeight, Me.GetDIBDC, 0, 0, m_dibWidth, m_dibHeight, bfParams
    
End Sub

'Note that this function is simply a wrapper to the AlphaBlend API call, meaning it expects premultiplied image data
Public Sub AlphaBlendToDCEx(ByVal dstDC As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, Optional ByVal customAlpha As Long = 255)
    
    Dim bfParams As Long
    
    If (m_dibColorDepth = 32) Then
    
        'Use the image's current alpha channel, and blend it with the supplied customAlpha value
        bfParams = customAlpha * &H10000 Or &H1000000
        
        'Raise a warning if premultiplication is not set
        If (Not m_IsAlphaPremultiplied) Then
            Debug.Print "WARNING!  Premultiplied alpha state unmarked, but alphaBlendToDCEx being called (" & m_dibWidth & ", " & m_dibHeight & ")!"
        End If
    
    Else
        'Ignore alpha channel, and only use the supplied customAlpha value
        bfParams = (customAlpha * &H10000)
    End If
    
    AlphaBlend dstDC, dstX, dstY, dstWidth, dstHeight, Me.GetDIBDC, srcX, srcY, srcWidth, srcHeight, bfParams
    
End Sub
