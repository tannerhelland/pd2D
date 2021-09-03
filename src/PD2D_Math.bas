Attribute VB_Name = "PD2D_Math"
Option Explicit

Private Const SIGN_BIT As Long = &H80000000

'Many drawing features lean on various geometry functions
Public Const PI As Double = 3.14159265358979
Public Const PI_HALF As Double = 1.5707963267949
Public Const PI_DOUBLE As Double = 6.28318530717958
Public Const PI_DIV_180 As Double = 0.017453292519943

'Return the arctangent of two values (rise / run); unlike VB's integrated Atn() function, this return is quadrant-specific.
' (It also circumvents potential DBZ errors when horizontal.)
Public Function Atan2(ByVal y As Double, ByVal x As Double) As Double
 
    If (y = 0) And (x = 0) Then
        Atan2 = 0
        Exit Function
    End If
 
    If y > 0 Then
        If x >= y Then
            Atan2 = Atn(y / x)
        ElseIf x <= -y Then
            Atan2 = Atn(y / x) + PI
        Else
            Atan2 = PI_HALF - Atn(x / y)
        End If
    Else
        If x >= -y Then
            Atan2 = Atn(y / x)
        ElseIf x <= y Then
            Atan2 = Atn(y / x) - PI
        Else
            Atan2 = -Atn(x / y) - PI_HALF
        End If
    End If
 
End Function

'Convert a width and height pair to a new max width and height, while preserving aspect ratio
' NOTE: by default, inclusive fitting is assumed, but the user can set that parameter to false.  That can be used to
'        fit an image into a new size with no blank space, but cropping overhanging edges as necessary.)
Public Sub ConvertAspectRatio(ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByRef newWidth As Long, ByRef newHeight As Long, Optional ByVal fitInclusive As Boolean = True)
    
    Dim srcAspect As Double, dstAspect As Double
    If (srcHeight > 0) And (dstHeight > 0) Then
        srcAspect = srcWidth / srcHeight
        dstAspect = dstWidth / dstHeight
    Else
        Exit Sub
    End If
    
    Dim aspectLarger As Boolean
    aspectLarger = (srcAspect > dstAspect)
    
    'Exclusive fitting fits the opposite dimension, so simply reverse the way the dimensions are calculated
    If (Not fitInclusive) Then aspectLarger = Not aspectLarger
    
    If aspectLarger Then
        newWidth = dstWidth
        newHeight = CDbl(srcHeight / srcWidth) * newWidth
    Else
        newHeight = dstHeight
        newWidth = CDbl(srcWidth / srcHeight) * newHeight
    End If
    
End Sub

Public Sub ConvertCartesianToPolar(ByVal srcX As Double, ByVal srcY As Double, ByRef dstRadius As Double, ByRef dstAngle As Double, Optional ByVal centerX As Double = 0#, Optional ByVal centerY As Double = 0#)
    dstRadius = Sqr((srcX - centerX) * (srcX - centerX) + (srcY - centerY) * (srcY - centerY))
    dstAngle = PD2D_Math.Atan2((srcY - centerY), (srcX - centerX))
End Sub

'This function operates in DEGREES by default; see the final parameter to change.
Public Sub ConvertPolarToCartesian(ByVal srcAngle As Double, ByVal srcRadius As Double, ByRef dstX As Double, ByRef dstY As Double, Optional ByVal centerX As Double = 0#, Optional ByVal centerY As Double = 0#, Optional ByVal angleIsInRadians As Boolean = False)
    
    If (Not angleIsInRadians) Then srcAngle = DegreesToRadians(srcAngle)
    
    'Calculate the new (x, y)
    dstX = srcRadius * Cos(srcAngle)
    dstY = srcRadius * Sin(srcAngle)
    
    'Offset by the supplied center (x, y)
    dstX = dstX + centerX
    dstY = dstY + centerY

End Sub

Public Function Frac(ByVal srcValue As Double) As Double
    Frac = srcValue - Int(srcValue)
End Function

Public Function GetBoundaryRectOfArbitraryPoints(ByRef listOfPoints() As PointFloat) As RectF
    
    Dim minX As Single, maxX As Single, minY As Single, maxY As Single
    minX = 9999999.9
    maxX = -9999999.9
    minY = 9999999.9
    maxY = -9999999.9
    
    Dim i As Long
    For i = LBound(listOfPoints) To UBound(listOfPoints)
        With listOfPoints(i)
        If (.x < minX) Then minX = .x
        If (.x > maxX) Then maxX = .x
        If (.y < minY) Then minY = .y
        If (.y > maxY) Then maxY = .y
        End With
    Next i
    
    With GetBoundaryRectOfArbitraryPoints
        .Left = minX
        .Top = minY
        .Width = maxX - minX
        .Height = maxY - minY
    End With

End Function

'Given a RectF object, enlarge the boundaries to produce an integer-only RectF that is guaranteed
' to encompass the entire original rect.  (Said another way, the modified rect will *never* be smaller
' than the original rect.)
Public Sub GetIntClampedRectF(ByRef srcRectF As RectF)
    Dim xOffset As Single, yOffset As Single
    xOffset = srcRectF.Left - Int(srcRectF.Left)
    yOffset = srcRectF.Top - Int(srcRectF.Top)
    srcRectF.Left = Int(srcRectF.Left)
    srcRectF.Top = Int(srcRectF.Top)
    srcRectF.Width = Int(srcRectF.Width + xOffset + 0.999999999999999)
    srcRectF.Height = Int(srcRectF.Height + yOffset + 0.999999999999999)
End Sub

'Max/min functions
Public Function Max2Float_Single(ByVal f1 As Single, ByVal f2 As Single) As Single
    If (f1 > f2) Then Max2Float_Single = f1 Else Max2Float_Single = f2
End Function

Public Function Max2Int(ByVal l1 As Long, ByVal l2 As Long) As Long
    If (l1 > l2) Then Max2Int = l1 Else Max2Int = l2
End Function

'Return the maximum value from an arbitrary list of floating point values
Public Function MaxArbitraryListF(ParamArray listOfValues() As Variant) As Double
    
    If (UBound(listOfValues) >= LBound(listOfValues)) Then
                    
        Dim i As Long, numOfPoints As Long
        numOfPoints = (UBound(listOfValues) - LBound(listOfValues)) + 1
        
        Dim maxValue As Double
        maxValue = listOfValues(0)
        
        If (numOfPoints > 1) Then
            For i = 1 To numOfPoints - 1
                If listOfValues(i) > maxValue Then maxValue = listOfValues(i)
            Next i
        Else
            MaxArbitraryListF = listOfValues(LBound(listOfValues))
        End If
        
        MaxArbitraryListF = maxValue
        
    Else
        Debug.Print "No points provided - MaxArbitraryListF() function failed!"
    End If
        
End Function

Public Function Min2Float_Single(ByVal f1 As Single, ByVal f2 As Single) As Single
    If (f1 < f2) Then Min2Float_Single = f1 Else Min2Float_Single = f2
End Function

Public Function Min2Int(ByVal l1 As Long, ByVal l2 As Long) As Long
    If (l1 < l2) Then Min2Int = l1 Else Min2Int = l2
End Function

'Return the minimum value from an arbitrary list of floating point values
Public Function MinArbitraryListF(ParamArray listOfValues() As Variant) As Double
    
    If (UBound(listOfValues) >= LBound(listOfValues)) Then
                    
        Dim i As Long, numOfPoints As Long
        numOfPoints = (UBound(listOfValues) - LBound(listOfValues)) + 1
        
        Dim minValue As Double
        minValue = listOfValues(0)
        
        If (numOfPoints > 1) Then
            For i = 1 To numOfPoints - 1
                If listOfValues(i) < minValue Then minValue = listOfValues(i)
            Next i
        Else
            MinArbitraryListF = listOfValues(LBound(listOfValues))
        End If
        
        MinArbitraryListF = minValue
        
    Else
        Debug.Print "No points provided - MinArbitraryListF() function failed!"
    End If
        
End Function

'This is a modified modulo function; it handles negative values specially to ensure they work with certain distort functions
Public Function Modulo(ByVal Quotient As Double, ByVal Divisor As Double) As Double
    Modulo = Quotient - Fix(Quotient / Divisor) * Divisor
    If Modulo < 0 Then Modulo = Modulo + Divisor
End Function

Public Function RadiansToDegrees(ByVal srcRadian As Double) As Double
    RadiansToDegrees = (srcRadian * 180) / PI
End Function

Public Function DegreesToRadians(ByVal srcDegrees As Double) As Double
    DegreesToRadians = (srcDegrees * PI) / 180
End Function

'Safe unsigned addition, regardless of compilation options (e.g. compiling to native code with
' overflow ignored negates the need for this, but we sometimes use it "just in case").
' With thanks to vbforums user Krool for the original implementation: http://www.vbforums.com/showthread.php?698563-CommonControls-(Replacement-of-the-MS-common-controls)
Public Function UnsignedAdd(ByVal baseValue As Long, ByVal amtToAdd As Long) As Long
    UnsignedAdd = ((baseValue Xor SIGN_BIT) + amtToAdd) Xor SIGN_BIT
End Function

