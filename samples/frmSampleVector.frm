VERSION 5.00
Begin VB.Form frmSample 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "pd2D Sample Project -- github.com/tannerhelland/pd2D"
   ClientHeight    =   8190
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14325
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   546
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   955
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTest 
      Caption         =   "Go!"
      Height          =   615
      Index           =   2
      Left            =   600
      TabIndex        =   17
      Top             =   5160
      Width           =   3615
   End
   Begin VB.CheckBox chkTest2Curvature 
      BackColor       =   &H80000005&
      Caption         =   "also randomize curvature"
      Height          =   255
      Left            =   600
      TabIndex        =   15
      Top             =   4320
      Width           =   3615
   End
   Begin VB.CommandButton cmdTest2AddPolygons 
      Caption         =   "Add more polygons"
      Height          =   615
      Left            =   600
      TabIndex        =   14
      Top             =   3600
      Width           =   3615
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Go!"
      Height          =   615
      Index           =   1
      Left            =   600
      TabIndex        =   13
      Top             =   2880
      Width           =   3615
   End
   Begin VB.CheckBox chkAntialiasing 
      BackColor       =   &H80000005&
      Caption         =   "use antialiasing"
      Height          =   255
      Left            =   6360
      TabIndex        =   11
      Top             =   180
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CommandButton cmdErase 
      Cancel          =   -1  'True
      Caption         =   "Erase!"
      Height          =   615
      Left            =   600
      TabIndex        =   10
      Top             =   7440
      Width           =   3615
   End
   Begin VB.CheckBox chkTest1Complete 
      BackColor       =   &H80000005&
      Caption         =   "don't fill the shape"
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   2040
      Width           =   3615
   End
   Begin VB.CheckBox chkTest1Lines 
      BackColor       =   &H80000005&
      Caption         =   "use lines instead of curves"
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   1680
      Width           =   3615
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Stop!"
      Height          =   615
      Index           =   3
      Left            =   600
      TabIndex        =   6
      Top             =   6360
      Width           =   3615
   End
   Begin VB.Timer tmrSample 
      Enabled         =   0   'False
      Interval        =   16
      Left            =   3840
      Top             =   120
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Go!"
      Height          =   615
      Index           =   0
      Left            =   600
      TabIndex        =   4
      Top             =   960
      Width           =   3615
   End
   Begin VB.PictureBox picOutput 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7440
      Left            =   4320
      ScaleHeight     =   494
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   652
      TabIndex        =   1
      Top             =   600
      Width           =   9810
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "animated compass"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   360
      TabIndex        =   16
      Top             =   4800
      Width           =   3810
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "animated polygons"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   360
      TabIndex        =   12
      Top             =   2520
      Width           =   3810
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "erase existing drawing"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   360
      TabIndex        =   9
      Top             =   7080
      Width           =   3330
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "stop current demo"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   360
      TabIndex        =   5
      Top             =   6000
      Width           =   3330
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "animated waveform"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Width           =   3810
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "pd2D samples:"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1605
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "pd2D output:"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   4320
      TabIndex        =   0
      Top             =   120
      Width           =   1470
   End
End
Attribute VB_Name = "frmSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'pd2D Basic Sample Project
'Copyright 2016 by Tanner Helland
'Created: 22/June/16
'Last updated: 23/June/16
'Last update: continued work on initial build
'
'This small form should help you "hit the ground running" when it comes to pd2D capabilities.  Here's what you
' need to know:
'
'1) When your project starts, you must initialize a pd2D backend before doing any painting tasks.  This involves
'   placing one line of code inside Form_Load or Sub Main():
'
'   Drawing2D.StartRenderingBackend P2_DefaultBackend
'
'2) When your project ends, you need to release the backend you started inside Form_Load or Sub Main(), e.g.:
'
'   Drawing2D.StopRenderingEngine P2_DefaultBackend
'
'3) pd2D is based on a simple drawing model: a PAINTER uses PENS and BRUSHES to draw on various SURFACES.
'   For simplicity, this project declares a single painter instance at form-level.  The painter's name
'   is "m_Painter", and it handles all interactions betweens PENS, BRUSHES, and SURFACES.
'
'4) Drawing occurs on surfaces (the "pd2DSurface" class).  There are two primary ways to create a surface:
'
'   - You can wrap a surface around an existing VB object, like a picture box or form.  This lets you paint
'     directly onto that object, but you must be aware of VB behavior with properties like .AutoRedraw.
'     (For example, if .AutoRedraw is set to TRUE, you must use a line of code like
'     "PictureBox.Picture = PictureBox.Image" to force the picture box to update.)
'
'   - You can also create an unlimited number of "in-memory" surfaces.  These surfaces are not tied to
'     any on-screen object, which makes them both very fast, and capable of supporting very large sizes.
'     However, to see the contents of an in-memory surface, you will eventually need to paint it onto a
'     surface that *is* tied to the screen (like a form or picture box surface created via the first method).
'
'   In this demo, I'll use a combination of these two methods to demonstrate how "back-buffering" works.
'   Specifically, I will create two surfaces:
'
'   - An in-memory surface at the same size as the sample form's large black picture box.  This surface is called
'   "m_BackBuffer", and I will perform all painting tasks on this surface.  Because this surface is not tied to
'   an on-screen object, it never needs to synchronize with the screen -- so painting to it is instantaneous.
'
'   - To show our painting results on-screen, I will periodically copy the contents of "m_BackBuffer" into a
'   second surface, called "m_TargetPictureBox".  This surface is created by wrapping a pd2Dsurface object around
'   the main form's black picture box, using the helpful Drawing2D.QuickCreateSurfaceFromDC() function.
'
'5) When performing drawing tasks, you'll probably create lots of pens and brushes.  You never need to worry
'   about destroying these resources.  pd2D takes care of this for you.  The same goes for in-memory surfaces,
'   because pd2D has full control over those.
'
'   The one exception to the "don't care about destroying resources" rule is surfaces that are wrapped around
'   normal, on-screen VB objects like picture boxes or forms.  These surfaces need to be destroyed before the
'   underlying object is destroyed, or you may run into trouble.  (This is not normally a problem unless you are
'   creating and destroying controls as run-time -- but please be aware of it!)
'
'6) To simplify the most common pd2D tasks, I've created a lot of helper functions inside the master Drawing2D
'   module.  These functions are prefixed with "Quick", e.g. "Drawing2D.QuickCreateSolidPen", which lets you
'   create a solid-colored pen for painting in just one line of code.  You'll probably want to make use of these,
'   as they can save you some trouble over manually instantiating pens and setting individual properties one line
'   at a time.
'
'I think that's everything!  Many thanks to other talented coders whose source code inspired this sample project.
' Specifically, thank you to:
'
' - Stefan Ceriu for the original inspiration for the "animated waveform" demo (https://github.com/stefanceriu/SCSiriWaveformView)
'   Stefan's original code (MIT-licensed) is available at the link.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************


Option Explicit

'This single painter object performs all the drawing you see in this sample project.  Most projects will only ever
' need a single painter object.  (Creating new painters is basically instantaneous, so you could also create
' painters on-demand, if you prefer that approach.)
Private m_Painter As pd2DPainter

'To prevent flickering, we're not going to draw directly onto the main form's picture box.  Instead, we're going to
' draw to an invisible "in-memory" surface.  After our drawing is complete, we'll copy the entire contents of the
' in-memory image to the screen in one fell swoop.  This approach is called "double-buffering".
'
'This "m_BackBuffer" surface is the in-memory image we'll be drawing to.
Private m_BackBuffer As pd2DSurface

'As a convenience, let's give our various test demonstrations some readable names
Private Enum PD_2D_Tests
    NoTestRunning = -1
    Test1_SplineDemo = 0
    Test2_PolygonDemo = 1
    Test3_CompassDemo = 2
End Enum

#If False Then
    Private Const NoTestRunning = -1, Test1_SplineDemo = 0, Test2_PolygonDemo = 1, Test3_CompassDemo = 2
#End If

'This is the test we're currently running (if any).  The "tmrSample" timer relies on this to know what animation
' tasks it needs to perform.
Private m_ActiveTest As PD_2D_Tests

'Some of the sample animations in this project use a collection of points and other mathematical data.
' You can change these constants to modify the animations.
Private Const DEFAULT_ANIMATION_SHAPES As Long = 20
Private m_NumOfPoints As Long

Private Const WAVE_FREQUENCY As Single = 2#
Private Const NUMBER_OF_WAVES As Long = 5
Private Const WAVE_PHASE_SHIFT As Single = -0.15
Private Const WAVE_DENSITY As Single = 25#
Private Const WAVE_PRIMARY_PEN_WIDTH As Single = 2.5
Private Const WAVE_SECONDARY_PEN_WIDTH As Single = 0.5
Private Const WAVE_AMP_ADJUSTMENT As Single = 0.005
Private m_WavePhase As Single, m_WaveAmplitude As Single, m_TargetAmplitude As Single

'For the compass demo, we use the API to grab the current position of the mouse cursor, in screen coordinates.
' We also use the API to correctly note the center of the sample picture box, in screen coordinates
Private m_CompassCenterScreen As POINTLONG, m_CompassCenterClient As POINTFLOAT, m_CompassRadius As Single
Private m_CompassLinesThick As pd2DPath, m_CompassLinesThin As pd2DPath, m_CompassArrow As pd2DPath
Private Declare Function GetCursorPos Lib "user32" (ByRef dstPointL As POINTLONG) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal srcHWnd As Long, ByRef targetPoint As POINTLONG) As Long

Private m_Test1UseLines As Boolean, m_Test1DontCloseShape As Boolean
Private m_Test2UseCurvature As Boolean

'pd2D supports subpixel positioning, which means that pixels no longer need to be defined as integers.
' The POINTFLOAT type is very simple: it simply includes an x and y coordinate, both defined as Singles.
Private m_listOfPoints() As POINTFLOAT

Private Type AnimatedPolygon
    PolygonBorderWidth As Single
    PolygonCenter As POINTFLOAT
    PolygonColorBorder As Long
    PolygonColorFill As Long
    PolygonCurvature As Single
    PolygonDirection As Single
    PolygonSides As Long
    PolygonSpeed As Single
    PolygonRadius As Single
    PolygonRotation As Single
    PolygonTransform As pd2DTransform
End Type

Private m_ListOfPolygons() As AnimatedPolygon

Private Sub Form_Load()

    'Before we can do any drawing, we must always start by initializing a drawing backend.
    ' (This approach is required by GDI+, because GDI+ offloads some processing tasks to a background thread.)
    '
    'For now, the default backend and GDI+ backends are identical, so it doesn't matter which one we pick.
    Drawing2D.StartRenderingBackend P2_DefaultBackend
    
    '(Note that you also need to *stop* this rendering backend inside Form_Unload().
    
    'Next, we want the drawing library to relay any relevant debug information to the immediate window.
    ' (You can set this value to whatever you want in your own projects; performance may see a tiny improvement if
    '  debug mode is turned off.)
    Drawing2D.SetLibraryDebugMode True
    
    'Next, we need a painter instance.  Most projects will only need one painter per project.  Just like real-life,
    ' a painter can work with any number of different pens, brushes, and surfaces.
    Drawing2D.QuickCreatePainter m_Painter
    
    'Next, let's create our in-memory surface, which I'm going to refer to as our "back buffer".  We will do all our
    ' painting on *this* surface.  (Note our use of the "Quick"-prefixed functions inside the Drawing2D module.
    ' These are a nice shorthand way to perform complicated instantiation tasks.)
    Drawing2D.QuickCreateBlankSurface m_BackBuffer, picOutput.ScaleWidth, picOutput.ScaleHeight, True, True, vbBlack, 0
    
    'When drawing onto an object, pd2D prefers pixel measurements.  I always recommend setting this at design-time,
    ' but just to be safe, but we can perform a failsafe check now.
    picOutput.ScaleMode = vbPixels
    
    'Normally, that's all you need to do inside Form_Load!
    
    'For this demo project, let's perform a few other quick tasks:
    ' First, let's reset VB's random number engine.  (Some of our animated demos rely on random numbers, and they'll
    ' always look the same if we don't do re-seed the random number generator.)
    Randomize Timer
    
    'Let's also set some default module-level variables
    m_ActiveTest = NoTestRunning
    
End Sub

'Whenever the sample form is resized, we want to resize the sample output window to match.  Note that we also need to
' resize our in-memory "back buffer" surface to match the new picture box size.
Private Sub Form_Resize()
    
    'Figure out new width/height values that fill most of the form
    Dim newOutputWidth As Long, newOutputHeight As Long
    newOutputWidth = frmSample.ScaleWidth - picOutput.Left - 8
    newOutputHeight = frmSample.ScaleHeight - picOutput.Top - 8
    
    'If the form hasn't been resized to something tiny, apply the new size immediately.
    If ((newOutputWidth > 0) And (newOutputHeight > 0)) Then
    
        picOutput.Move picOutput.Left, picOutput.Top, newOutputWidth, newOutputHeight
        
        'Because we use a back buffer for drawing, we also need to recreate it to match the new picture box size.
        Drawing2D.QuickCreateBlankSurface m_BackBuffer, picOutput.ScaleWidth, picOutput.ScaleHeight, True, True, vbBlack, 0
        
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    'If an animated demo is still running, make sure to turn it off!
    tmrSample.Enabled = False
    
    'Before we shut down the rendering backend, we need to release any remaining pd2D objects.
    Set m_Painter = Nothing
    Set m_BackBuffer = Nothing
    Set m_CompassLinesThick = Nothing
    Set m_CompassLinesThin = Nothing
    Set m_CompassArrow = Nothing
    
    'Note that pd2DTransforms are hiding inside the polygon collection (each polygon has its own transform object)
    Erase m_ListOfPolygons
    
    'As the final step at shutdown time, release the rendering backend we started inside Form_Load
    Drawing2D.StopRenderingEngine P2_DefaultBackend
    
End Sub

'When switching between demos, we want to erase both our in-memory buffer and the target picture box.  Surfaces provide
' a convenient function for this, called "EraseSurfaceContents".
Private Sub EraseAllBuffers()
    m_BackBuffer.EraseSurfaceContents vbBlack, 0#
    picOutput.Cls
End Sub

'The user can also click the on-screen "erase" button to erase whenever they want
Private Sub cmdErase_Click()
    EraseAllBuffers
End Sub

'Each animated demo is tied to a different "cmdTest" button.
Private Sub cmdTest_Click(Index As Integer)
    
    Dim i As Long
    
    'When we initialize various animation properties, we're going to randomize things like coordinates
    ' across the surface of the target picture box.  As such, we'll be accessing the picture box's
    ' coordinates many times.  Cache those values locally.
    Dim picWidth As Single, picHeight As Single
    picWidth = picOutput.ScaleWidth
    picHeight = picOutput.ScaleHeight
    
    'Some tests use the "tmrSample" timer on the main window.  We mark the active test here, before we
    ' start the timer, so the timer knows which animation actions to perform.
    m_ActiveTest = Index
    
    Select Case Index
    
        'The first demo is a basic waveform demo, inspired by an MIT-licensed project originally shared on GitHub:
        ' https://github.com/stefanceriu/SCSiriWaveformView
        ' Thank you to Stefan Ceriu for the original inspiration for this sample animation
        Case Test1_SplineDemo
            
            m_TargetAmplitude = 1#
            
            'Start the animation timer!
            tmrSample.Enabled = True
            
        Case Test2_PolygonDemo
            
            'Prepare an initial set of points for the spline animation
            m_NumOfPoints = DEFAULT_ANIMATION_SHAPES
            ReDim m_ListOfPolygons(0 To m_NumOfPoints - 1) As AnimatedPolygon
            
            For i = 0 To m_NumOfPoints - 1
                Test2RandomizePolygon m_ListOfPolygons(i), i
            Next i
            
            'Start the animation timer!
            tmrSample.Enabled = True
        
        Case Test3_CompassDemo
        
            'Store the center coordinates of the sample picture box (in screen coordinates)
            m_CompassCenterScreen.x = picOutput.ScaleWidth / 2
            m_CompassCenterScreen.y = picOutput.ScaleHeight / 2
            ClientToScreen picOutput.hWnd, m_CompassCenterScreen
            
            'We're also going to create the initial "compass" image(s); once created, we only have to transform these,
            ' not re-create them from scratch.
            Test3InitializeCompass
            
            'Start the animation timer!
            tmrSample.Enabled = True
        
        Case Else
            m_ActiveTest = NoTestRunning
            tmrSample.Enabled = False
    
    End Select

End Sub

Private Sub Test2RandomizePolygon(ByRef dstPolygon As AnimatedPolygon, ByVal polygonIndex As Long)
    With dstPolygon
        .PolygonBorderWidth = 0.5 + (Rnd * 5)
        .PolygonCenter.x = Rnd * m_BackBuffer.GetSurfaceWidth
        .PolygonCenter.y = Rnd * m_BackBuffer.GetSurfaceHeight
        .PolygonColorBorder = Rnd * 16777216
        .PolygonColorFill = Rnd * 16777216
        .PolygonCurvature = Rnd
        .PolygonDirection = Rnd * 360
        .PolygonRadius = 10# + (Rnd * 50)
        .PolygonRotation = (Rnd * 2 - 1#) * 5
        .PolygonSides = 3 + (polygonIndex Mod 6)
        .PolygonSpeed = 0.001 + (Rnd * 5)
        Set .PolygonTransform = New pd2DTransform
        .PolygonTransform.Reset
    End With
End Sub

Private Sub Test3InitializeCompass()
    
    'Start by finding the smallest dimension of the output picture box.  We'll use this to define the "radius" of our compass.
    If (picOutput.ScaleWidth < picOutput.ScaleHeight) Then m_CompassRadius = picOutput.ScaleWidth / 2 Else m_CompassRadius = picOutput.ScaleHeight / 2
    m_CompassRadius = m_CompassRadius - 1
    
    m_CompassCenterClient.x = picOutput.ScaleWidth / 2
    m_CompassCenterClient.y = picOutput.ScaleHeight / 2
    
    Set m_CompassArrow = New pd2DPath
    
    'The arrow image is easiest: just a small arrow, pointing at angle 0
    Dim arrowPoints() As Double
    ReDim arrowPoints(0 To 5) As Double
    DrawingMath.ConvertPolarToCartesian 0#, m_CompassRadius - 1#, arrowPoints(0), arrowPoints(1), m_CompassCenterClient.x, m_CompassCenterClient.y, False
    DrawingMath.ConvertPolarToCartesian -3#, m_CompassRadius - 12#, arrowPoints(2), arrowPoints(3), m_CompassCenterClient.x, m_CompassCenterClient.y, False
    DrawingMath.ConvertPolarToCartesian 3#, m_CompassRadius - 12#, arrowPoints(4), arrowPoints(5), m_CompassCenterClient.x, m_CompassCenterClient.y, False
    With m_CompassArrow
        .AddTriangle arrowPoints(0), arrowPoints(1), arrowPoints(2), arrowPoints(3), arrowPoints(4), arrowPoints(5)
    End With
    
    'To ensure that our arrow sits "outside" the compass lines, shrink the compass radius by the arrow's size (plus some padding)
    m_CompassRadius = m_CompassRadius - (12# + 10#)
    
    'Compass lines themselves are a bit more complicated; we're basically going to use polar coordinates to simplify their creation
    Set m_CompassLinesThick = New pd2DPath
    Set m_CompassLinesThin = New pd2DPath
    
    Dim lineX1 As Double, lineY1 As Double, lineX2 As Double, lineY2 As Double
    Dim lineLengthThick As Single, lineLengthThin As Single
    lineLengthThick = m_CompassRadius * 0.35
    lineLengthThin = m_CompassRadius * 0.3
    
    Dim i As Long, j As Long
    For i = 0 To 359 Step 30
        DrawingMath.ConvertPolarToCartesian i, m_CompassRadius, lineX1, lineY1, m_CompassCenterClient.x, m_CompassCenterClient.y
        DrawingMath.ConvertPolarToCartesian i, m_CompassRadius - lineLengthThick, lineX2, lineY2, m_CompassCenterClient.x, m_CompassCenterClient.y
        
        'Normally, a path object auto-connects neighboring figures.  To prevent this, we clearly mark each line as an
        ' independent figure, which prevents the auto-connect behavior.
        m_CompassLinesThick.StartNewFigure
        m_CompassLinesThick.AddLine lineX1, lineY1, lineX2, lineY2
        m_CompassLinesThick.CloseCurrentFigure
        
        'While here, let's also fill-in the thin compass lines.  These are important for demonstrating the benefits of antialiasing.
        For j = i To i + 29 Step 3
            DrawingMath.ConvertPolarToCartesian j, m_CompassRadius, lineX1, lineY1, m_CompassCenterClient.x, m_CompassCenterClient.y
            DrawingMath.ConvertPolarToCartesian j, m_CompassRadius - lineLengthThin, lineX2, lineY2, m_CompassCenterClient.x, m_CompassCenterClient.y
            m_CompassLinesThin.StartNewFigure
            m_CompassLinesThin.AddLine lineX1, lineY1, lineX2, lineY2
            m_CompassLinesThin.CloseCurrentFigure
        Next j
        
    Next i
    
    'Finally, let's add a small cross in the center of the compass, to help orient the user
    m_CompassLinesThin.StartNewFigure
    m_CompassLinesThin.AddLine m_CompassCenterClient.x - lineLengthThin / 2, m_CompassCenterClient.y, m_CompassCenterClient.x + lineLengthThin / 2, m_CompassCenterClient.y
    m_CompassLinesThin.CloseCurrentFigure
    
    m_CompassLinesThin.StartNewFigure
    m_CompassLinesThin.AddLine m_CompassCenterClient.x, m_CompassCenterClient.y - lineLengthThin / 2, m_CompassCenterClient.x, m_CompassCenterClient.y + lineLengthThin / 2
    m_CompassLinesThin.CloseCurrentFigure
    
End Sub

'During animated demonstrations, this sample timer will perform animation tasks (like moving polygon points around).
Private Sub tmrSample_Timer()
    
    'Inside this demo, we're going to be checking a lot of coordinates against the sample picture box's dimensions.
    ' Let's cache those dimensions up front, to save us some processing effort.  (Note that we can pull the dimensions
    ' from either the picture box, or the surface we've wrapped around it - doesn't matter.)
    Dim picWidth As Single, picHeight As Single
    picWidth = picOutput.ScaleWidth
    picHeight = picOutput.ScaleHeight
    
    Dim i As Long
    Dim cPen As pd2DPen, cBrush As pd2DBrush
    
    Select Case m_ActiveTest
        
        'Olaf Schimdt's "animated polygon/curve" screensaver
        Case Test1_SplineDemo
            
            'Calculate some intermediary drawing values
            Dim halfHeight As Single, halfWidth As Single
            halfHeight = picHeight * 0.5
            halfWidth = picWidth * 0.5
            
            Dim maxAmplitude As Single
            maxAmplitude = halfHeight - 4#
            
            Dim drawProgress As Single, normedAmplitude As Single, multiplier As Single
            Dim drawOpacity As Single, drawScaling As Single, penWidth As Single, penColor As Long
            Dim x As Single, y As Single
            
            'Increment our current phase
            m_WavePhase = m_WavePhase + WAVE_PHASE_SHIFT
            
            'Move toward a target amplitude (either 1.0 or -1.0), and when we reach the target, reverse direction
            If (m_WaveAmplitude < m_TargetAmplitude) Then
                m_WaveAmplitude = m_WaveAmplitude + WAVE_AMP_ADJUSTMENT
            Else
                m_WaveAmplitude = m_WaveAmplitude - WAVE_AMP_ADJUSTMENT
            End If
            
            If (Abs(m_WaveAmplitude - m_TargetAmplitude) < WAVE_AMP_ADJUSTMENT) Then m_TargetAmplitude = -1 * m_TargetAmplitude
            
            'Prepare a list of points.  These points will describe our waveform
            m_NumOfPoints = (picWidth + WAVE_DENSITY) / WAVE_DENSITY + 1
            ReDim m_listOfPoints(0 To m_NumOfPoints) As POINTFLOAT
            
            Dim curPoint As Long
            
            'Erase any existing drawing on the backbuffer
            m_BackBuffer.EraseSurfaceContents 0, 0
            
            'We're now going to draw a series of basic sine waves.  Each wave will have equal phases, but their opacity
            ' and amplitude will reduce incrementally.
            For i = 0 To NUMBER_OF_WAVES
                
                '"Progress" is a floating-point value that we use to modify a number of wave values.  Remember that only
                ' the first wave is drawn at maximum amplitude and opacity; each successive one is made smaller and more translucent.
                drawProgress = 1# - (i / NUMBER_OF_WAVES)
                normedAmplitude = (1.5 * drawProgress - 0.5) * m_WaveAmplitude
                multiplier = (drawProgress / 3# * 2#) + (1# / 3#)
                If multiplier > 1 Then multiplier = 1
                drawOpacity = multiplier
                
                'Next, we're going to calculate the actual points of this waveform.  Don't mind the math involved -
                ' it's just a bit of basic trig, scaled to fit the target picture box
                x = 0#
                curPoint = 0
                Do While x < (picWidth + WAVE_DENSITY)
                
                    'Use a parabola to scale the sine wave (we use a parabola because we want the wave's peak to
                    ' occur in the middle of the output window, then taper as it approaches either end)
                    drawScaling = 1 / halfWidth * (x - halfWidth)
                    drawScaling = -1 * (drawScaling * drawScaling) + 1
                    
                    'Calculate a y-value that corresponds with the calculate x-value
                    If (x < picWidth) Then
                        y = drawScaling * maxAmplitude * normedAmplitude * Sin(PI_DOUBLE * (x / picWidth) * WAVE_FREQUENCY + m_WavePhase) + halfHeight
                    Else
                        y = halfHeight
                    End If
                    
                    'Add this newly calculated point to our running point collection
                    m_listOfPoints(curPoint).x = x
                    m_listOfPoints(curPoint).y = y
                    
                    'Move to the next point!
                    curPoint = curPoint + 1
                    x = x + WAVE_DENSITY
                    
                Loop
                
                'Calculate a variable pen width based on which wave we are drawing.  The first wave receives full thickness;
                ' the others are incrementally smaller until they reach "WAVE_SECONDARY_PEN_WIDTH"
                If (i = 0) Then penWidth = WAVE_PRIMARY_PEN_WIDTH Else penWidth = WAVE_SECONDARY_PEN_WIDTH + ((NUMBER_OF_WAVES - i) / (NUMBER_OF_WAVES)) * (WAVE_PRIMARY_PEN_WIDTH - WAVE_SECONDARY_PEN_WIDTH)
                
                'For now, we'll use the standard rainbow colors (ROYGBIV) to render each line
                Select Case i
                    Case 0
                        penColor = RGB(210, 0, 0)
                    Case 1
                        penColor = RGB(250, 100, 35)
                    Case 2
                        penColor = RGB(250, 200, 35)
                    Case 3
                        penColor = RGB(50, 220, 0)
                    Case 4
                        penColor = RGB(20, 50, 250)
                    Case 5
                        penColor = RGB(250, 25, 250)
                    Case 6
                        penColor = RGB(90, 0, 90)
                    Case Else
                        penColor = RGB(50, 0, 50)
                
                End Select
                
                'If the user wants us to fill the shape, let's draw the fill first (at 50% opacity).  Note that the "FillPolygon"
                ' function automatically connects the first and last points in the wave, which is how we form a solid shape from
                ' an abstract set of points.
                If (Not m_Test1DontCloseShape) Then
                    Drawing2D.QuickCreateSolidBrush cBrush, penColor, drawOpacity * 50
                    m_Painter.FillPolygonF_FromPtF m_BackBuffer, cBrush, curPoint, VarPtr(m_listOfPoints(0)), Not m_Test1UseLines, , P2_FR_OddEven
                End If
                
                'And finally, trace the path outline using the pen color and width we calculated previously
                Drawing2D.QuickCreateSolidPen cPen, penWidth, penColor, drawOpacity * 100, P2_LJ_Round, P2_LC_Round
                m_Painter.DrawLinesF_FromPtF m_BackBuffer, cPen, curPoint, VarPtr(m_listOfPoints(0)), Not m_Test1UseLines
                
            Next i
            
            'Finally, copy the full contents of the "back buffer" surface onto the on-screen picture box.
            ' (Because the picture box's .AutoRedraw property is set to FALSE, we do not need to forcibly
            ' refresh the picture box after copying.)
            m_BackBuffer.CopySurfaceToDC picOutput.hDC
            
            
        Case Test2_PolygonDemo
            
            'To move each polygon, we're going to use a "transform" object.  Transform objects "add together"
            ' transformations over time, which allows you to apply very complex motion with very little code.
            Dim newCenter As POINTFLOAT, oldCenter As POINTFLOAT, collisionAngle As Double
        
            For i = 0 To m_NumOfPoints - 1
                
                With m_ListOfPolygons(i)
                
                    'Update this polygon's running transformation
                    oldCenter = .PolygonCenter
                    .PolygonTransform.ApplyTranslation_Polar .PolygonDirection, .PolygonSpeed, True
                    
                    'Find the current (x, y) centerpoint of the polygon, with all translations applied
                    newCenter = .PolygonCenter
                    .PolygonTransform.ApplyTransformToPointF newCenter
                    
                    'Use the new centerpoint to rotate the polygon around its center according to its
                    ' randomized "rotation" speed
                    .PolygonTransform.ApplyRotation .PolygonRotation, newCenter.x, newCenter.y
                    
                    'If the polygon is going to fly off an edge of the screen, adjust its angle by 90 degrees
                    If (newCenter.x + .PolygonRadius > picWidth) Then
                        .PolygonTransform.ApplyTranslation picWidth - (newCenter.x + .PolygonRadius), 0#
                        If (.PolygonDirection < 180) Then
                            .PolygonDirection = .PolygonDirection + 90
                        Else
                            .PolygonDirection = .PolygonDirection - 90
                        End If
                        .PolygonDirection = DrawingMath.Modulo(.PolygonDirection, 360#)
                    ElseIf (newCenter.y + .PolygonRadius > picHeight) Then
                        .PolygonTransform.ApplyTranslation 0#, picHeight - (newCenter.y + .PolygonRadius)
                        If (.PolygonDirection < 90) Or (.PolygonDirection > 270) Then
                            .PolygonDirection = .PolygonDirection - 90
                        Else
                            .PolygonDirection = .PolygonDirection + 90
                        End If
                        .PolygonDirection = DrawingMath.Modulo(.PolygonDirection, 360#)
                    ElseIf (newCenter.x - .PolygonRadius < 0) Then
                        .PolygonTransform.ApplyTranslation -1 * newCenter.x + .PolygonRadius, 0#
                        If (.PolygonDirection > 180) Then
                            .PolygonDirection = .PolygonDirection + 90
                        Else
                            .PolygonDirection = .PolygonDirection - 90
                        End If
                        .PolygonDirection = DrawingMath.Modulo(.PolygonDirection, 360#)
                    ElseIf (newCenter.y - .PolygonRadius < 0) Then
                        .PolygonTransform.ApplyTranslation 0#, -1 * newCenter.y + .PolygonRadius
                        If (.PolygonDirection < 90) Or (.PolygonDirection > 270) Then
                            .PolygonDirection = .PolygonDirection + 90
                        Else
                            .PolygonDirection = .PolygonDirection - 90
                        End If
                        .PolygonDirection = DrawingMath.Modulo(.PolygonDirection, 360#)
                    End If
                End With
            Next i
            
            'The actual draw operation is handled by a separate function, listed immediately below this one
            RenderTest2Animation
            
            'Finally, copy the full contents of the "back buffer" surface onto the on-screen picture box.
            ' (Because the picture box's .AutoRedraw property is set to FALSE, we do not need to forcibly
            ' refresh the picture box after performing the copy operation.)
            m_BackBuffer.CopySurfaceToDC picOutput.hDC
            
        
        Case Test3_CompassDemo
        
            'Grab the current mouse cursor position, in screen coordinates
            Dim mousePosition As POINTLONG
            GetCursorPos mousePosition
            
            'Calculate the angle between the mouse position and the center of the output picture box
            Dim compassAngle As Single
            compassAngle = DrawingMath.Atan2(mousePosition.y - m_CompassCenterScreen.y, mousePosition.x - m_CompassCenterScreen.x)
            compassAngle = DrawingMath.RadiansToDegrees(compassAngle)
            
            'Create a transformation object that describes this angle
            Dim cTransform As pd2DTransform
            Set cTransform = New pd2DTransform
            cTransform.ApplyRotation compassAngle, m_CompassCenterClient.x, m_CompassCenterClient.y
            
            'Clear the current back buffer image
            m_BackBuffer.EraseSurfaceContents 0, 0
            
            'Render the compass arrow, with this rotation applied
            Drawing2D.QuickCreateSolidBrush cBrush, vbRed, 100#
            m_Painter.FillPath_Transformed m_BackBuffer, cBrush, m_CompassArrow, cTransform
            
            'Render the compass lines, also with this rotation applied
            Drawing2D.QuickCreateSolidPen cPen, 2.5, vbWhite
            m_Painter.DrawPath_Transformed m_BackBuffer, cPen, m_CompassLinesThick, cTransform
            
            Drawing2D.QuickCreateSolidPen cPen, 0.6, vbWhite, 50#
            m_Painter.DrawPath_Transformed m_BackBuffer, cPen, m_CompassLinesThin, cTransform
            
            'Finally, copy the full contents of the "back buffer" surface onto the on-screen picture box.
            ' (Because the picture box's .AutoRedraw property is set to FALSE, we do not need to forcibly
            ' refresh the picture box after performing the copy operation.)
            m_BackBuffer.CopySurfaceToDC picOutput.hDC
            
        
        Case Else
    
    End Select
    
End Sub

Private Sub RenderTest2Animation()
    
    Dim cBrush As pd2DBrush, cPen As pd2DPen, cPath As pd2DPath
    Set cPath = New pd2DPath
    
    'Completely erase the back buffer.
    m_BackBuffer.EraseSurfaceContents 0, 0
    
    'Iterate through our current polygon collection, and draw each one in turn
    Dim i As Long
    For i = 0 To m_NumOfPoints - 1
        
        With m_ListOfPolygons(i)
            
            'Create a path object from this polygon's set of points
            cPath.ResetPath
            cPath.AddPolygon_Regular .PolygonSides, .PolygonRadius, .PolygonCenter.x, .PolygonCenter.y, m_Test2UseCurvature, .PolygonCurvature
            
            'Apply this polygon's transformation.  (The transformation is a running sum of all move and rotate operations
            ' we've applied to this polygon.)
            cPath.ApplyTransformation .PolygonTransform
            
            'Fill the polygon area using the polygon's random fill color at a fraction of its original opacity
            Drawing2D.QuickCreateSolidBrush cBrush, .PolygonColorFill, 25#
            m_Painter.FillPath m_BackBuffer, cBrush, cPath
            
            'Finish by tracing the polygon outline using the polygon's random border color at 100% opacity
            Drawing2D.QuickCreateSolidPen cPen, .PolygonBorderWidth, .PolygonColorBorder, 100#
            m_Painter.DrawPath m_BackBuffer, cPen, cPath
            
        End With
        
    Next i
        
    'Note that we don't have to free any objects here - pd2D always takes care of that for us.

End Sub

'If an animated demo supports different rendering options (via checkbox or other UI), we handle those options here.
Private Sub chkAntialiasing_Click()
    If CBool(chkAntialiasing.Value = vbChecked) Then m_BackBuffer.SetSurfaceAntialiasing P2_AA_HighQuality Else m_BackBuffer.SetSurfaceAntialiasing P2_AA_None
End Sub

Private Sub chkTest1Complete_Click()
    m_Test1DontCloseShape = CBool(chkTest1Complete.Value = vbChecked)
End Sub

Private Sub chkTest1Lines_Click()
    m_Test1UseLines = CBool(chkTest1Lines.Value = vbChecked)
End Sub

Private Sub chkTest2Curvature_Click()
    m_Test2UseCurvature = CBool(chkTest2Curvature.Value = vbChecked)
End Sub

'Clicking the "add polygons" button will add 20 more random polygons to the demonstration
Private Sub cmdTest2AddPolygons_Click()
    
    Dim oldNumOfPoints As Long
    oldNumOfPoints = m_NumOfPoints
    
    m_NumOfPoints = m_NumOfPoints + DEFAULT_ANIMATION_SHAPES
    ReDim Preserve m_ListOfPolygons(0 To m_NumOfPoints - 1) As AnimatedPolygon
    
    Dim i As Long
    For i = oldNumOfPoints To m_NumOfPoints - 1
        Test2RandomizePolygon m_ListOfPolygons(i), i
    Next i

End Sub

