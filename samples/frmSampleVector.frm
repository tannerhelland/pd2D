VERSION 5.00
Begin VB.Form frmSample 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "pd2D Sample Project -- github.com/tannerhelland/pd2D"
   ClientHeight    =   7680
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
   ScaleHeight     =   512
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   955
   StartUpPosition =   3  'Windows Default
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
      Top             =   6960
      Width           =   3615
   End
   Begin VB.CheckBox chkTest1Complete 
      BackColor       =   &H80000005&
      Caption         =   "automatically complete the shape"
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
      Index           =   2
      Left            =   600
      TabIndex        =   6
      Top             =   5880
      Width           =   3615
   End
   Begin VB.Timer tmrSample 
      Enabled         =   0   'False
      Interval        =   16
      Left            =   120
      Top             =   7080
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
      Height          =   6960
      Left            =   4320
      ScaleHeight     =   462
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   652
      TabIndex        =   1
      Top             =   600
      Width           =   9810
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
      Top             =   6600
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
      Top             =   5520
      Width           =   3330
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "curve demo (by Olaf Schmidt)"
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
'I think that's everything!  Many thanks to other talented VB coders whose source code inspired this sample project.
' Specifically, thank you to:
'
' - Olaf Schmidt for the "polygon curves" sample code (http://www.vbforums.com/showthread.php?727765-BSpline-based-quot-Bezier-Art-quot)
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
End Enum

#If False Then
    Private Const NoTestRunning = -1, Test1_SplineDemo = 0, Test2_PolygonDemo = 1
#End If

'This is the test we're currently running (if any).  The "tmrSample" timer relies on this to know what animation
' tasks it needs to perform.
Private m_ActiveTest As PD_2D_Tests

'Some of the sample animations in this project use a collection of points and other mathematical data.
' You can change this constants to modify the animations.
Private Const DEFAULT_ANIMATION_POINTS As Long = 6
Private m_NumOfPoints As Long

Private m_Test1UseLines As Boolean, m_Test1CloseShape As Boolean
Private m_Test1ColorPhase As Long, m_Test1ColorIncrement As Single
Private m_Test2UseCurvature As Boolean

'pd2D supports subpixel positioning, which means that pixels no longer need to be defined as integers.
' The POINTFLOAT type is very simple: it simply includes an x and y coordinate, both defined as Singles.
Private m_listOfPoints() As POINTFLOAT
Private m_listOfSignsX() As Long, m_listOfSignsY() As Long

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

Private Sub chkTest2Curvature_Click()
    EraseAllBuffers
    m_Test2UseCurvature = CBool(chkTest2Curvature.Value = vbChecked)
End Sub

Private Sub chkAntialiasing_Click()
    EraseAllBuffers
    If CBool(chkAntialiasing.Value = vbChecked) Then m_BackBuffer.SetSurfaceAntialiasing P2_AA_HighQuality Else m_BackBuffer.SetSurfaceAntialiasing P2_AA_None
End Sub

'Add more polygons to the demonstration
Private Sub cmdTest2AddPolygons_Click()
    
    Dim oldNumOfPoints As Long
    oldNumOfPoints = m_NumOfPoints
    
    m_NumOfPoints = m_NumOfPoints + 5
    ReDim Preserve m_ListOfPolygons(0 To m_NumOfPoints - 1) As AnimatedPolygon
    
    Dim i As Long
    For i = oldNumOfPoints To m_NumOfPoints - 1
        Test2RandomizePolygon m_ListOfPolygons(i), i
    Next i

End Sub

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
    
    'Before we exit, release the rendering backend we started inside Form_Load
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
    
        'The first demo is an animated polygon/curve demo, originally shared by Olaf Schmidt at this link:
        ' http://www.vbforums.com/showthread.php?727765-BSpline-based-quot-Bezier-Art-quot
        ' Thank you to Olaf for sharing this pretty "screensaver"-type animation.
        Case Test1_SplineDemo
        
            'Prepare an initial set of points for the spline animation
            m_NumOfPoints = DEFAULT_ANIMATION_POINTS
            ReDim m_listOfPoints(0 To m_NumOfPoints - 1) As POINTFLOAT
            ReDim m_listOfSignsX(0 To m_NumOfPoints - 1) As Long: ReDim m_listOfSignsY(0 To m_NumOfPoints - 1) As Long
            
            'These points are randomly scattered across the picture box's available area.  Note that we can read
            ' the picture box's property directly from the surface object wrapped around it.
            For i = 0 To m_NumOfPoints - 1
                m_listOfPoints(i).x = Rnd * picWidth
                m_listOfPoints(i).y = Rnd * picHeight
                
                'This animation also stores a "sign", either +1 or -1, for each point in the polygon.  When a
                ' given point hits the edge of the picture box, we'll reverse its direction.
                m_listOfSignsX(i) = IIf(i Mod 2, 1, -1)
                m_listOfSignsY(i) = IIf(i Mod 2, -1, 1)
            Next i
            
            'Start the animation timer!
            tmrSample.Enabled = True
            
        Case Test2_PolygonDemo
            
            'Prepare an initial set of points for the spline animation
            m_NumOfPoints = DEFAULT_ANIMATION_POINTS
            ReDim m_ListOfPolygons(0 To m_NumOfPoints - 1) As AnimatedPolygon
            
            For i = 0 To m_NumOfPoints - 1
                Test2RandomizePolygon m_ListOfPolygons(i), i
            Next i
            
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

'During animated demonstrations, this sample timer will perform animation tasks (like moving polygon points around).
Private Sub tmrSample_Timer()
    
    'Inside this demo, we're going to be checking a lot of coordinates against the sample picture box's dimensions.
    ' Let's cache those dimensions up front, to save us some processing effort.  (Note that we can pull the dimensions
    ' from either the picture box, or the surface we've wrapped around it - doesn't matter.)
    Dim picWidth As Single, picHeight As Single
    picWidth = picOutput.ScaleWidth
    picHeight = picOutput.ScaleHeight
    
    'Wrap a temporary pd2D surface around the output picture box.  We will release this surface at the end of this sub,
    ' because the underlying hDC is not owned by us - it's owned by VB, so we shouldn't monopolize it any longer than
    ' we absolutely have to.
    Dim targetPictureBox As pd2DSurface
    Drawing2D.QuickCreateSurfaceFromDC targetPictureBox, picOutput.hDC, False
    
    Dim i As Integer, animationSteps As Long
    
    Select Case m_ActiveTest
        
        'Olaf Schimdt's "animated polygon/curve" screensaver
        Case Test1_SplineDemo
            
            'Because animation steps are so fast, lets perform a whole bunch of them inside a timer event.
            ' This results in a much smoother effect than waiting for VB's timer object (which can be very erratic).
            For animationSteps = 0 To 127
                
                'Iterate through each point in our animated polygon and slightly move its position
                For i = 0 To m_NumOfPoints - 1
                
                    m_listOfPoints(i).x = m_listOfPoints(i).x + m_listOfSignsX(i) * 0.0004 * Abs(m_listOfPoints(i).y - m_listOfPoints(i).x)
                    m_listOfPoints(i).y = m_listOfPoints(i).y + m_listOfSignsY(i) * 0.1 / Abs((33 - m_listOfPoints(i).y) / (77 + m_listOfPoints(i).x))
                  
                    'If a sample point leaves the screen, reverse its direction
                    If m_listOfPoints(i).x < 0 Then m_listOfSignsX(i) = 1: m_listOfPoints(i).x = 0
                    If m_listOfPoints(i).x > picWidth Then m_listOfSignsX(i) = -1: m_listOfPoints(i).x = picWidth
                    If m_listOfPoints(i).y < 0 Then m_listOfSignsY(i) = 1: m_listOfPoints(i).y = 0
                    If m_listOfPoints(i).y > picHeight Then m_listOfSignsY(i) = -1: m_listOfPoints(i).y = picHeight
                    
                Next i
                
                'Gradually cycle between colors
                m_Test1ColorIncrement = m_Test1ColorIncrement + 0.34
                If m_Test1ColorIncrement > 255 Then
                    m_Test1ColorIncrement = 0
                    m_Test1ColorPhase = m_Test1ColorPhase + 1
                    If m_Test1ColorPhase > 5 Then m_Test1ColorPhase = 0
                End If
                
                Select Case m_Test1ColorPhase
                    Case 0: DrawDemo RGB(m_Test1ColorIncrement, 255 - m_Test1ColorIncrement, 255)
                    Case 1: DrawDemo RGB(255, m_Test1ColorIncrement, 255 - m_Test1ColorIncrement)
                    Case 2: DrawDemo RGB(255 - m_Test1ColorIncrement, 255, m_Test1ColorIncrement)
                    Case 3: DrawDemo RGB(0, 255 - m_Test1ColorIncrement, m_Test1ColorIncrement)
                    Case 4: DrawDemo RGB(255 - m_Test1ColorIncrement, m_Test1ColorIncrement, 0)
                    Case 5: DrawDemo RGB(0, 0, 255 - m_Test1ColorIncrement)
                End Select
                
                'Every 16 animation steps, copy the full contents of the "back buffer" surface onto the
                ' on-screen picture box.  (Because the picture box's .AutoRedraw property is set to FALSE,
                ' we do not need to forcibly refresh the picture box after copying the image data over.)
                If (animationSteps And 15) = 0 Then m_Painter.CopySurfaceI targetPictureBox, 0, 0, m_BackBuffer
                
            Next animationSteps
        
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
            
            DrawDemo
            
            'Finally, copy the full contents of the "back buffer" surface onto the on-screen picture box.
            ' (Because the picture box's .AutoRedraw property is set to FALSE, we do not need to forcibly
            ' refresh the picture box after performing the copy operation.)
            m_Painter.CopySurfaceI targetPictureBox, 0, 0, m_BackBuffer
            
        Case Else
    
    End Select
    
End Sub

Private Sub DrawDemo(Optional ByVal drawColor As Long = vbBlack)
    
    'In each of these tests, we'll be performing a series of simple steps:
    ' 1) Wrapping a pd2D surface around the picture box we want to paint on
    ' 2) Creating whatever pens, brushes, and other rendering tools we need
    ' 3) Painting onto our temporary surface.  (Because this surface just "wraps" the picture box, all of our
    '     paint operations will appear immediately on the screen.)
    
    Dim cBrush As pd2DBrush, cPen As pd2DPen, cPath As pd2DPath
    Set cPath = New pd2DPath
    
    Select Case m_ActiveTest
    
        Case Test1_SplineDemo
            
            'Create a pen matching the color we were passed.  We're also going to make the pen semi-transparent,
            ' and we're going to use "round" junctions where two lines meet (instead of mitered or beveled junctions).
            Drawing2D.QuickCreateSolidPen cPen, 0.5, drawColor, 5#, P2_LJ_Round, P2_LC_Round
            
            'Paint this path onto the output picture box
            If m_Test1CloseShape Then
                m_Painter.DrawPolygonF m_BackBuffer, cPen, m_NumOfPoints, VarPtr(m_listOfPoints(0)), Not m_Test1UseLines, 0.5
            Else
                m_Painter.DrawLinesF_FromPtF m_BackBuffer, cPen, m_NumOfPoints, VarPtr(m_listOfPoints(0)), Not m_Test1UseLines, 0.5
            End If
            
        Case Test2_PolygonDemo
            
            'Completely erase the back buffer.
            m_BackBuffer.EraseSurfaceContents 0, 0
            
            Dim i As Long
            For i = 0 To m_NumOfPoints - 1
                
                With m_ListOfPolygons(i)
                    
                    'Create a path from this polygon
                    cPath.ResetPath
                    cPath.AddPolygon_Regular .PolygonSides, .PolygonRadius, .PolygonCenter.x, .PolygonCenter.y, m_Test2UseCurvature, .PolygonCurvature
                    
                    'Apply this polygon's transformation
                    cPath.ApplyTransformation .PolygonTransform
                    
                    'First, fill the polygon area using the polygon's random fill color at 50% opacity
                    Drawing2D.QuickCreateSolidBrush cBrush, .PolygonColorFill, 50#
                    m_Painter.FillPath m_BackBuffer, cBrush, cPath
                    
                    'Then, trace the polygon outline using the polygon's random border color at 100% opacity
                    Drawing2D.QuickCreateSolidPen cPen, .PolygonBorderWidth, .PolygonColorBorder, 100#
                    m_Painter.DrawPath m_BackBuffer, cPen, cPath
                    
                End With
                
            Next i
        
    End Select
    
    'Note that we don't have to free any objects here - pd2D always takes care of that for us.

End Sub

'If an animated demo supports different rendering options, we'll set those options here.  Option changes always
' trigger an "erase" action, to make it easier to see what's changed.
Private Sub chkTest1Complete_Click()
    EraseAllBuffers
    m_Test1CloseShape = CBool(chkTest1Complete.Value = vbChecked)
End Sub

Private Sub chkTest1Lines_Click()
    EraseAllBuffers
    m_Test1UseLines = CBool(chkTest1Lines.Value = vbChecked)
End Sub
