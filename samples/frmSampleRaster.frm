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
   Begin VB.CommandButton cmdErase 
      Cancel          =   -1  'True
      Caption         =   "Erase!"
      Height          =   615
      Left            =   600
      TabIndex        =   6
      Top             =   6960
      Width           =   3615
   End
   Begin VB.CommandButton cmdLoadImage 
      Caption         =   "Select image..."
      Height          =   615
      Left            =   600
      TabIndex        =   4
      Top             =   960
      Width           =   3615
   End
   Begin VB.PictureBox picOutput 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      TabIndex        =   5
      Top             =   6600
      Width           =   3330
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "load a test image"
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
      Caption         =   "pd2D raster samples:"
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
      Width           =   2295
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
'pd2D Basic Raster (Bitmap) Sample Project
'Copyright 2016 by Tanner Helland
'Created: 01/July/16
'Last updated: 01/July/16
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
'I think that's everything!
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

Private Sub cmdLoadImage_Click()
    
    'pd2D makes image loading fast and convenient.  Let's start by displaying a common dialog with filters
    ' for all supported image formats.
    '
    '(Note that this pdOpenSaveDialog class is not part of pd2D, but you're welcome to use it; it came from
    ' the PhotoDemon project and it is BSD-licensed, just like pd2D.)
    Dim cFileOpen As pdOpenSaveDialog
    Set cFileOpen = New pdOpenSaveDialog
    
    Dim supportedImageFiles As String
    supportedImageFiles = "Supported images|*.bmp;*.emf;*.gif;*.ico;*.jpg;*.jpeg;*.png;*.tif;*.tiff;*.wmf|All files|*.*"
    
    Dim imgFilename As String
    If cFileOpen.GetOpenFileName(imgFilename, "", True, False, supportedImageFiles, , GetSampleImageFolder, "Please select an image file", , frmSample.hWnd) Then
        
        'pd2D provides a simplified function for loading images - just one line of code!
        Drawing2D.QuickLoadPicture picOutput, imgFilename
        
    End If
    
End Sub

Private Function GetSampleImageFolder() As String
    GetSampleImageFolder = App.Path
    If (StrComp(Right$(GetSampleImageFolder, 1), "\", vbBinaryCompare) <> 0) Then GetSampleImageFolder = GetSampleImageFolder & "\"
    GetSampleImageFolder = GetSampleImageFolder & "test images\"
End Function

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
