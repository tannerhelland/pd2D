VERSION 5.00
Begin VB.Form frmTutorial 
   AutoRedraw      =   -1  'True
   Caption         =   "Introduction to pd2D surfaces"
   ClientHeight    =   8010
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12825
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   534
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   855
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Reset form"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   4080
      Width           =   4215
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Paint PNG to Form at random opacity (and position)"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   4215
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Paint PNG to Form at random size"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   4215
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Paint PNG to Form at original size"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   4215
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load PNG file into pdSurface object"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "frmTutorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'pd2D Introduction to Surfaces
'Copyright 2021 by Tanner Helland
'Created: 12/October/21
'Last updated: 12/October/21
'Last update: initial build
'
'Surfaces (hopefully you figured this out from the name) are the primary drawing target of pd2D operations.
' Surfaces in pd2D are created as RGBA by default - where RGBA stands for "Red, Green, Blue, and Alpha".
' The "Alpha" component is used to store transparency data, so surfaces can have both opaque and transparent
' (or semi-transparent) pixels.
'
'Surfaces can be created from existing image files.  They can also be created as blank surfaces, which you
' can then paint on with whatever tools you want.
'
'In this brief demo, I'll show you how to load a transparent PNG image, then paint it onto a VB form.
'
'You really don't need to go into this tutorial with much existing knowledge, except for this:
' drawing commands in pd2D ALWAYS USE PIXELS.  To simplify this tutorial, the underlying form (frmTutorial)
' has its .ScaleMode set to Pixels as well, so we can reference positions on the form without confusion.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'This pdSurface object will hold the PNG data.  (We're going to store the PNG in a module-level object,
' so that the user can hit the "paint" button multiple times without us needing to re-load the PNG file.)
Private m_PNG As pd2DSurface

Private Sub Form_Load()
    
    'You always have to start pd2D before using it!
    PD2D.StartRenderingEngine
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    'You need to be a good citizen and always shut down pd2D before your program exits!
    PD2D.StopRenderingEngine
    
End Sub

Private Sub cmdLoad_Click(Index As Integer)
    
    'Keep reading to see how we use this "temporary" surface for drawing
    Dim temporarySurface As pd2DSurface
    
    Select Case Index
    
        'Command button 1: Load PNG file into pdSurface object
        Case 0
            
            'Loading an image from file (in any format) is this simple:
            Dim targetImageLocation As String
            targetImageLocation = App.Path & "\sunburst.png"
            
            Set m_PNG = New pd2DSurface
            If m_PNG.CreateSurfaceFromFile(targetImageLocation) Then
                Debug.Print "Successfully loaded " & targetImageLocation
            Else
                Debug.Print "WARNING: failed to load " & targetImageLocation
            End If
            
            'Once you've done this, m_PNG has safely stored the file contents locally.  The original PNG file
            ' is no longer needed, and you don't need to push the Load button again during this session.
            
        'Command button 2: Paint pdSurface onto the form at its original size
        Case 1
            
            'First, clear any existing drawing and select a random backcolor
            ' (this will help demonstrate PNG transparency)
            Set frmTutorial.Picture = LoadPicture(vbNullString)
            Randomize Timer
            frmTutorial.BackColor = RGB(Rnd * 256, Rnd * 256, Rnd * 256)
            
            'Next, make sure the user hit the "load PNG file button"
            If (m_PNG Is Nothing) Then
                Debug.Print "You need to load the PNG file first!"
            
            Else
                
                'pd2D can only paint onto pd2D surfaces.  So how do we paint onto a VB object?
                ' Easy: we wrap a (temporary) pd2D surface around it.
                Set temporarySurface = New pd2DSurface
                temporarySurface.WrapSurfaceAroundDC frmTutorial.hDC
                
                'Anything we draw onto "temporarySurface" will now get drawn directly onto frmTutorial!
                
                '"Wrapping" surfaces around objects like this is very powerful and it results in a much
                ' simpler drawing API, because pd2D drawing commands only require a destination surface -
                ' but pd2D doesn't care *what* that surface is wrapped around (a form, a picture box,
                ' a blank in-memory surface, an image loaded from file - all work identically!)
                
                'Now we just need to call the appropriate pd2D draw function to draw our PNG surface.
                ' In this example, we'll draw the image just to the right of the command button.
                PD2D.DrawSurfaceI temporarySurface, cmdLoad(0).Left + cmdLoad(0).Width + 20, cmdLoad(0).Top, m_PNG
                
                'With our drawing complete, we don't need our temporary surface object anymore
                Set temporarySurface = Nothing
                
                'When AutoRedraw is TRUE, frmTutorial will have no idea that something new has been drawn to it
                ' (because we didn't use built-in VB drawing commands that AutoRedraw can auto-detect).
                ' So, we need to manually notify the form that new stuff has been drawn, so that it knows to
                ' refresh its on-screen appearance.
                If frmTutorial.AutoRedraw Then
                    Set frmTutorial.Picture = frmTutorial.Image
                    frmTutorial.Refresh
                End If
                
            End If
            
        'Command button 3: Paint pdSurface onto the form at a random size
        Case 2
            
            'This time, let's force the form backcolor to black, just for fun
            Set frmTutorial.Picture = LoadPicture(vbNullString)
            frmTutorial.BackColor = RGB(0, 0, 0)
            
            'Again, make sure the user hit the "load PNG file button"
            If (m_PNG Is Nothing) Then
                Debug.Print "You need to load the PNG file first!"
            
            Else
                
                'This time, let's use the "quick" helper function to wrap a pd2D surface around a form
                ' with one line of code.
                PD2D.QuickWrapSurfaceAroundDC temporarySurface, frmTutorial.hDC
                
                'Easy, right?
                
                'This time, we want to resize the PNG when we paint it.  Note that this does *NOT* affect
                ' the source surface - the stretch is "silently" performed as part of the draw command,
                ' and the surface object remains in pristine condition.
                '
                'To resize the image when drawing it, we need to use a different function - one that allows
                ' us to specify a new width and/or height (instead of just using the original size).
                '
                '(Note that the TOP and LEFT position of this drawing are the same as the previous button.)
                '
                'For purposes of this demonstration, we'll use a random size on the range [100, 500]
                Randomize Timer
                PD2D.DrawSurfaceResizedI temporarySurface, cmdLoad(0).Left + cmdLoad(0).Width + 20, cmdLoad(0).Top, _
                                           100 + Rnd * 400, 100 + Rnd * 400, m_PNG
                
                'With our drawing complete, we don't need our temporary surface object anymore
                Set temporarySurface = Nothing
                
                'Same as before, AutoRedraw requires manual notification of new drawing operations
                If frmTutorial.AutoRedraw Then
                    Set frmTutorial.Picture = frmTutorial.Image
                    frmTutorial.Refresh
                End If
                
            End If
            
        'Command button 4: Paint pdSurface onto the form at a random opacity
        Case 3
            
            'Let's stick with the whole "black background" thing
            Set frmTutorial.Picture = LoadPicture(vbNullString)
            frmTutorial.BackColor = RGB(0, 0, 0)
            
            'Still gotta make sure the user hit the "load PNG file button"
            If (m_PNG Is Nothing) Then
                Debug.Print "You need to load the PNG file first!"
            
            Else
                
                'You know the drill by now: first, wrap a pd2D surface around whatever we want to paint on
                PD2D.QuickWrapSurfaceAroundDC temporarySurface, frmTutorial.hDC
                
                'This time, we want to change the PNG's opacity when we paint it.  As with resizing,
                ' note that this does *NOT* affect the source surface - the opacity change is "silently"
                ' performed as part of the draw command, and the surface object remains in pristine condition.
                '
                'To help make any opacity changes more noticeable, let's also paint the image at a random
                ' position.  We can do this by varying the .Left and .Top parameters of the draw command.
                '
                'Also note that opacity is always supplied on the range [0, 100], where 0 = fully transparency
                ' and 100 = fully opaque. We need our random opacity to stay in that range too.
                Randomize Timer
                PD2D.DrawSurfaceI temporarySurface, cmdLoad(0).Left + cmdLoad(0).Width + 20 + Rnd * 200, _
                                  Rnd * 300, m_PNG, Rnd * 100
                
                'You should understand the following lines without comments by now
                Set temporarySurface = Nothing
                If frmTutorial.AutoRedraw Then
                    Set frmTutorial.Picture = frmTutorial.Image
                    frmTutorial.Refresh
                End If
                
            End If
            
        'Command button 5: reset the form (erase any drawing)
        Case 4
            
            Set frmTutorial.Picture = LoadPicture(vbNullString)
            frmTutorial.BackColor = vbButtonFace
            
    End Select
    
End Sub
