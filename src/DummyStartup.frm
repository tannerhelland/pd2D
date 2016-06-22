VERSION 5.00
Begin VB.Form DummyForm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "This form is a dummy startup form"
   ClientHeight    =   4065
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8460
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   271
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   564
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblExplanation 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Please ignore this form.  It exists only to test compilation."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   8175
   End
End
Attribute VB_Name = "DummyForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

