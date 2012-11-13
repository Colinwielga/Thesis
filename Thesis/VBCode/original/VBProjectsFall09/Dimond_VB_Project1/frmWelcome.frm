VERSION 5.00
Begin VB.Form frmWelcome 
   Caption         =   "Welcome to the QuickHelp Play-Calling!"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   8145
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H8000000E&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmdEnter 
      BackColor       =   &H8000000E&
      Caption         =   "Enter QuickHelp Play-Calling"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
   Begin VB.Image imgFootball 
      Height          =   7200
      Left            =   0
      Picture         =   "frmWelcome.frx":0000
      Top             =   0
      Width           =   9600
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this is a title Screen welcoming the user to the program

Private Sub cmdEnter_Click()
    
    'This button hides the welcome screen and shows the first form
    frmQuarter.Show
    frmWelcome.Hide
    
End Sub

Private Sub cmdQuit_Click()
    'quit button with a friendly message
    MsgBox "Go get em'!", , "Good Luck!"
    End
End Sub
