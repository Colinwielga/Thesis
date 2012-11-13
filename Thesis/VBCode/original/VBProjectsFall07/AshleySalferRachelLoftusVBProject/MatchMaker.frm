VERSION 5.00
Begin VB.Form frmMainPage 
   Caption         =   "Main Page"
   ClientHeight    =   9165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12600
   LinkTopic       =   "Form1"
   Picture         =   "MatchMaker.frx":0000
   ScaleHeight     =   9165
   ScaleWidth      =   12600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Quit!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton cmdClickToStart 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Click Here To Find Your Celebrity Soulmate!"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   4080
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5280
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "WELCOME TO CELEBRITY MATCHMAKER!!! "
      BeginProperty Font 
         Name            =   "Rage Italic"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   2895
      Left            =   2640
      TabIndex        =   0
      Top             =   0
      Width           =   5895
   End
End
Attribute VB_Name = "frmMainPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub cmdClickToStart_Click()
    'Switches to registration form
    frmMainPage.Hide
    frmRegistration.Show
    
End Sub

Private Sub cmdQuit_Click()
    'Ends program
    End
End Sub


Private Sub Form_Load()

End Sub
