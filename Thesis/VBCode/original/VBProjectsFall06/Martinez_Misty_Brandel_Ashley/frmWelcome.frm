VERSION 5.00
Begin VB.Form frmWelcome 
   Caption         =   "Welcome"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   Picture         =   "frmWelcome.frx":0000
   ScaleHeight     =   5205
   ScaleWidth      =   7560
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optNo 
      Caption         =   "Option2"
      Height          =   255
      Left            =   4560
      TabIndex        =   5
      Top             =   2400
      Width           =   255
   End
   Begin VB.OptionButton optYes 
      Caption         =   "Option1"
      Height          =   255
      Left            =   4560
      TabIndex        =   4
      Top             =   2040
      Width           =   255
   End
   Begin VB.CommandButton cmdGo 
      BackColor       =   &H80000005&
      Caption         =   "Let's Play!"
      Height          =   855
      Left            =   2040
      MaskColor       =   &H00FF80FF&
      TabIndex        =   3
      Top             =   4080
      Width           =   2775
   End
   Begin VB.TextBox txtName 
      Height          =   735
      Left            =   360
      TabIndex        =   2
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Label lblHat 
      Caption         =   "Look for the Hat to EXIT!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   9
      Top             =   4800
      Width           =   2415
   End
   Begin VB.Label lblRead 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "Have You Read Dr. Seuss books before?"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3840
      TabIndex        =   8
      Top             =   1680
      Width           =   3375
   End
   Begin VB.Label lblNo 
      Caption         =   "No"
      Height          =   255
      Left            =   5040
      TabIndex        =   7
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label lblYes 
      Caption         =   "Yes"
      Height          =   255
      Left            =   5040
      TabIndex        =   6
      Top             =   2040
      Width           =   495
   End
   Begin VB.Image imgExit 
      Height          =   705
      Left            =   5880
      Picture         =   "frmWelcome.frx":287E6
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "Type your name here: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label lblWelcome 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "Welcome to the Phat Red Hat!!"
      BeginProperty Font 
         Name            =   "@MS UI Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   6855
   End
   Begin VB.Image Image1 
      Height          =   5235
      Left            =   0
      Picture         =   "frmWelcome.frx":28CEB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7575
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGo_Click()
    YourName = txtName.Text         'the player enters their name here
    frmWelcome.Visible = False      'Welcome level disappears
    frmLevel1.Visible = True        'Level 1 appears
    
End Sub


Private Sub imgExit_Click()         'Exit Button is to exit program
End
End Sub


Private Sub optNo_Click()
    Yes = False                     'The player finds out their result at the end of the game
    No = True                       'Allows program to know what to print
End Sub

Private Sub optYes_Click()
    Yes = True                      'The player finds out their result at the end of the game
    No = False                      'Allows program to know what to print
End Sub
