VERSION 5.00
Begin VB.Form frm1Screen 
   BackColor       =   &H0000FFFF&
   Caption         =   "Welcome!"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7065
   ForeColor       =   &H0000FFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9135
   ScaleWidth      =   7065
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picStartup 
      Height          =   7455
      Left            =   480
      ScaleHeight     =   7395
      ScaleWidth      =   5955
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   6015
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   1095
      Left            =   5040
      TabIndex        =   2
      Top             =   7680
      Width           =   1935
   End
   Begin VB.CommandButton cmdInstruct 
      Caption         =   "Instructions"
      Height          =   1095
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start the game"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   7680
      Width           =   2055
   End
End
Attribute VB_Name = "frm1Screen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdInstruct_Click()
    picStartup.Visible = True
    picStartup.Picture = LoadPicture(App.Path & "\lemonpeel.jpg")
    frmInstr.Show
    frm1Screen.Hide
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdStart_Click()
    EnterName = InputBox("Enter your name")
    frm1Screen.Hide
    frmMainScreen.Show
    Cash = 10
    Fame = 1
    Lemons = 8
    Sugar = 3
    Ice = 3
    Cups = 300
    Pitchers = 1
    RecipeL = 3
    RecipeS = 1
    RecipeI = 1
    charged = 1
End Sub

Private Sub picStartup_Click()
    picStartup.Picture = LoadPicture(App.Path & "\lemonpeel.jpg")
End Sub
