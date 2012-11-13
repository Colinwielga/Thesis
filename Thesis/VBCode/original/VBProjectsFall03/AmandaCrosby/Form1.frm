VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF8080&
   Caption         =   "Form1"
   ClientHeight    =   3930
   ClientLeft      =   3345
   ClientTop       =   4095
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   5685
   Begin VB.PictureBox picTotalScore 
      BeginProperty Font 
         Name            =   "NIST Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      ScaleHeight     =   435
      ScaleWidth      =   825
      TabIndex        =   11
      Top             =   3120
      Width           =   885
   End
   Begin VB.PictureBox picLevelTwoScore 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "NIST Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      ScaleHeight     =   435
      ScaleWidth      =   675
      TabIndex        =   7
      Top             =   2280
      Width           =   735
   End
   Begin VB.PictureBox picLevelOneScore 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "NIST Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      ScaleHeight     =   435
      ScaleWidth      =   675
      TabIndex        =   4
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "NIST Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton cmdLevelTwo 
      BackColor       =   &H0080C0FF&
      Caption         =   "Level Two"
      BeginProperty Font 
         Name            =   "NIST Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton cmdLevelOne 
      BackColor       =   &H008080FF&
      Caption         =   "Level One"
      BeginProperty Font 
         Name            =   "NIST Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
      Caption         =   "Designed by: Amanda Crosby"
      BeginProperty Font 
         Name            =   "NIST Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   12
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Total Score Is:"
      BeginProperty Font 
         Name            =   "NIST Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   10
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lblLevelOneCompleted 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   $"Form1.frx":0000
      BeginProperty Font 
         Name            =   "NIST Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   1575
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblLevelTwoCompleted 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   $"Form1.frx":00B8
      BeginProperty Font 
         Name            =   "NIST Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   1575
      Left            =   3240
      TabIndex        =   8
      Top             =   600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Click to See Score:"
      BeginProperty Font 
         Name            =   "NIST Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   6
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Click to See Score:"
      BeginProperty Font 
         Name            =   "NIST Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Music Learning Program"
      BeginProperty Font 
         Name            =   "NIST Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Music Learning Program (M:\CS130\AmandaCrosbyMusicLearningProgram.vbp)
'Forms: Form1, LevelOne, LevelTwo (M:\CS130\AmandaCrosby\Form1.frm), (M:\CS130\AmandaCrosby\LevelOne.frm), (M:\CS130\AmandaCrosby\LevelTwo.frm)
'Amanda Crosby
'Written October 20, 2003
'This is a program to help kids learn some basic music symbols
Option Explicit

Private Sub cmdLevelOne_Click()
    Form1.Hide
    LevelOne.Show
    LevelTwo.Hide
End Sub

Private Sub cmdLevelThree_Click()
    Form1.Hide
    LevelOne.Hide
    LevelTwo.Hide
    LevelThree.Show
End Sub

Private Sub cmdLevelTwo_Click()
    Form1.Hide
    LevelOne.Hide
    LevelTwo.Show
End Sub

Private Sub cmdQuit_Click()
    End
End Sub




Private Sub Label5_Click()
    picLevelOneScore.Cls
    picLevelTwoScore.Cls
    picLevelOneScore.Print LevelOneScore
    picLevelTwoScore.Print LevelTwoScore
    If LevelOneScore >= 30 Then
        lblLevelOneCompleted.Visible = True
    End If
    If LevelTwoScore >= 30 Then
        lblLevelTwoCompleted.Visible = True
    End If
    Dim TotalScore As Integer
    TotalScore = LevelOneScore + LevelTwoScore
    picTotalScore.Cls
    picTotalScore.Print TotalScore
End Sub

Private Sub Label6_Click()
    picLevelOneScore.Cls
    picLevelTwoScore.Cls
    picLevelOneScore.Print LevelOneScore
    picLevelTwoScore.Print LevelTwoScore
    If LevelOneScore >= 30 Then
        lblLevelOneCompleted.Visible = True
    End If
    If LevelTwoScore >= 30 Then
        lblLevelTwoCompleted.Visible = True
    End If
    Dim TotalScore As Integer
    TotalScore = LevelOneScore + LevelTwoScore
    picTotalScore.Cls
    picTotalScore.Print TotalScore
End Sub
