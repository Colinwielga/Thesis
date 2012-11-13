VERSION 5.00
Begin VB.Form frmWelcome 
   BackColor       =   &H000000FF&
   Caption         =   "Welcome to Children's Jeopardy!"
   ClientHeight    =   10785
   ClientLeft      =   390
   ClientTop       =   450
   ClientWidth     =   14370
   FillColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10785
   ScaleWidth      =   14370
   Begin VB.CommandButton cmdQuit 
      Height          =   1455
      Left            =   7920
      Picture         =   "frmWelcome.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "EXIT!"
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton cmdHighScores 
      Enabled         =   0   'False
      Height          =   1455
      Left            =   6480
      Picture         =   "frmWelcome.frx":145E
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "HIGH SCORES!"
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton cmdPlay 
      Height          =   1455
      Left            =   4920
      Picture         =   "frmWelcome.frx":2B17
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "PLAY!"
      Top             =   5640
      Width           =   1575
   End
   Begin VB.PictureBox Picture8 
      Height          =   3615
      Left            =   0
      Picture         =   "frmWelcome.frx":4118
      ScaleHeight     =   3555
      ScaleWidth      =   4755
      TabIndex        =   7
      Top             =   0
      Width           =   4815
   End
   Begin VB.PictureBox Picture7 
      Height          =   3615
      Left            =   9600
      Picture         =   "frmWelcome.frx":878F
      ScaleHeight     =   3555
      ScaleWidth      =   4755
      TabIndex        =   6
      Top             =   7200
      Width           =   4815
   End
   Begin VB.PictureBox Picture6 
      Height          =   3615
      Left            =   4800
      Picture         =   "frmWelcome.frx":CE06
      ScaleHeight     =   3555
      ScaleWidth      =   4755
      TabIndex        =   5
      Top             =   7200
      Width           =   4815
   End
   Begin VB.PictureBox Picture5 
      Height          =   3615
      Left            =   0
      Picture         =   "frmWelcome.frx":1147D
      ScaleHeight     =   3555
      ScaleWidth      =   4755
      TabIndex        =   4
      Top             =   7200
      Width           =   4815
   End
   Begin VB.PictureBox Picture4 
      Height          =   3615
      Left            =   9600
      Picture         =   "frmWelcome.frx":15AF4
      ScaleHeight     =   3555
      ScaleWidth      =   4755
      TabIndex        =   3
      Top             =   0
      Width           =   4815
   End
   Begin VB.PictureBox Picture3 
      Height          =   3615
      Left            =   9600
      Picture         =   "frmWelcome.frx":1A16B
      ScaleHeight     =   3555
      ScaleWidth      =   4755
      TabIndex        =   2
      Top             =   3600
      Width           =   4815
   End
   Begin VB.PictureBox Picture2 
      Height          =   3615
      Left            =   0
      Picture         =   "frmWelcome.frx":1E7E2
      ScaleHeight     =   3555
      ScaleWidth      =   4755
      TabIndex        =   1
      Top             =   3600
      Width           =   4815
   End
   Begin VB.PictureBox Picture1 
      Height          =   3615
      Left            =   4800
      Picture         =   "frmWelcome.frx":22E59
      ScaleHeight     =   3555
      ScaleWidth      =   4755
      TabIndex        =   0
      Top             =   0
      Width           =   4815
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   7920
      TabIndex        =   18
      Top             =   5520
      Width           =   135
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   6360
      TabIndex        =   17
      Top             =   5520
      Width           =   135
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Height          =   3375
      Left            =   9480
      TabIndex        =   13
      Top             =   3720
      Width           =   135
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Height          =   3375
      Left            =   4800
      TabIndex        =   12
      Top             =   3720
      Width           =   135
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   4800
      TabIndex        =   11
      Top             =   5520
      Width           =   4815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   4800
      TabIndex        =   10
      Top             =   7080
      Width           =   4815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   4800
      TabIndex        =   9
      Top             =   3600
      Width           =   4815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Welcome to Children's Jeopardy!"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   4920
      TabIndex        =   8
      Top             =   3720
      Width           =   4575
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdHighScores_Click()
    'hide welcome screen and show high score form
    frmWelcome.Hide
    frmHighScores.Show
End Sub

Private Sub cmdPlay_Click()
    'hide the main program naviagtion screen
    frmWelcome.Hide
    'show the character selection screen
    frmCharacter.Show
End Sub

Private Sub cmdQuit_Click()
    'terminate the program
    End
End Sub

