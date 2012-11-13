VERSION 5.00
Begin VB.Form FrmGameTheory 
   BackColor       =   &H000000FF&
   Caption         =   "Fun with Matrices"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   8190
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdMinMax 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click here to explore  minimax/maximin strategies"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3360
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000000FF&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   4455
      Index           =   1
      Left            =   10920
      Picture         =   "FrmGameTheory.frx":0000
      ScaleHeight     =   4395
      ScaleWidth      =   4275
      TabIndex        =   4
      Top             =   1200
      Width           =   4335
   End
   Begin VB.CommandButton CmdQuit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton CmdDominance 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click here to explore dominance relations"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   1815
   End
   Begin VB.PictureBox PicResults 
      BackColor       =   &H00FFFFFF&
      Height          =   9375
      Left            =   2040
      ScaleHeight     =   9315
      ScaleMode       =   0  'User
      ScaleWidth      =   8175
      TabIndex        =   1
      Top             =   120
      Width           =   8775
   End
   Begin VB.CommandButton CmdIntroduction 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click here to learn about Game Theory"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label LblNames 
      BackColor       =   &H000000FF&
      Caption         =   "By Carson Sievert and Aaron Skelly"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   240
      TabIndex        =   6
      Top             =   6600
      Width           =   1575
   End
   Begin VB.Label LblGAMETHEORY 
      BackColor       =   &H000000FF&
      Caption         =   "INTRODUCTION TO GAME THEORY"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1200
      Left            =   720
      TabIndex        =   5
      Top             =   9600
      Width           =   14010
   End
End
Attribute VB_Name = "FrmGameTheory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Introduction to Game Theory
'Form: FrmGameTheory
'Carson Sievert
'Aaron Skelly
'March 23, 2008

'This VB Project provides basic information about Game Theory. The program
'deals only with 2x2 matrices for simplicity, but is most help for beginners
'to easily see how the concepts work. The program not only provides examples
'but also allows the user to input his/her own values. This feature allows the
'user to contruct as many examples as possible as well as check results.

Private Sub CmdDominance_Click()
    FrmGameTheory.Hide  'this will bring you to the page about dominance
    FrmDominance.Show
End Sub

Private Sub CmdIntroduction_Click() 'Here is the loop that loads the file
    Dim info(1 To 100) As String    'containing the text introduction to
    Dim CTR As Integer              'simple game theory
    Open App.Path & "\game theory.txt" For Input As #1
    Do While Not EOF(1)
        CTR = CTR + 1
        Input #1, info(CTR)
        PicResults.Print info(CTR)
    Loop
End Sub

Private Sub CmdMinMax_Click()
    FrmGameTheory.Hide 'this will bring you to the page about maximin/minimax
    Frmminmax.Show     'strategies
End Sub

Private Sub CmdQuit_Click()
End
End Sub

