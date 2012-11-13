VERSION 5.00
Begin VB.Form frmFinal 
   BackColor       =   &H80000000&
   Caption         =   "Final Score"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9750
   FillColor       =   &H000000FF&
   BeginProperty Font 
      Name            =   "Bodoni MT Black"
      Size            =   20.25
      Charset         =   0
      Weight          =   900
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "frmFinal.frx":0000
   ScaleHeight     =   6045
   ScaleWidth      =   9750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLast 
      Caption         =   "Move on"
      Height          =   1095
      Left            =   6120
      TabIndex        =   6
      Top             =   0
      Width           =   3255
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "Print the final score"
      Height          =   1095
      Left            =   1320
      TabIndex        =   5
      Top             =   120
      Width           =   3615
   End
   Begin VB.PictureBox picTeam2 
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   48
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   7200
      ScaleHeight     =   1635
      ScaleWidth      =   2475
      TabIndex        =   1
      Top             =   4320
      Width           =   2535
   End
   Begin VB.PictureBox picTeam1 
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   48
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   1200
      ScaleHeight     =   1635
      ScaleWidth      =   2430
      TabIndex        =   0
      Top             =   4320
      Width           =   2490
   End
   Begin VB.Label lblTeam2 
      Height          =   1335
      Left            =   6000
      TabIndex        =   4
      Top             =   1920
      Width           =   3495
   End
   Begin VB.Label lblDash 
      Alignment       =   2  'Center
      Caption         =   "-----"
      Height          =   615
      Left            =   4560
      TabIndex        =   3
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label lblTeam1 
      Height          =   1335
      Left            =   1200
      TabIndex        =   2
      Top             =   1920
      Width           =   3735
   End
End
Attribute VB_Name = "frmFinal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this screen shows the two scores next to each other so the user can see who won

Private Sub cmdLast_Click()
frmFinal.Visible = False
frmDone.Visible = True

End Sub

Private Sub cmdprint_Click()

picTeam1.Print Team1Points
picTeam2.Print Team2Points
'simply shows the points scored

End Sub

Private Sub Form_Load()
 
lblTeam1.Caption = Nteam1(1)
lblTeam2.Caption = NTeam2(1)
'changes the labels to the name of the teams

End Sub


