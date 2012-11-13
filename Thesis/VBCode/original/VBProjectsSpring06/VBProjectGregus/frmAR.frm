VERSION 5.00
Begin VB.Form frmAR 
   BackColor       =   &H80000007&
   Caption         =   "Player Profile"
   ClientHeight    =   8010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10020
   LinkTopic       =   "Form1"
   ScaleHeight     =   8010
   ScaleWidth      =   10020
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H000000FF&
      Caption         =   "Back to Bio Page"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold Condensed"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      MaskColor       =   &H000000FF&
      TabIndex        =   6
      Top             =   6000
      Width           =   2535
   End
   Begin VB.PictureBox picBoxAR1 
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   120
      Picture         =   "frmAR.frx":0000
      ScaleHeight     =   3735
      ScaleWidth      =   3015
      TabIndex        =   2
      Top             =   1920
      Width           =   3015
   End
   Begin VB.PictureBox picBoxAR2 
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   4320
      Picture         =   "frmAR.frx":2B50
      ScaleHeight     =   2775
      ScaleWidth      =   2775
      TabIndex        =   1
      Top             =   2400
      Width           =   2775
   End
   Begin VB.PictureBox picBoxSJU 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      Picture         =   "frmAR.frx":90A4
      ScaleHeight     =   855
      ScaleWidth      =   4455
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
   Begin VB.Label lblCredit 
      BackColor       =   &H00000000&
      Caption         =   "Project by: Dan Gregus"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold Condensed"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5040
      TabIndex        =   8
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label lblDislikes 
      BackColor       =   &H80000012&
      Caption         =   "Lifelong goal: To one day leave the witness protection program and rejoin modern society."
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   7320
      TabIndex        =   7
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label lblAR 
      BackColor       =   &H00000000&
      Caption         =   "Adam ""Squirrel"" Rietz"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      TabIndex        =   5
      Top             =   840
      Width           =   6855
   End
   Begin VB.Label lblLikes 
      BackColor       =   &H00000000&
      Caption         =   "Likes: Classified   Dislikes: Classified"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   7320
      TabIndex        =   4
      Top             =   2160
      Width           =   3135
   End
   Begin VB.Label LBLpLUS 
      BackColor       =   &H00000000&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   90
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1575
      Left            =   3120
      TabIndex        =   3
      Top             =   2880
      Width           =   1095
   End
End
Attribute VB_Name = "frmAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'SJU Lacrosse Guide (Final Project 1.VBP)
'frmAR (frmAR.frm)
'Dan Gregus
'3/22/06
'Objective: To create a player profile page for Adam Rietz that can be linked to the team bio page

Private Sub cmdBack_Click()
frmAR.Visible = False
frmBio.Visible = True
End Sub
