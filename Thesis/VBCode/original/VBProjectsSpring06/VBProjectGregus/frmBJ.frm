VERSION 5.00
Begin VB.Form frmBJ 
   BackColor       =   &H80000012&
   Caption         =   "Player Profile"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
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
      TabIndex        =   3
      Top             =   6720
      Width           =   2535
   End
   Begin VB.PictureBox picBoxBJ 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   120
      Picture         =   "frmBJ.frx":0000
      ScaleHeight     =   4335
      ScaleWidth      =   4455
      TabIndex        =   2
      Top             =   1800
      Width           =   4455
   End
   Begin VB.PictureBox picBoxSJU 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      Picture         =   "frmBJ.frx":300C
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
      TabIndex        =   6
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label lblDream 
      BackColor       =   &H00000000&
      Caption         =   "Lifelong dream: To play one injury-free season of lacrosse for SJU"
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
      Height          =   1455
      Left            =   5160
      TabIndex        =   5
      Top             =   5400
      Width           =   2775
   End
   Begin VB.Label lblLikes 
      BackColor       =   &H00000000&
      Caption         =   $"frmBJ.frx":4137
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
      Height          =   3495
      Left            =   5040
      TabIndex        =   4
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label lblBJ 
      BackColor       =   &H00000000&
      Caption         =   "Brian Jensen"
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
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   3615
   End
End
Attribute VB_Name = "frmBJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'SJU Lacrosse Guide (Final Project 1.VBP)
'frmBJ (frmBJ.frm)
'Dan Gregus
'3/22/06
'Objective: To create a player profile page for Brian Jensen that can be linked to the team bio page

Private Sub cmdBack_Click()
    frmBJ.Visible = False
    frmBio.Visible = True
End Sub
