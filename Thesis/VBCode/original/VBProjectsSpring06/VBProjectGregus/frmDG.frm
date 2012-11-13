VERSION 5.00
Begin VB.Form frmDG 
   BackColor       =   &H80000007&
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
      Left            =   120
      MaskColor       =   &H000000FF&
      TabIndex        =   5
      Top             =   6960
      Width           =   2535
   End
   Begin VB.PictureBox picBoxDG 
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   120
      Picture         =   "frmDG.frx":0000
      ScaleHeight     =   5055
      ScaleWidth      =   5895
      TabIndex        =   2
      Top             =   1680
      Width           =   5895
   End
   Begin VB.PictureBox picBoxSJU 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      Picture         =   "frmDG.frx":65A5
      ScaleHeight     =   855
      ScaleWidth      =   4215
      TabIndex        =   0
      Top             =   0
      Width           =   4215
   End
   Begin VB.Label Label3 
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
      Left            =   6480
      TabIndex        =   6
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label lblLife 
      BackColor       =   &H00000000&
      Caption         =   "Lifelong goal: Rule his own island nation with an iron fist"
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
      Left            =   6120
      TabIndex        =   4
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label lblGet 
      BackColor       =   &H00000000&
      Caption         =   "Gets by: With a little help from his friends"
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
      Left            =   6120
      TabIndex        =   3
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label lblDG 
      BackColor       =   &H00000000&
      Caption         =   "Dan ""Gray Goose"" Gregus"
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
      Width           =   6855
   End
End
Attribute VB_Name = "frmDG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'SJU Lacrosse Guide (Final Project 1.VBP)
'frmDG (frmDG.frm)
'Dan Gregus
'3/22/06
'Objective: To create a player profile page for Dan Gregus that can be linked to the team bio page


Private Sub cmdBack_Click()
    frmDG.Visible = False
    frmBio.Visible = True
End Sub

