VERSION 5.00
Begin VB.Form frmJB 
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
      Top             =   7080
      Width           =   2535
   End
   Begin VB.PictureBox picBoxJB 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   5175
      Left            =   0
      Picture         =   "frmJB.frx":0000
      ScaleHeight     =   5175
      ScaleWidth      =   6855
      TabIndex        =   2
      Top             =   1800
      Width           =   6855
   End
   Begin VB.PictureBox picBoxSJU 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      Picture         =   "frmJB.frx":53C5
      ScaleHeight     =   855
      ScaleWidth      =   4455
      TabIndex        =   0
      Top             =   0
      Width           =   4455
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
      Left            =   6840
      TabIndex        =   6
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label lblImportant 
      BackColor       =   &H00000000&
      Caption         =   "Important Advice: ""Always keep a towel handy.  Trust me on this one."""
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
      Height          =   2055
      Left            =   7080
      TabIndex        =   5
      Top             =   4200
      Width           =   2295
   End
   Begin VB.Label lblCrazy 
      BackColor       =   &H00000000&
      Caption         =   "Craziest Job: Joe once played a male model as an extra in the movie Zoolander"
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
      Height          =   2055
      Left            =   7080
      TabIndex        =   4
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label lblJB 
      BackColor       =   &H00000000&
      Caption         =   "Joe Boone"
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
      Top             =   960
      Width           =   3015
   End
End
Attribute VB_Name = "frmJB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'SJU Lacrosse Guide (Final Project 1.VBP)
'frmJB (frmJB.frm)
'Dan Gregus
'3/22/06
'Objective: To create a player profile page for Joe Boone that can be linked to the team bio page

Private Sub cmdBack_Click()
    frmJB.Visible = False
    frmBio.Visible = True
End Sub
