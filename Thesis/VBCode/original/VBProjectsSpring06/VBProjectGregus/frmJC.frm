VERSION 5.00
Begin VB.Form frmJC 
   BackColor       =   &H80000007&
   Caption         =   "Player Profile"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10650
   LinkTopic       =   "Form1"
   ScaleHeight     =   8415
   ScaleWidth      =   10650
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
      Left            =   360
      MaskColor       =   &H000000FF&
      TabIndex        =   4
      Top             =   6720
      Width           =   2535
   End
   Begin VB.PictureBox picBoxSJU 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      Picture         =   "frmJC.frx":0000
      ScaleHeight     =   855
      ScaleWidth      =   4455
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
   Begin VB.PictureBox picBoxJC 
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   360
      Picture         =   "frmJC.frx":112B
      ScaleHeight     =   4335
      ScaleWidth      =   3855
      TabIndex        =   0
      Top             =   1920
      Width           =   3855
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
      Left            =   4800
      TabIndex        =   5
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label lblJC 
      BackColor       =   &H00000000&
      Caption         =   "John ""Pinball"" Carlson"
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
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   5655
   End
   Begin VB.Label lblExplain 
      BackColor       =   &H00000000&
      Caption         =   $"frmJC.frx":3792
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
      Height          =   3975
      Left            =   4800
      TabIndex        =   2
      Top             =   2040
      Width           =   3135
   End
End
Attribute VB_Name = "frmJC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'SJU Lacrosse Guide (Final Project 1.VBP)
'frmJC (frmJC.frm)
'Dan Gregus
'3/22/06
'Objective: To create a player profile page for John Carlson that can be linked to the team bio page

Private Sub cmdBack_Click()
    frmJC.Visible = False
    frmBio.Visible = True

End Sub
