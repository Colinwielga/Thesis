VERSION 5.00
Begin VB.Form frmJBr 
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
      Left            =   6000
      MaskColor       =   &H000000FF&
      TabIndex        =   3
      Top             =   6960
      Width           =   2535
   End
   Begin VB.PictureBox picBoxJB 
      BorderStyle     =   0  'None
      Height          =   6375
      Left            =   360
      Picture         =   "frmJBr.frx":0000
      ScaleHeight     =   6375
      ScaleWidth      =   5415
      TabIndex        =   2
      Top             =   1800
      Width           =   5415
   End
   Begin VB.PictureBox picBoxSJU 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      Picture         =   "frmJBr.frx":A928
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
      Left            =   6120
      TabIndex        =   5
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label lblInteresting 
      BackColor       =   &H00000000&
      Caption         =   "Interesting Fact:  There's absolutely nothing interesting about John Broich at all."
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
      Height          =   1815
      Left            =   6120
      TabIndex        =   4
      Top             =   1920
      Width           =   3255
   End
   Begin VB.Label lblJB 
      BackColor       =   &H00000000&
      Caption         =   "John ""Cheese"" Broich"
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
      Width           =   5655
   End
End
Attribute VB_Name = "frmJBr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'SJU Lacrosse Guide (Final Project 1.VBP)
'frmJBr (frmJBr.frm)
'Dan Gregus
'3/22/06
'Objective: To create a player profile page for John Broich that can be linked to the team bio page

Private Sub cmdBack_Click()
    frmJBr.Visible = False
    frmBio.Visible = True
End Sub
