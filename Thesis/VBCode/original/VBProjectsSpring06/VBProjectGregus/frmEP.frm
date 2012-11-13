VERSION 5.00
Begin VB.Form frmEP 
   BackColor       =   &H80000007&
   Caption         =   "Player Profile"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   ScaleHeight     =   8445
   ScaleWidth      =   10785
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
      TabIndex        =   4
      Top             =   6720
      Width           =   2535
   End
   Begin VB.PictureBox picBoxSJU 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      Picture         =   "frmEP.frx":0000
      ScaleHeight     =   855
      ScaleWidth      =   4455
      TabIndex        =   2
      Top             =   0
      Width           =   4455
   End
   Begin VB.PictureBox picBoxEP 
      BorderStyle     =   0  'None
      Height          =   4455
      Left            =   240
      Picture         =   "frmEP.frx":112B
      ScaleHeight     =   4455
      ScaleWidth      =   4215
      TabIndex        =   0
      Top             =   2040
      Width           =   4215
   End
   Begin VB.Label Label2 
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
      Left            =   5880
      TabIndex        =   5
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label lblEP 
      BackColor       =   &H00000000&
      Caption         =   "Erick ""Buffalo"" Peterson"
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
      TabIndex        =   3
      Top             =   960
      Width           =   6855
   End
   Begin VB.Label lblWhy 
      BackColor       =   &H00000000&
      Caption         =   $"frmEP.frx":49F7
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
      Height          =   3855
      Left            =   4920
      TabIndex        =   1
      Top             =   2160
      Width           =   2895
   End
End
Attribute VB_Name = "frmEP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'SJU Lacrosse Guide (Final Project 1.VBP)
'frmEP (frmEP.frm)
'Dan Gregus
'3/22/06
'Objective: To create a player profile page for Erick Peterson that can be linked to the team bio page

Private Sub cmdBack_Click()
frmEP.Visible = False
frmBio.Visible = True
End Sub
