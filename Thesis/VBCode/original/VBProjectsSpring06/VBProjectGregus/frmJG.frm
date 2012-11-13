VERSION 5.00
Begin VB.Form frmJG 
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
      Height          =   1095
      Left            =   240
      MaskColor       =   &H000000FF&
      TabIndex        =   3
      Top             =   7440
      Width           =   2535
   End
   Begin VB.PictureBox picBoxJG 
      BorderStyle     =   0  'None
      Height          =   5535
      Left            =   240
      Picture         =   "frmJG.frx":0000
      ScaleHeight     =   5535
      ScaleWidth      =   4335
      TabIndex        =   2
      Top             =   1800
      Width           =   4335
   End
   Begin VB.PictureBox picBoxSJU 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      Picture         =   "frmJG.frx":76BD
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
      Left            =   4920
      TabIndex        =   6
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label lblLife 
      BackColor       =   &H00000000&
      Caption         =   "Lifelong goal: Find out if the Hokey pokey really is what it's all about."
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
      Height          =   1695
      Left            =   4800
      TabIndex        =   5
      Top             =   3240
      Width           =   2895
   End
   Begin VB.Label lbllikes 
      BackColor       =   &H00000000&
      Caption         =   "Likes: The great outdoors. Dislikes: Sleeves"
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
      Height          =   855
      Left            =   4800
      TabIndex        =   4
      Top             =   1800
      Width           =   3255
   End
   Begin VB.Label lblJG 
      BackColor       =   &H00000000&
      Caption         =   "Justin Gervais"
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
      Width           =   3735
   End
End
Attribute VB_Name = "frmJG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'SJU Lacrosse Guide (Final Project 1.VBP)
'frmJG (frmJG.frm)
'Dan Gregus
'3/22/06
'Objective: To create a player profile page for Justin Gervais that can be linked to the team bio page

Private Sub cmdBack_Click()
    frmJG.Visible = False
    frmBio.Visible = True
End Sub
