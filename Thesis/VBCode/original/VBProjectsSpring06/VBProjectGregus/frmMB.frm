VERSION 5.00
Begin VB.Form frmMB 
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
      Height          =   735
      Left            =   240
      MaskColor       =   &H000000FF&
      TabIndex        =   3
      Top             =   7680
      Width           =   2535
   End
   Begin VB.PictureBox picBoxMB 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   5655
      Left            =   240
      Picture         =   "frmMB.frx":0000
      ScaleHeight     =   5655
      ScaleWidth      =   5775
      TabIndex        =   2
      Top             =   1920
      Width           =   5775
   End
   Begin VB.PictureBox picBoxSJU 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      Picture         =   "frmMB.frx":592E
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
      Left            =   6240
      TabIndex        =   6
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label lblMajor 
      BackColor       =   &H00000000&
      Caption         =   "Major: Interpretative Dance Minor: Management "
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
      Height          =   1575
      Left            =   6360
      TabIndex        =   5
      Top             =   3480
      Width           =   2535
   End
   Begin VB.Label lbllikes 
      BackColor       =   &H00000000&
      Caption         =   "Likes: Lacrosse Dislikes: Everything that's not Lacrosse"
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
      Height          =   1335
      Left            =   6360
      TabIndex        =   4
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label lblMB 
      BackColor       =   &H00000000&
      Caption         =   "Mark Bachand"
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
      TabIndex        =   1
      Top             =   1080
      Width           =   3855
   End
End
Attribute VB_Name = "frmMB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'SJU Lacrosse Guide (Final Project 1.VBP)
'frmMB (frmMB.frm)
'Dan Gregus
'3/22/06
'Objective: To create a player profile page for Mark Bachand that can be linked to the team bio page

Private Sub cmdBack_Click()
    frmMB.Visible = False
    frmBio.Visible = True
End Sub
