VERSION 5.00
Begin VB.Form frmWD 
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
      Left            =   5040
      MaskColor       =   &H000000FF&
      TabIndex        =   3
      Top             =   6960
      Width           =   2535
   End
   Begin VB.PictureBox picBoxWD 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   6735
      Left            =   240
      Picture         =   "frmWD.frx":0000
      ScaleHeight     =   6735
      ScaleWidth      =   4575
      TabIndex        =   2
      Top             =   1680
      Width           =   4575
   End
   Begin VB.PictureBox picBoxSJU 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      Picture         =   "frmWD.frx":5EE6
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
   Begin VB.Label lblDislikes 
      BackColor       =   &H00000000&
      Caption         =   "Will be Studying Abroad / Serving his country: Possibly Iraq or Iran, but the smart money is on Syria"
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
      Height          =   2415
      Left            =   5160
      TabIndex        =   5
      Top             =   3240
      Width           =   2655
   End
   Begin VB.Label lblLikes 
      BackColor       =   &H00000000&
      Caption         =   "Likes: Jesus      Dislikes: Terrorists"
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
      Left            =   5160
      TabIndex        =   4
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Label lblWD 
      BackColor       =   &H00000000&
      Caption         =   "Will ""Supernatural"" Durbin"
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
      Width           =   6735
   End
End
Attribute VB_Name = "frmWD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'SJU Lacrosse Guide (Final Project 1.VBP)
'frmWD (frmWD.frm)
'Dan Gregus
'3/22/06
'Objective: To create a player profile page for Will Durbin that can be linked to the team bio page

Private Sub cmdBack_Click()
    frmWD.Visible = False
    frmBio.Visible = True
End Sub
