VERSION 5.00
Begin VB.Form frmAK 
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
      Height          =   975
      Left            =   240
      MaskColor       =   &H000000FF&
      TabIndex        =   3
      Top             =   7440
      Width           =   2535
   End
   Begin VB.PictureBox picBoxAK 
      BorderStyle     =   0  'None
      Height          =   5535
      Left            =   240
      Picture         =   "frmAK.frx":0000
      ScaleHeight     =   5535
      ScaleWidth      =   6615
      TabIndex        =   2
      Top             =   1680
      Width           =   6615
   End
   Begin VB.PictureBox picBoxSJU 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      Picture         =   "frmAK.frx":5F31
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
      Left            =   7200
      TabIndex        =   6
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label lblDream 
      BackColor       =   &H00000000&
      Caption         =   "Lifelong Dream: To gain 1,000,000 friends on facebook"
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
      Height          =   1215
      Left            =   7080
      TabIndex        =   5
      Top             =   4560
      Width           =   2415
   End
   Begin VB.Label lblLikes 
      BackColor       =   &H00000000&
      Caption         =   "Likes: Glasses, Side burns, pointing at the camera.           Dislikes:  People who don't think he is awesome"
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
      Left            =   7080
      TabIndex        =   4
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label lblAK 
      BackColor       =   &H00000000&
      Caption         =   "  Alex Kady"
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
      Width           =   3135
   End
End
Attribute VB_Name = "frmAK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'SJU Lacrosse Guide (Final Project 1.VBP)
'frmAK (frmAK.frm)
'Dan Gregus
'3/22/06
'Objective: To create a player profile page for Alex Kady that can be linked to the team bio page


'
Private Sub cmdBack_Click()
    frmAK.Visible = False
    frmBio.Visible = True
End Sub

