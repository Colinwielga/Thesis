VERSION 5.00
Begin VB.Form frmTR 
   BackColor       =   &H80000012&
   Caption         =   "Tina Rytel"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Back to Profile Page"
      BeginProperty Font 
         Name            =   "Minion Pro"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   6600
      TabIndex        =   5
      Top             =   6480
      Width           =   4455
   End
   Begin VB.PictureBox Picture1 
      Height          =   8055
      Left            =   0
      Picture         =   "frmTR.frx":0000
      ScaleHeight     =   7995
      ScaleWidth      =   6195
      TabIndex        =   0
      Top             =   120
      Width           =   6255
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000012&
      Caption         =   "* How she's going to save the world: Educate the mases."
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   6360
      TabIndex        =   4
      Top             =   2520
      Width           =   5175
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      Caption         =   "Favorite Bands: Green Day; Earth, Wind, and  Fire; Rage Against the Machine"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   6360
      TabIndex        =   3
      Top             =   1080
      Width           =   5535
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "Favorite Font: Jazz Text"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   6360
      TabIndex        =   2
      Top             =   240
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "* Not Actual Picture"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   8280
      Width           =   5895
   End
End
Attribute VB_Name = "frmTR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The Solar Project (Final Project 1.VBP)
'frmDG (frmDG.frm)
'Dan Gregus
'3/27/08
'Objective: To create a  profile page for Tina Rytel that can be linked to the team bio page


Private Sub cmdBack_Click()
    frmTR.Visible = False
    frmBio.Visible = True
End Sub
