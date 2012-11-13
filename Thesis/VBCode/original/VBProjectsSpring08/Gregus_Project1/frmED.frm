VERSION 5.00
Begin VB.Form frmED 
   BackColor       =   &H80000012&
   Caption         =   "Factuly Profile"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGoBack 
      Caption         =   "Back to Profile Page"
      BeginProperty Font 
         Name            =   "Minion Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   240
      TabIndex        =   3
      Top             =   4920
      Width           =   4095
   End
   Begin VB.PictureBox Picture1 
      Height          =   4575
      Left            =   240
      Picture         =   "frmED.frx":0000
      ScaleHeight     =   4515
      ScaleWidth      =   4035
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "How he's going to save the world: Inspire a generation of students to work for a better tomorrow."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1935
      Left            =   4680
      TabIndex        =   2
      Top             =   2880
      Width           =   4575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Ernie is the Faculty Advisor for the solar project.  He acts as a go-between for the students and the administration."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   2295
      Left            =   4680
      TabIndex        =   1
      Top             =   240
      Width           =   4575
   End
End
Attribute VB_Name = "frmED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The Solar Project
'frmED (frmDG.frm)
'Dan Gregus
'3/27/08
'Objective: To create a  profile page for Ernie Diedrich that can be linked to the team bio page


Private Sub cmdGoBack_Click()
    frmED.Visible = False
    frmBio.Visible = True
End Sub
