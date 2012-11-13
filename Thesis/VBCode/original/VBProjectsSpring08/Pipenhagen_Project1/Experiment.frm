VERSION 5.00
Begin VB.Form frmhumanistic 
   BackColor       =   &H00C000C0&
   Caption         =   "Form1"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14880
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   14880
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   4935
      Left            =   4320
      Picture         =   "Experiment.frx":0000
      ScaleHeight     =   4875
      ScaleWidth      =   6435
      TabIndex        =   5
      Top             =   3600
      Width           =   6495
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H000000FF&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2040
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Theories"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "Followers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H000000FF&
      Caption         =   "Quiz"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C000C0&
      Caption         =   "Humanistic Psychology"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   48
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2640
      TabIndex        =   0
      Top             =   360
      Width           =   9975
   End
End
Attribute VB_Name = "frmhumanistic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: A Review of Theoretical Orientations in Clinical Psychology
'Form name: frmhumanistic
'Author: Calvin Pipenhagen
'Date Written: March 25, 2008
'Objective: Serves as the home page for all information related to a humanistic orientation.
Option Explicit
Private Sub Command1_Click() 'loads form containing humanistic theories
frmhumanistic.Hide
frmhumanistictheories.Show
End Sub

Private Sub Command2_Click() 'loads form containing humanistic followers
frmhumanistic.Hide
frmhumanisticfollowers.Show
End Sub

Private Sub Command3_Click() 'loads the humanistic quiz
frmhumanistic.Hide
frmhumanisticquiz.Show
End Sub

Private Sub Command4_Click() 'returns to the main menu
frmhumanistic.Hide
frmselectschool.Show
End Sub
