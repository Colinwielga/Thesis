VERSION 5.00
Begin VB.Form frmcognitivebehavioral 
   Caption         =   "Quiz"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12480
   LinkTopic       =   "Form1"
   Picture         =   "frmcognitivebehavioral.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   12480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
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
      Height          =   555
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   10200
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H000000FF&
      Caption         =   "Stroop Task"
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
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9120
      Width           =   3015
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
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9120
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
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9120
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9120
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cognitive-Behavioral "
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3000
      TabIndex        =   5
      Top             =   240
      Width           =   8655
   End
End
Attribute VB_Name = "frmcognitivebehavioral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: A Review of Theoretical Orientations in Clinical Psychology
'Form name: frmcognitivebehavioral
'Author: Calvin Pipenhagen
'Date Written: March 10, 2008
'Objective: The main page for the cognitive-behavioral orientation, it facilitates the naviagation of this topic.
Private Sub Command1_Click() 'leads to form presenting theories
frmcognitivetheories.Show
frmcognitivebehavioral.Hide
End Sub

Private Sub Command2_Click() 'leads to form presenting notable followers of this orientation
frmcognitivefollowers.Show
frmcognitivebehavioral.Hide
End Sub

Private Sub Command3_Click() 'leads to quiz
frmcognitivebehavioral.Hide
frmcognitivequiz.Show
End Sub

Private Sub Command4_Click()
frmstroop.Show
frmcognitivebehavioral.Hide
End Sub

Private Sub Command5_Click() 'returns to the main menu
frmcognitivebehavioral.Hide
frmselectschool.Show
End Sub
