VERSION 5.00
Begin VB.Form FormSelectProject 
   BackColor       =   &H80000017&
   Caption         =   "Form1"
   ClientHeight    =   10020
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18525
   FillColor       =   &H000000C0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10020
   ScaleWidth      =   18525
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   5535
      Left            =   5280
      Picture         =   "FormSelectProject.frx":0000
      ScaleHeight     =   5475
      ScaleWidth      =   6195
      TabIndex        =   2
      Top             =   2400
      Width           =   6255
   End
   Begin VB.CommandButton cmdIndividualProjects 
      Caption         =   "Select for Individual Projects"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   960
      TabIndex        =   1
      Top             =   3360
      Width           =   3615
   End
   Begin VB.CommandButton cmdGroupProjects 
      Caption         =   "Select for Group Projects"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   13560
      TabIndex        =   0
      Top             =   3360
      Width           =   3615
   End
   Begin VB.Label lblTitle 
      Caption         =   "Project Grader for CS130 Projects"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      TabIndex        =   3
      Top             =   480
      Width           =   7815
   End
End
Attribute VB_Name = "FormSelectProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGroupProjects_Click()
'if this is selected, the user will be taken to the group project grading pages
FormGroupInputData1.Show
FormSelectProject.Hide
End Sub

Private Sub cmdIndividualProjects_Click()
'if this is selected, the user will be take to the individual project grading pages
FormInputData1.Show
FormSelectProject.Hide
End Sub
