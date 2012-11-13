VERSION 5.00
Begin VB.Form FormStart 
   BackColor       =   &H00C0C000&
   Caption         =   "Form1"
   ClientHeight    =   11985
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13050
   LinkTopic       =   "Form1"
   ScaleHeight     =   11985
   ScaleWidth      =   13050
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStart 
      Caption         =   "Begin Grading!!!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   1320
      TabIndex        =   2
      Top             =   7920
      Width           =   9855
   End
   Begin VB.Label lblTitle 
      Caption         =   "Computer Science 130 Project Grader"
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
      Left            =   2040
      TabIndex        =   1
      Top             =   240
      Width           =   8655
   End
   Begin VB.Label lblDescription 
      Caption         =   $"FormStart.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   1080
      TabIndex        =   0
      Top             =   1680
      Width           =   10215
   End
End
Attribute VB_Name = "FormStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdStart_Click()
'this button takes the user from the description page to the beginning of the project
FormStart.Hide
FormProfessorSelect.Show
End Sub
