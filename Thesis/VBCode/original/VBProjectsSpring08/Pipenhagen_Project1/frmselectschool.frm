VERSION 5.00
Begin VB.Form frmselectschool 
   BackColor       =   &H00FF8080&
   Caption         =   "Form1"
   ClientHeight    =   11010
   ClientLeft      =   -555
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Quit"
      Height          =   375
      Left            =   6480
      TabIndex        =   13
      Top             =   10320
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "See Past Scores"
      Height          =   255
      Left            =   4440
      TabIndex        =   12
      Top             =   9960
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      Caption         =   "References"
      Height          =   255
      Left            =   7680
      TabIndex        =   11
      Top             =   9960
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF8080&
      Height          =   5895
      Left            =   10560
      Picture         =   "frmselectschool.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3600
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Height          =   5895
      Left            =   5520
      Picture         =   "frmselectschool.frx":5590
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Width           =   4095
   End
   Begin VB.CommandButton cmdpsychodynamic 
      Height          =   5895
      Left            =   480
      Picture         =   "frmselectschool.frx":55262
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3600
      Width           =   4095
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FF8080&
      Caption         =   "Rogers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   12120
      TabIndex        =   10
      Top             =   9600
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF8080&
      Caption         =   "Skinner"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   7080
      TabIndex        =   9
      Top             =   9600
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF8080&
      Caption         =   "Freud"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   8
      Top             =   8040
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF8080&
      Caption         =   "Freud"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   9600
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
      Caption         =   "Click on a Psychologist to Begin"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1455
      Left            =   840
      TabIndex        =   6
      Top             =   480
      Width           =   13335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   "Humanistic"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11520
      TabIndex        =   4
      Top             =   2760
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "Cognitive Behavioral"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      TabIndex        =   3
      Top             =   2760
      Width           =   3975
   End
   Begin VB.Label lblpsychodynamic 
      BackColor       =   &H00FF8080&
      Caption         =   "Psychodynamic"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   1
      Top             =   2760
      Width           =   3975
   End
End
Attribute VB_Name = "frmselectschool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: A Review of Theoretical Orientations in Clinical Psychology
'Form name: frmselectschool (this refers to school of theoretical orientation)
'Author: Calvin Pipenhagen
'Date Written: March 7, 2008
'Objective: The purpose of this project is to be a review of basic psychological concepts. In an expanded form,
            'it could serve as a review tool for the psychology GRE. It basically presents information and then
            'quizes the user on that information. There is also a replication of the famous psychological experiment
            'the Stroop task. This is more just for fun. It adds an interactive dimension to the project.
        'Form:
            'This form is the central interface from which the user navigates amongst the various theoretical
            'orientations.
Option Explicit

Private Sub cmdpsychodynamic_Click() 'Displays psychodynamic page
    frmselectschool.Hide
    frmpsychodynamic.Show
    
End Sub

Private Sub Command1_Click() 'Displays cognitive-behavioral page
    frmselectschool.Hide
    frmcognitivebehavioral.Show
End Sub

Private Sub Command2_Click() 'Displays humanistic page
    frmselectschool.Hide
    frmhumanistic.Show
End Sub

Private Sub Command3_Click() 'show reference form
frmreferences.Show
frmselectschool.Hide
End Sub

Private Sub Command4_Click() 'show form for past scores
frmpastscores.Show
frmselectschool.Hide
End Sub

Private Sub Command5_Click() 'quit program
End
End Sub

Private Sub Form_Load() 'This prevents the main page from appearing before the user enters their name (or any other string)
frmselectschool.Hide
frmpsychodynamic.Hide
Dim found As Boolean
found = False

    
    
Do While found = False
names = InputBox("Enter your name", "Name")
    If Len(names) = 0 Then  'insures that the user enters some string
        MsgBox "Please enter your name", , "You didn't submit a name"
    Else
        MsgBox names & ", Welcome to a Review of Theoretical Orientations in Clinical Psychology", , "Welcome!"
        found = True
End If
Loop
 frmselectschool.Show 'finally, the main page is displayed
End Sub
