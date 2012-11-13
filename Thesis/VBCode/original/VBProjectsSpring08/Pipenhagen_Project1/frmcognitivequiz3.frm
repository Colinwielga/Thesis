VERSION 5.00
Begin VB.Form frmcognitivequizthree 
   BackColor       =   &H00C0C000&
   Caption         =   "Form1"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C000&
      Height          =   3135
      Left            =   360
      TabIndex        =   20
      Top             =   6600
      Width           =   14415
      Begin VB.OptionButton Option9 
         BackColor       =   &H00C0C000&
         Caption         =   "Option1"
         Height          =   255
         Left            =   3480
         TabIndex        =   28
         Top             =   2520
         Width           =   255
      End
      Begin VB.OptionButton Option8 
         BackColor       =   &H00C0C000&
         Caption         =   "Option1"
         Height          =   255
         Left            =   3480
         TabIndex        =   27
         Top             =   1920
         Width           =   255
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H00C0C000&
         Caption         =   "Option1"
         Height          =   255
         Left            =   3480
         TabIndex        =   26
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0C000&
         Caption         =   "Change thought processes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3840
         TabIndex        =   25
         Top             =   1920
         Width           =   8175
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0C000&
         Caption         =   "Both A and B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3840
         TabIndex        =   24
         Top             =   2520
         Width           =   8175
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         Caption         =   "Modify behavior"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3840
         TabIndex        =   23
         Top             =   1200
         Width           =   8175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         Caption         =   "Contingency management holds that by controlling consequences they can?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   2520
         TabIndex        =   22
         Top             =   120
         Width           =   11535
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "9."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   21
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C000&
      Height          =   2895
      Left            =   360
      TabIndex        =   11
      Top             =   3360
      Width           =   14415
      Begin VB.OptionButton Option6 
         BackColor       =   &H00C0C000&
         Caption         =   "Option1"
         Height          =   255
         Left            =   3480
         TabIndex        =   19
         Top             =   2280
         Width           =   255
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00C0C000&
         Caption         =   "Option1"
         Height          =   255
         Left            =   3480
         TabIndex        =   18
         Top             =   1560
         Width           =   255
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0C000&
         Caption         =   "Option1"
         Height          =   255
         Left            =   3480
         TabIndex        =   17
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0C000&
         Caption         =   "Systematic desensitization"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3840
         TabIndex        =   16
         Top             =   2280
         Width           =   8175
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C000&
         Caption         =   "Almost all psychodynamic oriented therapies"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3840
         TabIndex        =   15
         Top             =   1560
         Width           =   8175
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C000&
         Caption         =   "Behavioral rehersal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3840
         TabIndex        =   14
         Top             =   960
         Width           =   8175
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C000&
         Caption         =   "The learning of relaxation techniques is prominent in which therapy?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2640
         TabIndex        =   13
         Top             =   0
         Width           =   9375
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C000&
         Caption         =   "8."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   12
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      Height          =   2895
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   14415
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0C000&
         Caption         =   "Option1"
         Height          =   255
         Left            =   3480
         TabIndex        =   10
         Top             =   2160
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C000&
         Caption         =   "Option1"
         Height          =   255
         Left            =   3480
         TabIndex        =   9
         Top             =   1560
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Option1"
         Height          =   255
         Left            =   3480
         TabIndex        =   8
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0C000&
         Caption         =   "Evaluating the effectivness of the treatment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3840
         TabIndex        =   7
         Top             =   2160
         Width           =   8175
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0C000&
         Caption         =   "Attempting to implement the learned behaviors in the real world"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3840
         TabIndex        =   6
         Top             =   1560
         Width           =   8175
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0C000&
         Caption         =   "Leaving the clinic."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3840
         TabIndex        =   5
         Top             =   840
         Width           =   8175
      End
      Begin VB.Label frmcognitivequiz3 
         BackColor       =   &H00C0C000&
         Caption         =   "What is the last step in Behavioral rehersal?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   4
         Top             =   240
         Width           =   9375
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0C000&
         Caption         =   "7."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   3
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdreturn 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Return to Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   614
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   10080
      Width           =   4335
   End
   Begin VB.CommandButton cmdscorecognitivequiz 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Score Quiz"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   614
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   10080
      Width           =   4335
   End
End
Attribute VB_Name = "frmcognitivequizthree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: A Review of Theoretical Orientations in Clinical Psychology
'Form name: frmcognitivequizthree
'Author: Calvin Pipenhagen
'Date Written: March 25, 2008
'Objective: This is portion of the quiz relating to the cognitive-behavioral theory. It determines
           'if the user selected the appropriate option.
Option Explicit

Private Sub cmdreturn_Click() 'returns to the main cognitive-behavioral form
Dim ctr As Integer
Dim n As Integer

frmcognitivequizthree.Hide
frmcognitivebehavioral.Show

End Sub

Private Sub cmdscorecognitivequiz_Click() 'Scores the quiz
Dim score As Single
Dim n As Integer
Dim pass As Integer
Dim pos As Integer
Dim temp As Integer
Dim tempname As String
Dim j As Integer

If Option2 = True Then 'If user selected appropriate answers, his or her score increases.
    cognitivequizsum = cognitivequizsum + 1
End If
If Option6 = True Then
    cognitivequizsum = cognitivequizsum + 1
End If
If Option7 = True Then
    cognitivequizsum = cognitivequizsum + 1
End If
score = cognitivequizsum / 9
MsgBox "you scored " & FormatPercent(score), , "Your Score" 'presents the user with the results of the quiz


End Sub

