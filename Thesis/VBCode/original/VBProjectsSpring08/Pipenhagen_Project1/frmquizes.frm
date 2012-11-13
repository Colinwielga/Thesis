VERSION 5.00
Begin VB.Form frmpsychodynamicquizthree 
   BackColor       =   &H0080C0FF&
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   FillColor       =   &H0080C0FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H0080C0FF&
      Height          =   2535
      Left            =   600
      TabIndex        =   20
      Top             =   6720
      Width           =   14295
      Begin VB.OptionButton Option9 
         BackColor       =   &H0080C0FF&
         Caption         =   "Option1"
         Height          =   255
         Left            =   3240
         TabIndex        =   28
         Top             =   2040
         Width           =   255
      End
      Begin VB.OptionButton Option8 
         BackColor       =   &H0080C0FF&
         Caption         =   "Option1"
         Height          =   255
         Left            =   3240
         TabIndex        =   27
         Top             =   1320
         Width           =   255
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H0080C0FF&
         Caption         =   "Option1"
         Height          =   255
         Left            =   3240
         TabIndex        =   26
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label15 
         BackColor       =   &H0080C0FF&
         Caption         =   "10 years"
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
         Left            =   3600
         TabIndex        =   25
         Top             =   1320
         Width           =   8175
      End
      Begin VB.Label Label11 
         BackColor       =   &H0080C0FF&
         Caption         =   "3 years"
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
         Left            =   3600
         TabIndex        =   24
         Top             =   720
         Width           =   8175
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080C0FF&
         Caption         =   "8 years"
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
         Left            =   3600
         TabIndex        =   23
         Top             =   2040
         Width           =   8175
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080C0FF&
         Caption         =   "At what age does the Phallic stage begin?"
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
         Left            =   2400
         TabIndex        =   22
         Top             =   120
         Width           =   9375
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080C0FF&
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
         Left            =   1920
         TabIndex        =   21
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080C0FF&
      Height          =   3015
      Left            =   600
      TabIndex        =   11
      Top             =   3360
      Width           =   14295
      Begin VB.OptionButton Option6 
         BackColor       =   &H0080C0FF&
         Caption         =   "Option1"
         Height          =   255
         Left            =   3360
         TabIndex        =   19
         Top             =   2160
         Width           =   255
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H0080C0FF&
         Caption         =   "Option1"
         Height          =   255
         Left            =   3360
         TabIndex        =   18
         Top             =   1440
         Width           =   255
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H0080C0FF&
         Caption         =   "Option1"
         Height          =   255
         Left            =   3360
         TabIndex        =   17
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label10 
         BackColor       =   &H0080C0FF&
         Caption         =   "A secretary"
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
         Left            =   3720
         TabIndex        =   16
         Top             =   2160
         Width           =   8175
      End
      Begin VB.Label Label9 
         BackColor       =   &H0080C0FF&
         Caption         =   "An executive"
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
         Left            =   3720
         TabIndex        =   15
         Top             =   1440
         Width           =   8175
      End
      Begin VB.Label Label8 
         BackColor       =   &H0080C0FF&
         Caption         =   "A conscience"
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
         Left            =   3720
         TabIndex        =   14
         Top             =   840
         Width           =   8175
      End
      Begin VB.Label Label7 
         BackColor       =   &H0080C0FF&
         Caption         =   "The superego is most analagous to what?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2520
         TabIndex        =   13
         Top             =   240
         Width           =   9375
      End
      Begin VB.Label Label6 
         BackColor       =   &H0080C0FF&
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
         Left            =   2040
         TabIndex        =   12
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Height          =   2895
      Left            =   600
      TabIndex        =   2
      Top             =   120
      Width           =   14295
      Begin VB.OptionButton Option3 
         BackColor       =   &H0080C0FF&
         Caption         =   "Option1"
         Height          =   255
         Left            =   3360
         TabIndex        =   10
         Top             =   2400
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Option1"
         Height          =   255
         Left            =   3360
         TabIndex        =   9
         Top             =   1800
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Option1"
         Height          =   255
         Left            =   3360
         TabIndex        =   8
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label17 
         BackColor       =   &H0080C0FF&
         Caption         =   "Reaction Formation"
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
         Left            =   3720
         TabIndex        =   7
         Top             =   2400
         Width           =   8175
      End
      Begin VB.Label Label16 
         BackColor       =   &H0080C0FF&
         Caption         =   "Regression"
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
         Left            =   3720
         TabIndex        =   6
         Top             =   1800
         Width           =   8175
      End
      Begin VB.Label Label14 
         BackColor       =   &H0080C0FF&
         Caption         =   "Fixation"
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
         Left            =   3720
         TabIndex        =   5
         Top             =   1200
         Width           =   8175
      End
      Begin VB.Label Label13 
         BackColor       =   &H0080C0FF&
         Caption         =   "Someone who is unable to advance to the next psychosexual stage is experiencing what?"
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
         Left            =   2760
         TabIndex        =   4
         Top             =   120
         Width           =   10335
      End
      Begin VB.Label Label12 
         BackColor       =   &H0080C0FF&
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
   Begin VB.CommandButton Command2 
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
      Height          =   615
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9720
      Width           =   4335
   End
   Begin VB.CommandButton cmdscore 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Score Your Quiz"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9720
      Width           =   4335
   End
End
Attribute VB_Name = "frmpsychodynamicquizthree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: A Review of Theoretical Orientations in Clinical Psychology
'Form name: frmpsychodynamicquizthree
'Author: Calvin Pipenhagen
'Date Written: March 25, 2008
'Objective: This is the third page of a three page quiz on psychodynamic theory. It purpose
           'is to get the user's answers and determine if they are right

Option Explicit
Private Sub cmdscore_click()
If Option1 = True Then     'checks to see if the user answered correctly
    psychodynamicquizsum = psychodynamicquizsum + 1 'adds to score if he or she did answer correctly
End If
If Option4 = True Then
    psychodynamicquizsum = psychodynamicquizsum + 1
End If
If Option9 = True Then
    psychodynamicquizsum = psychodynamicquizsum + 1
End If
MsgBox "you scored " & FormatPercent(psychodynamicquizsum / 9), , "Your Score" 'presents the user with the results of the quiz
End Sub

Private Sub Command2_Click()


frmpsychodynamic.Show
frmpsychodynamicquizthree.Hide

End Sub


