VERSION 5.00
Begin VB.Form frmpsychodynamicquiz 
   BackColor       =   &H0080C0FF&
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdnextpage 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Continue"
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
      Left            =   4800
      MaskColor       =   &H0080FFFF&
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   9480
      Width           =   5535
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0080C0FF&
      Height          =   2895
      Left            =   480
      TabIndex        =   23
      Top             =   6240
      Width           =   15855
      Begin VB.OptionButton Option9 
         BackColor       =   &H0080C0FF&
         Caption         =   "Option1"
         Height          =   495
         Left            =   4080
         TabIndex        =   31
         Top             =   2040
         Width           =   255
      End
      Begin VB.OptionButton Option8 
         BackColor       =   &H0080C0FF&
         Caption         =   "Option1"
         Height          =   495
         Left            =   4080
         TabIndex        =   30
         Top             =   1440
         Width           =   255
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H0080C0FF&
         Caption         =   "Option1"
         Height          =   495
         Left            =   4080
         TabIndex        =   29
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080C0FF&
         Caption         =   "3."
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
         Left            =   2640
         TabIndex        =   28
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080C0FF&
         Caption         =   "Which of the following best characterizes the ego?"
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
         Left            =   3120
         TabIndex        =   27
         Top             =   240
         Width           =   9375
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080C0FF&
         Caption         =   "It is the executive of the personality."
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
         Left            =   4560
         TabIndex        =   26
         Top             =   840
         Width           =   8175
      End
      Begin VB.Label Label11 
         BackColor       =   &H0080C0FF&
         Caption         =   "The ego is most analagous to a conscience"
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
         Left            =   4560
         TabIndex        =   25
         Top             =   2160
         Width           =   8175
      End
      Begin VB.Label Label15 
         BackColor       =   &H0080C0FF&
         Caption         =   "It obeys the pleasure principle"
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
         Left            =   4560
         TabIndex        =   24
         Top             =   1560
         Width           =   8175
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080C0FF&
      Height          =   2775
      Left            =   480
      TabIndex        =   14
      Top             =   3120
      Width           =   15855
      Begin VB.OptionButton Option6 
         BackColor       =   &H0080C0FF&
         Caption         =   "Option1"
         Height          =   495
         Left            =   4080
         TabIndex        =   22
         Top             =   1920
         Width           =   255
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H0080C0FF&
         Caption         =   "Option1"
         Height          =   495
         Left            =   4080
         TabIndex        =   21
         Top             =   1320
         Width           =   255
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H0080C0FF&
         Caption         =   "Option1"
         Height          =   495
         Left            =   4080
         TabIndex        =   20
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label6 
         BackColor       =   &H0080C0FF&
         Caption         =   "2."
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
         Left            =   2760
         TabIndex        =   19
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label7 
         BackColor       =   &H0080C0FF&
         Caption         =   "When is the Superego formed?"
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
         Left            =   3240
         TabIndex        =   18
         Top             =   240
         Width           =   9375
      End
      Begin VB.Label Label8 
         BackColor       =   &H0080C0FF&
         Caption         =   "The superego exists from the moment we are born."
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
         Left            =   4560
         TabIndex        =   17
         Top             =   840
         Width           =   8175
      End
      Begin VB.Label Label9 
         BackColor       =   &H0080C0FF&
         Caption         =   "The superego is not part of psychoanalytic theory."
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
         Left            =   4560
         TabIndex        =   16
         Top             =   1440
         Width           =   8175
      End
      Begin VB.Label Label10 
         BackColor       =   &H0080C0FF&
         Caption         =   "The superego is formed during the Oedipus Complex."
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
         Left            =   4560
         TabIndex        =   15
         Top             =   2040
         Width           =   8175
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Height          =   2895
      Left            =   480
      TabIndex        =   5
      Top             =   120
      Width           =   15855
      Begin VB.OptionButton Option3 
         BackColor       =   &H0080C0FF&
         Caption         =   "Option1"
         Height          =   495
         Left            =   4200
         TabIndex        =   13
         Top             =   2040
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Option1"
         Height          =   495
         Left            =   4200
         TabIndex        =   12
         Top             =   1440
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Option1"
         Height          =   495
         Left            =   4200
         TabIndex        =   11
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label12 
         BackColor       =   &H0080C0FF&
         Caption         =   "1."
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
         Left            =   2880
         TabIndex        =   10
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label13 
         BackColor       =   &H0080C0FF&
         Caption         =   "How did Freud conceptualize the structure of personality?"
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
         Left            =   3360
         TabIndex        =   9
         Top             =   240
         Width           =   9375
      End
      Begin VB.Label Label14 
         BackColor       =   &H0080C0FF&
         Caption         =   "Freud concieved of personality as two seperate urges, Eros and Thantos."
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
         Left            =   4560
         TabIndex        =   8
         Top             =   840
         Width           =   8175
      End
      Begin VB.Label Label16 
         BackColor       =   &H0080C0FF&
         Caption         =   "Freud theorized the exsistence of an id, ego and superego."
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
         Left            =   4560
         TabIndex        =   7
         Top             =   1560
         Width           =   8175
      End
      Begin VB.Label Label17 
         BackColor       =   &H0080C0FF&
         Caption         =   "Freud assumed that personality is solely a product of out unconscious urges."
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
         Left            =   4560
         TabIndex        =   6
         Top             =   2160
         Width           =   8175
      End
   End
   Begin VB.CommandButton cmdcontinue 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Continue"
      Height          =   614
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   13200
      Width           =   4335
   End
   Begin VB.OptionButton optb 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   -1680
      TabIndex        =   3
      Top             =   2040
      Width           =   735
   End
   Begin VB.OptionButton optc 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   -1680
      TabIndex        =   2
      Top             =   2760
      Width           =   735
   End
   Begin VB.OptionButton opta 
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   -1680
      TabIndex        =   1
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "1."
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
      Left            =   -2520
      TabIndex        =   0
      Top             =   480
      Width           =   375
   End
End
Attribute VB_Name = "frmpsychodynamicquiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: A Review of Theoretical Orientations in Clinical Psychology
'Form name: frmpsychodynamicquiz
'Author: Calvin Pipenhagen
'Date Written: March 25, 2008
'Objective: This is the first page of a three page quiz on psychodynamic theory. It purpose
           'is to get the user's answers and determine if they are right
Option Explicit
Private Sub cmdnextpage_Click() 'each question is placed within its own frame so that the option
psychodynamicquizsum = 0                                'buttons are only active for that frame
If Option2 = True Then 'this is determining if the right option was selected. If it was one is
                       'added to the global variable quizsum.
    psychodynamicquizsum = psychodynamicquizsum + 1
End If
If Option6 = True Then
    psychodynamicquizsum = psychodynamicquizsum + 1
End If
If Option7 = True Then
    psychodynamicquizsum = psychodynamicquizsum + 1
End If
frmpsychodynamicquiz.Hide 'moves to the next page of the quiz
frmpsychodynamicquiztwo.Show

End Sub

