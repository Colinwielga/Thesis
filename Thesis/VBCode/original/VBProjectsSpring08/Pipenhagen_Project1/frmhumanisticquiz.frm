VERSION 5.00
Begin VB.Form frmhumanisticquiz 
   BackColor       =   &H00FF8080&
   Caption         =   "Form1"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton opta 
      BackColor       =   &H00FF8080&
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
      Left            =   -2160
      TabIndex        =   31
      Top             =   1080
      Width           =   735
   End
   Begin VB.OptionButton optc 
      BackColor       =   &H00FF8080&
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
      Left            =   -2160
      TabIndex        =   30
      Top             =   2520
      Width           =   735
   End
   Begin VB.OptionButton optb 
      BackColor       =   &H00FF8080&
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
      Left            =   -2160
      TabIndex        =   29
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton cmdcontinue 
      BackColor       =   &H00FF8080&
      Caption         =   "Continue"
      Height          =   614
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   12960
      Width           =   4335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Height          =   2895
      Left            =   0
      TabIndex        =   19
      Top             =   -120
      Width           =   15855
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF8080&
         Caption         =   "Option1"
         Height          =   495
         Left            =   4200
         TabIndex        =   22
         Top             =   840
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FF8080&
         Caption         =   "Option1"
         Height          =   495
         Left            =   4200
         TabIndex        =   21
         Top             =   1440
         Width           =   255
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FF8080&
         Caption         =   "Option1"
         Height          =   495
         Left            =   4200
         TabIndex        =   20
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FF8080&
         Caption         =   "Congruence"
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
         TabIndex        =   27
         Top             =   2160
         Width           =   8175
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FF8080&
         Caption         =   "Insight"
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
         Top             =   1560
         Width           =   8175
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FF8080&
         Caption         =   "Knowledge"
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
         Top             =   960
         Width           =   8175
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FF8080&
         Caption         =   "Which of the following is a core component of client centered therapy?"
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
         TabIndex        =   24
         Top             =   240
         Width           =   11175
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FF8080&
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
         TabIndex        =   23
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF8080&
      Height          =   2775
      Left            =   0
      TabIndex        =   10
      Top             =   2880
      Width           =   15855
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FF8080&
         Caption         =   "Option1"
         Height          =   495
         Left            =   4080
         TabIndex        =   13
         Top             =   720
         Width           =   255
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FF8080&
         Caption         =   "Option1"
         Height          =   495
         Left            =   4080
         TabIndex        =   12
         Top             =   1320
         Width           =   255
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00FF8080&
         Caption         =   "Option1"
         Height          =   495
         Left            =   4080
         TabIndex        =   11
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FF8080&
         Caption         =   "Total and complete respect"
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
         TabIndex        =   18
         Top             =   2040
         Width           =   8175
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FF8080&
         Caption         =   "Honesty"
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
         Top             =   1440
         Width           =   8175
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FF8080&
         Caption         =   "Genuineness"
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
         Top             =   840
         Width           =   8175
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FF8080&
         Caption         =   "Congruence is best defined as what?"
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
         TabIndex        =   15
         Top             =   240
         Width           =   9375
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FF8080&
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
         TabIndex        =   14
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF8080&
      Height          =   2895
      Left            =   0
      TabIndex        =   1
      Top             =   6000
      Width           =   15855
      Begin VB.OptionButton Option7 
         BackColor       =   &H00FF8080&
         Caption         =   "Option1"
         Height          =   495
         Left            =   4080
         TabIndex        =   4
         Top             =   720
         Width           =   255
      End
      Begin VB.OptionButton Option8 
         BackColor       =   &H00FF8080&
         Caption         =   "Option1"
         Height          =   495
         Left            =   4080
         TabIndex        =   3
         Top             =   1440
         Width           =   255
      End
      Begin VB.OptionButton Option9 
         BackColor       =   &H00FF8080&
         Caption         =   "Option1"
         Height          =   495
         Left            =   4080
         TabIndex        =   2
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FF8080&
         Caption         =   "Find meaning in a sometimes harsh world"
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
         TabIndex        =   9
         Top             =   1560
         Width           =   8175
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FF8080&
         Caption         =   "Discover their own innate potential"
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
         Top             =   2160
         Width           =   8175
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FF8080&
         Caption         =   "Deal with their problems in the present"
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
         Top             =   840
         Width           =   8175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FF8080&
         Caption         =   "Logotherapy encourages patients to do what?"
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
         TabIndex        =   6
         Top             =   240
         Width           =   9375
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF8080&
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
         TabIndex        =   5
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdnextpage 
      BackColor       =   &H00C0C0C0&
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
      Left            =   4320
      MaskColor       =   &H0080FFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9240
      Width           =   5535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
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
      Left            =   -3000
      TabIndex        =   32
      Top             =   240
      Width           =   375
   End
End
Attribute VB_Name = "frmhumanisticquiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: A Review of Theoretical Orientations in Clinical Psychology
'Form name: frmhumanisticquiz
'Author: Calvin Pipenhagen
'Date Written: March 25, 2008
'Objective: To quiz the user about information presented about humanistic theory.
Option Explicit
Private Sub cmdnextpage_Click() 'determines if the user's answer is right and adjusts score appropriately
humanisticquizsum = 0 'resets the global variable quizsum so it can be used by the other quizes
If Option3 = True Then
    humanisticquizsum = humanisticquizsum + 1
End If
If Option4 = True Then
    humanisticquizsum = humanisticquizsum + 1
End If
If Option8 = True Then
    humanisticquizsum = humanisticquizsum + 1
End If
frmhumanisticquiz.Hide 'moves to the next page of the quiz
frmhumanisticquiztwo.Show
End Sub
