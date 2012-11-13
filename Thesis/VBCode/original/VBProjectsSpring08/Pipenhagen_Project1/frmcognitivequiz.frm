VERSION 5.00
Begin VB.Form frmcognitivequiz 
   BackColor       =   &H00C0C000&
   Caption         =   "Form1"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C000&
      Height          =   2655
      Left            =   600
      TabIndex        =   19
      Top             =   6600
      Width           =   14055
      Begin VB.OptionButton Option9 
         BackColor       =   &H00C0C000&
         Caption         =   "Option1"
         Height          =   255
         Left            =   2280
         TabIndex        =   27
         Top             =   2160
         Width           =   255
      End
      Begin VB.OptionButton Option8 
         BackColor       =   &H00C0C000&
         Caption         =   "Option1"
         Height          =   255
         Left            =   2280
         TabIndex        =   26
         Top             =   1440
         Width           =   255
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H00C0C000&
         Caption         =   "Option1"
         Height          =   255
         Left            =   2280
         TabIndex        =   25
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0C000&
         Caption         =   "Contingeny management"
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
         Left            =   2640
         TabIndex        =   24
         Top             =   1440
         Width           =   8175
      End
      Begin VB.Label Label11 
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
         Height          =   735
         Left            =   2640
         TabIndex        =   23
         Top             =   2160
         Width           =   8175
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         Caption         =   "Exposure therapy"
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
         Left            =   2640
         TabIndex        =   22
         Top             =   720
         Width           =   8175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         Caption         =   "Shaping is a type of what?"
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
         Left            =   1560
         TabIndex        =   21
         Top             =   120
         Width           =   9375
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
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
         Left            =   1080
         TabIndex        =   20
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C000&
      Height          =   2895
      Left            =   600
      TabIndex        =   10
      Top             =   3360
      Width           =   14055
      Begin VB.OptionButton Option6 
         BackColor       =   &H00C0C000&
         Caption         =   "Option1"
         Height          =   255
         Left            =   2400
         TabIndex        =   18
         Top             =   2160
         Width           =   255
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00C0C000&
         Caption         =   "Option1"
         Height          =   255
         Left            =   2400
         TabIndex        =   17
         Top             =   1440
         Width           =   255
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0C000&
         Caption         =   "Option1"
         Height          =   255
         Left            =   2400
         TabIndex        =   16
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0C000&
         Caption         =   "Freud's conception of personality"
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
         Left            =   2760
         TabIndex        =   15
         Top             =   2160
         Width           =   8175
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C000&
         Caption         =   "Classical conditioning"
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
         Left            =   2760
         TabIndex        =   14
         Top             =   1440
         Width           =   8175
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C000&
         Caption         =   "Operant Conditioning"
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
         Left            =   2760
         TabIndex        =   13
         Top             =   840
         Width           =   8175
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C000&
         Caption         =   "Contingeny management is based on what?"
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
         Left            =   1680
         TabIndex        =   12
         Top             =   120
         Width           =   9375
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C000&
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
         Left            =   1080
         TabIndex        =   11
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      Height          =   2895
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   14055
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0C000&
         Caption         =   "Option1"
         Height          =   255
         Left            =   2520
         TabIndex        =   9
         Top             =   2160
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C000&
         Caption         =   "Option1"
         Height          =   255
         Left            =   2520
         TabIndex        =   8
         Top             =   1560
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Option1"
         Height          =   255
         Left            =   2520
         TabIndex        =   7
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0C000&
         Caption         =   "The client visualizes an anxiety provoking scene"
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
         Left            =   2880
         TabIndex        =   6
         Top             =   2160
         Width           =   8175
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0C000&
         Caption         =   "A hierarchy of situations is formed"
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
         Left            =   2880
         TabIndex        =   5
         Top             =   1560
         Width           =   8175
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0C000&
         Caption         =   "A client is exposed to an anxiety causing situation"
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
         Left            =   2880
         TabIndex        =   4
         Top             =   840
         Width           =   8175
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0C000&
         Caption         =   "What occurs first during the Systematic Desensitization"
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
         Left            =   1800
         TabIndex        =   3
         Top             =   240
         Width           =   9375
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0C000&
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
         Left            =   1200
         TabIndex        =   2
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdcontinue 
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
      Height          =   614
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9840
      Width           =   4335
   End
End
Attribute VB_Name = "frmcognitivequiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: A Review of Theoretical Orientations in Clinical Psychology
'Form name: frmcognitivequiz
'Author: Calvin Pipenhagen
'Date Written: March 25, 2008
'Objective: This is portion of the quiz relating to the cognitive-behavioral theory. It determines
           'if the user selected the appropriate option.
Option Explicit

Private Sub cmdcontinue_Click()
cognitivequizsum = 0 'resets the global variable quizsum to 0 so it can be accessed by other quizes.
If Option2 = True Then 'the three options listed in conditional statements are the right answers.
    cognitivequizsum = cognitivequizsum + 1 'if a user selects one of these 1 is added to their score.
End If
If Option4 = True Then
    cognitivequizsum = cognitivequizsum + 1
End If
If Option8 = True Then
    cognitivequizsum = cognitivequizsum + 1
End If
frmcognitivequiz.Hide 'moves to next page
frmcognitivequiztwo.Show
End Sub


