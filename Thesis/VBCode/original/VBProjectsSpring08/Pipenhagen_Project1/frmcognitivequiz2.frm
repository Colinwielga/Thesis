VERSION 5.00
Begin VB.Form frmcognitivequiztwo 
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
      Height          =   2655
      Left            =   360
      TabIndex        =   19
      Top             =   6480
      Width           =   14535
      Begin VB.OptionButton Option9 
         BackColor       =   &H00C0C000&
         Caption         =   "Option1"
         Height          =   255
         Left            =   2520
         TabIndex        =   27
         Top             =   2280
         Width           =   255
      End
      Begin VB.OptionButton Option8 
         BackColor       =   &H00C0C000&
         Caption         =   "Option1"
         Height          =   255
         Left            =   2520
         TabIndex        =   26
         Top             =   1680
         Width           =   255
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H00C0C000&
         Caption         =   "Option1"
         Height          =   255
         Left            =   2520
         TabIndex        =   25
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0C000&
         Caption         =   "Self-actualization"
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
         TabIndex        =   24
         Top             =   1680
         Width           =   8175
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0C000&
         Caption         =   "Rational reconstruction"
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
         TabIndex        =   23
         Top             =   2280
         Width           =   8175
      End
      Begin VB.Label Label4 
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
         Left            =   2880
         TabIndex        =   22
         Top             =   960
         Width           =   8175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         Caption         =   "The monitoring of thoughts is particularly important for which therapy?"
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
         Left            =   1680
         TabIndex        =   21
         Top             =   240
         Width           =   12615
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "6."
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
         TabIndex        =   20
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C000&
      Height          =   2655
      Left            =   360
      TabIndex        =   10
      Top             =   3480
      Width           =   14535
      Begin VB.OptionButton Option6 
         BackColor       =   &H00C0C000&
         Caption         =   "Option1"
         Height          =   255
         Left            =   2640
         TabIndex        =   18
         Top             =   2040
         Width           =   255
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00C0C000&
         Caption         =   "Option1"
         Height          =   255
         Left            =   2640
         TabIndex        =   17
         Top             =   1320
         Width           =   255
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0C000&
         Caption         =   "Option1"
         Height          =   255
         Left            =   2640
         TabIndex        =   16
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0C000&
         Caption         =   "To establish rapport between the client and therapist"
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
         Left            =   3000
         TabIndex        =   15
         Top             =   2040
         Width           =   8175
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C000&
         Caption         =   "To reduce anxiety"
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
         Left            =   3000
         TabIndex        =   14
         Top             =   1320
         Width           =   8175
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C000&
         Caption         =   "To provide the client with more coping options."
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
         Left            =   3000
         TabIndex        =   13
         Top             =   720
         Width           =   8175
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C000&
         Caption         =   "What is the goal of behavioral rehersal?"
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
         Left            =   1800
         TabIndex        =   12
         Top             =   120
         Width           =   9375
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C000&
         Caption         =   "5."
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
         Left            =   1320
         TabIndex        =   11
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      Height          =   3015
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   14535
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0C000&
         Caption         =   "Option1"
         Height          =   255
         Left            =   2640
         TabIndex        =   9
         Top             =   2400
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C000&
         Caption         =   "Option1"
         Height          =   255
         Left            =   2640
         TabIndex        =   8
         Top             =   1800
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Option1"
         Height          =   255
         Left            =   2640
         TabIndex        =   7
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0C000&
         Caption         =   "If is one of the only effective treatments for OCD"
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
         Left            =   3000
         TabIndex        =   6
         Top             =   2400
         Width           =   8175
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0C000&
         Caption         =   "It is a blatant pseudoscience"
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
         Left            =   3000
         TabIndex        =   5
         Top             =   1800
         Width           =   8175
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0C000&
         Caption         =   "Results are promising, but more research is needed to form any conclusions."
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
         Left            =   3000
         TabIndex        =   4
         Top             =   1080
         Width           =   8175
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0C000&
         Caption         =   "Which of the following is true about Exposure plus response prevention therapy?"
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
         Left            =   1800
         TabIndex        =   3
         Top             =   120
         Width           =   10455
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0C000&
         Caption         =   "4."
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
         Left            =   1320
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
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9840
      Width           =   4335
   End
End
Attribute VB_Name = "frmcognitivequiztwo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: A Review of Theoretical Orientations in Clinical Psychology
'Form name: frmcognitivequiztwo
'Author: Calvin Pipenhagen
'Date Written: March 25, 2008
'Objective: This is portion of the quiz relating to the cognitive-behavioral theory. It determines
           'if the user selected the appropriate option.
Option Explicit

Private Sub cmdcontinue_Click()

If Option3 = True Then 'checks to see if the user answered correctly. Determines his or her score based on this.
    cognitivequizsum = cognitivequizsum + 1
End If
If Option4 = True Then
    cognitivequizsum = cognitivequizsum + 1
End If
If Option9 = True Then
    cognitivequizsum = cognitivequizsum + 1
End If

frmcognitivequiztwo.Hide 'proceeds to last page of quiz
frmcognitivequizthree.Show
End Sub
