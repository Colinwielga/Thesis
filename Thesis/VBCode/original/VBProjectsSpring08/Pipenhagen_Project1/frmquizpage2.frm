VERSION 5.00
Begin VB.Form frmpsychodynamicquiztwo 
   BackColor       =   &H0080C0FF&
   Caption         =   "Form1"
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
      Height          =   2775
      Left            =   1080
      TabIndex        =   19
      Top             =   6120
      Width           =   13815
      Begin VB.OptionButton Option9 
         BackColor       =   &H0080C0FF&
         Caption         =   "Option1"
         Height          =   375
         Left            =   2520
         TabIndex        =   27
         Top             =   2160
         Width           =   255
      End
      Begin VB.OptionButton Option8 
         BackColor       =   &H0080C0FF&
         Caption         =   "Option1"
         Height          =   375
         Left            =   2520
         TabIndex        =   26
         Top             =   1440
         Width           =   255
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H0080C0FF&
         Caption         =   "Option1"
         Height          =   375
         Left            =   2520
         TabIndex        =   25
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label15 
         BackColor       =   &H0080C0FF&
         Caption         =   "A byproduct of the Oedipus complex"
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
         Top             =   1440
         Width           =   8175
      End
      Begin VB.Label Label11 
         BackColor       =   &H0080C0FF&
         Caption         =   "An natural instinct to blame others for our problems"
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
         Top             =   2160
         Width           =   8175
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080C0FF&
         Caption         =   "A defense mechanism"
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
         Top             =   840
         Width           =   8175
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080C0FF&
         Caption         =   "Projection is an example of what?"
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
         TabIndex        =   21
         Top             =   240
         Width           =   9375
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080C0FF&
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
      BackColor       =   &H0080C0FF&
      Height          =   2655
      Left            =   1080
      TabIndex        =   10
      Top             =   3120
      Width           =   13815
      Begin VB.OptionButton Option6 
         BackColor       =   &H0080C0FF&
         Caption         =   "Option1"
         Height          =   375
         Left            =   2640
         TabIndex        =   18
         Top             =   1920
         Width           =   255
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H0080C0FF&
         Caption         =   "Option1"
         Height          =   375
         Left            =   2640
         TabIndex        =   17
         Top             =   1320
         Width           =   255
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H0080C0FF&
         Caption         =   "Option1"
         Height          =   375
         Left            =   2640
         TabIndex        =   16
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label10 
         BackColor       =   &H0080C0FF&
         Caption         =   "Say whatever comes to mind regardsless of its content"
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
         BackColor       =   &H0080C0FF&
         Caption         =   "Say whatever comes to mind, provided the content is appropriate"
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
         BackColor       =   &H0080C0FF&
         Caption         =   "Talk in a structured format about their unconscious"
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
         BackColor       =   &H0080C0FF&
         Caption         =   "What do patients do during free association?"
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
         BackColor       =   &H0080C0FF&
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
      BackColor       =   &H0080C0FF&
      Height          =   2775
      Left            =   1080
      TabIndex        =   1
      Top             =   0
      Width           =   13815
      Begin VB.OptionButton Option3 
         BackColor       =   &H0080C0FF&
         Caption         =   "Option1"
         Height          =   375
         Left            =   2760
         TabIndex        =   9
         Top             =   2040
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Option1"
         Height          =   375
         Left            =   2760
         TabIndex        =   8
         Top             =   1440
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Option1"
         Height          =   375
         Left            =   2760
         TabIndex        =   7
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label17 
         BackColor       =   &H0080C0FF&
         Caption         =   "Anal"
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
         Left            =   3120
         TabIndex        =   6
         Top             =   2040
         Width           =   8175
      End
      Begin VB.Label Label16 
         BackColor       =   &H0080C0FF&
         Caption         =   "Phallic"
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
         Left            =   3120
         TabIndex        =   5
         Top             =   1440
         Width           =   8175
      End
      Begin VB.Label Label14 
         BackColor       =   &H0080C0FF&
         Caption         =   "Oral"
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
         Left            =   3120
         TabIndex        =   4
         Top             =   840
         Width           =   8175
      End
      Begin VB.Label Label13 
         BackColor       =   &H0080C0FF&
         Caption         =   "What is the second psychosexual stage?"
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
         Left            =   1920
         TabIndex        =   3
         Top             =   240
         Width           =   9375
      End
      Begin VB.Label Label12 
         BackColor       =   &H0080C0FF&
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
         Left            =   1440
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
      Height          =   615
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9480
      Width           =   4335
   End
End
Attribute VB_Name = "frmpsychodynamicquiztwo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: A Review of Theoretical Orientations in Clinical Psychology
'Form name: frmpsychodynamicquiztwo
'Author: Calvin Pipenhagen
'Date Written: March 25, 2008
'Objective: This is the second page of a three page quiz on psychodynamic theory. It purpose
           'is to get the user's answers and determine if they are right
Option Explicit


Private Sub cmdcontinue_Click()
If Option3 = True Then 'if the user answered correctly his or her score is added to.
    psychodynamicquizsum = psychodynamicquizsum + 1
End If
If Option6 = True Then
    psychodynamicquizsum = psychodynamicquizsum + 1
End If
If Option7 = True Then
    psychodynamicquizsum = psychodynamicquizsum + 1
End If
frmpsychodynamicquiztwo.Hide ' moves to last page of quiz
frmpsychodynamicquizthree.Show
End Sub



