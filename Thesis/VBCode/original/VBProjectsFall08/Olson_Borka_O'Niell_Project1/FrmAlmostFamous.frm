VERSION 5.00
Begin VB.Form FrmForm2 
   BackColor       =   &H00800080&
   Caption         =   "Almost Famous"
   ClientHeight    =   8865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13770
   LinkTopic       =   "Form2"
   Picture         =   "FrmAlmostFamous.frx":0000
   ScaleHeight     =   8865
   ScaleWidth      =   13770
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Menu"
      Height          =   975
      Left            =   600
      TabIndex        =   21
      Top             =   6840
      Width           =   2295
   End
   Begin VB.Frame Frame5 
      Caption         =   "5. Whose hand writes the opening credits?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3840
      TabIndex        =   17
      Top             =   5160
      Width           =   6135
      Begin VB.OptionButton Option15 
         Caption         =   "Patrick Fugit"
         Height          =   495
         Left            =   240
         TabIndex        =   20
         Top             =   960
         Width           =   3015
      End
      Begin VB.OptionButton Option14 
         Caption         =   "Cameron Crowe"
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   600
         Width           =   2895
      End
      Begin VB.OptionButton Option13 
         Caption         =   "Billy Crudup"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "4. Who gave guitar lessons to the memebers of Stillwater on the set?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3840
      TabIndex        =   13
      Top             =   3840
      Width           =   6135
      Begin VB.OptionButton Option12 
         Caption         =   "Peter Frampton"
         Height          =   495
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   2775
      End
      Begin VB.OptionButton Option11 
         Caption         =   "Eric Clapton"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   600
         Width           =   2775
      End
      Begin VB.OptionButton Option10 
         Caption         =   "Jimmy Page"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "3. Where does Penny Lane want to move?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3840
      TabIndex        =   9
      Top             =   2520
      Width           =   6135
      Begin VB.OptionButton Option9 
         Caption         =   "New York"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   3015
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Morocco"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   2775
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Argentina"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "2. What group is the Stillwater tour based on?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3840
      TabIndex        =   5
      Top             =   1320
      Width           =   6135
      Begin VB.OptionButton Option6 
         Caption         =   "The Allman Brothers"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   2415
      End
      Begin VB.OptionButton Option5 
         Caption         =   "The Who"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   3075
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Led Zepplin"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.CommandButton cmdScoreFamous 
      Caption         =   "Calculate Score"
      Height          =   855
      Left            =   600
      TabIndex        =   4
      Top             =   5760
      Width           =   2295
   End
   Begin VB.Frame Question1 
      Caption         =   "1. The film is semi-autobiographical; whose life is it based upon?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3840
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin VB.OptionButton Option3 
         Caption         =   "Kate Hudson"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   3615
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Patrick Fugit"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   3855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Cameron Crowe"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4095
      End
   End
End
Attribute VB_Name = "FrmForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X As Integer
Dim Ctr As Integer
'Movie Trivia
'FrmAlmostFamous
'Amber Olson, Emily Borka, Shannon O'Neill
'11-1-08
' The purpose of this form is to allow the user to click their answer in Radio Buttons and then get a score, like a quiz.
Private Sub cmdReturn_Click()
'This is to clear the Radio buttons and to get back to the main page and hide all the rest.
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
Option5.Value = False
Option6.Value = False
Option7.Value = False
Option8.Value = False
Option9.Value = False
Option10.Value = False
Option11.Value = False
Option12.Value = False
Option13.Value = False
Option14.Value = False
Option15.Value = False
FrmForm1.Show
FrmForm2.Hide
FrmForm3.Hide
FrmForm4.Hide
Frmsearch.Hide

End Sub

Private Sub cmdScoreFamous_Click()
'These are the answers to the radio buttons. Using boolean values to give the user a score after they are done with the quiz.
'And print them out.
Ctr = 0
If Option1 = True Then
   Ctr = Ctr + 1
End If

If Option6 = True Then
    Ctr = Ctr + 1
End If

If Option8 = True Then
    Ctr = Ctr + 1
End If
If Option12 = True Then
    Ctr = Ctr + 1
End If
If Option14 = True Then
    Ctr = Ctr + 1
End If

MsgBox UserName & " you got " & Ctr & " right out of 5.", , "Score"


End Sub

