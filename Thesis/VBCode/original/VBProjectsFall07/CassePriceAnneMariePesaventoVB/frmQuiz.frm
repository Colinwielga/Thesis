VERSION 5.00
Begin VB.Form frmQuestion1 
   BackColor       =   &H003D30AD&
   Caption         =   "Question 1"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   11970
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   4935
      Left            =   7200
      Picture         =   "frmQuiz.frx":0000
      ScaleHeight     =   4875
      ScaleWidth      =   4395
      TabIndex        =   6
      Top             =   240
      Width           =   4455
   End
   Begin VB.OptionButton OptQ1 
      BackColor       =   &H003D30AD&
      Caption         =   "Romancing the Ladies"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0084C11E&
      Height          =   495
      Index           =   3
      Left            =   960
      TabIndex        =   5
      Top             =   4680
      Width           =   4215
   End
   Begin VB.OptionButton OptQ1 
      BackColor       =   &H003D30AD&
      Caption         =   "Being Macho"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0084C11E&
      Height          =   375
      Index           =   2
      Left            =   960
      TabIndex        =   4
      Top             =   3840
      Width           =   4215
   End
   Begin VB.OptionButton OptQ1 
      BackColor       =   &H003D30AD&
      Caption         =   "Looking Into Magical Mirrors"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0084C11E&
      Height          =   615
      Index           =   1
      Left            =   960
      TabIndex        =   2
      Top             =   2760
      Width           =   5295
   End
   Begin VB.OptionButton OptQ1 
      BackColor       =   &H003D30AD&
      Caption         =   "Reading"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0084C11E&
      Height          =   615
      Index           =   0
      Left            =   960
      MaskColor       =   &H003D30AD&
      TabIndex        =   1
      Top             =   1800
      Width           =   2535
   End
   Begin VB.CommandButton cmdName 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Next>>"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   960
      TabIndex        =   0
      Top             =   5520
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H003D30AD&
      Caption         =   "Which of the following activities                do you most enjoy?"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0084C11E&
      Height          =   1095
      Left            =   720
      TabIndex        =   3
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "frmQuestion1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdName_Click()

'if statement to add 1 to the a specific counter depending on which option is selected

If OptQ1(0) = True Then
    CtrA = CtrA + 1
ElseIf OptQ1(1) = True Then
    CtrB = CtrB + 1
ElseIf OptQ1(2) = True Then
    CtrC = CtrC + 1
ElseIf OptQ1(3) = True Then
    CtrD = CtrD + 1
End If

'moves to next form
frmQuestion1.Hide
frmQuestion2.Show

End Sub


