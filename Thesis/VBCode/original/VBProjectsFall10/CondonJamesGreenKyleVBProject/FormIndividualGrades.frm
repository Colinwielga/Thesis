VERSION 5.00
Begin VB.Form FormIndividualGrades 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   12165
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14430
   FillColor       =   &H00C00000&
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   ScaleHeight     =   12165
   ScaleWidth      =   14430
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit the Grading Program"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6840
      TabIndex        =   5
      Top             =   9480
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      Height          =   4695
      Left            =   6240
      Picture         =   "FormIndividualGrades.frx":0000
      ScaleHeight     =   4635
      ScaleWidth      =   4875
      TabIndex        =   4
      Top             =   960
      Width           =   4935
   End
   Begin VB.PictureBox picGrade 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   480
      ScaleHeight     =   2115
      ScaleWidth      =   12915
      TabIndex        =   3
      Top             =   6720
      Width           =   12975
   End
   Begin VB.CommandButton cmdComputeGrade 
      Caption         =   "Compute The Student's Grade"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1320
      TabIndex        =   2
      Top             =   4080
      Width           =   2775
   End
   Begin VB.CommandButton cmdName 
      Caption         =   "Student Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      TabIndex        =   1
      Top             =   2040
      Width           =   2175
   End
   Begin VB.PictureBox picName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   600
      ScaleHeight     =   1035
      ScaleWidth      =   3195
      TabIndex        =   0
      Top             =   480
      Width           =   3255
   End
End
Attribute VB_Name = "FormIndividualGrades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdComputeGrade_Click()
picGrade.Print StudentName; " recieved:"
'this assings a letter grade to the student depending on the amount of points recieved through out the project and prints them out in a picture box
Select Case RunningTotal
    Case Is > 90
        picGrade.Print RunningTotal; " points out of a possible 100 points which means the student received a "; FormatPercent(RunningTotal / 100, 2); " and recieved an A"
    Case 80 To 89
        picGrade.Print RunningTotal; " points out of a possible 100 points which means the student received a "; FormatPercent(RunningTotal / 100, 2); " and recieved a B"
    Case 70 To 79
         picGrade.Print RunningTotal; " points out of a possible 100 points which means the student received a "; FormatPercent(RunningTotal / 100, 2); " and recieved a C"
    Case 60 To 69
         picGrade.Print RunningTotal; " points out of a possible 100 points which means the student received a "; FormatPercent(RunningTotal / 100, 2); " and recieved a D"
    Case Is < 60
         picGrade.Print RunningTotal; " points out of a possible 100 points which means the student received a "; FormatPercent(RunningTotal / 100, 2); " and recieved an F"
End Select

End Sub

Private Sub cmdName_Click()
picName.Print StudentName
End Sub

Private Sub cmdQuit_Click()
End

End Sub


