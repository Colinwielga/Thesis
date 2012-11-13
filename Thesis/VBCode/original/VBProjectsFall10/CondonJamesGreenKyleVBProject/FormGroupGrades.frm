VERSION 5.00
Begin VB.Form FormGroupGrades 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   11205
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14235
   FillColor       =   &H000000C0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   11205
   ScaleWidth      =   14235
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
      Height          =   975
      Left            =   5760
      TabIndex        =   8
      Top             =   9960
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   5895
      Left            =   5400
      Picture         =   "FormGroupGrades.frx":0000
      ScaleHeight     =   5835
      ScaleWidth      =   7275
      TabIndex        =   7
      Top             =   360
      Width           =   7335
   End
   Begin VB.PictureBox picGrades 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   600
      ScaleHeight     =   2955
      ScaleWidth      =   11955
      TabIndex        =   6
      Top             =   6480
      Width           =   12015
   End
   Begin VB.CommandButton cmdGrades 
      Caption         =   "ComputeStudent Grades"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1200
      TabIndex        =   5
      Top             =   4200
      Width           =   2535
   End
   Begin VB.CommandButton cmdNames 
      Caption         =   "Student Names"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1440
      TabIndex        =   4
      Top             =   2280
      Width           =   1935
   End
   Begin VB.PictureBox picName2 
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
      Left            =   2520
      ScaleHeight     =   1035
      ScaleWidth      =   1995
      TabIndex        =   1
      Top             =   960
      Width           =   2055
   End
   Begin VB.PictureBox picName1 
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
      Left            =   240
      ScaleHeight     =   1035
      ScaleWidth      =   2115
      TabIndex        =   0
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label lbl 
      Caption         =   "Student 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label lblStudent1 
      Caption         =   "Student 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "FormGroupGrades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGrades_Click()

picGrades.Print StudentName; " and "; StudentName2; " recieved:"
'this assings a letter grade to the student depending on the amount of points recieved through out the project and prints them out in a picture box
Select Case RunningTotal
    Case Is >= 126
        picGrades.Print RunningTotal; " points out of a possible 140 points which is a "; FormatPercent(RunningTotal / 140, 2); " and means the Students recieved an A"
    Case 112 To 125
        picGrades.Print RunningTotal; " points out of a possible 140 points which is a "; FormatPercent(RunningTotal / 140, 2); " and means the Students recieved a B"
    Case 98 To 111
         picGrades.Print RunningTotal; " points out of a possible 140 points which is a "; FormatPercent(RunningTotal / 140, 2); " and means the Students recieved a C"
    Case 84 To 110
         picGrades.Print RunningTotal; " points out of a possible 140 points which is a "; FormatPercent(RunningTotal / 140, 2); " and means the Students recieved a D"
    Case Is < 84
        picGrades.Print RunningTotal; " points out of a possible 140 points which is a "; FormatPercent(RunningTotal / 140, 2); " and means the Students recieved an F"
End Select

End Sub




Private Sub cmdNames_Click()
picName1.Print StudentName
picName2.Print StudentName2
End Sub

Private Sub cmdQuit_Click()
End

End Sub
