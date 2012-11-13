VERSION 5.00
Begin VB.Form frmTwo 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Compute Scores and Averages"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6660
   LinkTopic       =   "Form2"
   ScaleHeight     =   4560
   ScaleWidth      =   6660
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H008080FF&
      Caption         =   "Quit"
      Height          =   735
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdComputeTotalScore 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Compute Students' Total Scores"
      Height          =   1215
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton cmdClassAverage 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Compute Class Average"
      Height          =   1215
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      Width           =   2175
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   120
      ScaleHeight     =   1035
      ScaleWidth      =   6195
      TabIndex        =   0
      Top             =   360
      Width           =   6255
   End
End
Attribute VB_Name = "frmTwo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Title Calculating Final Grades for a High School English Class
'Lynn Paradis
'October 31-November 1, 2006
'Objectives - The overall objective of this project was to give my
        'room mate, who is student teaching, an easier way of calculating
        'final grades after a project, and to find the class average.
        'I've changed the names of the students
        'to keep the students identities anonymous, and used fewer than a usual class
        'to keep the picResult boxes smaller and easier to read.
    'This from is to Find the Class Average and the Students' final grade.
'Subroutines
    ' The first subroutine works by adding the grades together so an average can be
        'calculated.  This subroutine reads the variables from a Module.  The grades are
        'then assigned a letter grade by using If/Then statements.
    'The second subroutine works by adding the grades together and dividing by the total number of points
        '(in this case everything was worth 20 points) to find the final grade.  And print
        'the final grades so the person running the program can view them.
    'The third subroutine is a Quit button.
Option Explicit

Private Sub cmdClassAverage_Click()
Dim Avg As Single
Dim Sum As Single, Pos As Integer
picResults.Cls
For Pos = 1 To Counter
    Sum = Sum + Grade1(Pos) + Grade2(Pos) + Grade3(Pos) + Grade4(Pos) + Grade5(Pos)
Next Pos

Avg = ((Sum / Counter) / 100)

If Avg >= 0.9 Then
    picResults.Print "The average final grade of Miss Leuthner's class is: " & FormatPercent(Avg)
    picResults.Print FormatPercent(Avg); " is an "; "A"
   
ElseIf Avg >= 0.8 And Avg < 0.9 Then
        picResults.Print "The average final grade of Miss Leuthner's class is: " & FormatPercent(Avg)
        picResults.Print FormatPercent(Avg); " is an "; "B"
    
ElseIf Avg >= 0.7 And Avg < 0.8 Then
        picResults.Print "The average final grade of Miss Leuthner's class is: " & FormatPercent(Avg)
        picResults.Print FormatPercent(Avg); " is an "; "C"
      
ElseIf Avg >= 0.6 And Avg < 0.7 Then
        picResults.Print "The average final grade of Miss Leuthner's class is: " & FormatPercent(Avg)
        picResults.Print FormatPercent(Avg); " is an "; "D"
       
ElseIf Avg < 0.59 Then
        picResults.Print "The average final grade of Miss Leuthner's class is: " & FormatPercent(Avg)
        picResults.Print FormatPercent(Avg); " is an "; "F"
        End If
End Sub


Private Sub cmdComputeTotalScore_Click()
Dim Sum As Integer, M As Integer
M = 0
For M = 1 To Counter
    Sum = 0
    Sum = Sum + Grade1(M) + Grade2(M) + Grade3(M) + Grade4(M) + Grade5(M)
    picResults.Print "Student, "; Names(M); " Final grade is " & Sum; "%"
Next M
End Sub

Private Sub cmdQuit_Click()
End
End Sub



