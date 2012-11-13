VERSION 5.00
Begin VB.Form frmTeacherGrades 
   BackColor       =   &H008080FF&
   Caption         =   "See Student Grades"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   9750
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDesigner 
      BackColor       =   &H00C0C0FF&
      Height          =   285
      Left            =   6600
      TabIndex        =   3
      Text            =   "Designed by Amanda Aamodt"
      Top             =   5640
      Width           =   2895
   End
   Begin VB.CommandButton cmdBest 
      Caption         =   "Who did the best on the test???  (Grades from top to bottom on the test)"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   1335
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00C0C0FF&
      Height          =   5175
      Left            =   1800
      ScaleHeight     =   5115
      ScaleWidth      =   7755
      TabIndex        =   1
      Top             =   240
      Width           =   7815
   End
   Begin VB.CommandButton cmdGradesAlpha 
      Caption         =   "View Student Grades by Alphabetical Order"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmTeacherGrades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    'declaring variables
    Dim I As Integer, ID(1 To 20) As Double
    Dim First(1 To 20) As String, Last(1 To 20) As String
    Dim Grade1(1 To 20) As Integer, Grade2(1 To 20) As Integer
    Dim TestGrade(1 To 20) As Integer
    Dim X As Single, FinalGrade As Integer
    Dim Pass As Integer, Temp As Integer
    Dim Temp2 As String, Temp3 As String

Private Sub cmdBest_Click()
    picResults.Cls
    Open App.Path & "\Grades.txt" For Input As #1   'opens the file with the students grades
    picResults.Print "Test Grade", "First Name", "Last Name"
    picResults.Print
    I = 0   'initialize counter I to zero, to be used for position in the array
    Do Until EOF(1)
        I = I + 1       'increment counter I each time throught the loop
                        'to move to the next postion in the array
                'Read next data set from the file into the array
        Input #1, ID(I), First(I), Last(I), Grade1(I), Grade2(I), TestGrade(I)
    Loop
    Close #1    'Close the file used for input
    
    For Pass = 1 To 19      'the bubble sort allows us to order the students by highest to lowest grade on the test
        For I = 1 To 20 - Pass
            If TestGrade(I) < TestGrade(I + 1) Then
                Temp = TestGrade(I)             'switches the test grades
                TestGrade(I) = TestGrade(I + 1)
                TestGrade(I + 1) = Temp
                Temp2 = First(I)                'switches the first names so they match the test grades
                First(I) = First(I + 1)
                First(I + 1) = Temp2
                Temp3 = Last(I)                 'switches the last names so they match the test grades
                Last(I) = Last(I + 1)
                Last(I + 1) = Temp3
            End If
        Next I
    Next Pass
    
    For I = 1 To 20
        picResults.Print TestGrade(I), First(I), Last(I)
    Next I
End Sub

Private Sub cmdGradesAlpha_Click()
    picResults.Cls  'clears the picturebox used for output
    
    Open App.Path & "\Grades.txt" For Input As #1   'opens the file with the students grades
    picResults.Print "ID", "First Name", "Last Name", "Grade 1", "Grade 2", "Test Grade"
    picResults.Print
    I = 0   'initialize counter I to zero, to be used for position in the array
    Do Until EOF(1)
        I = I + 1       'increment counter I each time throught the loop
                        'to move to the next postion in the array
                'Read next data set from the file into the array
        Input #1, ID(I), First(I), Last(I), Grade1(I), Grade2(I), TestGrade(I)
        picResults.Print ID(I), First(I), Last(I), Grade1(I), Grade2(I), TestGrade(I)
    Loop
    Close #1    'Close the file used for input
End Sub


