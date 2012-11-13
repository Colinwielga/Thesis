VERSION 5.00
Begin VB.Form frmSearchAndSort 
   BackColor       =   &H0080C0FF&
   Caption         =   "Search And Sort"
   ClientHeight    =   8655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10860
   LinkTopic       =   "Form1"
   ScaleHeight     =   8655
   ScaleWidth      =   10860
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdSortByAverage 
      Caption         =   "Sort by Average"
      Height          =   375
      Left            =   480
      TabIndex        =   11
      Top             =   5040
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdSortbyName 
      Caption         =   "Sort by Name"
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   4440
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Main Menu"
      Height          =   975
      Left            =   4800
      TabIndex        =   9
      Top             =   7080
      Width           =   2655
   End
   Begin VB.PictureBox picScore 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   4080
      ScaleHeight     =   555
      ScaleWidth      =   2595
      TabIndex        =   8
      Top             =   1560
      Width           =   2655
   End
   Begin VB.CommandButton cmdMoreData 
      Caption         =   "Enter More Data"
      Height          =   975
      Left            =   360
      TabIndex        =   6
      Top             =   7080
      Width           =   3615
   End
   Begin VB.PictureBox picAverageScore 
      BackColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   4080
      ScaleHeight     =   2955
      ScaleWidth      =   4155
      TabIndex        =   5
      Top             =   3600
      Width           =   4215
   End
   Begin VB.PictureBox picRounds 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      ScaleHeight     =   555
      ScaleWidth      =   1275
      TabIndex        =   3
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdAnalyze 
      Caption         =   "Analyze Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   9855
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   975
      Left            =   8160
      TabIndex        =   0
      Top             =   7080
      Width           =   2415
   End
   Begin VB.Label lblScore 
      BackColor       =   &H0080C0FF&
      Caption         =   "Most recent Score:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   1560
      Width           =   3135
   End
   Begin VB.Label lblAverageScore 
      BackColor       =   &H0080C0FF&
      Caption         =   "Average Total Score:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   3720
      Width           =   3255
   End
   Begin VB.Label lblRounds 
      BackColor       =   &H0080C0FF&
      Caption         =   "Total Rounds Played:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2640
      Width           =   3375
   End
End
Attribute VB_Name = "frmSearchAndSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: GolfGuide
':Form Name:  frmSearchAndSort
':Author:   Tyler Cash
':Date written:  March 22, 2009


'This form allows the user to display information about the data in their text file.
'The user clicks the analyze button and the program computes and displays various
'statistics.

Option Explicit

'Declaring variables
Dim CTR As Integer
Dim Rounds(1 To 3) As Integer
Dim Courses(1 To 3) As String
Dim Average(1 To 3) As Single

Private Sub cmdAnalyze_Click()
'This button calculates several statistics about the data in the users text file.

'Declaring Variables
Dim J As Integer

Dim RoundsPlayed As Integer
Dim AverageTotalScore(1 To 3) As Single
Dim Sum(1 To 3) As Integer


'Loading the Courses Array with course names
Courses(1) = "Rich-Spring Golf Club"
Courses(2) = "Albany Golf Club"
Courses(3) = "Blackberry Ridge Golf Course"

'Finding the total number of rounds played.  We divide by 18 because when we read the
'text file, we counted how many total holes were played.
    RoundsPlayed = (Rounds(1) + Rounds(2) + Rounds(3)) / 18
    
'Computing the sum of all scores for each of the three courses
    For J = 1 To CTR
        Sum(1) = Sum(1) + Score1(J)
        Sum(2) = Sum(2) + Score2(J)
        Sum(3) = Sum(3) + Score3(J)
    Next J
    
 'If the course has at least one round played on it, we compute the average
 'score for that course.
    For J = 1 To 3
        If Rounds(J) > 0 Then
            Average(J) = Sum(J) / (Rounds(J) / 18)
        End If
    Next J
'Displaying the total score of the inputted data and telling the user their score in
'relation to par.
    If TotalScore - 72 > 0 Then
        picScore.Print TotalScore, "That's "; Abs(TotalScore - 72); " over par."
    ElseIf TotalScore - 72 < 0 Then
        picScore.Print TotalScore, "That's "; Abs(TotalScore - 72); " under par."
    ElseIf TotalScore - 72 = 0 Then
        picScore.Print TotalScore, "You scored par!"
    End If
    
'Displaying the number of rounds played
    picRounds.Print RoundsPlayed
    
'Displaying the average scores for each course and all courses in a well formatted table
    picAverageScore.Print "Course"; Tab(30); "Average Total Score"
    picAverageScore.Print "----------------------------------------------------------------------------"
    For J = 1 To 3
        picAverageScore.Print Courses(J); Tab(40); FormatNumber(Average(J), 1)
    Next J
    
'Disabling the analyze button so the user can't click it again.
    cmdAnalyze.Enabled = False
    
'Enabling the Sorting Buttons
    cmdSortbyName.Visible = True
    cmdSortByAverage.Visible = True
End Sub


Private Sub cmdExit_Click()
'This button changes forms to the main menu.

'Changing forms
    frmTitle.Show
    
'This form is unloaded in case the user tries to enter another data set.
'There are certain tasks that need to be performed each time this form is accessed.
    Unload Me
    
End Sub

Private Sub cmdMoreData_Click()
'This button changes forms to the form allowing for another set of scoring input.
    frmScoring.Show

'This form is unloaded in case the user tries to enter another data set.
'There are certain tasks that need to be performed each time this form is accessed.
    Unload Me
    
End Sub

Private Sub cmdQuit_Click()
'This button ends the program
    End
End Sub

Private Sub cmdSortByAverage_Click()
'This sub sorts the courses and average arrays by average
'It also keeps the correct course and average paired together

'Declaring variables
Dim Pass As Integer
Dim Pos As Integer
Dim J As Integer
Dim tempName As String
Dim tempAverage As Single

'Sort the Averages
    For Pass = 1 To 2
        For Pos = 1 To 3 - Pass
            If Average(Pos) > Average(Pos + 1) Then
                tempName = Courses(Pos)
                Courses(Pos) = Courses(Pos + 1)
                Courses(Pos + 1) = tempName
                tempAverage = Average(Pos)
                Average(Pos) = Average(Pos + 1)
                Average(Pos + 1) = tempAverage
            End If
        Next Pos
    Next Pass
    
'Print the Course Names and Averages
    picAverageScore.Cls
    picAverageScore.Print "Course"; Tab(30); "Average Total Score"
    picAverageScore.Print "----------------------------------------------------------------------------"
    
    For J = 1 To 3
        picAverageScore.Print Courses(J); Tab(40); FormatNumber(Average(J), 1)
    Next J
End Sub

Private Sub cmdSortbyName_Click()
'This sub sorts the courses and Average arrays by course name
'It also keeps the correct course and average paired together

'Declaring variables
Dim Pass As Integer
Dim Pos As Integer
Dim J As Integer
Dim tempName As String
Dim tempAverage As Single

'Sort the names
    For Pass = 1 To 2
        For Pos = 1 To 3 - Pass
            If Courses(Pos) > Courses(Pos + 1) Then
                tempName = Courses(Pos)
                Courses(Pos) = Courses(Pos + 1)
                Courses(Pos + 1) = tempName
                tempAverage = Average(Pos)
                Average(Pos) = Average(Pos + 1)
                Average(Pos + 1) = tempAverage
            End If
        Next Pos
    Next Pass
    
'Print the Course names and averages
    picAverageScore.Cls
    picAverageScore.Print "Course"; Tab(30); "Average Total Score"
    picAverageScore.Print "----------------------------------------------------------------------------"
    
    For J = 1 To 3
        picAverageScore.Print Courses(J); Tab(40); FormatNumber(Average(J), 1)
    Next J
        
End Sub

Private Sub Form_Load()
'When this form loads, this sub reads the scoring data from the text file into three
'arrays (one for each course).

'Setting some initial values in case the user enters multiple data sets
CTR = 0
Rounds(1) = 0
Rounds(2) = 0
Rounds(3) = 0

'Opening the text file determined by the user previously in the program
    Open FileName For Input As #1
    
'Loading the data from the text file into arrays
    Do Until EOF(1)
    CTR = CTR + 1
        Input #1, Score1(CTR), Score2(CTR), Score3(CTR)
    
'Counting how many holes have been played on each course.
        If Score1(CTR) > 0 Then
            Rounds(1) = Rounds(1) + 1
        End If
        If Score2(CTR) > 0 Then
            Rounds(2) = Rounds(2) + 1
        End If
        If Score3(CTR) > 0 Then
            Rounds(3) = Rounds(3) + 1
        End If
    Loop
    
'Closing the text file.
    Close #1
    
'Enabling the Analyze button in case it was disabled by a previous analyzing session.
    cmdAnalyze.Enabled = True
    
End Sub
