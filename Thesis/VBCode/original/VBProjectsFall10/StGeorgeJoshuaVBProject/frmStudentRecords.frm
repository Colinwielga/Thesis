VERSION 5.00
Begin VB.Form frmStudentRecords 
   BackColor       =   &H00000080&
   Caption         =   "Lingua Vivens - Admin Options - Student Grades"
   ClientHeight    =   10665
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   ScaleHeight     =   10665
   ScaleWidth      =   11400
   Begin VB.Frame fraSortBy 
      BackColor       =   &H00000080&
      Caption         =   "Sort by ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3015
      Left            =   7920
      TabIndex        =   1
      Top             =   3000
      Width           =   2775
      Begin VB.OptionButton optScoreDown 
         BackColor       =   &H00000080&
         Caption         =   "Scores Descending"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   2400
         Width           =   2295
      End
      Begin VB.OptionButton optScoreUp 
         BackColor       =   &H00000080&
         Caption         =   "Scores Ascending"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1920
         Width           =   2295
      End
      Begin VB.OptionButton optClass 
         BackColor       =   &H00000080&
         Caption         =   "Class"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   1815
      End
      Begin VB.OptionButton optAlphaDescending 
         BackColor       =   &H00000080&
         Caption         =   "Last Name Decending"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   2415
      End
      Begin VB.OptionButton optAlphaAscending 
         BackColor       =   &H00000080&
         Caption         =   "Last Name Ascending"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Value           =   -1  'True
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdLogOut 
      BackColor       =   &H00808080&
      Caption         =   "LogOut"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9360
      Width           =   2775
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00808080&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8520
      Width           =   2775
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00808080&
      Caption         =   "Return to Administrator Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7680
      Width           =   2775
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9735
      Left            =   720
      ScaleHeight     =   9675
      ScaleWidth      =   6315
      TabIndex        =   7
      Top             =   360
      Width           =   6375
   End
   Begin VB.CommandButton cmdShowStudents 
      BackColor       =   &H00000080&
      Caption         =   "Show Student Grades "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6360
      Width           =   2775
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Student Records"
      BeginProperty Font 
         Name            =   "Roman"
         Size            =   48
         Charset         =   255
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   7320
      TabIndex        =   11
      Top             =   480
      Width           =   3375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   10335
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   11055
   End
End
Attribute VB_Name = "frmStudentRecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdLogOut_Click()
    frmStudentRecords.Hide
    Call LogOut
End Sub

Private Sub cmdQuit_Click()
    Call MakeConcurrent
    End
End Sub

Private Sub cmdReturn_Click()
    'returns to the administrator pane
    Call MakeConcurrent
    frmStudentRecords.Hide
    frmAdmin.Show
End Sub

Private Sub cmdShowStudents_Click()
    'Variables need for sorting using the bubble sort method
    Dim pos As Integer
    Dim Pass As Integer
    'Stores the values of each option button
    Dim AlphaUP As Boolean
    Dim AlphaDown As Boolean
    Dim ScoreUp As Boolean
    Dim ScoreDown As Boolean
    Dim ClassSort As Boolean
    'Gets the values of each option button
    AlphaUP = optAlphaAscending.Value
    AlphaDown = optAlphaDescending.Value
    ScoreUp = optScoreUp.Value
    ScoreDown = optScoreDown.Value
    ClassSort = optClass.Value
    'Test to see which button is true and initiates a unique sort for each button
    If AlphaUP = True Then
        For Pass = 1 To loginCtr - 1
            For pos = 1 To loginCtr - Pass
                If lastName(pos) > lastName(pos + 1) Then
                'Uses the swapstring subroutine to swap the various entries in the arrays
                    Call SwapString(firstName(pos), firstName(pos + 1))
                    Call SwapString(lastName(pos), lastName(pos + 1))
                    Call SwapString(userName(pos), userName(pos + 1))
                    Call SwapString(ClassEnrolled(pos), ClassEnrolled(pos + 1))
                    Call SwapString(studentGradeName(pos), studentGradeName(pos + 1))
                    Call SwapSingle(StudentGrade(pos), StudentGrade(pos + 1))
                    Call SwapInteger(studentCorrect(pos), studentCorrect(pos + 1))
                    Call SwapInteger(studentWrong(pos), studentWrong(pos + 1))
                    Call SwapInteger(StudentAttempted(pos), StudentAttempted(pos + 1))
                End If
            Next pos
        Next Pass
    ElseIf AlphaDown = True Then
        For Pass = 1 To loginCtr - 1
            For pos = 1 To loginCtr - Pass
                If lastName(pos) < lastName(pos + 1) Then
                    Call SwapString(firstName(pos), firstName(pos + 1))
                    Call SwapString(lastName(pos), lastName(pos + 1))
                    Call SwapString(userName(pos), userName(pos + 1))
                    Call SwapString(ClassEnrolled(pos), ClassEnrolled(pos + 1))
                    Call SwapString(studentGradeName(pos), studentGradeName(pos + 1))
                    Call SwapSingle(StudentGrade(pos), StudentGrade(pos + 1))
                    Call SwapInteger(studentCorrect(pos), studentCorrect(pos + 1))
                    Call SwapInteger(studentWrong(pos), studentWrong(pos + 1))
                    Call SwapInteger(StudentAttempted(pos), StudentAttempted(pos + 1))
                End If
            Next pos
        Next Pass
    ElseIf ScoreUp = True Then
        For Pass = 1 To loginCtr - 1
            For pos = 1 To loginCtr - Pass
                If StudentGrade(pos) < StudentGrade(pos + 1) Then
                    Call SwapString(firstName(pos), firstName(pos + 1))
                    Call SwapString(lastName(pos), lastName(pos + 1))
                    Call SwapString(userName(pos), userName(pos + 1))
                    Call SwapString(ClassEnrolled(pos), ClassEnrolled(pos + 1))
                    Call SwapString(studentGradeName(pos), studentGradeName(pos + 1))
                    Call SwapSingle(StudentGrade(pos), StudentGrade(pos + 1))
                    Call SwapInteger(studentCorrect(pos), studentCorrect(pos + 1))
                    Call SwapInteger(studentWrong(pos), studentWrong(pos + 1))
                    Call SwapInteger(StudentAttempted(pos), StudentAttempted(pos + 1))
                End If
            Next pos
        Next Pass
    ElseIf ScoreDown = True Then
        For Pass = 1 To loginCtr - 1
            For pos = 1 To loginCtr - Pass
                If StudentGrade(pos) > StudentGrade(pos + 1) Then
                    Call SwapString(firstName(pos), firstName(pos + 1))
                    Call SwapString(lastName(pos), lastName(pos + 1))
                    Call SwapString(userName(pos), userName(pos + 1))
                    Call SwapString(ClassEnrolled(pos), ClassEnrolled(pos + 1))
                    Call SwapString(studentGradeName(pos), studentGradeName(pos + 1))
                    Call SwapSingle(StudentGrade(pos), StudentGrade(pos + 1))
                    Call SwapInteger(studentCorrect(pos), studentCorrect(pos + 1))
                    Call SwapInteger(studentWrong(pos), studentWrong(pos + 1))
                    Call SwapInteger(StudentAttempted(pos), StudentAttempted(pos + 1))
                End If
            Next pos
        Next Pass
    ElseIf ClassSort = True Then
        For Pass = 1 To loginCtr - 1
            For pos = 1 To loginCtr - Pass
                If ClassEnrolled(pos) > ClassEnrolled(pos + 1) Then
                    Call SwapString(firstName(pos), firstName(pos + 1))
                    Call SwapString(lastName(pos), lastName(pos + 1))
                    Call SwapString(userName(pos), userName(pos + 1))
                    Call SwapString(ClassEnrolled(pos), ClassEnrolled(pos + 1))
                    Call SwapString(studentGradeName(pos), studentGradeName(pos + 1))
                    Call SwapSingle(StudentGrade(pos), StudentGrade(pos + 1))
                    Call SwapInteger(studentCorrect(pos), studentCorrect(pos + 1))
                    Call SwapInteger(studentWrong(pos), studentWrong(pos + 1))
                    Call SwapInteger(StudentAttempted(pos), StudentAttempted(pos + 1))
                End If
            Next pos
        Next Pass
    End If
    
    picResults.Cls
    picResults.Print "First Name"; Tab(20); "Last Name"; Tab(40); "Class Enrolled"; Tab(60); "Grade"
    picResults.Print "***************************************************************************************************************************"
    For pos = 1 To loginCtr
        picResults.Print firstName(pos); Tab(20); lastName(pos); Tab(40); ClassEnrolled(pos); Tab(60); FormatPercent(StudentGrade(pos))
    Next pos
End Sub
