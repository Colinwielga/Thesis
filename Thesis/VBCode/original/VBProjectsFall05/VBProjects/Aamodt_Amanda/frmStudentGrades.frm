VERSION 5.00
Begin VB.Form frmStudentGrades 
   BackColor       =   &H00FFFF80&
   Caption         =   "Student Grades"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   8595
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDesigner 
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Left            =   240
      TabIndex        =   17
      Text            =   "Designed by Amanda Aamodt"
      Top             =   4320
      Width           =   2535
   End
   Begin VB.PictureBox picLetterGrade 
      BackColor       =   &H00FFFFC0&
      Height          =   735
      Left            =   5040
      ScaleHeight     =   675
      ScaleWidth      =   2955
      TabIndex        =   16
      Top             =   240
      Width           =   3015
   End
   Begin VB.TextBox txtDminus 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   5760
      TabIndex        =   15
      Text            =   "62-63.99% = D-"
      Top             =   3600
      Width           =   2295
   End
   Begin VB.TextBox txtD 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   5760
      TabIndex        =   14
      Text            =   "64-69.99% = D"
      Top             =   3000
      Width           =   2295
   End
   Begin VB.TextBox txtDplus 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   5760
      TabIndex        =   13
      Text            =   "70-71.99% = D+"
      Top             =   2400
      Width           =   2295
   End
   Begin VB.TextBox txtCminus 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   5760
      TabIndex        =   12
      Text            =   "72-73.99% = C-"
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox txtC 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   5760
      TabIndex        =   11
      Text            =   "74-79.99% = C"
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox txtCplus 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   3240
      TabIndex        =   10
      Text            =   "80-81.99% = C+"
      Top             =   4200
      Width           =   2295
   End
   Begin VB.TextBox txtBminus 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      Text            =   "82-83.99% = B-"
      Top             =   3600
      Width           =   2295
   End
   Begin VB.TextBox txtB 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Height          =   405
      Left            =   3240
      TabIndex        =   8
      Text            =   "84-89.99% = B"
      Top             =   3000
      Width           =   2295
   End
   Begin VB.TextBox txtBplus 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Text            =   "90-91.99% = B+"
      Top             =   2400
      Width           =   2295
   End
   Begin VB.TextBox txtAminus 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Text            =   "92-93.99% = A-"
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox txtA 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Text            =   "94% and above = A"
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox txtF 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   5760
      TabIndex        =   4
      Text            =   "below 62% = F"
      Top             =   4200
      Width           =   2295
   End
   Begin VB.CommandButton cmdSeeGrade 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Enter the Info Below Then Press To See Your Grade"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      MaskColor       =   &H00FFFFC0&
      TabIndex        =   3
      Top             =   1560
      UseMaskColor    =   -1  'True
      Width           =   2895
   End
   Begin VB.TextBox txtID 
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   3480
      Width           =   2535
   End
   Begin VB.PictureBox picSeeGrade 
      BackColor       =   &H00FFFFC0&
      Height          =   735
      Left            =   120
      ScaleHeight     =   675
      ScaleWidth      =   4635
      TabIndex        =   0
      Top             =   240
      Width           =   4695
   End
   Begin VB.Label lblIdNumber 
      BackColor       =   &H00C0C000&
      Caption         =   "Enter Your ID Number"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2880
      Width           =   2655
   End
End
Attribute VB_Name = "frmStudentGrades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this part of the program is where students can view their individual grades
'they must have their ID number to view their grade

Private Sub cmdSeeGrade_Click()
    'declaring variables
    Dim I As Integer, ID(1 To 20) As Single
    Dim First(1 To 20) As String, Last(1 To 20) As String
    Dim Grade1(1 To 20) As Integer, Grade2(1 To 20) As Integer
    Dim TestGrade(1 To 20) As Integer, RealGrade As Single
    Dim X As Single, FinalGrade As Single
    
    Open App.Path & "\Grades.txt" For Input As #1   'opens the file with the students grades
    I = 0   'initialize counter I to zero, to be used for position in the array
    Do Until EOF(1)
        I = I + 1       'increment counter I each time throught the loop
                        'to move to the next postion in the array
                'Read next data set from the file into the array
        Input #1, ID(I), First(I), Last(I), Grade1(I), Grade2(I), TestGrade(I)
    Loop
    Close #1    'Close the file used for input
    
    X = txtID.Text  'gets the student ID number from the textbox
    For I = 1 To 20
        If X = ID(I) Then
        'if the ID number inputed in the text box is equal to the ID number in the list, then
            FinalGrade = (Grade1(I) + Grade2(I) + TestGrade(I)) / 120
            'the final grade is the total number earned divided by the total number of points possible
            picSeeGrade.Print First(I); "  "; Last(I), FormatPercent(FinalGrade)
            'prints the students first name, last name, and grade (in percentage form)
        End If
    Next I
    
    RealGrade = 100 * FinalGrade  'we need to percentage to be multiplied by 100 for the Select Case to work
    Select Case RealGrade   'the Select Case allows us to compare the single grade with several options
        Case Is >= 94
            picLetterGrade.Print "Your current grade is an A"
        Case Is >= 92
            picLetterGrade.Print "Your current grade is an A-"
        Case Is >= 90
            picLetterGrade.Print "Your current grade is a B+"
        Case Is >= 84
            picLetterGrade.Print "Your current grade is a B"
        Case Is >= 82
            picLetterGrade.Print "Your current grade is a B-"
        Case Is >= 80
            picLetterGrade.Print "Your current grade is a C+"
        Case Is >= 74
            picLetterGrade.Print "Your current grade is a C"
        Case Is >= 72
            picLetterGrade.Print "Your current grade is a C-"
        Case Is >= 70
            picLetterGrade.Print "Your current grade is a D+"
        Case Is >= 64
            picLetterGrade.Print "Your current grade is a D"
        Case Is >= 62
            picLetterGrade.Print "Your current grade is a D-"
        Case Is < 62
            picLetterGrade.Print "Your current grade is an F"
        Case Else
            picLetterGrade.Print "Error"
    End Select
    
End Sub
