VERSION 5.00
Begin VB.Form frmCalculators 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Calculator"
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7710
   LinkTopic       =   "Form2"
   ScaleHeight     =   4995
   ScaleWidth      =   7710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdInput 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Step 1:     Enter your Birthday!"
      BeginProperty Font 
         Name            =   "Gigi"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H008080FF&
      Caption         =   "Back to Main Menu"
      BeginProperty Font 
         Name            =   "Gigi"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton cmdLeapyear 
      BackColor       =   &H008080FF&
      Caption         =   "Were you born on a leapyear?"
      BeginProperty Font 
         Name            =   "Gigi"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   1815
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00C0E0FF&
      Height          =   2295
      Left            =   3120
      ScaleHeight     =   2235
      ScaleWidth      =   3675
      TabIndex        =   1
      Top             =   720
      Width           =   3735
   End
   Begin VB.CommandButton cmdAge 
      BackColor       =   &H008080FF&
      Caption         =   "How old are you?"
      BeginProperty Font 
         Name            =   "Gigi"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3480
      Width           =   1815
   End
End
Attribute VB_Name = "frmCalculators"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit 'spell check, error check, Cited: Lecture 11

'This button computes the age of the user
Private Sub cmdAge_Click()
    Select Case Currentdate
        Case Is < birthdate
            Age = (Currentyear - Birthyear) - 1
        Case Is > birthdate
            Age = Currentyear - Birthyear
        Case Is = birthdate
            Age = Currentyear - Birthyear
                picResults.Print "Today is your birthday!"; Tab(2); "Happy Birthday!"
        End Select
    'Select Case to compute the age of the user
    'Cited: Lab 8, Problem 2
    picResults.Print "You are "; Age; " years old."
    'displays the age of the user in a picture box
    'Cited: Lecture 11
End Sub


'This button asks the user to enter his/her birthdate, birthyear, the current date, and the current year in a series of input boxes
'Cited: Lecture 12
Private Sub cmdInput_Click()
    Currentyear = InputBox("Enter the current year", "Current Year")
    Birthyear = InputBox("Enter your birth year", "Birth Year")
    birthdate = InputBox("Enter your birthdate. (ex: 0101 for January 1)", "Birthdate")
    Currentdate = InputBox("Enter the current date. (ex: 0101 for January 1)", "Current Date")
End Sub


'This button determines whether the user was born on a leapyear. Cited: http://en.wikipedia.org/wiki/Leap_year, http://kalender-365.de/leap-years.php
Private Sub cmdLeapyear_Click()
    Dim Leapyear(1 To 100) As Single
    'Declares the array Leapyear as variable type single
    'Cited: Imad's VB Programs, "RunnersANDTimesProject"
    Dim year As Boolean
    'Declared variable
    'Cited: Lecture 11
    year = False
    'initialized variable
    'Cited: TA Chris Kerber
    
    Open App.Path & "\leapyear.txt" For Input As #2
    'opens data file
    CTR = 0
    'initializes CTR as zero
    Do Until EOF(2)
        CTR = CTR + 1
        Input #2, Leapyear(CTR)
    Loop
    'reads data from file into the array Leapyear
    Close #2
    'closes data file leapyear.txt
    'Above section cited: Imad's VB Programs, "Average Exam Scores"
    
    Pos = 0
    'initializes Pos as zero
    Do While year = False And Pos < CTR
        Pos = Pos + 1
        If Birthyear = Leapyear(Pos) Then
            year = True
        End If
    Loop
    If year = True Then
        picResults.Print "You were born on a Leap Year!"
    Else
        picResults.Print "You were not born on a Leap Year."
    End If
    'performs a Match & Stop search to find if the user's birthyear is located in the list of leap years found in the data file leapyear.txt
    'Cited TA Chris Kerber and Lecture 16
End Sub


'This button takes the user back to the form frmHome
'Cited: Lecture 18
Private Sub cmdQuit_Click()
    frmCalculators.Hide
    'makes the form frmCalculators invisible to the user
    frmHome.Show
    'makes the form frmHome visible to the user
End Sub
