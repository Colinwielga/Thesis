VERSION 5.00
Begin VB.Form Other 
   BackColor       =   &H0080C0FF&
   Caption         =   "Form5"
   ClientHeight    =   8925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7080
   LinkTopic       =   "Form5"
   ScaleHeight     =   8925
   ScaleWidth      =   7080
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   3375
      Left            =   4800
      Picture         =   "Other.frx":0000
      ScaleHeight     =   3315
      ScaleWidth      =   2115
      TabIndex        =   8
      Top             =   3240
      Width           =   2175
   End
   Begin VB.CommandButton cmdCourses 
      Caption         =   "Courses Available"
      Height          =   1215
      Left            =   5400
      TabIndex        =   6
      Top             =   240
      Width           =   1575
   End
   Begin VB.PictureBox picResults 
      Height          =   7095
      Left            =   360
      ScaleHeight     =   7035
      ScaleWidth      =   4155
      TabIndex        =   5
      Top             =   1680
      Width           =   4215
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "Calculate Handicap"
      Height          =   1335
      Left            =   5040
      TabIndex        =   4
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox txtCourse 
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   3015
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   855
      Left            =   6000
      TabIndex        =   1
      Top             =   7920
      Width           =   975
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Starting Page"
      Height          =   975
      Left            =   4800
      TabIndex        =   0
      Top             =   6840
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080C0FF&
      Caption         =   "If you would like a list of the courses available, please click on the button to the right."
      Height          =   975
      Left            =   3720
      TabIndex        =   7
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Please enter the name of the course you played with the tees you played from.     (Ex. Pebble Creek Middle Tees)"
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "Other"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Names(1 To 100) As String
Dim Slope(1 To 100) As Single
Dim Rating(1 To 100) As Single
Dim PATH As String
Dim J As Integer
Dim K As Integer
Dim Course As String
Dim Score As Single
Dim Found As Boolean
Dim CTR As Integer
Dim Position As Integer
Dim HandicapDifferential As Single
Dim Handicap As Single
Sub FormLoad()
    PATH = "M:\CS130\Miscellaneous"
End Sub

'This section is inputing the course information and then searching for the course that the user inputs
'When a match is found, you can then input your score, and the hanicap will be calculated

Private Sub cmdCalc_Click()

CTR = 0
Open PATH & "Courses.txt" For Input As #1
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Names(CTR), Slope(CTR), Rating(CTR)
Loop
Found = False
Course = txtCourse.Text
Do While Not Found
    
    For J = 1 To CTR
        Position = Position + 1
        If Names(J) = Course Then
            Found = True
        End If
    Next J
Loop
If Found Then
    Score = InputBox("Please enter your score")
    
        If Score > 0 Then
            HandicapDifferential = ((Score - Rating(Position)) * 113 / Slope(Position))
            Handicap = FormatNumber(HandicapDifferential, 1) * 0.96
            picResults.Cls
            picResults.Print "Your handicap is "; FormatNumber(Handicap, 1)
        Else
            MsgBox "Sorry but you must enter a positive number", , "Error"
        End If
End If
If Not Found Then
    picResults.Cls
    picResults.Print "Sorry but the course you're looking for was not found.  If you'd like to see a list of courses that are available, please click on the Courses Available Button."
End If
Close #1
    
End Sub

'This part allows the user to see which courses are in the array
'The courses are printes, so they know what they can choose from

Private Sub cmdCourses_Click()
Open PATH & "Courses.txt" For Input As #1
CTR = 0
Do While Not EOF(1)
    CTR = CTR + 1
    Input #1, Names(CTR), Slope(CTR), Rating(CTR)
Loop
For K = 1 To CTR Step 1
    picResults.Cls
    picResults.Print Names(K)
Next K
Close #1
End Sub

Private Sub cmdQuit_Click()
End
End Sub

'This allows the user to return to the previous forms/screens

Private Sub cmdReturn_Click()
Form1.Show
Other.Hide
End Sub
