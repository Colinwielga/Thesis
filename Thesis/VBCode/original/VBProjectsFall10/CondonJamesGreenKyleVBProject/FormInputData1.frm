VERSION 5.00
Begin VB.Form FormInputData1 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   13815
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18960
   FillColor       =   &H00C00000&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   ScaleHeight     =   13815
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPic 
      Height          =   3495
      Left            =   10440
      Picture         =   "FormInputData1.frx":0000
      ScaleHeight     =   3435
      ScaleWidth      =   5715
      TabIndex        =   18
      Top             =   2520
      Width           =   5775
   End
   Begin VB.CommandButton cmdForm2 
      Caption         =   "Go to Next Grading Form"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   8520
      TabIndex        =   17
      Top             =   10200
      Width           =   3375
   End
   Begin VB.CommandButton cmdStudentName 
      Caption         =   "First Enter Name Of Student"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      TabIndex        =   15
      Top             =   360
      Width           =   3375
   End
   Begin VB.PictureBox picForm1Total 
      Height          =   2895
      Left            =   6480
      ScaleHeight     =   2835
      ScaleWidth      =   12435
      TabIndex        =   14
      Top             =   6120
      Width           =   12495
   End
   Begin VB.CommandButton cmdCompute1 
      Caption         =   "Lastly Compute Form 1 Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7680
      TabIndex        =   13
      Top             =   3840
      Width           =   2655
   End
   Begin VB.PictureBox picNameofStudent 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      ScaleHeight     =   795
      ScaleWidth      =   5115
      TabIndex        =   12
      Top             =   1560
      Width           =   5175
   End
   Begin VB.CommandButton cmdPoor 
      Caption         =   "Poor Quality"
      Height          =   855
      Left            =   10320
      TabIndex        =   11
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdMediocre 
      Caption         =   "Mediocre Quality"
      Height          =   855
      Left            =   8640
      TabIndex        =   10
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdExceptional 
      Caption         =   "Exceptional Quality"
      Height          =   855
      Left            =   6960
      TabIndex        =   9
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox txtFeatures 
      Height          =   1335
      Left            =   4440
      TabIndex        =   7
      Top             =   8640
      Width           =   1815
   End
   Begin VB.TextBox txtDifficulty 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4440
      TabIndex        =   4
      Top             =   6840
      Width           =   1815
   End
   Begin VB.TextBox txtCreativity 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4440
      TabIndex        =   2
      Top             =   4920
      Width           =   1935
   End
   Begin VB.TextBox txtCommandButtons 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4440
      TabIndex        =   0
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label lblthird 
      Caption         =   "Third:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   16
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label lblUserInterface 
      Caption         =   "Second Click on the Appropriate User Interface Quality"
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
      Left            =   7440
      TabIndex        =   8
      Top             =   360
      Width           =   4575
   End
   Begin VB.Label lblFeatures 
      Caption         =   "Enter Total Number of Interesting Features not Learned in Class"
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
      TabIndex        =   6
      Top             =   8760
      Width           =   3735
   End
   Begin VB.Label lbldifficulty 
      Caption         =   "Enter Overall Difficulty Here (1-10)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   5
      Top             =   6960
      Width           =   4095
   End
   Begin VB.Label lblCreativity 
      Caption         =   "Enter Overall Creativity Here (1-10)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   3
      Top             =   5280
      Width           =   3735
   End
   Begin VB.Label lblCommandButtons 
      Caption         =   "Enter Number of Command Buttons on Project (Excluding Quit Buttons)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   600
      TabIndex        =   1
      Top             =   3600
      Width           =   3615
   End
End
Attribute VB_Name = "FormInputData1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim InterfacePoints As Integer

Private Sub cmdCompute1_Click()
Dim CommandButtons As Integer, Creativity As Integer, Difficulty As Integer, Features As Integer
Dim CommandPoints As Integer, CreativityPoints As Integer, DifficultyPoints As Integer
Dim FeaturesPoints As Integer
'this clears running total and the pic box in case the user made a mistake in entering the grades
CommandButtons = 0
Creativity = 0
Difficulty = 0
Features = 0
RunningTotal = 0
'sets the input from text boxes equal to variables
CommandButtons = txtCommandButtons
Creativity = txtCreativity
Difficulty = txtDifficulty
Features = txtFeatures

'All the if thens are giving the student points depending on the input from the user
'if the user inputs a value that cannot be correct, a error message will pop up and tell the user to re-input the data and re-compute


If CommandButtons < 0 Then
    MsgBox "Invalid Command Buttons Entry, Re-Enter Entry and Re-compute"
ElseIf CommandButtons = 0 Then
    CommandPoints = 0
ElseIf CommandButtons = 1 Then
    CommandPoints = 1
ElseIf CommandButtons = 2 Then
    CommandPoints = 2
ElseIf CommandButtons = 3 Then
    CommandPoints = 3
ElseIf CommandButtons = 4 Then
    CommandPoints = 4
ElseIf CommandButtons = 5 Then
    CommandPoints = 5
ElseIf CommandButtons = 6 Then
    CommandPoints = 6
ElseIf CommandButtons = 7 Then
    CommandPoints = 7
ElseIf CommandButtons = 8 Then
    CommandPoints = 9
ElseIf CommandButtons > 8 Then
    CommandPoints = 10
End If


Select Case Creativity
    Case Is < 0
        MsgBox "Invalid Creativity Entry, Re-Enter Entry and Re-compute"
    Case 0
        CreativityPoints = CreativityPoints + 0
    Case 1
        CreativityPoints = CreativityPoints + 1
    Case 2
        CreativityPoints = CreativityPoints + 2
    Case 3
        CreativityPoints = CreativityPoints + 3
    Case 4
        CreativityPoints = CreativityPoints + 4
    Case 5
        CreativityPoints = CreativityPoints + 5
    Case 6
        CreativityPoints = CreativityPoints + 6
    Case 7
        CreativityPoints = CreativityPoints + 7
    Case 8
        CreativityPoints = CreativityPoints + 8
    Case 9
        CreativityPoints = CreativityPoints + 9
    Case 10
        CreativityPoints = CreativityPoints + 10
    Case Is > 10
        MsgBox "Invalid Creativity Entry, Re-Enter Entry and Re-compute"
End Select

If Difficulty < 0 Then
    MsgBox "Invalid Difficulty Entry, Re-Enter Entry and Re-compute"
ElseIf Difficulty = 0 Then
    DifficultyPoints = DifficultyPoints + 0
ElseIf Difficulty = 1 Then
    DifficultyPoints = DifficultyPoints + 1
ElseIf Difficulty = 2 Then
    DifficultyPoints = DifficultyPoints + 2
ElseIf Difficulty = 3 Then
    DifficultyPoints = DifficultyPoints + 3
ElseIf Difficulty = 4 Then
    DifficultyPoints = DifficultyPoints + 4
ElseIf Difficulty = 5 Then
    DifficultyPoints = DifficultyPoints + 5
ElseIf Difficulty = 6 Then
    DifficultyPoints = DifficultyPoints + 6
ElseIf Difficulty = 7 Then
    DifficultyPoints = DifficultyPoints + 7
ElseIf Difficulty = 8 Then
    DifficultyPoints = DifficultyPoints + 8
ElseIf Difficulty = 9 Then
    DifficultyPoints = DifficultyPoints + 9
ElseIf Difficulty = 10 Then
    DifficultyPoints = DifficultyPoints + 10
ElseIf Difficulty > 10 Then
    MsgBox "Invalid Difficulty Entry, Re-Enter Entry and Re-compute"
End If

Select Case Features
    Case Is < 0
        MsgBox "Invalid Interesting Features Entry, Re-Enter Entry and Re-compute"
    Case 0
        FeaturesPoints = FeaturesPoints + 0
    Case 1
        FeaturesPoints = FeaturesPoints + 3
    Case 2
        FeaturesPoints = FeaturesPoints + 7
    Case Is > 2
        FeaturesPoints = FeaturesPoints + 10
End Select
'this adds up all the points and assings it to the Running Total
'we did it this way so that you can show where the student recieved points in the pic box
RunningTotal = InterfacePoints + CommandPoints + CreativityPoints + DifficultyPoints + FeaturesPoints + RunningTotal
'this prints off all the different variables into a picture box
picForm1Total.Cls
picForm1Total.Print "                                                                                                        Total Points From Form 1"
picForm1Total.Print "                                                                                                      Number of Points Recieved From"
picForm1Total.Print "Number of Command Buttons       Overall Creativity         Overall Difficulty       Number of Interesting Features        InterFace Quality"
picForm1Total.Print "       "; CommandPoints; "out of 10                                          "; CreativityPoints; " out of 10                     "; DifficultyPoints; " out of 10                           "; FeaturesPoints; "out of 10                                    "; InterfacePoints; "out of 10"
picForm1Total.Print "*****************************************************************************************************************************************************************************************************************"
picForm1Total.Print StudentName; " recieved a total of "; RunningTotal; " out of 50 possible points on the first section"

cmdForm2.Visible = True
End Sub

Private Sub cmdExceptional_Click()
'all the buttons like this when clicked give the student points for having this in their project
'it then pops up a msgbox and tells the user how many points the student will recieve
InterfacePoints = 10
MsgBox "The Student Will Recieve 10 out of 10 Points for this Catagory"
cmdMediocre.Visible = False
cmdPoor.Visible = False
End Sub

Private Sub cmdForm2_Click()
FormInputData1.Hide
FormInputData2.Show
End Sub

Private Sub cmdMediocre_Click()
InterfacePoints = 5
MsgBox "The Student Will Recieve 5 out of 10 Points for this Catagory"
cmdExceptional.Visible = False
cmdPoor.Visible = False
End Sub

Private Sub cmdPoor_Click()
InterfacePoints = 1
MsgBox "The Student Will Recieve 1 out of 10 Points for this Catagory"
cmdExceptional.Visible = False
cmdMediocre.Visible = False
End Sub

Private Sub cmdStudentName_Click()
StudentName = InputBox("Enter the Student's Name")
picNameofStudent.Print StudentName
End Sub

Private Sub Form_Load()
cmdForm2.Visible = False
End Sub
