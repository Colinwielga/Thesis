VERSION 5.00
Begin VB.Form FormGroupInputData1 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   14145
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18375
   FillColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   14145
   ScaleWidth      =   18375
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdForm2 
      Caption         =   "Go To Next Grading Form"
      Height          =   735
      Left            =   6360
      TabIndex        =   22
      Top             =   13200
      Width           =   3735
   End
   Begin VB.CommandButton cmdCompute 
      Caption         =   "Compute Form One"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6840
      TabIndex        =   21
      Top             =   7440
      Width           =   3495
   End
   Begin VB.PictureBox picForm1Total 
      Height          =   4095
      Left            =   120
      ScaleHeight     =   4035
      ScaleWidth      =   17955
      TabIndex        =   20
      Top             =   8640
      Width           =   18015
   End
   Begin VB.TextBox txtDescription 
      Height          =   975
      Left            =   7920
      TabIndex        =   19
      Top             =   5880
      Width           =   1335
   End
   Begin VB.TextBox txtFeatures 
      Height          =   1215
      Left            =   7800
      TabIndex        =   17
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmdStudentName1 
      Caption         =   " First Enter the Students Names"
      Height          =   615
      Left            =   480
      TabIndex        =   13
      Top             =   120
      Width           =   2295
   End
   Begin VB.PictureBox picStudentName2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3480
      ScaleHeight     =   1035
      ScaleWidth      =   2235
      TabIndex        =   12
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox txtModules 
      Height          =   975
      Left            =   10560
      TabIndex        =   11
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox txtIfThen 
      Height          =   975
      Left            =   12720
      TabIndex        =   9
      Top             =   5640
      Width           =   1815
   End
   Begin VB.TextBox txtDifficulty 
      Height          =   1335
      Left            =   12840
      TabIndex        =   6
      Top             =   4200
      Width           =   1575
   End
   Begin VB.TextBox txtCreativity 
      Height          =   1215
      Left            =   3480
      TabIndex        =   5
      Top             =   6000
      Width           =   1695
   End
   Begin VB.TextBox txtCommandButtons 
      Height          =   1095
      Left            =   3480
      TabIndex        =   3
      Top             =   4320
      Width           =   1695
   End
   Begin VB.PictureBox picStudentName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      ScaleHeight     =   1035
      ScaleWidth      =   2235
      TabIndex        =   0
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label lblDescription 
      Caption         =   "Enter the Clarity and Overall Quality of the Description of the Project (1-10)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5520
      TabIndex        =   18
      Top             =   6000
      Width           =   2055
   End
   Begin VB.Label lblFeatures 
      Caption         =   "Number of Interesting Features Not Learned In Class"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5640
      TabIndex        =   16
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label lblSecondStudent 
      Caption         =   "Second Student"
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
      Left            =   3480
      TabIndex        =   15
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label lblFirstStudent 
      Caption         =   "First Student"
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
      Left            =   480
      TabIndex        =   14
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label lblModules 
      Caption         =   "Enter the Number of Forms the students Used in the Project"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6960
      TabIndex        =   10
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Label lblIfThens 
      Caption         =   "Enter Number of If Then Statements Present Throughout the Project"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9600
      TabIndex        =   8
      Top             =   5640
      Width           =   2655
   End
   Begin VB.Label lblDifficulty 
      Caption         =   "Enter Overall Difficulty Of Project (1-10)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9600
      TabIndex        =   7
      Top             =   4200
      Width           =   2655
   End
   Begin VB.Label lblCreativity 
      Caption         =   "Enter overall Creativity of the Project (1-10)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      TabIndex        =   4
      Top             =   6120
      Width           =   2535
   End
   Begin VB.Label lblSecond 
      Caption         =   "Second:"
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
      Left            =   600
      TabIndex        =   2
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label lblCommandButtons 
      Caption         =   "Enter Number of Command Buttons used  Throughout the Project"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      TabIndex        =   1
      Top             =   4440
      Width           =   2415
   End
End
Attribute VB_Name = "FormGroupInputData1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCompute_Click()
Dim Forms As Integer, FormPoints As Integer, CommandButtons As Integer, CommandPoints As Integer
Dim Creativity As Integer, CreativityPoints As Integer, Features As Integer, FeaturesPoints As Integer
Dim Description As Integer, DescriptionPoints As Integer, Difficulty As Integer, DifficultyPoints As Integer
Dim IfThen As Integer, IfThenPoints As Integer
'this clears running total and the pic box in case the user made a mistake in entering the grades
picForm1Total.Cls
RunningTotal = 0
'sets the input from text boxes equal to variables
Forms = txtModules
CommandButtons = txtCommandButtons
Creativity = txtCreativity
Description = txtDescription
Features = txtFeatures
Difficulty = txtDifficulty
IfThen = txtIfThen

'All the if thens are giving the student points depending on the input from the user
'if the user inputs a value that cannot be correct, a error message will pop up and tell the user to re-input the data and re-compute
If IfThen < 0 Then
    MsgBox ("Invalid Number of If Then Statements Entry, Please Re-Enter the Data and Re-Compute")
ElseIf IfThen = 0 Then
    IfThenPoints = 0
ElseIf IfThen = 1 Then
    IfThenPoints = 3
ElseIf IfThen = 2 Then
    IfThenPoints = 5
ElseIf IfThen = 3 Then
    IfThenPoints = 7
ElseIf IfThen = 4 Then
    IfThenPoints = 9
ElseIf IfThen >= 5 Then
    IfThenPoints = 10
End If



'this case statement assings points depending on the description of the project
Select Case Description
    Case Is < 0
        MsgBox ("Invalid Description of Project Entry, Please Re-Enter the Data and Re-Compute")
    Case 1 To 3
        DescriptionPoints = 2
    Case 4 To 6
        DescriptionPoints = 5
    Case 7 To 8
       DescriptionPoints = 8
    Case 9
        DescriptionPoints = 9
    Case 10
        DescriptionPoints = 10
    Case Is > 10
        MsgBox ("Invalid Description of Project Entry, Please Re-Enter the Data and Re-Compute")
End Select

If Difficulty < 0 Then
    MsgBox "Invalid Difficulty Entry, Re-Enter Entry and Re-compute"
ElseIf Difficulty = 0 Then
    DifficultyPoints = 0
ElseIf Difficulty = 1 Then
    DifficultyPoints = 1
ElseIf Difficulty = 2 Then
    DifficultyPoints = 2
ElseIf Difficulty = 3 Then
    DifficultyPoints = 3
ElseIf Difficulty = 4 Then
    DifficultyPoints = 4
ElseIf Difficulty = 5 Then
    DifficultyPoints = 5
ElseIf Difficulty = 6 Then
    DifficultyPoints = 6
ElseIf Difficulty = 7 Then
    DifficultyPoints = 7
ElseIf Difficulty = 8 Then
    DifficultyPoints = 8
ElseIf Difficulty = 9 Then
    DifficultyPoints = 9
ElseIf Difficulty = 10 Then
    DifficultyPoints = 10
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

If Forms < 0 Then
    MsgBox "Invalid Entry of Number of Forms, Please Re-Enter The information and Re-Compute"
ElseIf Forms = 0 Then
    FormPoints = 0
ElseIf Forms = 1 Then
    FormPoints = 1
ElseIf Forms = 2 Then
    FormPoints = 5
ElseIf Forms = 3 Then
    FormPoints = 7
ElseIf Forms = 4 Then
    FormPoints = 9
ElseIf Forms >= 5 Then
    FormPoints = 10
End If

If CommandButtons < 0 Then
    MsgBox "Invalid Command Buttons Entry, Please Re-Enter Entry and Re-compute"
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
'this adds up all the points and assings it to the Running Total
'we did it this way so that you can show where the student recieved points in the pic box
RunningTotal = FormPoints + DescriptionPoints + IfThenPoints + CommandPoints + CreativityPoints + DifficultyPoints + FeaturesPoints + RunningTotal
'this prints off all the different variables into a picture box
picForm1Total.Cls
picForm1Total.Print "                                                                                                                                             Total Points From Form 1"
picForm1Total.Print "                                                                                                                                           Number of Points Recieved From"
picForm1Total.Print "Number of Command Buttons     Overall Creativity       Overall Difficulty       Number of Interesting Features        Points from Number of Forms       Points From Description of Project        Points from Number of If/Then Statements"
picForm1Total.Print "       "; CommandPoints; "out of 10                         "; CreativityPoints; " out of 10                 "; DifficultyPoints; " out of 10                           "; FeaturesPoints; "out of 10                                    "; FormPoints; "out of 10                                   "; DescriptionPoints; " out of 10                                     "; IfThenPoints; " out of 10"
picForm1Total.Print "**************************************************************************************************************************************************************************************************************************************"
picForm1Total.Print StudentName; " and "; StudentName2; " recieved a total of "; RunningTotal; " out of 70 possible points on the first section"
cmdForm2.Visible = True
End Sub

Private Sub cmdForm2_Click()
FormGroupInputData1.Hide
FormGroupInputData2.Show
End Sub

Private Sub cmdStudentName1_Click()
'This is where you enter the students name
StudentName = InputBox("Enter The First Student's Name")
StudentName2 = InputBox("Enter The Second Student's Name")
picStudentName.Print StudentName
picStudentName2.Print StudentName2
End Sub

Private Sub Form_Load()
'does not allow you to go to the next form without calculating the data on this form
cmdForm2.Visible = False
End Sub
