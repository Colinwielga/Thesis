VERSION 5.00
Begin VB.Form frmTrivia 
   BackColor       =   &H00C000C0&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMenu 
      BackColor       =   &H00FF80FF&
      Caption         =   "Back to Menu"
      Height          =   735
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      Width           =   1695
   End
   Begin VB.PictureBox picResults 
      Height          =   735
      Left            =   960
      ScaleHeight     =   675
      ScaleWidth      =   2475
      TabIndex        =   1
      Top             =   1080
      Width           =   2535
   End
   Begin VB.CommandButton cmdPlay 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Play"
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmTrivia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdMenu_Click()
frmTrivia.Hide
frmMenu.Show
End Sub

Private Sub cmdPlay_Click()
Dim answer As Integer
Dim answer2 As Integer
Dim answer3 As Integer
Dim answer4 As String
Dim answer5 As String
Dim answer6 As String
Dim answer7 As String
Dim answer8 As String
Dim answer9 As String
answer = InputBox("How many years did Kristie play basketball?", "Question 1")
    Select Case answer
        Case 1 To 3
            picResults.Print "Too low!"
        Case Is = 4
            picResults.Print "Correct!"
        Case Else
            picResults.Print "Too high!"
        End Select
answer2 = InputBox("How many years was Kristie in Chamber Choir?", "Question 2")
picResults.Cls
    Select Case answer2
        Case Is = 1
            picResults.Print "Correct!"
        Case Else
            picResults.Print "Sorry, that is incorrect!"
        End Select
answer3 = InputBox("How many sports was Kristie involved in?", "Question 3")
picResults.Cls
    Select Case answer3
        Case 1 To 3
            picResults.Print "Too low!"
        Case Is = 4
            picResults.Print "Correct!"
        Case Else
            picResults.Print "Too high!"
    End Select
answer4 = InputBox("What is Kristie's favorite TV Show?", "Question 4")
picResults.Cls
    Select Case answer4
        Case Is = "Greys Anatomy"
            picResults.Print "Correct!"
        Case Else
            picResults.Print "Sorry, that is incorrect!"
        End Select
answer5 = InputBox("Who is Kristie's favorite band?", "Question 5")
picResults.Cls
    Select Case answer5
        Case Is = "Staind"
            picResults.Print "Correct!"
        Case Else
            picResults.Print "Sorry, that is incorrect!"
    End Select
answer6 = InputBox("Which choir was Kristie involved in for the most amount of years?", "Question 6")
    picResults.Cls
        Select Case answer6
            Case Is = "Concert Choir"
                picResults.Print "Correct!"
            Case Else
                picResults.Print "Sorry, that is incorrect!"
        End Select
answer7 = InputBox("What is Kristie's favorite movie?", "Question 7")
    picResults.Cls
        Select Case answer7
            Case Is = "Step Up"
                picResults.Print "Correct!"
            Case Else
                picResults.Print "Sorry, that is incorrect!"
        End Select
answer8 = InputBox("What sports did Kristie play in 10th grade?", "Question 8")
    picResults.Cls
        Select Case answer8
            Case Is = "Basketball"
                picResults.Print "Correct!"
            Case Else
                picResults.Print "Sorry, that is incorrect"
        End Select
answer9 = InputBox("How many years was Kristie a SAAD officer?", "Question9")
    picResults.Cls
        Select Case answer9
            Case Is = 1
                picResults.Print "Correct!"
            Case Else
                picResults.Print "Sorry, that is incorrect!"
        End Select
            

        
                
            
        
        
        
        

End Sub



