VERSION 5.00
Begin VB.Form frmTrivia
   Caption         =   "Form1"
   ClientHeight    =   8970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13740
   LinkTopic       =   "Form1"
   Picture         =   "frmTrivia.frx":0000
   ScaleHeight     =   8970
   ScaleWidth      =   13740
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEasyAsnwers
      Caption         =   "Easy Answers"
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   2520
      Width           =   1935
   End
   Begin VB.PictureBox picResults
      BackColor       =   &H00FFFFFF&
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3015
      Left            =   3240
      ScaleHeight     =   2955
      ScaleWidth      =   9915
      TabIndex        =   8
      Top             =   2640
      Width           =   9975
   End
   Begin VB.CommandButton cmdHardAnswers
      Caption         =   "Hard Answers"
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   6240
      Width           =   1935
   End
   Begin VB.CommandButton cmdInterAnswers
      Caption         =   "Intermediate Answers"
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   4320
      Width           =   1935
   End
   Begin VB.CommandButton cmdClear
      Caption         =   "Clear"
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   6840
      Width           =   1695
   End
   Begin VB.CommandButton cmdHard
      Caption         =   "Hard Questions"
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   5160
      Width           =   2415
   End
   Begin VB.CommandButton cmdQuit
      Caption         =   "Quit"
      Height          =   615
      Left            =   9000
      TabIndex        =   3
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdIntermediate
      Caption         =   "Intermediate Questions"
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   2415
   End
   Begin VB.CommandButton cmdEasy
      Caption         =   "Easy Question"
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   2415
   End
   Begin VB.CommandButton cmdReturn
      Caption         =   "Return to Main Menu"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmTrivia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEasy_Click()
    Dim InputAnswer As String 'sets data type for InputAnswer as String
    Dim Score As Integer 'sets data type for Score as Integer
    Dim Question(1 To 5) As String 'sets data type for Question array as String
    Dim Answer(1 To 5) As String 'sets data type for Answer array as String
    Dim CTR As Integer, Pos As Integer 'sets data type for CTR and Pos as Integer

        CTR = 0 'sets value of CTR to 0
    Open App.Path & "\easy.txt" For Input As #1
    Do Until EOF(1) 'reads Easy.txt until End of File
        CTR = CTR + 1 'adds 1 to value of CTR to keep track of number of data lines in file
        Input #1, Question(CTR), Answer(CTR) 'stores data from line of text in Easy.txt as Question(CTR) and Answer(CTR)
    Loop 'returns to the file to repeat the steps above
    Score = 0 'sets value of Score to 0
    Close #1 'closes file Easy.txt
    For Pos = 1 To CTR 'sets number of times to perform the following process
        InputAnswer = InputBox(Question(Pos), "Question") 'asks user a question and sets InputAnswer as input provided by user via textbox
        If LCase(InputAnswer) = LCase(Answer(Pos)) Then 'allows for lowercase lettering and creates the condition that the InputAnswer is equal to Answer(pos) from file
            Score = Score + 1 'if InputAnswer=Answer(pos) then 1 is added to Score
        Else 'if InputAnswer does not match Answer(pos) then
            MsgBox "Sorry.  Your Wrong."
        End If 'ends nested if
    Next Pos 'returns to pos and asks the next question
    If Score > 3 Then 'if user gets more than three answers correct...
        MsgBox "You got " & Score & " correct!"
    Else 'if user gets three of less answer correct...
        MsgBox "You got " & Score & " correct!"
    End If 'ends If conditional
End Sub

Private Sub cmdEasyAsnwers_Click()
    picResults.Cls
    picResults.Print "There are 6 seasons completed"
    picResults.Print "Ari Gold is Vince's agent"
    picResults.Print "The group is from Queens, New York"
    picResults.Print "Johnny Drama is the oldest of the group"
    picResults.Print "Aquaman had the number one opening of all time"
End Sub

Private Sub cmdHard_Click()
    Dim Question(1 To 5) As String
    Dim InputAnswer As String 'sets data type for InputAnswer as String
    Dim Answer(1 To 5) As String 'sets data type for Answer array as String
    Dim CTR As Integer, Pos As Integer 'sets data type for CTR and Pos as Integer
    Dim Score As Integer 'sets data type for Score as Integer

    Open App.Path & "\hard.txt" For Input As #1
        CTR = 0 'sets value of CTR to 0
    Do Until EOF(1) 'reads Hard.txt until End of File
        CTR = CTR + 1 'adds 1 to value of CTR to keep track of number of data lines in file
        Input #1, Question(CTR), Answer(CTR) 'stores data from line of text in Hard.txt as Question(CTR) and Answer(CTR)
    Loop 'returns to the file to repeat the steps above
    Close #1 'closes file Hard.txt
    Score = 0 'sets value of Score to 0
    For Pos = 1 To CTR 'sets number of times to perform the following process
        InputAnswer = InputBox(Question(Pos), "Question") 'asks user a question and sets InputAnswer as input provided by user via textbox
        If LCase(InputAnswer) = LCase(Answer(Pos)) Then 'allows for lowercase lettering and creates the condition that the InputAnswer is equal to Answer(pos) from file
            Score = Score + 1 'if InputAnswer=Answer(pos) then 1 is added to Score
        Else 'if InputAnswer does not match Answer(pos) then
            MsgBox "Sorry.  Your Wrong."
        End If 'ends nested if
    Next Pos 'returns to pos and asks the next question
    If Score > 3 Then 'if user gets more than three answers correct...
        MsgBox "You got " & Score & " correct!"
    Else 'if user gets three of less answer correct...
        MsgBox "You got " & Score & " correct!"
    End If 'ends If conditional

End Sub

Private Sub cmdInterAnswers_Click()
    picResults.Cls
    picResults.Print "Turtle's dog's name is Arnold"
    picResults.Print "Ari turns down Warner Bros studio head position"
    picResults.Print "Barbara Miller becomes Ari's partner at Miller Gold Agency"
    picResults.Print "Vince does Jimmy Kimmel in the first season of Entourage"
    picResults.Print "Five Towns is the name of Drama's hit TV series"
End Sub

Private Sub cmdHardAnswers_Click()
    picResults.Cls
    picResults.Print "Turtle's real name is Sal"
    picResults.Print "Vince proposed to Mandy Moore when he was younger"
    picResults.Print "Rex Lee plays Lloyd"
    picResults.Print "Medillin is purchased for $1 after it tanks at Cannes Film Festival"
    picResults.Print "Vince gave her a Robert Niche painting"
End Sub

Private Sub cmdIntermediate_Click()
    Dim Question(1 To 5) As String 'sets data type for Question array as String
    Dim Answer(1 To 5) As String 'sets data type for Answer array as String
    Dim CTR As Integer, Pos As Integer 'sets data type for CTR and Pos as Integer
    Dim InputAnswer As String 'sets data type for InputAnswer as String
    Dim Score As Integer 'sets data type for Score as Integer

    Open App.Path & "\intermediate.txt" For Input As #1
        CTR = 0 'sets value of CTR to 0
    Do Until EOF(1) 'reads Middling.txt until End of File
        CTR = CTR + 1 'adds 1 to value of CTR to keep track of number of data lines in file
        Input #1, Question(CTR), Answer(CTR) 'stores data from line of text in Middling.txt as Question(CTR) and Answer(CTR)
    Loop 'returns to the file to repeat the steps above
    Close #1 'closes file Middling.txt
    Score = 0 'sets value of Score to 0
    For Pos = 1 To CTR 'sets number of times to perform the following process
        InputAnswer = InputBox(Question(Pos), "Question") 'asks user a question and sets InputAnswer as input provided by user via textbox
        If LCase(InputAnswer) = LCase(Answer(Pos)) Then 'allows for lowercase lettering and creates the condition that the InputAnswer is equal to Answer(pos) from file
            Score = Score + 1 'if InputAnswer=Answer(pos) then 1 is added to Score
        Else 'if InputAnswer does not match Answer(pos) then
            MsgBox "Sorry.  Your Wrong."
        End If 'ends nested if
    Next Pos 'returns to pos and asks the next question
    If Score > 3 Then 'if user gets more than three answers correct...
        MsgBox "You got " & Score & " correct!"
    Else 'if user gets three of less answer correct...
        MsgBox "You got " & Score & " correct!"
    End If 'ends If conditional
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdReturn_Click()
    frmTrivia.Hide
    frmMain.Show
End Sub

