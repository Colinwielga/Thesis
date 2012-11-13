VERSION 5.00
Begin VB.Form frmPuzzle 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   7050
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11385
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   11385
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Submit!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      TabIndex        =   25
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Check Answers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   24
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdCheat 
      Caption         =   "Can't figure it out?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   23
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdMultiple 
      Caption         =   "Answer One"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   22
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdFreebie 
      Caption         =   "Freebie"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   21
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdInput 
      Caption         =   "Input"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8040
      TabIndex        =   20
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Display the Puzzle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.PictureBox picPuzzle 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   600
      ScaleHeight     =   6195
      ScaleWidth      =   6315
      TabIndex        =   0
      Top             =   480
      Width           =   6375
      Begin VB.Line Line4 
         X1              =   0
         X2              =   6360
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   6360
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line Line2 
         X1              =   4080
         X2              =   4080
         Y1              =   0
         Y2              =   6480
      End
      Begin VB.Line Line1 
         X1              =   2040
         X2              =   2040
         Y1              =   0
         Y2              =   6480
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   17
      Left            =   6360
      TabIndex        =   19
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   16
      Left            =   5640
      TabIndex        =   18
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   15
      Left            =   4920
      TabIndex        =   17
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   14
      Left            =   4320
      TabIndex        =   16
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   13
      Left            =   3600
      TabIndex        =   15
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   12
      Left            =   2880
      TabIndex        =   14
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   11
      Left            =   2160
      TabIndex        =   13
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   10
      Left            =   1440
      TabIndex        =   12
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   9
      Left            =   720
      TabIndex        =   11
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   8
      Left            =   240
      TabIndex        =   10
      Top             =   6240
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   7
      Left            =   240
      TabIndex        =   9
      Top             =   5640
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   6
      Left            =   240
      TabIndex        =   8
      Top             =   4920
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   5
      Left            =   240
      TabIndex        =   7
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   4
      Left            =   240
      TabIndex        =   6
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   5
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   255
   End
End
Attribute VB_Name = "frmPuzzle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Guess As String, Found As Boolean, apple As Integer, Done As Integer
'This is the puzzle form were the user attempts to solve the puzzle
'based on the check boxes they selected... it may have multiple command buttons
'to aid the user, there are 3 different possible puzzles each sorted by difficulty


Private Sub cmdCheat_Click() 'reveals a solved version of the puzzle for a few seconds
    frmPuzzle.Hide              'before returning to the original puzzle
    frmCheat.Show
    For Pos = 1 To 81
        frmCheat.picCheat.Print Correct(Pos); "    "; Correct(Pos + 1); "    "; Correct(Pos + 2); "    "; Correct(Pos + 3); "    "; Correct(Pos + 4); "    "; Correct(Pos + 5); "    "; Correct(Pos + 6); "    "; Correct(Pos + 7); "    "; Correct(Pos + 8)
        frmCheat.picCheat.Print Tab(50)
        Pos = Pos + 8
    Next Pos
End Sub

Private Sub cmdCheck_Click()
    Dim tracker As Integer, Cycle As Boolean, Spot As Integer
    Dim Icing(1 To 81) As String
If Spot < 81 Then                           'checks the values inputed by the user and determines
    Cycle = False                       'if they are correct. A position is kept through the loop
    For Pos = 1 To ctr                  'so it may be remembered and used after the function reports its first/consecutive
        If Puz(Pos) > 0 Then            'values.
            If Puz(Pos) <> Correct(Pos) Then
                Icing(Pos) = Grid(Pos)
            End If
        End If
    Next Pos
    If Spot = 0 Then
        Spot = 1
    End If
    For Pos = Spot To ctr Or Cycle = True
        If Len(Icing(Pos)) = 2 Then
            MsgBox "Position " & Grid(Pos) & " is incorrect."
            Cycle = True
        End If
    Next Pos
End If
        
    
End Sub

Private Sub cmdFreebie_Click()
    Dim num As Integer
    Randomize
    num = 0
    
If Found = False Then
    Do Until num > 0 Or Done = apple        'this function allows the user a free answer filled in
                                            'uses boolean to only allow one answer
        num = Int(Rnd * 81) + 1
        
        If Puz(num) = Correct(num) Then
            num = 0
        Else
            Puz(num) = Correct(num)
            Done = Done + 1
            Found = True
        End If
    Loop
    picPuzzle.Cls
    
    For Pos = 1 To 81
        picPuzzle.Print Puz(Pos); "    "; Puz(Pos + 1); "    "; Puz(Pos + 2); "    "; Puz(Pos + 3); "    "; Puz(Pos + 4); "    "; Puz(Pos + 5); "    "; Puz(Pos + 6); "    "; Puz(Pos + 7); "    "; Puz(Pos + 8)
        picPuzzle.Print Tab(50)
        Pos = Pos + 8
    Next Pos
Else
    MsgBox User & " ,you already used this once!"
End If
    
    
End Sub

Private Sub cmdInput_Click()    'this function is responsible for having the user input values to change
   Found = False                'the puzzle... it won't allow an original value to be changed
   Guess = InputBox(User & " ,enter the position you wish to change (letterNumber).")
   For Pos = 1 To ctr
        If Grid(Pos) = Guess Then
            If Puz(Pos) = Correct(Pos) Then
                MsgBox User & ", you may not change this number!"
            Else
                apple = Pos
            End If
        End If
    Next Pos
    
    If Puz(Pos) <> Correct(Pos) Or Puz(Pos) = 0 Then
        Puz(apple) = InputBox(User & " ,enter a new value for position " & Guess & ".")
        picPuzzle.Cls
        For Pos = 1 To 81
            picPuzzle.Print Puz(Pos); "    "; Puz(Pos + 1); "    "; Puz(Pos + 2); "    "; Puz(Pos + 3); "    "; Puz(Pos + 4); "    "; Puz(Pos + 5); "    "; Puz(Pos + 6); "    "; Puz(Pos + 7); "    "; Puz(Pos + 8)
            picPuzzle.Print Tab(50)
            Pos = Pos + 8
        Next Pos
        
    End If
End Sub

Private Sub cmdMultiple_Click()
    Dim num As Integer
    Do Until num > 0 Or Done = apple        'same as freebie but the boolean is removed... usable multiple times
        
        num = Int(Rnd * 81) + 1
        
        If Puz(num) = Correct(num) Then
            num = 0
        Else
            Puz(num) = Correct(num)
            Done = Done + 1
            
        End If
    Loop
    picPuzzle.Cls
    
    For Pos = 1 To 81
        picPuzzle.Print Puz(Pos); "    "; Puz(Pos + 1); "    "; Puz(Pos + 2); "    "; Puz(Pos + 3); "    "; Puz(Pos + 4); "    "; Puz(Pos + 5); "    "; Puz(Pos + 6); "    "; Puz(Pos + 7); "    "; Puz(Pos + 8)
        picPuzzle.Print Tab(50)
        Pos = Pos + 8
    Next Pos
End Sub

Private Sub cmdTotal_Click()    'brings up the finals statistics screen
    frmFinal.Show
    frmPuzzle.Hide
End Sub

Private Sub Command1_Click()    'inputs the correct puzzle based off the users choice
If Level = "Easy" Then
    Open App.Path & "\EasyPuzzleIncomplete.txt" For Input As #1
    Do Until EOF(1)
        ctr = ctr + 1
        Input #1, Puz(ctr), Grid(ctr), Correct(ctr)
        If Puz(ctr) = 0 Then
            apple = apple + 1
        End If
    Loop
    Close #1
    
    For Pos = 1 To 81
        picPuzzle.Print Puz(Pos); "    "; Puz(Pos + 1); "    "; Puz(Pos + 2); "    "; Puz(Pos + 3); "    "; Puz(Pos + 4); "    "; Puz(Pos + 5); "    "; Puz(Pos + 6); "    "; Puz(Pos + 7); "    "; Puz(Pos + 8)
        picPuzzle.Print Tab(50)
        Pos = Pos + 8
    Next Pos
End If
If Level = "Intermediate" Then 'Creates a puzzle and table for the Intermediate sudoku puzzle
    Open App.Path & "\IntermediatePuzzle.txt" For Input As #1
        Do Until EOF(1)
            ctr = ctr + 1
            Input #1, Puz(ctr), Grid(ctr), Correct(ctr)
            If Puz(ctr) = 0 Then
            apple = apple + 1
            End If
        Loop
        Close #1
    
        For Pos = 1 To 81
            frmPuzzle.picPuzzle.Print Puz(Pos); "    "; Puz(Pos + 1); "    "; Puz(Pos + 2); "    "; Puz(Pos + 3); "    "; Puz(Pos + 4); "    "; Puz(Pos + 5); "    "; Puz(Pos + 6); "    "; Puz(Pos + 7); "    "; Puz(Pos + 8)
            frmPuzzle.picPuzzle.Print Tab(50)
            Pos = Pos + 8
        Next Pos
End If
If Level = "Difficult" Then 'Creates a puzzle and table for the difficult sudoku puzzle
    Open App.Path & "\DifficultPuzzle.txt" For Input As #1
        Do Until EOF(1)
            ctr = ctr + 1
            Input #1, Puz(ctr), Grid(ctr), Correct(ctr)
            If Puz(ctr) = 0 Then
            apple = apple + 1
            End If
        Loop
        Close #1
    
        For Pos = 1 To 81
            frmPuzzle.picPuzzle.Print Puz(Pos); "    "; Puz(Pos + 1); "    "; Puz(Pos + 2); "    "; Puz(Pos + 3); "    "; Puz(Pos + 4); "    "; Puz(Pos + 5); "    "; Puz(Pos + 6); "    "; Puz(Pos + 7); "    "; Puz(Pos + 8)
            frmPuzzle.picPuzzle.Print Tab(50)
            Pos = Pos + 8
        Next Pos
End If
    
End Sub

Private Sub Picture1_Click()

End Sub



Private Sub Multiple_Click()
    Dim num As Integer
    Randomize
    num = 0
    Do Until num > 0 Or Done = apple
        
        num = Int(Rnd * 81) + 1
        
        If Puz(num) = Correct(num) Then
            num = 0
        Else
            Puz(num) = Correct(num)
            Done = Done + 1
            
        End If
    Loop
    picPuzzle.Cls
    picTest.Print Puz(num)
    For Pos = 1 To 81
        picPuzzle.Print Puz(Pos); "    "; Puz(Pos + 1); "    "; Puz(Pos + 2); "    "; Puz(Pos + 3); "    "; Puz(Pos + 4); "    "; Puz(Pos + 5); "    "; Puz(Pos + 6); "    "; Puz(Pos + 7); "    "; Puz(Pos + 8)
        picPuzzle.Print Tab(50)
        Pos = Pos + 8
    Next Pos
End Sub

