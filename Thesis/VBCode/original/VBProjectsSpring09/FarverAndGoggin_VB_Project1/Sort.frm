VERSION 5.00
Begin VB.Form frmSort 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11505
   LinkTopic       =   "Form1"
   ScaleHeight     =   8205
   ScaleWidth      =   11505
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00C0C000&
      Caption         =   "Return To Game"
      Height          =   855
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6840
      Width           =   2175
   End
   Begin VB.CommandButton cmdAverage 
      BackColor       =   &H008080FF&
      Caption         =   "Find The Average Score"
      Height          =   855
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5760
      Width           =   2175
   End
   Begin VB.CommandButton cmdSearch1 
      BackColor       =   &H00C000C0&
      Caption         =   "Search For A Score"
      Height          =   855
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton cmdSortNames 
      BackColor       =   &H00FF8080&
      Caption         =   "Sort By Name"
      Height          =   855
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4680
      Width           =   2175
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H0000FFFF&
      Caption         =   "Clear"
      Height          =   855
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6960
      Width           =   2055
   End
   Begin VB.CommandButton cmdSortScores 
      BackColor       =   &H0080C0FF&
      Caption         =   "Sort By Score"
      Height          =   855
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3600
      Width           =   2175
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0000FFFF&
      Caption         =   "Quit"
      Height          =   855
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6960
      Width           =   2055
   End
   Begin VB.CommandButton cmdSearchScore 
      BackColor       =   &H0000FF00&
      Caption         =   "Search For Top Scores"
      Height          =   855
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton cmdHallofFame 
      BackColor       =   &H00FF0000&
      Caption         =   "Hall of Fame"
      Height          =   855
      Left            =   600
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin VB.PictureBox picresults 
      BackColor       =   &H0080FF80&
      Height          =   5655
      Left            =   5160
      ScaleHeight     =   5595
      ScaleWidth      =   4515
      TabIndex        =   0
      Top             =   960
      Width           =   4575
   End
End
Attribute VB_Name = "frmSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim player(1 To 20) As String
Dim score(1 To 20) As Integer

Private Sub cmdAverage_Click()
    ' This subroutine will average the scores in the top ten, then print the outcome
    
    Dim Sum As Integer, Average As Single, Pos As Integer
    Open App.Path & "\HallofFame.txt" For Input As #1
    Ctr = 0
    Pos = 0
    Sum = 0
    Do Until EOF(1)
        Ctr = Ctr + 1
        Pos = Pos + 1
        Input #1, player(Ctr), score(Ctr)
        Sum = Sum + score(Pos)
    Loop
    
    'Print the Average
    Average = Sum / Ctr
    MsgBox ("The Average Score is " & Average & ".")
    
        
    
End Sub

Private Sub cmdClear_Click()
    picresults.Cls
End Sub

Private Sub cmdHallofFame_Click()
    
    'Code to Read High scores from file into Program
    
    Dim Pass As Integer
    Dim Pos As Integer
    Dim Tempname As String
    Dim Tempscore As Integer
   
   'Clear PicResults
   picresults.Cls
   
   picresults.Print "Top Ten All Time Scores"
   picresults.Print "*****************************"
   
    'Open File to obtain High Scores
    Open App.Path & "\HallofFame.txt" For Input As #1
    
    Ctr = 0
    Do Until EOF(1)
        Ctr = Ctr + 1
        Input #1, player(Ctr), score(Ctr)
        picresults.Print player(Ctr); Tab(20); score(Ctr)
    Loop
    Close #1

End Sub
Private Sub cmdQuit_Click()
    'Quit button
    End
End Sub
Private Sub cmdReturn_Click()
    'Move From Form to Form
    frmMatchingGame.Visible = True
    frmSort.Visible = False
End Sub

Private Sub cmdSearch1_Click()
    'Declare Variables
    Dim Found As Boolean, Pos As Integer, SearchValue As Integer
    Found = False
    Pos = 0
    SearchValue = InputBox("Please Enter A Score", "Enter A Score")
    
    'Match and Stop Search
    Do While Found = False And Pos < 20
        Pos = Pos + 1
        If score(Pos) = SearchValue Then
            Found = True
        End If
    Loop
    
    'Print the Score
    If Found = True Then
         MsgBox ("The player named " & player(Pos) & " has the score of " & score(Pos) & " in the record book")
    Else
         MsgBox ("The input is not currently in the record books")
    End If
End Sub

Private Sub cmdSearchScore_Click()
    Dim Found As Boolean
    Dim inputscore As Integer
    Dim betterscore As Integer
    Dim Pos As Integer
    
    'Clear picture box
    picresults.Cls
    
     'Open File to obtain High Scores
    Open App.Path & "\HallofFame.txt" For Input As #1
    
    Ctr = 0
    
    Do Until EOF(1)
        Ctr = Ctr + 1
        Input #1, player(Ctr), score(Ctr)
    Loop
    
    picresults.Print "These players achieved Better Scores"
    picresults.Print "*********************************************"
            'Exhaustive Search.
        'This loop will search for a score and print all scores lower than the input.
     inputscore = InputBox("Please enter a score you wish to find", "Score")
   For Pos = 1 To Ctr
        If inputscore >= score(Pos) Then
            picresults.Print player(Pos); Tab(20); score(Pos)
        End If
    Next Pos
    
    
    Close #1
    
End Sub

Private Sub cmdSortNames_Click()
    'Declare Variables in Bubble Sort
    Dim Pass As Integer, Pos As Integer, Tempscore As Integer, Tempname As String
    
    'Clear the picture box
    picresults.Cls
    
    'Bubble Sort by Names
    For Pass = 1 To Ctr - 1
        For Pos = 1 To Ctr - Pass
            If player(Pos) > player(Pos + 1) Then
                Tempscore = score(Pos)
                score(Pos) = score(Pos + 1)
                score(Pos + 1) = Tempscore
                
                Tempname = player(Pos)
                player(Pos) = player(Pos + 1)
                player(Pos + 1) = Tempname
                
            End If
        Next Pos
    Next Pass
    
    'Print a Heading
    picresults.Print "Scores Sorted By Name"
    picresults.Print "********************************"
    
    'Print Sorted Scores
    For Pos = 1 To Ctr
        picresults.Print player(Pos); Tab(20); score(Pos)
    Next Pos
                
End Sub

Private Sub cmdSortScores_Click()
    'Declare Variables in Bubble Sort
    Dim Pass As Integer, Pos As Integer, Tempscore As Integer, Tempname As String
    
    'Clear the picture box
    picresults.Cls
    
    'Bubble Sort by Score
    For Pass = 1 To Ctr - 1
        For Pos = 1 To Ctr - Pass
            If score(Pos) > score(Pos + 1) Then
                Tempscore = score(Pos)
                score(Pos) = score(Pos + 1)
                score(Pos + 1) = Tempscore
                
                Tempname = player(Pos)
                player(Pos) = player(Pos + 1)
                player(Pos + 1) = Tempname
                
            End If
        Next Pos
    Next Pass
    
    'Print a Heading
    picresults.Print "Scores Sorted By Score"
    picresults.Print "*********************************"
    
    
    'Print Sorted Scores
    For Pos = 1 To Ctr
        picresults.Print player(Pos); Tab(20); score(Pos)
    Next Pos
                
            
    
    
    
End Sub

