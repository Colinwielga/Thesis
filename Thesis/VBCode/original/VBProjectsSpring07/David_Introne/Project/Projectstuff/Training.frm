VERSION 5.00
Begin VB.Form Training 
   Caption         =   "Training"
   ClientHeight    =   8325
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12330
   LinkTopic       =   "Form1"
   ScaleHeight     =   8325
   ScaleWidth      =   12330
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton TrainIN 
      Caption         =   "Train!"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   600
      TabIndex        =   8
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CommandButton AlphaSort 
      Caption         =   "Alphabetical Sort and Display"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9960
      TabIndex        =   7
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton Back 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   360
      TabIndex        =   6
      Top             =   6360
      Width           =   2655
   End
   Begin VB.PictureBox PicDisplay 
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   7200
      ScaleHeight     =   5715
      ScaleWidth      =   5115
      TabIndex        =   5
      Top             =   2760
      Width           =   5175
   End
   Begin VB.CommandButton CmdTrainOp 
      Caption         =   "Display Training Options"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7200
      TabIndex        =   4
      Top             =   960
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   8475
      Left            =   0
      Picture         =   "Training.frx":0000
      Top             =   0
      Width           =   12735
   End
   Begin VB.Label lblshep 
      Caption         =   "Train pup as a herding and working dog. "
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   9960
      TabIndex        =   3
      Top             =   8400
      Width           =   3375
   End
   Begin VB.Label Lblpit 
      Caption         =   "Train pup as a loving, "
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   9720
      TabIndex        =   2
      Top             =   7920
      Width           =   3135
   End
   Begin VB.Label LblDuch 
      Caption         =   "Train pup as a hunting dog. "
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   9720
      TabIndex        =   1
      Top             =   8040
      Width           =   3375
   End
   Begin VB.Label lblMtn 
      Caption         =   "Train Pup as a pulling and working dog."
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   9720
      TabIndex        =   0
      Top             =   8040
      Width           =   3135
   End
End
Attribute VB_Name = "Training"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ctr As Integer, cnt As Integer
Dim Pos As Integer, Train(1 To 10) As String, Number(1 To 10) As String


Private Sub AlphaSort_Click()
    Dim Pass As Integer 'sets variables
    Dim Pos As Integer
    Dim Train2 As String
    
    Dim Number2 As String
   
    PicDisplay.Print "********************************************************"
    

    'Sorting alphabetically
    
    For Pass = 1 To (ctr - 1)
    
        For Pos = 1 To (ctr - Pass)
        
            If Train(Pos) > Train(Pos + 1) Then
                Train2 = Train(Pos)
                Train(Pos) = Train(Pos + 1)
                Train(Pos + 1) = Train2
                
                'Swap 2nd array
                Number2 = Number(Pos)
                Number(Pos) = Number(Pos + 1)
                Number(Pos + 1) = Number2
            End If
        Next Pos
    Next Pass
    
    'print out all the new stuff
    For Pos = 1 To ctr
    
        PicDisplay.Print Train(Pos); Tab(35); Number(Pos)
    Next Pos
End Sub

Private Sub Back_Click()
Select Case puppick 'sends you back based on the dog chosen
            Case 11 'moves on to next case if previous did not work
                Training.Hide
                ProShep.Show
            Case 12 'moves on to next case if previous did not work
                Training.Hide
                ProPit.Show
            Case 13 'moves on to next case if previous did not work
                Training.Hide
                ProMtn.Show
            Case 14
                Training.Hide
                Produch.Show
    End Select
End Sub

Private Sub CmdTrainOp_Click()

    'This loads file into parallel arrays
        PicDisplay.Cls 'this clears the Pic and sets up a table
        PicDisplay.Print "Training"; Tab(30); "Designation #"
        PicDisplay.Print "********************************************************"
        Open App.Path & "\training.txt" For Input As #1
        ctr = 0
        Do Until EOF(1)
            ctr = ctr + 1
            Input #1, Train(ctr), Number(ctr)
        Loop
        Close #1
        
        'This prints the training options out.
        For Pos = 1 To ctr
            PicDisplay.Print Train(Pos); Tab(35); Number(Pos)
        Next Pos
End Sub

Private Sub TrainIN_Click()
Dim Des As Integer

cnt = cnt + 1



If cnt <= 3 Then 'when cnt is under three these next things can happen
3
Des = InputBox("Please input a number that cooresponds to the training.")
    Select Case Des
        Case 1 'moves on to next case if previous did not work
            Select Case puppick
                Case 11 'moves on to next case if previous did not work
                    MsgBox "Shepards make great gaurd dogs. 5 points for Training", , "Results"
                    MsgBox "You gain 7 points"
                    Score = Score + 7
                Case 12 'moves on to next case if previous did not work
                    MsgBox "Pit bull terrier mix's are feirce gaurd dogs.", , "Results"
                    MsgBox "Plus 10 points", , Score
                    Score = Score + 10
                Case 13 'moves on to next case if previous did not work
                    MsgBox "Berneese Mtn dogs are actaully very nice, but are intimidating due to their size large.", , "Results"
                    MsgBox "You gained 5 points.", , Score
                    Score = Score + 2
                Case 14
                    MsgBox "Although loud and proportionaly feirce, Mini's are better suited for their original purpose, hunting.", , "Results"
                    MsgBox "You gained only 2 points.", , Score
                    Score = Score + 2
            End Select
        Case 2 'moves on to next case if previous did not work
            Select Case puppick
                Case 11 'moves on to next case if previous did not work
                    MsgBox "Shepards are decent working dogs", , "Results"
                    MsgBox "You gain 5 points", , Score
                    Score = Score + 5
                Case 12 'moves on to next case if previous did not work
                    MsgBox "Pit bull's are useful for short amounts of work (they tend to tire after a while), they were bread to deal with Bulls, originally.", , "Results"
                    MsgBox "Plus 5 points", , Score
                    Score = Score + 5
                Case 13 'moves on to next case if previous did not work
                    MsgBox "Berneese Mtn dogs are perfect working dogs, especialy in whether.", , "Results"
                    MsgBox "You gained a full 10 points.", , Score
                    Score = Score + 10
                Case 14 'moves on to next case if previous did not work
                    MsgBox "Duchshunds are not good with work.", , "Results"
                    MsgBox "You gained only 2 points.", , Score
                    Score = Score + 2
                End Select
        Case 3 'moves on to next case if previous did not work
            Select Case puppick
                Case 11 'moves on to next case if previous did not work
                    MsgBox "Shepards are great herding animals", , "Results"
                    MsgBox "You gain 10 points", , Score
                    Score = Score + 10
                Case 12 'moves on to next case if previous did not work
                    MsgBox "Pit bulls aren't typical herders.", , "Results"
                    MsgBox "Plus 5 points", , Score
                    Score = Score + 5
                Case 13 'moves on to next case if previous did not work
                    MsgBox "Berneese Mtn dogs not very good with herding, they kind of lumber along.", , "Results"
                    MsgBox "You gained only 2 points.", , Score
                    Score = Score + 5
                Case 14
                    MsgBox "Duchshunds are not a typical herding dogs.", , "Results"
                    MsgBox "You gained only 2 points.", , Score
                    Score = Score + 2
                End Select
         Case 4 'moves on to next case if previous did not work
            Select Case puppick
                Case 11 'moves on to next case if previous did not work
                    MsgBox "Shepards are only decent hunting dogs, and not typical.", , "Results"
                    MsgBox "You gain only 2 points", , Score
                    Score = Score + 2
                Case 12 'moves on to next case if previous did not work
                    MsgBox "Pit bulls aren't great for hunting.", , "Results"
                    MsgBox "Plus only 2 points", , Score
                    Score = Score + 2
                Case 13 'moves on to next case if previous did not work
                    MsgBox "Berneese Mtn dogs found hunting, but usualy end up carrying the game.", , "Results"
                    MsgBox "You gained 7 points.", , Score
                    Score = Score + 7
                Case 14
                    MsgBox "Duchshunds are bread for hunting.", , "Results"
                    MsgBox "You gained a full 10 points.", , Score
                    Score = Score + 10
                End Select
        Case 5 'moves on to next case if previous did not work
            Select Case puppick
                Case 11 'moves on to next case if previous did not work
                    MsgBox "Shepards are used for rescue training all over the world.", , "Results"
                    MsgBox "You gain 10 points", , Score
                    Score = Score + 2
                Case 12 'moves on to next case if previous did not work
                    MsgBox "Pit's are rarely used for this job", , "Results"
                    MsgBox "Plus only 2 points", , Score
                    Score = Score + 2
                Case 13 'moves on to next case if previous did not work
                    MsgBox "Berneese Mtn dogs are great for cold whether rescue.", , "Results"
                    MsgBox "You gained 10 points.", , Score
                    Score = Score + 7
                Case 14
                    MsgBox "Duchshunds are not usually used for rescue missions.", , "Results"
                    MsgBox "You gained only 2 points.", , Score
                    Score = Score + 10
                End Select
        Case 6 'moves on to next case if previous did not work
            MsgBox "All dogs with good attitudes love to be shown.", , "Results"
            MsgBox "You gained 5 points.", , Score
            Score = Score + 5
        Case 7 'moves on to next case if previous did not work
            Select Case puppick
                Case 11 'moves on to next case if previous did not work
                    MsgBox "Shepards are good family dogs when properly trained.", , "Results"
                    MsgBox "You gain 7 points", , Score
                    Score = Score + 7
                Case 12 'moves on to next case if previous did not work
                    MsgBox "Pit's aren't recommended family dogs, although if your going to won one this would be the best way to train them.", , "Results"
                    MsgBox "Plus 5 points", , Score
                    Score = Score + 5
                Case 13 'moves on to next case if previous did not work
                    MsgBox "Berneese Mtn dogs are great family dogs, sweet, loyal and fluffy.", , "Results"
                    MsgBox "You gained 10 points.", , Score
                    Score = Score + 10
                Case 14
                    MsgBox "Duchshunds are good family dogs when well adjusted to a home.", , "Results"
                    MsgBox "You gained 7 points.", , Score
                    Score = Score + 7
                End Select
          Case 8 'moves on to next case if previous did not work
            Select Case puppick
                Case 11 'moves on to next case if previous did not work
                    MsgBox "Shepards don't make good lap dogs.", , "Results"
                    MsgBox "You gain only 2 points", , Score
                    Score = Score + 2
                Case 12 'moves on to next case if previous did not work
                    MsgBox "Pit's are heavy despite their meduim build.", , "Results"
                    MsgBox "Plus only 2 points", , Score
                    Score = Score + 2
                Case 13 'moves on to next case if previous did not work
                    MsgBox "Berneese Mtn dogs are heavy and drooly, not ideal lap dogs.", , "Results"
                    MsgBox "You gained only 2 points.", , Score
                    Score = Score + 10
                Case 14
                    MsgBox "Duchshunds are great lap dogs when they sit still.", , "Results"
                    MsgBox "You gained 10 points.", , Score
                    Score = Score + 10
            End Select
        End Select
    Else 'if If did not work else presents other statement
    MsgBox Pupname, , " can only be trained in 3 areas."
    End If
    End Sub
