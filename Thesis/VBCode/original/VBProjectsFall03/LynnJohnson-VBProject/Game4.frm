VERSION 5.00
Begin VB.Form FrmGame4 
   BackColor       =   &H00800080&
   Caption         =   "Game 4 (Lynn Johnson)"
   ClientHeight    =   10560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12765
   LinkTopic       =   "Form1"
   ScaleHeight     =   10560
   ScaleWidth      =   12765
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsort 
      Caption         =   "Sort Alphabetically"
      Height          =   615
      Left            =   10920
      TabIndex        =   43
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdfind 
      Caption         =   "Click Here First"
      Height          =   615
      Left            =   9480
      TabIndex        =   41
      Top             =   3960
      Width           =   1215
   End
   Begin VB.PictureBox scoreresults 
      BackColor       =   &H00FFC0FF&
      Height          =   1215
      Left            =   9360
      ScaleHeight     =   1155
      ScaleWidth      =   3195
      TabIndex        =   40
      Top             =   8640
      Width           =   3255
   End
   Begin VB.CommandButton cmdscore 
      Caption         =   "Calculate Score"
      Height          =   495
      Left            =   10200
      TabIndex        =   39
      Top             =   8040
      Width           =   1575
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear Box"
      Height          =   495
      Left            =   10320
      TabIndex        =   38
      Top             =   6960
      Width           =   1095
   End
   Begin VB.PictureBox results 
      BackColor       =   &H00FFC0FF&
      Height          =   2055
      Left            =   9840
      ScaleHeight     =   1995
      ScaleWidth      =   1995
      TabIndex        =   37
      Top             =   4680
      Width           =   2055
   End
   Begin VB.PictureBox pbxresults 
      BackColor       =   &H00FFC0FF&
      Height          =   495
      Left            =   5760
      ScaleHeight     =   435
      ScaleWidth      =   4275
      TabIndex        =   36
      Top             =   360
      Width           =   4335
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   11040
      TabIndex        =   34
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdplay 
      Caption         =   "Play again"
      Height          =   735
      Left            =   9840
      TabIndex        =   33
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to Menu"
      Height          =   735
      Left            =   10440
      TabIndex        =   32
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdthirteen 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   7200
      TabIndex        =   31
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton cmdsixteen 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   4920
      TabIndex        =   30
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton cmdeleven 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   2640
      TabIndex        =   29
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton cmdone 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   360
      TabIndex        =   28
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton cmdtwo 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   360
      TabIndex        =   27
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton cmdsix 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   2640
      TabIndex        =   26
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton cmdnine 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   4920
      TabIndex        =   25
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton cmdeight 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   7200
      TabIndex        =   24
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton cmdtwelve 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   7200
      TabIndex        =   23
      Top             =   6000
      Width           =   1935
   End
   Begin VB.CommandButton cmdfive 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   4920
      TabIndex        =   22
      Top             =   6000
      Width           =   1935
   End
   Begin VB.CommandButton cmdten 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   2640
      TabIndex        =   21
      Top             =   6000
      Width           =   1935
   End
   Begin VB.CommandButton cmdseven 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   360
      TabIndex        =   20
      Top             =   6000
      Width           =   1935
   End
   Begin VB.CommandButton cmdfour 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   360
      TabIndex        =   19
      Top             =   8160
      Width           =   1935
   End
   Begin VB.CommandButton cmdfifteen 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   2640
      TabIndex        =   18
      Top             =   8160
      Width           =   1935
   End
   Begin VB.CommandButton cmdfourteen 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   4920
      TabIndex        =   17
      Top             =   8160
      Width           =   1935
   End
   Begin VB.CommandButton cmdthree 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   7200
      TabIndex        =   16
      Top             =   8160
      Width           =   1935
   End
   Begin VB.PictureBox picresults16 
      Height          =   1815
      Left            =   4920
      Picture         =   "Game4.frx":0000
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   15
      Top             =   1680
      Width           =   1935
   End
   Begin VB.PictureBox picresults14 
      Height          =   1815
      Left            =   4920
      Picture         =   "Game4.frx":4C84
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   14
      Top             =   8160
      Width           =   1935
   End
   Begin VB.PictureBox picresults13 
      Height          =   1815
      Left            =   7200
      Picture         =   "Game4.frx":9912
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   13
      Top             =   1680
      Width           =   1935
   End
   Begin VB.PictureBox picresults5 
      Height          =   1815
      Left            =   4920
      Picture         =   "Game4.frx":E5A0
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   12
      Top             =   6000
      Width           =   1935
   End
   Begin VB.PictureBox picresults11 
      Height          =   1815
      Left            =   2640
      Picture         =   "Game4.frx":137C7
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   11
      Top             =   1680
      Width           =   1935
   End
   Begin VB.PictureBox picresults10 
      Height          =   1815
      Left            =   2640
      Picture         =   "Game4.frx":17F46
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   10
      Top             =   6000
      Width           =   1935
   End
   Begin VB.PictureBox picresults8 
      Height          =   1815
      Left            =   7200
      Picture         =   "Game4.frx":1E94F
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   9
      Top             =   3840
      Width           =   1935
   End
   Begin VB.PictureBox picresults12 
      Height          =   1815
      Left            =   7200
      Picture         =   "Game4.frx":24344
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   8
      Top             =   6000
      Width           =   1935
   End
   Begin VB.PictureBox picresults15 
      Height          =   1815
      Left            =   2640
      Picture         =   "Game4.frx":28AC3
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   7
      Top             =   8160
      Width           =   1935
   End
   Begin VB.PictureBox picresults6 
      Height          =   1815
      Left            =   2640
      Picture         =   "Game4.frx":2D747
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   6
      Top             =   3840
      Width           =   1935
   End
   Begin VB.PictureBox picresults9 
      Height          =   1815
      Left            =   4920
      Picture         =   "Game4.frx":3296E
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   5
      Top             =   3840
      Width           =   1935
   End
   Begin VB.PictureBox picresults4 
      Height          =   1815
      Left            =   360
      Picture         =   "Game4.frx":39377
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   4
      Top             =   8160
      Width           =   1935
   End
   Begin VB.PictureBox picresults3 
      Height          =   1815
      Left            =   7200
      Picture         =   "Game4.frx":3EB31
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   3
      Top             =   8160
      Width           =   1935
   End
   Begin VB.PictureBox picresults7 
      Height          =   1815
      Left            =   360
      Picture         =   "Game4.frx":442EB
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   2
      Top             =   6000
      Width           =   1935
   End
   Begin VB.PictureBox picresults2 
      Height          =   1815
      Left            =   360
      Picture         =   "Game4.frx":49CE0
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   1
      Top             =   3840
      Width           =   1935
   End
   Begin VB.PictureBox picresults1 
      Height          =   1815
      Left            =   360
      Picture         =   "Game4.frx":4EA3F
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   0
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      Caption         =   "Three points will be subtracted from 100 for every mismatch."
      Height          =   495
      Left            =   9480
      TabIndex        =   44
      Top             =   9960
      Width           =   3015
   End
   Begin VB.Label lblclick 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Do you want to know what you're looking for?  Click below and find out!"
      Height          =   495
      Left            =   9480
      TabIndex        =   42
      Top             =   3240
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0FF&
      Caption         =   $"Game4.frx":5379E
      Height          =   735
      Left            =   480
      TabIndex        =   35
      Top             =   240
      Width           =   4815
   End
End
Attribute VB_Name = "FrmGame4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Purpose: This form allows the user to play a game of memory
        'by clicking on the cards.  It allows the user to get
        'a list of the pictures on the back of the cards.
        'It also creates a score for the game and shows
        'the player what kind of cards they are looking
        'for.  Following the game, this form allows
        'the user to play again using a different form,
        'return to the menu, or quit.

Option Explicit

Private Sub cmdclear_Click()
    'clears the card information box
    results.Cls
    
End Sub

Private Sub cmdeight_Click()
    'counts the amount of times this button is pushed
    'prepares for scoring
    icount = icount + 1
    pbxresults.Cls
    'cards flip over to show picture
    cmdeight.Visible = False
    picresults8.Visible = True
    
    If cmdseven.Visible = False Then
        pbxresults.Print "You found another dog! You found a match!"
    Else
        pbxresults.Print "You found a dog!"
    End If
    'these same actions take place for all sixteen "Memory Card" command buttons

End Sub

Private Sub cmdeleven_Click()
    icount = icount + 1
    pbxresults.Cls
    cmdeleven.Visible = False
    picresults11.Visible = True
    
    If cmdtwelve.Visible = False Then
        pbxresults.Print "You found another cow! You found a match!"
    Else
        pbxresults.Print "You found a cow!"
    End If
End Sub

Private Sub cmdfifteen_Click()
    icount = icount + 1
    pbxresults.Cls
    cmdfifteen.Visible = False
    picresults15.Visible = True
    
    If cmdsixteen.Visible = False Then
        pbxresults.Print "You found another chick! You found a match!"
    Else
        pbxresults.Print "You found a baby chick!"
    End If
End Sub

Private Sub cmdfive_Click()
    icount = icount + 1
    pbxresults.Cls
    cmdfive.Visible = False
    picresults5.Visible = True
    
    If cmdsix.Visible = False Then
        pbxresults.Print "You found the other bird! You found a match!"
    Else
        pbxresults.Print "You found a bird!"
    End If
End Sub

Private Sub cmdfour_Click()
    icount = icount + 1
    pbxresults.Cls
    cmdfour.Visible = False
    picresults4.Visible = True
    
    If cmdthree.Visible = False Then
        pbxresults.Print "You found the other elephant! You found a match!"
    Else
        pbxresults.Print "You found an elephant!"
    End If
End Sub

Private Sub cmdfourteen_Click()
    icount = icount + 1
    pbxresults.Cls
    cmdfourteen.Visible = False
    picresults14.Visible = True
    
    If cmdthirteen.Visible = False Then
        pbxresults.Print "You found another fish! You found a match!"
    Else
        pbxresults.Print "You found a fish!"
    End If
End Sub

Private Sub cmdnine_Click()
    icount = icount + 1
    pbxresults.Cls
    cmdnine.Visible = False
    picresults9.Visible = True
    
    If cmdten.Visible = False Then
        pbxresults.Print "You found another butterfly! You found a match!"
    Else
        pbxresults.Print "You found a butterfly!"
    End If
End Sub

Private Sub cmdone_Click()
    icount = icount + 1
    pbxresults.Cls
    cmdone.Visible = False
    picresults1.Visible = True
    
    If cmdtwo.Visible = False Then
        pbxresults.Print "You found the other frog! You found a match!"
    Else
        pbxresults.Print "You found a frog!"
    End If
    
End Sub

Private Sub cmdplay_Click()
    'Move from one game form to another game form
    Dim game As Integer
    game = InputBox("Enter a game number between 1 and 4")
    
    'sets up all games ready to play
    cmdone.Visible = True
    cmdtwo.Visible = True
    cmdthree.Visible = True
    cmdfour.Visible = True
    cmdfive.Visible = True
    cmdsix.Visible = True
    cmdseven.Visible = True
    cmdeight.Visible = True
    cmdnine.Visible = True
    cmdten.Visible = True
    cmdeleven.Visible = True
    cmdtwelve.Visible = True
    cmdthirteen.Visible = True
    cmdfourteen.Visible = True
    cmdfifteen.Visible = True
    cmdsixteen.Visible = True
    picresults1.Visible = False
    picresults2.Visible = False
    picresults3.Visible = False
    picresults4.Visible = False
    picresults5.Visible = False
    picresults6.Visible = False
    picresults7.Visible = False
    picresults8.Visible = False
    picresults9.Visible = False
    picresults10.Visible = False
    picresults11.Visible = False
    picresults12.Visible = False
    picresults13.Visible = False
    picresults14.Visible = False
    picresults15.Visible = False
    picresults16.Visible = False
        
    'clears previous score
    icount = 0
    cmatch = 0
    mismatch = 0
    sum = 0
    score = 0
    
    Select Case game
        Case 1
            frmGame1.Show
            FrmGame4.Hide
        Case 2
            frmGame2.Show
            FrmGame4.Hide
        Case 3
            frmGame3.Show
            FrmGame4.Hide
        Case 4
            FrmGame5.Show
            FrmGame4.Hide
        Case Else
            MsgBox "That number is not between 1 and 4.  Pick another number", , "Error"
    End Select
    
End Sub

Private Sub cmdquit_Click()
    End
    
End Sub

Private Sub cmdreturn_Click()
    'returns to main menu
    frmMenu.Show
    FrmGame4.Hide
    
    'makes sure all games are ready to play
    cmdone.Visible = True
    cmdtwo.Visible = True
    cmdthree.Visible = True
    cmdfour.Visible = True
    cmdfive.Visible = True
    cmdsix.Visible = True
    cmdseven.Visible = True
    cmdeight.Visible = True
    cmdnine.Visible = True
    cmdten.Visible = True
    cmdeleven.Visible = True
    cmdtwelve.Visible = True
    cmdthirteen.Visible = True
    cmdfourteen.Visible = True
    cmdfifteen.Visible = True
    cmdsixteen.Visible = True
    picresults1.Visible = False
    picresults2.Visible = False
    picresults3.Visible = False
    picresults4.Visible = False
    picresults5.Visible = False
    picresults6.Visible = False
    picresults7.Visible = False
    picresults8.Visible = False
    picresults9.Visible = False
    picresults10.Visible = False
    picresults11.Visible = False
    picresults12.Visible = False
    picresults13.Visible = False
    picresults14.Visible = False
    picresults15.Visible = False
    picresults16.Visible = False
        
    'clears previous scores of any kind
    icount = 0
    cmatch = 0
    mismatch = 0
    sum = 0
    score = 0
    
End Sub

Private Sub cmdscore_Click()
    'scores are found and displayed based on the amount
    'of matches found and amount of mismatches found.
    scoreresults.Cls
    sum = 0
    'counts amount of pairs selected
    sum = icount / 2
    'finds amount of mismatches found
    mismatch = sum - cmatch
    'calculates score
    score = 100 - 3 * mismatch
    
    If score = 100 And cmatch = 8 Then
        scoreresults.Print "You got a perfect score! Good job!"
    ElseIf icount = 0 Then
        scoreresults.Cls
    ElseIf cmatch < 8 Then
        scoreresults.Print "You have not finished your game."
        scoreresults.Print "Your tentative score is:"; mismatch; "mismatch(es)."
        scoreresults.Print "You have cleared"; cmatch; "pair(s)"
        scoreresults.Print "Your score Is"; score
        scoreresults.Print "To finish game continue play."
    Else
        scoreresults.Print "You flipped"; mismatch; "mismatch(es)."
        scoreresults.Print "You flipped"; cmatch; "pairs."
        scoreresults.Print "Your score Is"; score
    End If
    
End Sub

Private Sub cmdseven_Click()
    icount = icount + 1
    pbxresults.Cls
    cmdseven.Visible = False
    picresults7.Visible = True
    
    If cmdeight.Visible = False Then
        pbxresults.Print "You found another dog! You found a match!"
    Else
        pbxresults.Print "You found a dog!"
    End If
End Sub

Private Sub cmdsix_Click()
    icount = icount + 1
    pbxresults.Cls
    cmdsix.Visible = False
    picresults6.Visible = True
    
    If cmdfive.Visible = False Then
        pbxresults.Print "You found the other bird! You found a match!"
    Else
        pbxresults.Print "You found a bird!"
    End If
End Sub

Private Sub cmdsixteen_Click()
    icount = icount + 1
    pbxresults.Cls
    cmdsixteen.Visible = False
    picresults16.Visible = True
    
    If cmdfifteen.Visible = False Then
        pbxresults.Print "You found another chick! You found a match!"
    Else
        pbxresults.Print "You found a baby chick!"
    End If
End Sub

Private Sub cmdfind_Click()
    Dim i As Integer
    'open and read in from data file ACards(i)
    Open strPath For Input As #1
    For i = 1 To 8
        Input #1, ACards(i)
    Next i
    Close #1
    
    'print array from data file in a picture box
        results.Print "The Cards"
        results.Print "********************"
    
    For i = 1 To 8
        results.Print ACards(i)
    Next i
    
End Sub

Private Sub cmdsort_Click()
    Dim pass As Integer
    Dim temp As String
    Dim i As Integer
    Dim n As Integer
    n = 8
    results.Cls
    'sort contents from data file alphabetically
    For pass = 1 To n - 1
        For i = 1 To n - pass
            If ACards(i) > ACards(i + 1) Then
                temp = ACards(i)
                ACards(i) = ACards(i + 1)
                ACards(i + 1) = temp
            End If
        Next i
    Next pass
    
    'print what was sorted
        results.Print "The Cards"
        results.Print "********************"
    
    For i = 1 To 8
        results.Print ACards(i)
    Next i
    
End Sub

Private Sub cmdten_Click()
    icount = icount + 1
    pbxresults.Cls
    cmdten.Visible = False
    picresults10.Visible = True
    
    If cmdnine.Visible = False Then
        pbxresults.Print "You found another butterfly! You found a match!"
    Else
        pbxresults.Print "You found a butterfly!"
    End If
End Sub

Private Sub cmdthirteen_Click()
    icount = icount + 1
    pbxresults.Cls
    cmdthirteen.Visible = False
    picresults13.Visible = True
    
    If cmdfourteen.Visible = False Then
        pbxresults.Print "You found another fish! You found a match!"
    Else
        pbxresults.Print "You found a fish!"
    End If
End Sub

Private Sub cmdthree_Click()
    icount = icount + 1
    pbxresults.Cls
    cmdthree.Visible = False
    picresults3.Visible = True
    
    If cmdfour.Visible = False Then
        pbxresults.Print "You found the other elephant! You found a match!"
    Else
        pbxresults.Print "You found an elephant!"
    End If
    
End Sub

Private Sub cmdtwelve_Click()
    icount = icount + 1
    pbxresults.Cls
    cmdtwelve.Visible = False
    picresults12.Visible = True
    
    If cmdeleven.Visible = False Then
        pbxresults.Print "You found another cow! You found a match!"
    Else
        pbxresults.Print "You found a cow!"
    End If
End Sub

Private Sub cmdtwo_Click()
    icount = icount + 1
    pbxresults.Cls
    cmdtwo.Visible = False
    picresults2.Visible = True
    
    If cmdone.Visible = False Then
        pbxresults.Print "You found the other frog!  You found a match!"
    Else
        pbxresults.Print "You found a frog!"
    End If

End Sub


Private Sub picresults10_Click()
    'buttons disappear if a match is found
    'buttons cover again if a match is not found
    If cmdnine.Visible = False And cmdten.Visible = False Then
        picresults9.Visible = False
        picresults10.Visible = False
        cmatch = cmatch + 1
    Else
        cmdten.Visible = True
        picresults10.Visible = False
    End If
        pbxresults.Cls
    'these same actions take place for all sixteen picture boxes
        
End Sub

Private Sub picresults1_Click()
    If cmdone.Visible = False And cmdtwo.Visible = False Then
        picresults1.Visible = False
        picresults2.Visible = False
        cmatch = cmatch + 1
    Else
        cmdone.Visible = True
        picresults1.Visible = False
    End If
        pbxresults.Cls
    
End Sub


Private Sub picresults11_Click()
    If cmdeleven.Visible = False And cmdtwelve.Visible = False Then
        picresults11.Visible = False
        picresults12.Visible = False
        cmatch = cmatch + 1
    Else
        cmdeleven.Visible = True
        picresults11.Visible = False
    End If
        pbxresults.Cls
    
End Sub

Private Sub picresults12_Click()
    If cmdeleven.Visible = False And cmdtwelve.Visible = False Then
        picresults11.Visible = False
        picresults12.Visible = False
        cmatch = cmatch + 1
    Else
        cmdtwelve.Visible = True
        picresults12.Visible = False
    End If
        pbxresults.Cls
    
End Sub

Private Sub picresults14_Click()
    If cmdthirteen.Visible = False And cmdfourteen.Visible = False Then
        picresults13.Visible = False
        picresults14.Visible = False
        cmatch = cmatch + 1
    Else
        cmdfourteen.Visible = True
        picresults14.Visible = False
    End If
        pbxresults.Cls
    
End Sub

Private Sub picresults15_Click()
    If cmdfifteen.Visible = False And cmdsixteen.Visible = False Then
        picresults15.Visible = False
        picresults16.Visible = False
        cmatch = cmatch + 1
    Else
        cmdfifteen.Visible = True
        picresults15.Visible = False
    End If
        pbxresults.Cls
    
End Sub

Private Sub picresults16_Click()
    If cmdfifteen.Visible = False And cmdsixteen.Visible = False Then
        picresults15.Visible = False
        picresults16.Visible = False
        cmatch = cmatch + 1
    Else
        cmdsixteen.Visible = True
        picresults16.Visible = False
    End If
        pbxresults.Cls
    
End Sub

Private Sub picresults2_Click()
    If cmdone.Visible = False And cmdtwo.Visible = False Then
        picresults1.Visible = False
        picresults2.Visible = False
        cmatch = cmatch + 1
    Else
        cmdtwo.Visible = True
        picresults2.Visible = False
    End If
        pbxresults.Cls
    
End Sub

Private Sub picresults3_Click()
    If cmdthree.Visible = False And cmdfour.Visible = False Then
        picresults3.Visible = False
        picresults4.Visible = False
        cmatch = cmatch + 1
    Else
        cmdthree.Visible = True
        picresults3.Visible = False
    End If
        pbxresults.Cls
    
End Sub

Private Sub picresults4_Click()
    If cmdthree.Visible = False And cmdfour.Visible = False Then
        picresults3.Visible = False
        picresults4.Visible = False
        cmatch = cmatch + 1
    Else
        cmdfour.Visible = True
        picresults4.Visible = False
    End If
        pbxresults.Cls
    
End Sub

Private Sub picresults5_Click()
    If cmdfive.Visible = False And cmdsix.Visible = False Then
        picresults5.Visible = False
        picresults6.Visible = False
        cmatch = cmatch + 1
    Else
        cmdfive.Visible = True
        picresults5.Visible = False
    End If
        pbxresults.Cls
    
End Sub

Private Sub picresults6_Click()
    If cmdfive.Visible = False And cmdsix.Visible = False Then
        picresults5.Visible = False
        picresults6.Visible = False
        cmatch = cmatch + 1
    Else
        cmdsix.Visible = True
        picresults6.Visible = False
    End If
        pbxresults.Cls
    
End Sub

Private Sub picresults7_Click()
    If cmdseven.Visible = False And cmdeight.Visible = False Then
        picresults7.Visible = False
        picresults8.Visible = False
        cmatch = cmatch + 1
    Else
        cmdseven.Visible = True
        picresults7.Visible = False
    End If
        pbxresults.Cls
    
End Sub

Private Sub picresults8_Click()
    If cmdseven.Visible = False And cmdeight.Visible = False Then
        picresults7.Visible = False
        picresults8.Visible = False
        cmatch = cmatch + 1
    Else
        cmdeight.Visible = True
        picresults8.Visible = False
    End If
        pbxresults.Cls
    
End Sub

Private Sub picresults9_Click()
    If cmdnine.Visible = False And cmdten.Visible = False Then
        picresults9.Visible = False
        picresults10.Visible = False
        cmatch = cmatch + 1
    Else
        cmdnine.Visible = True
        picresults9.Visible = False
    End If
        pbxresults.Cls
    
End Sub

Private Sub picresults13_Click()
    If cmdthirteen.Visible = False And cmdfourteen.Visible = False Then
        picresults13.Visible = False
        picresults14.Visible = False
        cmatch = cmatch + 1
    Else
        cmdthirteen.Visible = True
        picresults13.Visible = False
    End If
        pbxresults.Cls
    
End Sub


