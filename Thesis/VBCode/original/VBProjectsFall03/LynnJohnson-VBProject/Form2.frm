VERSION 5.00
Begin VB.Form frmGame1 
   BackColor       =   &H008080FF&
   Caption         =   "Game 1"
   ClientHeight    =   10380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12615
   LinkTopic       =   "Form1"
   ScaleHeight     =   10380
   ScaleWidth      =   12615
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox scoreresults 
      Height          =   495
      Left            =   10320
      ScaleHeight     =   435
      ScaleWidth      =   1995
      TabIndex        =   42
      Top             =   8880
      Width           =   2055
   End
   Begin VB.CommandButton cmdscore 
      Caption         =   "Calculate Score"
      Height          =   735
      Left            =   9120
      TabIndex        =   41
      Top             =   8760
      Width           =   855
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear Box"
      Height          =   495
      Left            =   10080
      TabIndex        =   40
      Top             =   6960
      Width           =   1455
   End
   Begin VB.PictureBox results 
      Height          =   2055
      Left            =   9720
      ScaleHeight     =   1995
      ScaleWidth      =   1995
      TabIndex        =   39
      Top             =   4800
      Width           =   2055
   End
   Begin VB.CommandButton cmdsixteen 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   6840
      TabIndex        =   38
      Top             =   7800
      Width           =   1935
   End
   Begin VB.CommandButton cmdone 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   4680
      TabIndex        =   37
      Top             =   7800
      Width           =   1935
   End
   Begin VB.CommandButton cmdfour 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   2520
      TabIndex        =   36
      Top             =   7800
      Width           =   1935
   End
   Begin VB.CommandButton cmdeleven 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   360
      TabIndex        =   35
      Top             =   7800
      Width           =   1935
   End
   Begin VB.CommandButton cmdtwelve 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   6840
      TabIndex        =   34
      Top             =   5760
      Width           =   1935
   End
   Begin VB.CommandButton cmdsix 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   4680
      TabIndex        =   33
      Top             =   5760
      Width           =   1935
   End
   Begin VB.CommandButton cmdseven 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   2520
      TabIndex        =   32
      Top             =   5760
      Width           =   1935
   End
   Begin VB.CommandButton cmdthirteen 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   360
      TabIndex        =   31
      Top             =   5760
      Width           =   1935
   End
   Begin VB.CommandButton cmdeight 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   6840
      TabIndex        =   30
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CommandButton cmdten 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   4680
      TabIndex        =   29
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CommandButton cmdtwo 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   2520
      TabIndex        =   28
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CommandButton cmdfive 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   360
      TabIndex        =   27
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CommandButton cmdfourteen 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   6840
      TabIndex        =   26
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton cmdthree 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   4680
      TabIndex        =   25
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton cmdnine 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   2520
      TabIndex        =   24
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton cmdfifteen 
      Caption         =   "Memory Card"
      Height          =   1815
      Left            =   360
      TabIndex        =   23
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton cmdfind 
      Caption         =   "Click Here"
      Height          =   735
      Left            =   9960
      TabIndex        =   22
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to Menu"
      Height          =   735
      Left            =   10200
      TabIndex        =   20
      Top             =   240
      Width           =   1215
   End
   Begin VB.PictureBox picresults16 
      Height          =   1815
      Left            =   6840
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   18
      Top             =   7800
      Width           =   1935
   End
   Begin VB.PictureBox picresults14 
      Height          =   1815
      Left            =   6840
      Picture         =   "Form2.frx":4C84
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   16
      Top             =   1680
      Width           =   1935
   End
   Begin VB.PictureBox picresults13 
      Height          =   1815
      Left            =   360
      Picture         =   "Form2.frx":9912
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   15
      Top             =   5760
      Width           =   1935
   End
   Begin VB.PictureBox picresults12 
      Height          =   1815
      Left            =   6840
      Picture         =   "Form2.frx":E5A0
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   14
      Top             =   5760
      Width           =   1935
   End
   Begin VB.PictureBox picresults11 
      Height          =   1815
      Left            =   360
      Picture         =   "Form2.frx":12D1F
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   13
      Top             =   7800
      Width           =   1935
   End
   Begin VB.PictureBox picresults10 
      Height          =   1815
      Left            =   4680
      Picture         =   "Form2.frx":1749E
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   12
      Top             =   3720
      Width           =   1935
   End
   Begin VB.PictureBox picresults9 
      Height          =   1815
      Left            =   2520
      Picture         =   "Form2.frx":1DEA7
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   11
      Top             =   1680
      Width           =   1935
   End
   Begin VB.PictureBox picresults8 
      Height          =   1815
      Left            =   6840
      Picture         =   "Form2.frx":248B0
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   10
      Top             =   3720
      Width           =   1935
   End
   Begin VB.PictureBox picresults7 
      Height          =   1815
      Left            =   2520
      Picture         =   "Form2.frx":2A2A5
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   9
      Top             =   5760
      Width           =   1935
   End
   Begin VB.PictureBox picresults6 
      Height          =   1815
      Left            =   4680
      Picture         =   "Form2.frx":2FC9A
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   8
      Top             =   5760
      Width           =   1935
   End
   Begin VB.PictureBox picresults5 
      Height          =   1815
      Left            =   360
      Picture         =   "Form2.frx":34EC1
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   7
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CommandButton cmdplay 
      Caption         =   "Play again"
      Height          =   735
      Left            =   9360
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.PictureBox pbxresults 
      BackColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   5520
      ScaleHeight     =   435
      ScaleWidth      =   3435
      TabIndex        =   5
      Top             =   360
      Width           =   3495
   End
   Begin VB.PictureBox picresults4 
      Height          =   1815
      Left            =   2520
      Picture         =   "Form2.frx":3A0E8
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   4
      Top             =   7800
      Width           =   1935
   End
   Begin VB.PictureBox picresults3 
      Height          =   1815
      Left            =   4680
      Picture         =   "Form2.frx":3F8A2
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   3
      Top             =   1680
      Width           =   1935
   End
   Begin VB.PictureBox picresults2 
      Height          =   1815
      Left            =   2520
      Picture         =   "Form2.frx":4505C
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   2
      Top             =   3720
      Width           =   1935
   End
   Begin VB.PictureBox picresults1 
      Height          =   1815
      Left            =   4680
      Picture         =   "Form2.frx":49DBB
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   1
      Top             =   7800
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   10920
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
   Begin VB.PictureBox picresults15 
      Height          =   1815
      Left            =   360
      Picture         =   "Form2.frx":4EB1A
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   17
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label lblclick 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Do you want to know what you're looking for?  Click below and find out!"
      Height          =   495
      Left            =   9360
      TabIndex        =   21
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label lblinstruction 
      BackColor       =   &H00C0E0FF&
      Caption         =   $"Form2.frx":5379E
      ForeColor       =   &H00404080&
      Height          =   615
      Left            =   240
      TabIndex        =   19
      Top             =   360
      Width           =   4815
   End
End
Attribute VB_Name = "frmGame1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Purpose: This form allows the user to play a game of memory
        'by clicking on the cards.  It also creates a score
        'for the game and shows the player what kind of cards
        'they ar looking for.  Following the game, this form
        'allows the user to play again using a different form,
        'return to the menu, or quit.

Option Explicit
Dim icount As Integer


Private Sub cmdclear_Click()
    results.Cls
    
End Sub

Private Sub cmdeight_Click()
    icount = icount + 1
    pbxresults.Cls
    cmdeight.Visible = False
    picresults8.Visible = True
    If cmdseven.Visible = False Then
        pbxresults.Print "You found another dog! You found a match!"
    Else
        pbxresults.Print "You found a dog!"
    End If
    

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
    Select Case game
        Case 1
            frmGame2.Show
            frmGame1.Hide
        Case 2
            frmGame3.Show
            frmGame1.Hide
        Case 3
            FrmGame4.Show
            frmGame1.Hide
        Case 4
            FrmGame5.Show
            frmGame1.Hide
        Case Else
            MsgBox "That number is not between 1 and 4.  Pick another number", , "Error"
    End Select
    
End Sub

Private Sub cmdquit_Click()
    End
    
End Sub

Private Sub cmdreturn_Click()
    frmMenu.Show
    frmGame1.Hide
    
End Sub

Private Sub cmdscore_Click()
    scoreresults.Cls
    score = 100 / 16
    If score > 0 Then
        scoreresults.Print "Your score is"; score; "missed"
    ElseIf score = 0 Then
        scoreresults.Print "You got a perfect score! Good job!"
    Else
        scoreresults.Print "You have not finished your game. Please finish to see your score."
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
    Dim ACards(1 To 8) As String
    Open "M:\CS130\VB Project\AnimalCards.txt" For Input As #1
    For i = 1 To 8
        Input #1, ACards(i)
    Next i
    
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
    If cmdnine.Visible = False And cmdten.Visible = False Then
        picresults9.Visible = False
        picresults10.Visible = False
    Else
        cmdten.Visible = True
        picresults10.Visible = False
    End If
        pbxresults.Cls
        
End Sub

Private Sub picresults1_Click()
    If cmdone.Visible = False And cmdtwo.Visible = False Then
        picresults1.Visible = False
        picresults2.Visible = False
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
    Else
        cmdthirteen.Visible = True
        picresults13.Visible = False
    End If
        pbxresults.Cls
    
End Sub
