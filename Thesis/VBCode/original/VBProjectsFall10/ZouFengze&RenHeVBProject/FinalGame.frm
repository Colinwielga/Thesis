VERSION 5.00
Begin VB.Form FinalGame 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Witch's hut"
   ClientHeight    =   7980
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   10710
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   10710
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   8055
      Left            =   4680
      ScaleHeight     =   7995
      ScaleWidth      =   4395
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.CommandButton btnMenuClose 
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      TabIndex        =   20
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame FrameMenu 
      BackColor       =   &H00000000&
      Caption         =   "Menu"
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   9240
      TabIndex        =   15
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
      Begin VB.CommandButton btnSave 
         Caption         =   "Save"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton btnLoad 
         Caption         =   "Load"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton btnMainMenu 
         Caption         =   "Main menu"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton btnExit 
         Caption         =   "Exit"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.CommandButton btnMenuOpen 
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      TabIndex        =   14
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton btnHelp 
      Caption         =   "Help"
      Height          =   495
      Left            =   2520
      TabIndex        =   12
      Top             =   3960
      Width           =   1095
   End
   Begin VB.PictureBox cell1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   33.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   3
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton btnPlay 
      Caption         =   "Play"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      ForeColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   360
      TabIndex        =   0
      Top             =   4560
      Width           =   3255
      Begin VB.PictureBox cell9 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   33.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2160
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   11
         Top             =   2040
         Width           =   855
      End
      Begin VB.PictureBox cell8 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   33.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1200
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   10
         Top             =   2040
         Width           =   855
      End
      Begin VB.PictureBox cell7 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   33.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   9
         Top             =   2040
         Width           =   855
      End
      Begin VB.PictureBox cell6 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   33.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2160
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   8
         Top             =   1200
         Width           =   855
      End
      Begin VB.PictureBox cell5 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   33.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1200
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   7
         Top             =   1200
         Width           =   855
      End
      Begin VB.PictureBox cell4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   33.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   6
         Top             =   1200
         Width           =   855
      End
      Begin VB.PictureBox cell3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   33.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2160
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
      Begin VB.PictureBox cell2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   33.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1200
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Label lblTime 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   3480
      TabIndex        =   13
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   $"FinalGame.frx":0000
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4095
   End
   Begin VB.Image Image1 
      Height          =   7680
      Left            =   4320
      Picture         =   "FinalGame.frx":013B
      Stretch         =   -1  'True
      Top             =   120
      Width           =   6240
   End
End
Attribute VB_Name = "FinalGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim topLeftCell, topMiddleCell, topRightCell As Integer
Dim midLeftCell, centerCell, midRightCell As Integer
Dim bottomLeftCell, bottomMiddleCell, bottomRightCell As Integer

Dim userTurn, gameOver As Boolean
Dim taken(1 To 9) As Boolean
Dim i As Integer

Private Sub btnHelp_Click()
MsgBox "The object of Tic Tac Toe is to get three in a row. You play on a three by three game board. The first player is known as X and the second is O. Players alternate placing Xs and Os on the game board until either oppent has three in a row or all nine squares are filled. X always goes first, and in the event that no one has three in a row, the stalemate is called a cat game."
End Sub

Private Sub btnMenuClose_Click()
    FrameMenu.Visible = False
    btnSave.Visible = False
    btnLoad.Visible = False
    btnMainMenu.Visible = False
    btnExit.Visible = False
    btnMenuClose.Visible = False
    btnMenuOpen.Visible = True

End Sub

Private Sub btnMenuOpen_Click()

    FrameMenu.Visible = True
    btnSave.Visible = True
    btnLoad.Visible = True
    btnMainMenu.Visible = True
    btnExit.Visible = True
    btnMenuOpen.Visible = False
    btnMenuClose.Visible = True
End Sub
Private Sub btnExit_Click()
    FormExit.Show
End Sub

Private Sub btnLoad_Click()
    Open App.Path & "\Saves\Save.txt" For Input As #4
        Input #4, stage
        If stage = 4 Then
        Form4Game1.Show
        FinalGame.Hide
        MsgBox "Loading complete."
        ElseIf stage = 5 Then
        Form5Game2.Show
        FinalGame.Hide
        MsgBox "Loading complete."
        ElseIf stage = 8 Then
        FinalGame.Show
        MsgBox "Loading complete."
        End If
    
    Close #4
    
End Sub

Private Sub btnMainMenu_Click()
    gotoMain8.Show
End Sub

Private Sub btnNext_Click()
    FinalGame.Hide
    Form9.Show
End Sub

Private Sub btnSave_Click()
    Open App.Path & "\Saves\Save.txt" For Input As #1
    Input #1, stage
    Close #1
    
    If stage = 0 Then
    Open App.Path & "\Saves\Save.txt" For Output As #2
    stage = 8
    Write #2, stage
    MsgBox "Game saved."
    Else: FormSaves8.Show
    End If
    
    Close #2
End Sub



Private Sub btnPlay_Click()

    cell1.Cls
    cell2.Cls
    cell3.Cls
    cell4.Cls
    cell5.Cls
    cell6.Cls
    cell7.Cls
    cell8.Cls
    cell9.Cls
    cell1.Enabled = True
    cell2.Enabled = True
    cell3.Enabled = True
    cell4.Enabled = True
    cell5.Enabled = True
    cell6.Enabled = True
    cell7.Enabled = True
    cell8.Enabled = True
    cell9.Enabled = True
    topLeftCell = 0
    topMiddleCell = 0
    topRightCell = 0
    midLeftCell = 0
    centerCell = 0
    midRightCell = 0
    bottomLeftCell = 0
    bottomMiddleCell = 0
    bottomRightCell = 0
    userTurn = True
    gameOver = False


    For i = 1 To 9
    taken(i) = False
    Next i


End Sub

Private Sub cell1_Click()
    Picture1.Print "gameOver is: "; gameOver
    If Not gameOver Then
    If topLeftCell = 0 Then
    Picture1.Print "I reach if if "
        taken(1) = True

        If userTurn Then
        cell1.Print " X"
        topLeftCell = 1
        cell1.Enabled = False
        userTurn = False
        If Not TTT Then
    Call complay
    End If
    Else
    cell1.Print " O"
    userTurn = True
    topLeftCell = 2
    TTT
    End If
     End If
     End If
End Sub

Private Sub cell2_Click()
    If Not gameOver Then
    If topMiddleCell = 0 Then
     Picture1.Print "I reach if if "
        taken(2) = True

         Picture1.Print "User Turn for cell 2 is "; userTurn
         If userTurn Then
        cell2.Print " X"
        topMiddleCell = 1
        cell2.Enabled = False
         userTurn = False
         If Not TTT Then
 Call complay
 End If
   Else
   cell2.Print " O"
   userTurn = True
   topMiddleCell = 2
   TTT
   End If
   End If
   End If
End Sub

Private Sub cell3_Click()
       If Not gameOver Then
    If topRightCell = 0 Then
     Picture1.Print "I reach if if "
        taken(3) = True
 
         If userTurn Then
        cell3.Print " X"
        topRightCell = 1
        cell3.Enabled = False
         userTurn = False
         If Not TTT Then
 Call complay
 End If
   Else
   cell3.Print " O"
   userTurn = True
   topRightCell = 2
   TTT
   End If
   End If
   End If
End Sub

Private Sub cell4_Click()
    If Not gameOver Then
    If midLeftCell = 0 Then
     Picture1.Print "I reach if if "
        taken(4) = True

         If userTurn Then
        cell4.Print " X"
        midLeftCell = 1
        cell4.Enabled = False
         userTurn = False
         If Not TTT Then
 Call complay
 End If
   Else
   cell4.Print " O"
   userTurn = True
   midLeftCell = 2
   TTT
   End If
   End If
   End If
End Sub

Private Sub cell5_Click()
    If Not gameOver Then
    If centerCell = 0 Then
     Picture1.Print "I reach if if "
        taken(5) = True

         If userTurn Then
        cell5.Print " X"
        centerCell = 1
        cell5.Enabled = False
         userTurn = False
         If Not TTT Then
 Call complay
 End If
   Else
   cell5.Print " O"
   userTurn = True
   centerCell = 2
   TTT
   End If
   End If
   End If
End Sub

Private Sub cell6_Click()
   If Not gameOver Then
    If midRightCell = 0 Then
     Picture1.Print "I reach if if "
        taken(6) = True

         If userTurn Then
        cell6.Print " X"
        midRightCell = 1
        cell6.Enabled = False
         userTurn = False
         If Not TTT Then
 Call complay
 End If
   Else
   cell6.Print " O"
   userTurn = True
   midRightCell = 2
   TTT
   End If
   End If
   End If
End Sub

    Private Sub cell7_Click()
      If Not gameOver Then
        If bottomLeftCell = 0 Then
         Picture1.Print "I reach if if "
            taken(7) = True

             If userTurn Then
            cell7.Print " X"
            bottomLeftCell = 1
            cell7.Enabled = False
             userTurn = False
             If Not TTT Then
     Call complay
       End If
       Else
       cell7.Print " O"
       userTurn = True
       bottomLeftCell = 2
       TTT
       End If
       End If
       End If
End Sub

Private Sub cell8_Click()
  If Not gameOver Then
    If bottomMiddleCell = 0 Then
     Picture1.Print "I reach if if "
        taken(8) = True

         If userTurn Then
        cell8.Print " X"
        bottomMiddleCell = 1
        cell8.Enabled = False
         userTurn = False
         If Not TTT Then
         
 Call complay
    End If
   Else
   cell8.Print " O"
   userTurn = True
   bottomMiddleCell = 2
   TTT
   End If
   End If
   End If
End Sub

Private Sub cell9_Click()
 If Not gameOver Then
    If bottomRightCell = 0 Then
     Picture1.Print "I reach if if "
        taken(9) = True

         If userTurn Then
        cell9.Print " X"
        bottomRightCell = 1
        cell9.Enabled = False
         userTurn = False
         If Not TTT Then
 Call complay
 End If
   Else
   cell9.Print " O"
   userTurn = True
   bottomRightCell = 2
   TTT
   End If
   End If
   End If
End Sub

Private Sub Form_Load()
    Randomize
    topLeftCell = 0
    topMiddleCell = 0
    topRightCell = 0
    midLeftCell = 0
    centerCell = 0
    midRightCell = 0
    bottomLeftCell = 0
    bottomMiddleCell = 0
    bottomRightCell = 0
    userTurn = True
    gameOver = True


    For i = 1 To 9
    taken(i) = False
    Next i

End Sub

Private Function cellsFull() As Boolean
  If topLeftCell = 0 _
    Or (topMiddleCell = 0 _
    Or topRightCell = 0 _
    Or midLeftCell = 0 _
    Or centerCell = 0 _
    Or midRightCell = 0 _
    Or bottomLeftCell = 0 _
    Or bottomMiddleCell = 0 _
    Or bottomRightCell = 0) Then
    cellsFull = False
    Else
    cellsFull = True
        MsgBox "Draw Game! Play another one.", , "Draw"
  End If
End Function


Private Sub complay()
 Picture1.Print "I reach beginning"
  Select Case compWinPosition
    Case 1
      Call cell1_Click
      Picture1.Print "I reach one"
    Case 2
      Call cell2_Click
    Picture1.Print "I reach two"
    Case 3
      Call cell3_Click
       Picture1.Print "I reach three"
    Case 4
      Call cell4_Click
    Case 5
      Call cell5_Click
    Case 6
      Call cell6_Click
    Case 7
      Call cell7_Click
    Case 8
      Call cell8_Click
    Case 9
      Call cell9_Click
    Case 0
     Picture1.Print "I reach finally"
      Dim choice As Integer
      Do
        choice = Int(Rnd * 9) + 1
        Picture1.Print "choice is"; choice
      Loop Until Not taken(choice)
  
      Select Case choice
        Case 1
         Picture1.Print "I reach case 1"
          Call cell1_Click
        Case 2
        Picture1.Print "I reach case 2"
          Call cell2_Click
        Case 3
        Picture1.Print "I reach case 3"
          Call cell3_Click
        Case 4
        Picture1.Print "I reach case 4"
          Call cell4_Click
        Case 5
        Picture1.Print "I reach case 5"
          Call cell5_Click
        Case 6
        Picture1.Print "I reach case 6"
          Call cell6_Click
        Case 7
        Picture1.Print "I reach case 7"
          Call cell7_Click
        Case 8
        Picture1.Print "I reach case 8"
          Call cell8_Click
        Case 9
        Picture1.Print "I reach case 9"
          Call cell9_Click
      End Select
  End Select
  
End Sub


Private Function compWinPosition() As Integer
  If topLeftCell = 0 And (topMiddleCell = 2 _
    And topRightCell = 2 _
    Or centerCell = 2 _
    And bottomRightCell = 2 _
    Or midLeftCell = 2 _
    And bottomLeftCell = 2) Then
      compWinPosition = 1
  ElseIf topMiddleCell = 0 _
    And (topLeftCell = 2 _
    And topRightCell = 2 _
    Or centerCell = 2 _
    And bottomMiddleCell = 2) Then
      compWinPosition = 2
  ElseIf topRightCell = 0 _
    And (topLeftCell = 2 _
    And topMiddleCell = 2 _
    Or bottomLeftCell = 2 _
    And centerCell = 2 _
    Or midRightCell = 2 _
    And bottomRightCell = 2) Then
      compWinPosition = 3
  ElseIf midLeftCell = 0 _
    And (topLeftCell = 2 _
    And bottomLeftCell = 2 _
    Or centerCell = 2 _
    And midRightCell = 2) Then
      compWinPosition = 4
  ElseIf centerCell = 0 _
    And (topLeftCell = 2 _
    And bottomRightCell = 2 _
    Or bottomLeftCell = 2 _
    And topRightCell = 2 _
    Or topMiddleCell = 2 _
    And bottomMiddleCell = 2 _
    Or midLeftCell = 2 _
    And midRightCell = 2) Then
      compWinPosition = 5
  ElseIf midRightCell = 0 _
    And (topRightCell = 2 _
    And bottomRightCell = 2 _
    Or midLeftCell = 2 _
    And centerCell = 2) Then
      compWinPosition = 6
  ElseIf bottomLeftCell = 0 _
    And (topLeftCell = 2 _
    And midLeftCell = 2 _
    Or centerCell = 2 _
    And topRightCell = 2 _
    Or bottomMiddleCell = 2 _
    And bottomRightCell = 2) Then
      compWinPosition = 7
  ElseIf bottomMiddleCell = 0 _
    And (topMiddleCell = 2 _
    And centerCellCell = 2 _
    Or bottomLeftCell = 2 _
    And bottomRightCell = 2) Then
      compWinPosition = 8
  ElseIf bottomRightCell = 0 _
    And (topLeftCell = 2 _
    And centerCell = 2 _
    Or topRightCell = 2 _
    And midRightCell = 2 _
    Or bottomLeftCell = 2 _
    And bottomMiddleCell = 2) Then
      compWinPosition = 9
  Else
    compWinPosition = xWinPosition
  End If
End Function

Private Function xWinPosition() As Integer
  If topLeftCell = 0 _
    And (topMiddleCell = 1 _
    And topRightCel = 1 _
    Or centerCell = 1 _
    And bottomRightCell = 1 _
    Or midLeftCell = 1 _
    And bottomLeftCell = 1) Then
      xWinPosition = 1
  ElseIf topMiddleCell = 0 _
    And (topLeftCell = 1 _
    And topRightCell = 1 _
    Or centerCell = 1 _
    And bottomMiddleCell = 1) Then
      xWinPosition = 2
  ElseIf topRightCell = 0 _
    And (topLeftCell = 1 _
    And topMiddleCell = 1 _
    Or bottomLeftCell = 1 _
    And centerCell = 1 _
    Or midRightCell = 1 _
    And bottomRightCell = 1) Then
      xWinPosition = 3
  ElseIf midLeftCell = 0 _
    And (topLeftCell = 1 _
    And bottomLeftCell = 1 _
    Or centerCell = 1 _
    And midRightCell = 1) Then
      xWinPosition = 4
  ElseIf centerCell = 0 _
    And (topLeftCell = 1 _
    And bottomRightCell = 1 _
    Or bottomLeftCell = 1 _
    And topRightCell = 1 _
    Or topMiddleCell = 1 _
    And bottomMiddleCell = 1 _
    Or midLeftCell = 1 _
    And midRightCell = 1) Then
      xWinPosition = 5
  ElseIf midRightCell = 0 _
    And (topRightCell = 1 _
    And bottomRightCell = 1 _
    Or midLeftCell = 1 _
    And centerCell = 1) Then
      xWinPosition = 6
  ElseIf bottomLeftCell = 0 _
    And (topLeftCell = 1 _
    And midLeftCell = 1 _
    Or centerCell = 1 _
    And topRightCell = 1 _
    Or bottomMiddleCell = 1 _
    And bottomRightCell = 1) Then
      xWinPosition = 7
  ElseIf bottomMiddleCell = 0 _
    And (topMiddleCell = 1 _
    And centerCell = 1 _
    Or bottomLeftCell = 1 _
    And bottomRightCell = 1) Then
      xWinPosition = 8
  ElseIf bottomRightCell = 0 _
    And (topLeftCell = 1 _
    And centerCell = 1 _
    Or topRightCell = 1 _
    And midRightCell = 1 _
    Or bottomLeftCell = 1 _
    And bottomMiddleCell = 1) Then
      xWinPosition = 9
  Else
    xWinPosition = 0
  End If
End Function


Private Function TTT() As Boolean
  If topLeftCell = 1 _
     And topMiddleCell = 1 _
     And topRightCell = 1 _
     Or midLeftCell = 1 _
     And centerCell = 1 _
     And midRightCell = 1 _
     Or bottomLeftCell = 1 _
     And bottomMiddleCell = 1 _
     And bottomRightCell = 1 _
     Or topLeftCell = 1 _
     And midLeftCell = 1 _
     And bottomLeftCell = 1 _
     Or topMiddleCell = 1 _
     And centerCell = 1 _
     And bottomMiddleCell = 1 _
     Or topRightCell = 1 _
     And midRightCell = 1 _
     And bottomRightCell = 1 _
     Or topLeftCell = 1 _
     And centerCell = 1 _
     And bottomRightCell = 1 _
     Or topRightCell = 1 _
     And centerCell = 1 _
     And bottomLeftCell = 1 Then
       gameOver = True
    
    MsgBox "You win! Now you can exit."
    FinalGame.Hide
    Form9.Show
    
  
  ElseIf topLeftCell = 2 _
     And topMiddleCell = 2 _
     And topRightCell = 2 _
     Or midLeftCell = 2 _
     And centerCell = 2 _
     And midRightCell = 2 _
     Or bottomLeftCell = 2 _
     And bottomMiddleCell = 2 _
     And bottomRightCell = 2 _
     Or topLeftCell = 2 _
     And midLeftCell = 2 _
     And bottomLeftCell = 2 _
     Or topMiddleCell = 2 _
     And centerCell = 2 _
     And bottomMiddleCell = 2 _
     Or topRightCell = 2 _
     And midRightCell = 2 _
     And bottomRightCell = 2 _
     Or topLeftCell = 2 _
     And centerCell = 2 _
     And bottomRightCell = 2 _
     Or topRightCell = 2 _
     And centerCell = 2 _
     And bottomLeftCell = 2 Then
       gameOver = True
    
    MsgBox "Sorry, please try again."
  
  ElseIf cellsFull Then
     gameOver = True
     

     
  End If
  TTT = gameOver

End Function


