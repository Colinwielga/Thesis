VERSION 5.00
Begin VB.Form frmQuiz 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Quiz"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   12180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picOptions 
      Height          =   1455
      Left            =   2280
      Picture         =   "frmQuiz.frx":0000
      ScaleHeight     =   1395
      ScaleWidth      =   7155
      TabIndex        =   13
      Top             =   5040
      Width           =   7215
   End
   Begin VB.PictureBox picQueryYakshi 
      Height          =   5055
      Left            =   4200
      Picture         =   "frmQuiz.frx":568F2
      ScaleHeight     =   4995
      ScaleWidth      =   3915
      TabIndex        =   12
      Top             =   0
      Width           =   3975
   End
   Begin VB.PictureBox picQueryDate 
      Height          =   495
      Left            =   7320
      Picture         =   "frmQuiz.frx":5B714
      ScaleHeight     =   435
      ScaleWidth      =   3915
      TabIndex        =   11
      Top             =   3120
      Width           =   3975
   End
   Begin VB.PictureBox picdatewrong4 
      Height          =   495
      Left            =   9840
      Picture         =   "frmQuiz.frx":6242A
      ScaleHeight     =   435
      ScaleWidth      =   1275
      TabIndex        =   10
      Top             =   5160
      Width           =   1335
   End
   Begin VB.PictureBox picdateright 
      Height          =   615
      Left            =   7680
      Picture         =   "frmQuiz.frx":63F00
      ScaleHeight     =   555
      ScaleWidth      =   1275
      TabIndex        =   9
      Top             =   5160
      Width           =   1335
   End
   Begin VB.PictureBox picdatewrong2 
      Height          =   615
      Left            =   9840
      Picture         =   "frmQuiz.frx":660C6
      ScaleHeight     =   555
      ScaleWidth      =   1155
      TabIndex        =   8
      Top             =   3960
      Width           =   1215
   End
   Begin VB.PictureBox picdatewrong 
      Height          =   615
      Left            =   7680
      Picture         =   "frmQuiz.frx":67F80
      ScaleHeight     =   555
      ScaleWidth      =   1275
      TabIndex        =   7
      Top             =   3960
      Width           =   1335
   End
   Begin VB.PictureBox piclc 
      Height          =   5535
      Left            =   720
      Picture         =   "frmQuiz.frx":69A72
      ScaleHeight     =   5475
      ScaleWidth      =   3915
      TabIndex        =   6
      Top             =   360
      Width           =   3975
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Who is in the picture?"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5040
      TabIndex        =   5
      Top             =   6600
      Width           =   1815
   End
   Begin VB.PictureBox picQuery 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   480
      ScaleHeight     =   1455
      ScaleWidth      =   11295
      TabIndex        =   4
      Top             =   6480
      Width           =   11295
   End
   Begin VB.PictureBox picshs 
      Height          =   2055
      Left            =   600
      Picture         =   "frmQuiz.frx":6E2E0
      ScaleHeight     =   1995
      ScaleWidth      =   1995
      TabIndex        =   3
      Top             =   480
      Width           =   2055
   End
   Begin VB.PictureBox picgm 
      Height          =   2895
      Left            =   600
      Picture         =   "frmQuiz.frx":7B04A
      ScaleHeight     =   2835
      ScaleWidth      =   2235
      TabIndex        =   2
      Top             =   2880
      Width           =   2295
   End
   Begin VB.PictureBox picgb 
      Height          =   2055
      Left            =   2880
      Picture         =   "frmQuiz.frx":7F225
      ScaleHeight     =   1995
      ScaleWidth      =   3555
      TabIndex        =   1
      Top             =   600
      Width           =   3615
   End
   Begin VB.PictureBox pickanish 
      Height          =   3255
      Left            =   3360
      Picture         =   "frmQuiz.frx":81592
      ScaleHeight     =   3195
      ScaleWidth      =   2115
      TabIndex        =   0
      Top             =   2760
      Width           =   2175
   End
End
Attribute VB_Name = "frmQuiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'the division between correct and incorrect answers for these three quiz questions are kept track of
Dim try As Integer, try2 As Integer, try3 As Integer


Private Sub cmdNext_Click() 'for the third question on the quiz, an input box is used to gather the users answer; the user is given two tries
Dim Usr_Input As String

Usr_Input = InputBox("Enter your selection, please", "Answer")

If Usr_Input = "B" Then
    Correct = Correct + 1
    MsgBox "Well done.", , "Correct."
    Review(6) = False
    Review(9) = False
    frmQuiz.Visible = False
    frmResults.Visible = True
Else
    If try3 < 2 Then
        try3 = try3 + 1
        MsgBox "Sorry, try again please", , "Incorrect"
    End If
    If try3 >= 2 Then
    cmdNext.Visible = False
        Incorrect = Incorrect + 1
        Review(6) = True
        Review(9) = True
        
        frmQuiz.Visible = False
        frmResults.Visible = True
        
    End If
End If

End Sub

Private Sub Form_Activate() 'when the form is opened up, the first question appears while the other objects are hidden
    picQuery.Print "For which of these objects "
    picQuery.Print "is the medium steatite?"
    pickanish.Visible = True
    picgm.Visible = True
    picgb.Visible = True
    picshs.Visible = True
    cmdNext.Visible = False
    piclc.Visible = False
    picdatewrong.Visible = False
    picdatewrong2.Visible = False
    picdatewrong4.Visible = False
    picdateright.Visible = False
    picQueryDate.Visible = False
    picOptions.Visible = False
    picQueryYakshi.Visible = False
    Correct = 0
    Incorrect = 0
End Sub

Private Sub picdateright_Click() 'for the second question, whichever text the user clicks is submitted as the answer; when either the correct answer is given or the user runs out of tries, the next question is displayed
    MsgBox "Well done.", , "Correct"
    Correct = Correct + 1
    
    pickanish.Visible = False
    picgm.Visible = False
    picgb.Visible = False
    picshs.Visible = False
    picQuery.Visible = False
    piclc.Visible = False
    picdatewrong.Visible = False
    picdatewrong2.Visible = False
    picdatewrong4.Visible = False
    picdateright.Visible = False
    picQueryDate.Visible = False
    picQueryYakshi.Visible = True
    picOptions.Visible = True
    cmdNext.Visible = True
    
    Review(5) = False

End Sub



Private Sub picdatewrong_Click() 'same as above
    If try2 < 2 Then
        MsgBox "Sorry, try again please.", , "Incorrect"
        try2 = try2 + 1
    End If
    
    If try2 = 2 Then
        Incorrect = Incorrect + 1
        Review(5) = True
        
        pickanish.Visible = False
        picgm.Visible = False
        picgb.Visible = False
        picshs.Visible = False
        picQuery.Visible = False
        piclc.Visible = False
        picdatewrong.Visible = False
        picdatewrong2.Visible = False
        picdatewrong4.Visible = False
        picdateright.Visible = False
        picQueryDate.Visible = False
        cmdNext.Visible = True
        picQueryYakshi.Visible = True
        picOptions.Visible = True
        
    End If
End Sub

Private Sub picdatewrong2_Click() 'same as above
If try2 < 2 Then
        MsgBox "Sorry, try again please.", , "Incorrect"
        try2 = try2 + 1
    End If
    
    If try2 = 2 Then
        Incorrect = Incorrect + 1
        Review(5) = True
        
        pickanish.Visible = False
        picgm.Visible = False
        picgb.Visible = False
        picshs.Visible = False
        picQuery.Visible = False
        piclc.Visible = False
        picdatewrong.Visible = False
        picdatewrong2.Visible = False
        picdatewrong4.Visible = False
        picdateright.Visible = False
        picQueryDate.Visible = False
        picQueryYakshi.Visible = True
        picOptions.Visible = True
        cmdNext.Visible = True
        
    End If
End Sub

Private Sub picdatewrong4_Click() 'same as above
If try2 < 2 Then
        MsgBox "Sorry, try again please.", , "Incorrect"
        try2 = try2 + 1
    End If
    
    If try2 = 2 Then
        Incorrect = Incorrect + 1
        Review(5) = True

        pickanish.Visible = False
        picgm.Visible = False
        picgb.Visible = False
        picshs.Visible = False
        picQuery.Visible = False
        piclc.Visible = False
        picdatewrong.Visible = False
        picdatewrong2.Visible = False
        picdatewrong4.Visible = False
        picdateright.Visible = False
        picQueryDate.Visible = False
        picQueryYakshi.Visible = True
        picOptions.Visible = True
        cmdNext.Visible = True
        
    End If
End Sub

Private Sub picgb_Click() 'for the second question, the picture that the user clicks is submitted as the answer; when either the correct answer is given or the user runs out of tries, the next question is displayed
    If try < 2 Then
        MsgBox "Sorry, try again please.", , "Incorrect"
        try = try + 1
    End If
    
    If try = 2 Then
        Incorrect = Incorrect + 1
        Review(1) = True
        Review(2) = True
        Review(10) = True
        Review(13) = True
        
        pickanish.Visible = False
        picgm.Visible = False
        picgb.Visible = False
        picshs.Visible = False
        cmdNext.Visible = False
        picQuery.Visible = False
        picdatewrong.Visible = True
        picdatewrong2.Visible = True
        picdatewrong4.Visible = True
        picdateright.Visible = True
        picQueryDate.Visible = True
        piclc.Visible = True
        picOptions.Visible = False
        picQueryYakshi.Visible = False

    End If
End Sub

Private Sub picgm_Click() 'same as above
    If try < 2 Then
        MsgBox "Sorry, try again please.", , "Incorrect"
        try = try + 1
    End If
    
    If try = 2 Then
        Incorrect = Incorrect + 1
        Review(1) = True
        Review(2) = True
        Review(10) = True
        Review(13) = True
        
        pickanish.Visible = False
        picgm.Visible = False
        picgb.Visible = False
        picshs.Visible = False
        cmdNext.Visible = False
        picQuery.Visible = False
        picdatewrong.Visible = True
        picdatewrong2.Visible = True
        picdatewrong4.Visible = True
        picdateright.Visible = True
        picQueryDate.Visible = True
        piclc.Visible = True
        picOptions.Visible = False
        picQueryYakshi.Visible = False
    End If
End Sub

Private Sub pickanish_Click() 'same as above
    If try < 2 Then
        MsgBox "Sorry, try again please.", , "Incorrect"
        try = try + 1
    End If
    
    If try = 2 Then
        Incorrect = Incorrect + 1
        Review(1) = True
        Review(2) = True
        Review(10) = True
        Review(13) = True
        
        pickanish.Visible = False
        picgm.Visible = False
        picgb.Visible = False
        picshs.Visible = False
        cmdNext.Visible = False
        picQuery.Visible = False
        picdatewrong.Visible = True
        picdatewrong2.Visible = True
        picdatewrong4.Visible = True
        picdateright.Visible = True
        picQueryDate.Visible = True
        piclc.Visible = True
        picOptions.Visible = False
        picQueryYakshi.Visible = False
    End If
End Sub

Private Sub picshs_Click() 'same as above
    MsgBox "Well done.", , "Correct"
    Correct = Correct + 1
    Review(1) = False
    Review(2) = False
    Review(10) = False
    Review(13) = False
    
    pickanish.Visible = False
    picgm.Visible = False
    picgb.Visible = False
    picshs.Visible = False
    cmdNext.Visible = False
    picQuery.Visible = False
    picdatewrong.Visible = True
    picdatewrong2.Visible = True
    picdatewrong4.Visible = True
    picdateright.Visible = True
    picQueryDate.Visible = True
    piclc.Visible = True
    picOptions.Visible = False
    picQueryYakshi.Visible = False
End Sub
