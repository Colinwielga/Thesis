VERSION 5.00
Begin VB.Form frmQ2 
   BackColor       =   &H80000010&
   Caption         =   "PhotoMind™ "
   ClientHeight    =   10485
   ClientLeft      =   900
   ClientTop       =   615
   ClientWidth     =   13950
   LinkTopic       =   "Form1"
   ScaleHeight     =   10485
   ScaleWidth      =   13950
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   9840
      Width           =   1935
   End
   Begin VB.CommandButton cmdA 
      Caption         =   "Edward Curtis"
      Height          =   975
      Left            =   360
      TabIndex        =   6
      Top             =   8520
      Width           =   3015
   End
   Begin VB.CommandButton cmdB 
      Caption         =   "Ansel Adams"
      Height          =   975
      Left            =   3720
      TabIndex        =   5
      Top             =   8520
      Width           =   3015
   End
   Begin VB.CommandButton cmdC 
      Caption         =   "Dr. Harold Edgerton"
      Height          =   975
      Left            =   7080
      TabIndex        =   4
      Top             =   8520
      Width           =   3015
   End
   Begin VB.CommandButton cmdD 
      Caption         =   "Annie Leobovitz"
      Height          =   975
      Left            =   10440
      TabIndex        =   3
      Top             =   8520
      Width           =   3015
   End
   Begin VB.CommandButton cmdFF 
      Caption         =   "50 : 50"
      Height          =   495
      Left            =   4920
      TabIndex        =   2
      Top             =   9840
      Width           =   1935
   End
   Begin VB.CommandButton cmdComputer 
      Caption         =   "Ask Computer"
      Height          =   495
      Left            =   6960
      TabIndex        =   1
      Top             =   9840
      Width           =   1935
   End
   Begin VB.CommandButton cmdGoogle 
      Caption         =   "Google It!"
      Height          =   495
      Left            =   9120
      TabIndex        =   0
      Top             =   9840
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Left            =   12720
      Top             =   9960
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   13920
      Y1              =   9720
      Y2              =   9720
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "40"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   13320
      TabIndex        =   10
      Top             =   9960
      Width           =   375
   End
   Begin VB.Label lblName 
      BackColor       =   &H80000010&
      Caption         =   "PhotoMind™   Question-2"
      BeginProperty Font 
         Name            =   "Castellar"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   12255
   End
   Begin VB.Label lblQuestion 
      BackColor       =   &H80000010&
      Caption         =   "Who most likely took this from the following artists?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   8
      Top             =   7680
      Width           =   13095
   End
   Begin VB.Image Image1 
      Height          =   7770
      Left            =   600
      Picture         =   "Form3_Q2.frx":0000
      Top             =   120
      Width           =   11190
   End
End
Attribute VB_Name = "frmQ2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Question forms have 4 answer choices that record which answer was selected in the publicly dimmed array. If the correct
'answer is selected 1 is added to the correct answer counter. When an answer is selected it stops the countdown and advances
'to the next form. There are also 3 hint options. The Fifty Fifty option hides two wrong answers choices and eliminates
'the possibility of clicking on the Fifty Fifty button on the following forms. Computer and Google hint buttons display
'message boxes indicating the more likely answer choice and eliminate the possibility of clicking those hint buttons on
'the following forms.  There is also a countdown timer that gives the user 40 seconds to answer the questions.

Private Sub cmdA_Click()
'incorrect answer -> save users answer in the Answers array; go to the next form
    CTR = CTR + 1
    Answers(CTR) = "A"
    Timer1.Enabled = False
    frmQ3.Label1.Caption = "40"
    frmQ3.Show
    frmQ2.Hide
End Sub

Private Sub cmdB_Click()
'Correct answer -> add 1 to correct answers (CTR); save answer in the Answers array go to the next form
    CTR = CTR + 1
    Answers(CTR) = "B"
    Timer1.Enabled = False
    frmQ3.Label1.Caption = "40"
    frmQ3.Show
    frmQ2.Hide
End Sub

Private Sub cmdC_Click()
'incorrect answer -> save users answer in the Answers array; go to the next form
    CTR = CTR + 1
    Answers(CTR) = "C"
    Right = Right + 1
    Timer1.Enabled = False
    frmQ3.Label1.Caption = "40"
    frmQ3.Show
    frmQ2.Hide
End Sub

Private Sub cmdComputer_Click()
'adds 1 to computer counter and checks if has been used before. If has then gives an error, if not gives a hint. Hides the button on all other forms
Computer = Computer + 1

If Computer = 1 Then

    MsgBox "There is a " & FormatPercent(0.76, 0) & " possibility that the answer is " & cmdB.Caption & " or " & cmdC.Caption & ".", , "CPU Output"

    cmdComputer.Enabled = False
    frmQ3.cmdComputer.Enabled = False
    frmQ4.cmdComputer.Enabled = False
    frmQ5.cmdComputer.Enabled = False
    frmQ6.cmdComputer.Enabled = False
Else
    MsgBox "You already used this option", , "Error"
    cmdComputer.Enabled = False
End If
End Sub

Private Sub cmdD_Click()
'incorrect answer -> save users answer in the Answers array; go to the next form
    CTR = CTR + 1
    Answers(CTR) = "D"
    Timer1.Enabled = False
    frmQ3.Label1.Caption = "40"
    frmQ3.Show
    frmQ2.Hide
End Sub

Private Sub cmdFF_Click()
'subtracts 1 from the FiftyFifty counter; checks if has been used; if yes then gives an error, if not then gives a hint. Hides the button on all other forms

FF = FF + 1

Select Case FF
    Case Is = 1
        cmdA.Visible = False
        cmdB.Visible = True
        cmdC.Visible = True
        cmdD.Visible = False
        cmdFF.Enabled = False
        frmQ3.cmdFF.Enabled = False
        frmQ4.cmdFF.Enabled = False
        frmQ5.cmdFF.Enabled = False
        frmQ6.cmdFF.Enabled = False
    Case Else
        MsgBox "This should not show up!", , "Error"
        cmdFF.Enabled = False
End Select

End Sub

Private Sub cmdGoogle_Click()
'adds 1 to google counter and checks if has been used before. If has then gives an error, if not gives a hint. Hides the button on all other forms

MsgBox "Google reports that " & Chr(13) & Chr(13) & cmdB.Caption & Chr(13) & cmdC.Caption & Chr(13) & Chr(13) & "had more clicks than " & Chr(13) & Chr(13) & cmdD.Caption & Chr(13) & cmdA.Caption, , "Google Search"

cmdGoogle.Enabled = False
frmQ2.cmdGoogle.Enabled = False
frmQ3.cmdGoogle.Enabled = False
frmQ4.cmdGoogle.Enabled = False
frmQ5.cmdGoogle.Enabled = False
frmQ6.cmdGoogle.Enabled = False


End Sub

Private Sub cmdQuit_Click()
 End
End Sub


Private Sub Form_Load()
'initializes timer (1000 = 1 sec & 40 = how many seconds there are)
'code for timer retrieved from http://www.daniweb.com/forums/thread14312.html
     Timer1.Enabled = True
     Timer1.Interval = 1000
     Label1.Caption = "40"
          
End Sub

Private Sub Timer1_Timer()
'code for timer retrieved from http://www.daniweb.com/forums/thread14312.html

     If Label1.Caption = 0 Then
        Timer1.Enabled = False
        MsgBox "Your 40 seconds are up. Click OK when ready for the next question.", , "Times up!"
     Else
        Label1.Caption = Label1.Caption - 1
     End If
     
End Sub
