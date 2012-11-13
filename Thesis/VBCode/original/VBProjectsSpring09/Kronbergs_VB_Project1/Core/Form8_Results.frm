VERSION 5.00
Begin VB.Form frmResults 
   BackColor       =   &H80000010&
   Caption         =   "PhotoMind™ "
   ClientHeight    =   10380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13935
   LinkTopic       =   "Form1"
   ScaleHeight     =   10380
   ScaleWidth      =   13935
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Look up a question"
      Enabled         =   0   'False
      Height          =   735
      Left            =   11160
      TabIndex        =   8
      Top             =   4440
      Width           =   1935
   End
   Begin VB.CommandButton cmdClearPicResults 
      Caption         =   "Clear displayed "
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   8640
      Width           =   9495
   End
   Begin VB.CommandButton cmdCloseCD 
      Caption         =   "Put away your new CUP HOLDER"
      Height          =   855
      Left            =   12240
      TabIndex        =   6
      Top             =   8640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdOpenCD 
      Caption         =   "Click here to receive your CUP HOLDER"
      Height          =   975
      Left            =   12000
      TabIndex        =   5
      Top             =   7440
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdResults 
      Cancel          =   -1  'True
      Caption         =   "Show Results"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   10440
      TabIndex        =   4
      Top             =   2040
      Width           =   2775
   End
   Begin VB.CommandButton cmdAgain 
      Caption         =   "Try Again"
      Height          =   495
      Left            =   11040
      TabIndex        =   3
      Top             =   9720
      Width           =   2775
   End
   Begin VB.PictureBox picResults 
      Height          =   7455
      Left            =   240
      ScaleHeight     =   7395
      ScaleWidth      =   9435
      TabIndex        =   2
      Top             =   1200
      Width           =   9495
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H80000001&
      Caption         =   "Quit"
      Height          =   495
      Left            =   120
      MaskColor       =   &H000000FF&
      TabIndex        =   0
      Top             =   9720
      UseMaskColor    =   -1  'True
      Width           =   1935
   End
   Begin VB.Label lblCongratulations 
      BackColor       =   &H80000010&
      Caption         =   "  Congratulations! You have won a cup holder!!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   615
      Left            =   10800
      TabIndex        =   9
      Top             =   6600
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label lblName 
      BackColor       =   &H80000010&
      Caption         =   "PhotoMind™ RESULTS"
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
      Left            =   600
      TabIndex        =   1
      Top             =   0
      Width           =   10095
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   13920
      Y1              =   9600
      Y2              =   9600
   End
End
Attribute VB_Name = "frmResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'In the results form I have a picture box where to print the results, a clear picture box button, quit button and a show results button.
'The latter loads a file with the correct answers and checks if they match with the users answers, then prints them. Wrong answers are indicated.
'It also prints the number of correct answers according to the running sum of the correct answers. It also enables some earlier hidden buttons.
'There is a look up answers button which gets a question number from the user via input box and then goes and fetches the question along with
'the correct answer and prints them. The "prize" buttons open and close the CD Rom. User can also try the game again, where he is brought to the intro form.

Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Option Explicit
Dim RightAnswers(1 To 6) As String, RCTR As Integer


Private Sub cmdAgain_Click()
'brings back to the intro form
frmIntro.Show
frmResults.Hide

End Sub

Private Sub cmdCheck_Click()
'checks the number user entered and prints out the question with answer
Dim Check As String
Check = InputBox("Enter the questions number (1-6) you want to check", "Question Lookup")

picResults.Print
picResults.Print "_____________________________________"
Select Case Check
    Case 6
        picResults.Print frmQ6.lblQuestion
        picResults.Print "The correct answer is: " & frmQ6.cmdA.Caption
    Case 5
        picResults.Print frmQ5.lblQuestion
        picResults.Print "The correct answer is: " & frmQ5.cmdB.Caption

    Case 4
        picResults.Print frmQ4.lblQuestion
        picResults.Print "The correct answer is: " & frmQ4.cmdB.Caption
    Case 3
        picResults.Print frmQ3.lblQuestion
        picResults.Print "The correct answer is: " & frmQ3.cmdD.Caption
    Case 2
        picResults.Print frmQ2.lblQuestion
        picResults.Print "The correct answer is: " & frmQ2.cmdC.Caption
    Case 1
        picResults.Print frmQ1.lblQuestion
        picResults.Print "The correct answer is: " & frmQ1.cmdA.Caption
    Case Else
        MsgBox "That is not a valid entry", , "Error"
End Select
    
End Sub

Private Sub cmdClearPicResults_Click()
'clear picResults
    picResults.Cls
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdResults_Click()
'Load the correct answers. Display the answers user chose by comparing them to the answers in the array. Prints out the number of correct answers user got.
Dim I As Integer
RCTR = 0

'load the correct answers
Open App.Path & "\Core\RightAnswers.txt" For Input As #1

Do While Not EOF(1)
    RCTR = RCTR + 1
    Input #1, RightAnswers(RCTR)
Loop
Close

picResults.Print PlayerName & ", you got " & Right & " answers right"
picResults.Print "________________________________________"
picResults.Print Tab(14); "You"; Tab(27); "Correct"

'Shows which answers user got wrong
For I = 1 To CTR
        picResults.Print Tab(1); "Question " & I; Tab(15); Answers(I); Tab(30); RightAnswers(I);
    If Answers(I) <> RightAnswers(I) Then
        picResults.Print Tab(35); "<--- Wrong answer"
    End If
Next I

'Enables the "prize" buttons
cmdCheck.Enabled = True
lblCongratulations.Visible = True
cmdOpenCD.Visible = True
cmdCloseCD.Visible = True
cmdCloseCD.Enabled = False

End Sub


Private Sub cmdOpenCD_Click()
'Code to open/close CD drive was taken from:
'http://www.devx.com/vb2themax/Tip/18552
'enables close CD Rom button
    mciSendString "Set CDAudio Door Open Wait", vbNullString, 0, 0
    cmdCloseCD.Enabled = True
End Sub

Private Sub cmdCloseCD_Click()
'close CD drive
    mciSendString "Set CDAudio Door Closed Wait", vbNullString, 0, 0
End Sub

