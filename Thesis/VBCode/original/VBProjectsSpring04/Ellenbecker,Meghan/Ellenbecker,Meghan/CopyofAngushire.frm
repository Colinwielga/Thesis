VERSION 5.00
Begin VB.Form Angushire 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Form2"
   ClientHeight    =   4860
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6810
   LinkTopic       =   "Form2"
   ScaleHeight     =   4860
   ScaleWidth      =   6810
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   2055
      Left            =   4080
      Picture         =   "Angushire.frx":0000
      ScaleHeight     =   1995
      ScaleWidth      =   1755
      TabIndex        =   6
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   855
      Left            =   5280
      TabIndex        =   5
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Return to Starting Page"
      Height          =   975
      Left            =   4920
      TabIndex        =   4
      Top             =   2520
      Width           =   1455
   End
   Begin VB.PictureBox picResults 
      Height          =   1335
      Left            =   240
      ScaleHeight     =   1275
      ScaleWidth      =   3675
      TabIndex        =   3
      Top             =   2880
      Width           =   3735
   End
   Begin VB.CommandButton cmdWomen 
      Caption         =   "Women"
      Height          =   975
      Left            =   2040
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdMen 
      Caption         =   "Men"
      Height          =   975
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "ANGUSHIRE"
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Please select the tees that you played from by clicking on the appropriate button below."
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   3135
   End
End
Attribute VB_Name = "Angushire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim HandicapDifferential As Single
Dim Handicap As Single
Dim Score As Single


Private Sub cmdMen_Click()
'This section allows the user to enter their score, then their hanicap is calculated for them
'This button is specifically for those who played from the mens tees
Score = InputBox("Please enter your score")
If Score > 0 Then
    HandicapDifferential = ((Score - 32.8) * 113 / 109)
    Handicap = FormatNumber(HandicapDifferential, 1) * 0.96
    picResults.Print "Your handicap is "; FormatNumber(Handicap, 1)
Else
    MsgBox "Sorry but you must enter a positive number", , "Error"
End If
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdStart_Click()
'This allows the user to go back to the previous form/screen

Form1.Show
Angushire.Hide
End Sub

Private Sub cmdWomen_Click()
'This section allows the user to enter their score, then their hanicap is calculated for them
'This button is specifically for those who played from the womens tees

Score = InputBox("Please enter your score")
Score = InputBox("Please enter your score")
If Score > 0 Then
    HandicapDifferential = ((Score - 34) * 113 / 113)
    Handicap = FormatNumber(HandicapDifferential, 1) * 0.96
    picResults.Print "Your handicap is "; FormatNumber(Handicap, 1)
Else
    MsgBox "Sorry but you must enter a positive number", , "Error"
End If
End Sub

Private Sub Command1_Click()

End Sub
