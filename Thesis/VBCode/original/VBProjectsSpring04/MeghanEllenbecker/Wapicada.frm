VERSION 5.00
Begin VB.Form Wapicada 
   BackColor       =   &H00FF8080&
   Caption         =   "Form3"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8670
   LinkTopic       =   "Form3"
   ScaleHeight     =   4830
   ScaleWidth      =   8670
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   2055
      Left            =   5760
      Picture         =   "Wapicada.frx":0000
      ScaleHeight     =   1995
      ScaleWidth      =   2715
      TabIndex        =   9
      Top             =   240
      Width           =   2775
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   855
      Left            =   7560
      TabIndex        =   7
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FF8080&
      Caption         =   "Return to Starting Page"
      Height          =   1095
      Left            =   5880
      TabIndex        =   6
      Top             =   3000
      Width           =   1215
   End
   Begin VB.PictureBox picResults 
      Height          =   1095
      Left            =   600
      ScaleHeight     =   1035
      ScaleWidth      =   3195
      TabIndex        =   5
      Top             =   3480
      Width           =   3255
   End
   Begin VB.CommandButton cmdGold 
      Caption         =   "Gold"
      Height          =   1095
      Left            =   4440
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdRed 
      Caption         =   "Red"
      Height          =   1095
      Left            =   3000
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdWhite 
      Caption         =   "White"
      Height          =   1095
      Left            =   1560
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdBlue 
      Caption         =   "Blue"
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
      Caption         =   "By Meghan Ellenbecker"
      Height          =   495
      Left            =   4920
      TabIndex        =   10
      Top             =   4200
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   "WAPICADA "
      Height          =   255
      Left            =   1680
      TabIndex        =   8
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "Please select the tees that you played from by clicking on the appropriate button below."
      Height          =   495
      Index           =   1
      Left            =   600
      TabIndex        =   0
      Top             =   1200
      Width           =   3855
   End
End
Attribute VB_Name = "Wapicada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project1(Golf Project.vbp)
'Wapicada(Wapicada.frm)
'Meghan Ellenbecker
'March 13, 2004
'This finds the handicap of a person who played gold at Wapicada
Dim HandicapDifferential As Single
Dim Handicap As Single
Dim Score As Single

Private Sub cmdBlue_Click()
'This section allows the user to enter their score, then their handicap is calculated for them
'This button is specifically for those who played from the blue tees

Score = InputBox("Please enter your score")
If Score > 0 Then
    HandicapDifferential = ((Score - 71.7) * 113 / 127)
    Handicap = FormatNumber(HandicapDifferential, 1) * 0.96
    picResults.Print "Your handicap is "; FormatNumber(Handicap, 1)
Else
    MsgBox "Sorry but you must enter a positive number", , "Error"
End If
End Sub

Private Sub cmdGold_Click()
'This section allows the user to enter their score, then their handicap is calculated for them
'This button is specifically for those who played from the gold tees


Score = InputBox("Please enter your score")
If Score > 0 Then
    HandicapDifferential = ((Score - 67.2) * 113 / 118)
    Handicap = FormatNumber(HandicapDifferential, 1) * 0.96
    picResults.Print "Your handicap is "; FormatNumber(Handicap, 1)
Else
    MsgBox "Sorry but you must enter a positive number", , "Error"
End If
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdRed_Click()
'This section allows the user to enter their score, then their handicap is calculated for them
'This button is specifically for those who played from the red tees

Score = InputBox("Please enter your score")
If Score > 0 Then
    HandicapDifferential = ((Score - 71.5) * 113 / 126)
    Handicap = FormatNumber(HandicapDifferential, 1) * 0.96
    picResults.Print "Your handicap is "; FormatNumber(Handicap, 1)
Else
    MsgBox "Sorry but you must enter a positive number", , "Error"
End If
End Sub

Private Sub cmdReturn_Click()
'This allows the user to return to the previous form/screen
Form1.Show
Wapicada.Hide
End Sub

Private Sub cmdWhite_Click()
'This section allows the user to enter their score, then their handicap is calculated for them
'This button is specifically for those who played from the white tees

Score = InputBox("Please enter your score")
If Score > 0 Then
    HandicapDifferential = ((Score - 70.1) * 113 / 124)
    Handicap = FormatNumber(HandicapDifferential, 1) * 0.96
    picResults.Print "Your handicap is "; FormatNumber(Handicap, 1)
Else
    MsgBox "Sorry but you must enter a positive number", , "Error"
End If
End Sub

