VERSION 5.00
Begin VB.Form SCCC 
   BackColor       =   &H00C0C000&
   Caption         =   "Form4"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7275
   LinkTopic       =   "Form4"
   ScaleHeight     =   4920
   ScaleWidth      =   7275
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   1695
      Left            =   4800
      Picture         =   "SCCC.frx":0000
      ScaleHeight     =   1635
      ScaleWidth      =   2235
      TabIndex        =   9
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   6000
      TabIndex        =   7
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Starting Page"
      Height          =   975
      Left            =   4080
      TabIndex        =   6
      Top             =   3240
      Width           =   1335
   End
   Begin VB.PictureBox picResults 
      Height          =   1335
      Left            =   360
      ScaleHeight     =   1275
      ScaleWidth      =   3195
      TabIndex        =   5
      Top             =   3240
      Width           =   3255
   End
   Begin VB.CommandButton cmdYellow 
      Caption         =   "Yellow"
      Height          =   975
      Left            =   4440
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdRed 
      Caption         =   "Red"
      Height          =   975
      Left            =   3000
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdWhite 
      Caption         =   "White"
      Height          =   975
      Left            =   1560
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdBlue 
      Caption         =   "Blue"
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C000&
      Caption         =   "By Meghan Ellenbecker"
      Height          =   375
      Left            =   3960
      TabIndex        =   10
      Top             =   4440
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      Caption         =   "SAINT CLOUD COUNTRY CLUB"
      Height          =   255
      Left            =   1080
      TabIndex        =   8
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Please select the tees that you played from by clicking on the appropriate button below."
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   4215
   End
End
Attribute VB_Name = "SCCC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project1(Golf Project.vbp)
'SCCC(SCCC.frm)
'Meghan Ellenbecker
'March 13, 2004
'This form finds the handicap of someone who played at the St. Cloud Country Club

Dim HandicapDifferential As Single
Dim Handicap As Single
Dim Score As Single

Private Sub cmdBlue_Click()
'This section allows the user to enter their score, then their handicap is calculated for them
'This button is specifically for those who played from the blue tees

Score = InputBox("Please enter your score")
If Score > 0 Then
    HandicapDifferential = ((Score - 72.6) * 113 / 133)
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
    HandicapDifferential = ((Score - 71.2) * 113 / 123)
    Handicap = FormatNumber(HandicapDifferential, 1) * 0.96
    picResults.Print "Your handicap is "; FormatNumber(Handicap, 1)
Else
    MsgBox "Sorry but you must enter a positive number", , "Error"
End If
End Sub

Private Sub cmdReturn_Click()
'This allows the user to return to the previous screen/form

Form1.Show
SCCC.Hide
End Sub

Private Sub cmdWhite_Click()
'This section allows the user to enter their score, then their handicap is calculated for them
'This button is specifically for those who played from the white tees

Score = InputBox("Please enter your score")
If Score > 0 Then
    HandicapDifferential = ((Score - 71.7) * 113 / 132)
    Handicap = FormatNumber(HandicapDifferential, 1) * 0.96
    picResults.Print "Your handicap is "; FormatNumber(Handicap, 1)
Else
    MsgBox "Sorry but you must enter a positive number", , "Error"
End If
End Sub

Private Sub cmdYellow_Click()
'This section allows the user to enter their score, then their handicap is calculated for them
'This button is specifically for those who played from the yellow tees

Score = InputBox("Please enter your score")
If Score > 0 Then
    HandicapDifferential = ((Score - 68.2) * 113 / 125)
    Handicap = FormatNumber(HandicapDifferential, 1) * 0.96
    picResults.Print "Your handicap is "; FormatNumber(Handicap, 1)
Else
    MsgBox "Sorry but you must enter a positive number", , "Error"
End If
End Sub
