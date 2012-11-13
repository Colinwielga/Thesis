VERSION 5.00
Begin VB.Form frmGameboard 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Jeopardy!"
   ClientHeight    =   4695
   ClientLeft      =   1050
   ClientTop       =   3210
   ClientWidth     =   13425
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form2"
   ScaleHeight     =   19.563
   ScaleMode       =   4  'Character
   ScaleWidth      =   111.875
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H0000FFFF&
      Caption         =   "Your are almost ready to play!  Click here to enter your name and begin."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Click here!"
      Top             =   120
      Width           =   10815
   End
   Begin VB.CommandButton cmdRank 
      BackColor       =   &H000000C0&
      Caption         =   "Finish and See Ranking"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   55
      ToolTipText     =   "Finish and See Ranking"
      Top             =   4080
      Width           =   2535
   End
   Begin VB.PictureBox picOptOne 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   10800
      ScaleHeight     =   375
      ScaleWidth      =   2295
      TabIndex        =   54
      Top             =   2880
      Width           =   2295
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   8880
      Picture         =   "frmGameboard.frx":0000
      ScaleHeight     =   2295
      ScaleWidth      =   855
      TabIndex        =   53
      ToolTipText     =   "Alex Trebek is the man!"
      Top             =   1800
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Option2"
      Height          =   255
      Index           =   3
      Left            =   10440
      TabIndex        =   47
      Top             =   2880
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Option2"
      Height          =   255
      Index           =   2
      Left            =   10440
      TabIndex        =   46
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton cmdSubmitAnswer 
      BackColor       =   &H000000C0&
      Caption         =   "Submit Answer"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "Submit Answer"
      Top             =   3600
      Width           =   2535
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Option2"
      Height          =   255
      Index           =   1
      Left            =   10440
      TabIndex        =   44
      Top             =   1920
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Option1"
      Height          =   255
      Index           =   0
      Left            =   10440
      TabIndex        =   43
      Top             =   1440
      Width           =   255
   End
   Begin VB.PictureBox picName 
      BackColor       =   &H80000002&
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   480
      ScaleHeight     =   555
      ScaleWidth      =   1515
      TabIndex        =   42
      ToolTipText     =   "Your name"
      Top             =   3840
      Width           =   1575
   End
   Begin VB.PictureBox picScore 
      BackColor       =   &H80000002&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      ScaleHeight     =   315
      ScaleWidth      =   1755
      TabIndex        =   41
      ToolTipText     =   "Score"
      Top             =   1920
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Left            =   360
      Picture         =   "frmGameboard.frx":082D
      ScaleHeight     =   1155
      ScaleWidth      =   1755
      TabIndex        =   39
      Top             =   2400
      Width           =   1815
   End
   Begin VB.PictureBox picPlayer 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   600
      ScaleHeight     =   1695
      ScaleWidth      =   1335
      TabIndex        =   37
      ToolTipText     =   "Your character"
      Top             =   120
      Width           =   1335
   End
   Begin VB.PictureBox picOptOne 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   10800
      ScaleHeight     =   375
      ScaleWidth      =   2295
      TabIndex        =   33
      Top             =   1920
      Width           =   2295
   End
   Begin VB.PictureBox picOptOne 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   10800
      ScaleHeight     =   375
      ScaleWidth      =   2295
      TabIndex        =   32
      Top             =   2400
      Width           =   2295
   End
   Begin VB.PictureBox picOptOne 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   10800
      ScaleHeight     =   375
      ScaleWidth      =   2295
      TabIndex        =   31
      Top             =   1440
      Width           =   2295
   End
   Begin VB.PictureBox picQuestion 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   8640
      ScaleHeight     =   435
      ScaleWidth      =   4515
      TabIndex        =   30
      Top             =   240
      Width           =   4575
   End
   Begin VB.CommandButton cmd1000B 
      Height          =   735
      Left            =   3720
      Picture         =   "frmGameboard.frx":1345
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmd1000C 
      Height          =   735
      Left            =   4920
      Picture         =   "frmGameboard.frx":251F
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmd1000D 
      Height          =   735
      Left            =   6120
      Picture         =   "frmGameboard.frx":36F9
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmd1000E 
      Height          =   735
      Left            =   7320
      Picture         =   "frmGameboard.frx":48D3
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmd1000A 
      Height          =   735
      Left            =   2520
      Picture         =   "frmGameboard.frx":5AAD
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmd800B 
      Height          =   735
      Left            =   3720
      Picture         =   "frmGameboard.frx":6C87
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmd800C 
      Height          =   735
      Left            =   4920
      Picture         =   "frmGameboard.frx":7DEE
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmd800D 
      Height          =   735
      Left            =   6120
      Picture         =   "frmGameboard.frx":8F55
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmd800E 
      Height          =   735
      Left            =   7320
      Picture         =   "frmGameboard.frx":A0BC
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmd800A 
      Height          =   735
      Left            =   2520
      Picture         =   "frmGameboard.frx":B223
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmd600B 
      Height          =   735
      Left            =   3720
      Picture         =   "frmGameboard.frx":C38A
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmd600C 
      Height          =   735
      Left            =   4920
      Picture         =   "frmGameboard.frx":D4F7
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmd600D 
      Height          =   735
      Left            =   6120
      Picture         =   "frmGameboard.frx":E664
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmd600E 
      Height          =   735
      Left            =   7320
      Picture         =   "frmGameboard.frx":F7D1
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmd600A 
      Height          =   735
      Left            =   2520
      Picture         =   "frmGameboard.frx":1093E
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmd400B 
      Height          =   735
      Left            =   3720
      Picture         =   "frmGameboard.frx":11AAB
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "$400 question"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmd400C 
      Height          =   735
      Left            =   4920
      Picture         =   "frmGameboard.frx":12BFD
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "$400 question"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmd400D 
      Height          =   735
      Left            =   6120
      Picture         =   "frmGameboard.frx":13D4F
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "$400 question"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmd400E 
      Height          =   735
      Left            =   7320
      Picture         =   "frmGameboard.frx":14EA1
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "$400 question"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmd400A 
      Height          =   735
      Left            =   2520
      Picture         =   "frmGameboard.frx":15FF3
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "$400 question"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmd200C 
      Height          =   735
      Left            =   4920
      Picture         =   "frmGameboard.frx":17145
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "$200 question"
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmd200D 
      Height          =   735
      Left            =   6120
      Picture         =   "frmGameboard.frx":182BA
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "$200 question"
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmd200E 
      Height          =   735
      Left            =   7320
      Picture         =   "frmGameboard.frx":1942F
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "$200 question"
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmd200B 
      Height          =   735
      Left            =   3720
      Picture         =   "frmGameboard.frx":1A5A4
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "$200 question"
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmd200A 
      Height          =   735
      Left            =   2520
      Picture         =   "frmGameboard.frx":1B719
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "$200 question"
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000FFFF&
      Height          =   855
      Left            =   8640
      TabIndex        =   57
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Height          =   255
      Left            =   8640
      TabIndex        =   56
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label lblCat5 
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "                       Science"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   7320
      TabIndex        =   52
      ToolTipText     =   "Category Five"
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblCat4 
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "                         Math"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   6120
      TabIndex        =   51
      ToolTipText     =   "Category Four"
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblCat3 
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "    Social         Studies"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   4920
      TabIndex        =   50
      ToolTipText     =   "Category Three"
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblCat2 
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "                     Presidents"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   3720
      TabIndex        =   49
      ToolTipText     =   "Category Two"
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblCat1 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "                     Holidays"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   2520
      TabIndex        =   48
      ToolTipText     =   "Category One"
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000FFFF&
      Height          =   4695
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Width           =   135
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00000040&
      FillStyle       =   0  'Solid
      Height          =   2775
      Left            =   120
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000FFFF&
      Height          =   135
      Left            =   8640
      TabIndex        =   36
      Top             =   840
      Width           =   5175
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FFFF&
      Height          =   4695
      Left            =   8520
      TabIndex        =   35
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Height          =   4815
      Left            =   13200
      TabIndex        =   34
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label lblLabelTop 
      BackColor       =   &H0000FFFF&
      Height          =   135
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   13935
   End
   Begin VB.Label lblLabelRight 
      BackColor       =   &H0000FFFF&
      Height          =   4575
      Left            =   8520
      TabIndex        =   28
      Top             =   120
      Width           =   135
   End
   Begin VB.Label lblLabelBottom 
      BackColor       =   &H0000FFFF&
      Height          =   135
      Left            =   0
      TabIndex        =   27
      Top             =   4560
      Width           =   13215
   End
   Begin VB.Label lblLabelLeft 
      BackColor       =   &H0000FFFF&
      Height          =   4695
      Left            =   2400
      TabIndex        =   26
      Top             =   0
      Width           =   135
   End
   Begin VB.Label lblBorderMiddle 
      BackColor       =   &H0000FFFF&
      Height          =   135
      Left            =   2400
      TabIndex        =   25
      Top             =   840
      Width           =   6135
   End
End
Attribute VB_Name = "frmGameboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' declare form variables
Dim Answer As Integer, Value As Integer, ButtonsClicked As Integer
Dim main(1 To 5, 1 To 25) As String
    
Private Sub cmd1000A_Click()
    cmdSubmitAnswer.Enabled = True
    cmdSubmitAnswer.ToolTipText = StrReverse(main(1, 21))
    Option1(0) = False
    Option1(1) = False
    Option1(2) = False
    Option1(3) = False
    picQuestion.Cls
    picOptOne(0).Cls
    picOptOne(1).Cls
    picOptOne(2).Cls
    picOptOne(3).Cls
    picQuestion.AutoRedraw = True
    picQuestion.Width = Len(main(1, 21))
    picQuestion.Print main(1, 21)
    picOptOne(0).Print main(1, 22)
    picOptOne(1).Print main(1, 23)
    picOptOne(2).Print main(1, 24)
    picOptOne(3).Print main(1, 25)
    Answer = 1
    Value = 1000
    cmd1000A.Enabled = False
End Sub

Private Sub cmd1000B_Click()
    cmdSubmitAnswer.Enabled = True
    cmdSubmitAnswer.ToolTipText = StrReverse(main(2, 22))
    Option1(0) = False
    Option1(1) = False
    Option1(2) = False
    Option1(3) = False
    picQuestion.Cls
    picOptOne(0).Cls
    picOptOne(1).Cls
    picOptOne(2).Cls
    picOptOne(3).Cls
    picQuestion.AutoRedraw = True
    picQuestion.Width = Len(main(2, 21))
    picQuestion.Print main(2, 21)
    picOptOne(0).Print main(2, 22)
    picOptOne(1).Print main(2, 23)
    picOptOne(2).Print main(2, 24)
    picOptOne(3).Print main(2, 25)
    Answer = 0
    Value = 1000
    cmd1000B.Enabled = False
End Sub

Private Sub cmd1000C_Click()
    cmdSubmitAnswer.Enabled = True
    cmdSubmitAnswer.ToolTipText = StrReverse(main(3, 23))
    Option1(0) = False
    Option1(1) = False
    Option1(2) = False
    Option1(3) = False
    picQuestion.Cls
    picOptOne(0).Cls
    picOptOne(1).Cls
    picOptOne(2).Cls
    picOptOne(3).Cls
    picQuestion.AutoRedraw = True
    picQuestion.Width = Len(main(3, 21))
    picQuestion.Print main(3, 21)
    picOptOne(0).Print main(3, 22)
    picOptOne(1).Print main(3, 23)
    picOptOne(2).Print main(3, 24)
    picOptOne(3).Print main(3, 25)
    Answer = 1
    Value = 1000
    cmd1000C.Enabled = False
End Sub

Private Sub cmd1000D_Click()
    cmdSubmitAnswer.Enabled = True
    cmdSubmitAnswer.ToolTipText = StrReverse(main(4, 24))
    Option1(0) = False
    Option1(1) = False
    Option1(2) = False
    Option1(3) = False
    picQuestion.Cls
    picOptOne(0).Cls
    picOptOne(1).Cls
    picOptOne(2).Cls
    picOptOne(3).Cls
    picQuestion.AutoRedraw = True
    picQuestion.Width = Len(main(4, 21))
    picQuestion.Print main(4, 21)
    picOptOne(0).Print main(4, 22)
    picOptOne(1).Print main(4, 23)
    picOptOne(2).Print main(4, 24)
    picOptOne(3).Print main(4, 25)
    Answer = 2
    Value = 1000
    cmd1000D.Enabled = False
End Sub

Private Sub cmd1000E_Click()
Dim DailyDouble As String
    cmdSubmitAnswer.Enabled = True
    cmdSubmitAnswer.ToolTipText = StrReverse(main(5, 25))
    Option1(0) = False
    Option1(1) = False
    Option1(2) = False
    Option1(3) = False
    picQuestion.Cls
    picOptOne(0).Cls
    picOptOne(1).Cls
    picOptOne(2).Cls
    picOptOne(3).Cls
    picQuestion.AutoRedraw = True
    picQuestion.Width = Len(main(5, 21))
    picQuestion.Print main(5, 21)
    picOptOne(0).Print main(5, 22)
    picOptOne(1).Print main(5, 23)
    picOptOne(2).Print main(5, 24)
    picOptOne(3).Print main(5, 25)
    Answer = 3
    Value = 2000
    DailyDouble = MsgBox("You have selected the Daily Double!! This question is worth $2000.", , "Daily Double!")
    cmd1000E.Enabled = False
End Sub

Private Sub cmd200A_Click()
    cmdSubmitAnswer.Enabled = True
    cmdSubmitAnswer.ToolTipText = StrReverse(main(1, 5))
    Option1(0) = False
    Option1(1) = False
    Option1(2) = False
    Option1(3) = False
    picQuestion.Cls
    picOptOne(0).Cls
    picOptOne(1).Cls
    picOptOne(2).Cls
    picOptOne(3).Cls
    picQuestion.AutoRedraw = True
    picQuestion.Width = Len(main(1, 1))
    picQuestion.Print main(1, 1)
    picOptOne(0).Print main(1, 2)
    picOptOne(1).Print main(1, 3)
    picOptOne(2).Print main(1, 4)
    picOptOne(3).Print main(1, 5)
    Answer = 3
    Value = 200
    cmd200A.Enabled = False
End Sub

Private Sub cmd200B_Click()
    cmdSubmitAnswer.Enabled = True
    cmdSubmitAnswer.ToolTipText = StrReverse(main(2, 5))
    Option1(0) = False
    Option1(1) = False
    Option1(2) = False
    Option1(3) = False
    picQuestion.Cls
    picOptOne(0).Cls
    picOptOne(1).Cls
    picOptOne(2).Cls
    picOptOne(3).Cls
    picQuestion.AutoRedraw = True
    picQuestion.Width = Len(main(2, 1))
    picQuestion.Print main(2, 1)
    picOptOne(0).Print main(2, 2)
    picOptOne(1).Print main(2, 3)
    picOptOne(2).Print main(2, 4)
    picOptOne(3).Print main(2, 5)
    Answer = 3
    Value = 200
    cmd200B.Enabled = False
End Sub

Private Sub cmd200C_Click()
    cmdSubmitAnswer.Enabled = True
    cmdSubmitAnswer.ToolTipText = StrReverse(main(3, 4))
    Option1(0) = False
    Option1(1) = False
    Option1(2) = False
    Option1(3) = False
    picQuestion.Cls
    picOptOne(0).Cls
    picOptOne(1).Cls
    picOptOne(2).Cls
    picOptOne(3).Cls
    picQuestion.AutoRedraw = True
    picQuestion.Width = Len(main(3, 1))
    picQuestion.Print main(3, 1)
    picOptOne(0).Print main(3, 2)
    picOptOne(1).Print main(3, 3)
    picOptOne(2).Print main(3, 4)
    picOptOne(3).Print main(3, 5)
    Answer = 2
    Value = 200
    cmd200C.Enabled = False
End Sub

Private Sub cmd200D_Click()
    cmdSubmitAnswer.Enabled = True
    cmdSubmitAnswer.ToolTipText = StrReverse(main(4, 5))
    Option1(0) = False
    Option1(1) = False
    Option1(2) = False
    Option1(3) = False
    picQuestion.Cls
    picOptOne(0).Cls
    picOptOne(1).Cls
    picOptOne(2).Cls
    picOptOne(3).Cls
    picQuestion.AutoRedraw = True
    picQuestion.Width = Len(main(4, 1))
    picQuestion.Print main(4, 1)
    picOptOne(0).Print main(4, 2)
    picOptOne(1).Print main(4, 3)
    picOptOne(2).Print main(4, 4)
    picOptOne(3).Print main(4, 5)
    Answer = 3
    Value = 200
    cmd200D.Enabled = False
End Sub

Private Sub cmd200E_Click()
    cmdSubmitAnswer.Enabled = True
    cmdSubmitAnswer.ToolTipText = StrReverse(main(5, 5))
    Option1(0) = False
    Option1(1) = False
    Option1(2) = False
    Option1(3) = False
    picQuestion.Cls
    picOptOne(0).Cls
    picOptOne(1).Cls
    picOptOne(2).Cls
    picOptOne(3).Cls
    picQuestion.AutoRedraw = True
    picQuestion.Width = Len(main(5, 1))
    picQuestion.Print main(5, 1)
    picOptOne(0).Print main(5, 2)
    picOptOne(1).Print main(5, 3)
    picOptOne(2).Print main(5, 4)
    picOptOne(3).Print main(5, 5)
    Answer = 3
    Value = 200
    cmd200E.Enabled = False
End Sub

Private Sub cmd400A_Click()
    cmdSubmitAnswer.Enabled = True
    cmdSubmitAnswer.ToolTipText = StrReverse(main(1, 9))
    Option1(0) = False
    Option1(1) = False
    Option1(2) = False
    Option1(3) = False
    picQuestion.Cls
    picOptOne(0).Cls
    picOptOne(1).Cls
    picOptOne(2).Cls
    picOptOne(3).Cls
    picQuestion.AutoRedraw = True
    picQuestion.Width = Len(main(1, 6))
    picQuestion.Print main(1, 6)
    picOptOne(0).Print main(1, 7)
    picOptOne(1).Print main(1, 8)
    picOptOne(2).Print main(1, 9)
    picOptOne(3).Print main(1, 10)
    Answer = 2
    Value = 400
    cmd400A.Enabled = False
End Sub

Private Sub cmd400B_Click()
    cmdSubmitAnswer.Enabled = True
    cmdSubmitAnswer.ToolTipText = StrReverse(main(2, 9))
    Option1(0) = False
    Option1(1) = False
    Option1(2) = False
    Option1(3) = False
    picQuestion.Cls
    picOptOne(0).Cls
    picOptOne(1).Cls
    picOptOne(2).Cls
    picOptOne(3).Cls
    picQuestion.AutoRedraw = True
    picQuestion.Width = Len(main(2, 6))
    picQuestion.Print main(2, 6)
    picOptOne(0).Print main(2, 7)
    picOptOne(1).Print main(2, 8)
    picOptOne(2).Print main(2, 9)
    picOptOne(3).Print main(2, 10)
    Answer = 2
    Value = 400
    cmd400B.Enabled = False
End Sub

Private Sub cmd400C_Click()
    cmdSubmitAnswer.Enabled = True
    cmdSubmitAnswer.ToolTipText = StrReverse(main(3, 8))
    Option1(0) = False
    Option1(1) = False
    Option1(2) = False
    Option1(3) = False
    picQuestion.Cls
    picOptOne(0).Cls
    picOptOne(1).Cls
    picOptOne(2).Cls
    picOptOne(3).Cls
    picQuestion.AutoRedraw = True
    picQuestion.Width = Len(main(3, 6))
    picQuestion.Print main(3, 6)
    picOptOne(0).Print main(3, 7)
    picOptOne(1).Print main(3, 8)
    picOptOne(2).Print main(3, 9)
    picOptOne(3).Print main(3, 10)
    Answer = 1
    Value = 400
    cmd400C.Enabled = False
End Sub

Private Sub cmd400D_Click()
    cmdSubmitAnswer.Enabled = True
    cmdSubmitAnswer.ToolTipText = StrReverse(main(4, 9))
    Option1(0) = False
    Option1(1) = False
    Option1(2) = False
    Option1(3) = False
    picQuestion.Cls
    picOptOne(0).Cls
    picOptOne(1).Cls
    picOptOne(2).Cls
    picOptOne(3).Cls
    picQuestion.AutoRedraw = True
    picQuestion.Width = Len(main(4, 6))
    picQuestion.Print main(4, 6)
    picOptOne(0).Print main(4, 7)
    picOptOne(1).Print main(4, 8)
    picOptOne(2).Print main(4, 9)
    picOptOne(3).Print main(4, 10)
    Answer = 2
    Value = 400
    cmd400D.Enabled = False
End Sub

Private Sub cmd400E_Click()
    cmdSubmitAnswer.Enabled = True
    cmdSubmitAnswer.ToolTipText = StrReverse(main(5, 10))
    Option1(0) = False
    Option1(1) = False
    Option1(2) = False
    Option1(3) = False
    picQuestion.Cls
    picOptOne(0).Cls
    picOptOne(1).Cls
    picOptOne(2).Cls
    picOptOne(3).Cls
    picQuestion.AutoRedraw = True
    picQuestion.Width = Len(main(5, 6))
    picQuestion.Print main(5, 6)
    picOptOne(0).Print main(5, 7)
    picOptOne(1).Print main(5, 8)
    picOptOne(2).Print main(5, 9)
    picOptOne(3).Print main(5, 10)
    Answer = 3
    Value = 400
    cmd400E.Enabled = False
End Sub

Private Sub cmd600A_Click()
    cmdSubmitAnswer.Enabled = True
    cmdSubmitAnswer.ToolTipText = StrReverse(main(1, 12))
    Option1(0) = False
    Option1(1) = False
    Option1(2) = False
    Option1(3) = False
    picQuestion.Cls
    picOptOne(0).Cls
    picOptOne(1).Cls
    picOptOne(2).Cls
    picOptOne(3).Cls
    picQuestion.AutoRedraw = True
    picQuestion.Width = Len(main(1, 11))
    picQuestion.Print main(1, 11)
    picOptOne(0).Print main(1, 12)
    picOptOne(1).Print main(1, 13)
    picOptOne(2).Print main(1, 14)
    picOptOne(3).Print main(1, 15)
    Answer = 0
    Value = 600
    cmd600A.Enabled = False
End Sub

Private Sub cmd600B_Click()
    cmdSubmitAnswer.Enabled = True
    cmdSubmitAnswer.ToolTipText = StrReverse(main(2, 14))
    Option1(0) = False
    Option1(1) = False
    Option1(2) = False
    Option1(3) = False
    picQuestion.Cls
    picOptOne(0).Cls
    picOptOne(1).Cls
    picOptOne(2).Cls
    picOptOne(3).Cls
    picQuestion.AutoRedraw = True
    picQuestion.Width = Len(main(2, 11))
    picQuestion.Print main(2, 11)
    picOptOne(0).Print main(2, 12)
    picOptOne(1).Print main(2, 13)
    picOptOne(2).Print main(2, 14)
    picOptOne(3).Print main(2, 15)
    Answer = 2
    Value = 600
    cmd600B.Enabled = False
End Sub

Private Sub cmd600C_Click()
    cmdSubmitAnswer.Enabled = True
    cmdSubmitAnswer.ToolTipText = StrReverse(main(3, 13))
    Option1(0) = False
    Option1(1) = False
    Option1(2) = False
    Option1(3) = False
    picQuestion.Cls
    picOptOne(0).Cls
    picOptOne(1).Cls
    picOptOne(2).Cls
    picOptOne(3).Cls
    picQuestion.AutoRedraw = True
    picQuestion.Width = Len(main(3, 11))
    picQuestion.Print main(3, 11)
    picOptOne(0).Print main(3, 12)
    picOptOne(1).Print main(3, 13)
    picOptOne(2).Print main(3, 14)
    picOptOne(3).Print main(3, 15)
    Answer = 1
    Value = 600
    cmd600C.Enabled = False
End Sub

Private Sub cmd600D_Click()
    cmdSubmitAnswer.Enabled = True
    cmdSubmitAnswer.ToolTipText = StrReverse(main(4, 13))
    Option1(0) = False
    Option1(1) = False
    Option1(2) = False
    Option1(3) = False
    picQuestion.Cls
    picOptOne(0).Cls
    picOptOne(1).Cls
    picOptOne(2).Cls
    picOptOne(3).Cls
    picQuestion.AutoRedraw = True
    picQuestion.Width = Len(main(4, 11))
    picQuestion.Print main(4, 11)
    picOptOne(0).Print main(4, 12)
    picOptOne(1).Print main(4, 13)
    picOptOne(2).Print main(4, 14)
    picOptOne(3).Print main(4, 15)
    Answer = 1
    Value = 600
    cmd600D.Enabled = False
End Sub

Private Sub cmd600E_Click()
    cmdSubmitAnswer.Enabled = True
    cmdSubmitAnswer.ToolTipText = StrReverse(main(5, 14))
    Option1(0) = False
    Option1(1) = False
    Option1(2) = False
    Option1(3) = False
    picQuestion.Cls
    picOptOne(0).Cls
    picOptOne(1).Cls
    picOptOne(2).Cls
    picOptOne(3).Cls
    picQuestion.AutoRedraw = True
    picQuestion.Width = Len(main(5, 11))
    picQuestion.Print main(5, 11)
    picOptOne(0).Print main(5, 12)
    picOptOne(1).Print main(5, 13)
    picOptOne(2).Print main(5, 14)
    picOptOne(3).Print main(5, 15)
    Answer = 2
    Value = 600
    cmd600E.Enabled = False
End Sub

Private Sub cmd800A_Click()
    cmdSubmitAnswer.Enabled = True
    cmdSubmitAnswer.ToolTipText = StrReverse(main(1, 19))
    Option1(0) = False
    Option1(1) = False
    Option1(2) = False
    Option1(3) = False
    picQuestion.Cls
    picOptOne(0).Cls
    picOptOne(1).Cls
    picOptOne(2).Cls
    picOptOne(3).Cls
    picQuestion.AutoRedraw = True
    picQuestion.Width = Len(main(1, 16))
    picQuestion.Print main(1, 16)
    picOptOne(0).Print main(1, 17)
    picOptOne(1).Print main(1, 18)
    picOptOne(2).Print main(1, 19)
    picOptOne(3).Print main(1, 20)
    Answer = 2
    Value = 800
    cmd800A.Enabled = False
End Sub

Private Sub cmd800B_Click()
    cmdSubmitAnswer.Enabled = True
    cmdSubmitAnswer.ToolTipText = StrReverse(main(2, 17))
    Option1(0) = False
    Option1(1) = False
    Option1(2) = False
    Option1(3) = False
    picQuestion.Cls
    picOptOne(0).Cls
    picOptOne(1).Cls
    picOptOne(2).Cls
    picOptOne(3).Cls
    picQuestion.AutoRedraw = True
    picQuestion.Width = Len(main(2, 16))
    picQuestion.Print main(2, 16)
    picOptOne(0).Print main(2, 17)
    picOptOne(1).Print main(2, 18)
    picOptOne(2).Print main(2, 19)
    picOptOne(3).Print main(2, 20)
    Answer = 0
    Value = 800
    cmd800B.Enabled = False
End Sub

Private Sub cmd800C_Click()
    cmdSubmitAnswer.Enabled = True
    cmdSubmitAnswer.ToolTipText = StrReverse(main(3, 17))
    Option1(0) = False
    Option1(1) = False
    Option1(2) = False
    Option1(3) = False
    picQuestion.Cls
    picOptOne(0).Cls
    picOptOne(1).Cls
    picOptOne(2).Cls
    picOptOne(3).Cls
    picQuestion.AutoRedraw = True
    picQuestion.Width = Len(main(3, 16))
    picQuestion.Print main(3, 16)
    picOptOne(0).Print main(3, 17)
    picOptOne(1).Print main(3, 18)
    picOptOne(2).Print main(3, 19)
    picOptOne(3).Print main(3, 20)
    Answer = 0
    Value = 800
    cmd800C.Enabled = False
End Sub

Private Sub cmd800D_Click()
    cmdSubmitAnswer.Enabled = True
    cmdSubmitAnswer.ToolTipText = StrReverse(main(4, 20))
    Option1(0) = False
    Option1(1) = False
    Option1(2) = False
    Option1(3) = False
    picQuestion.Cls
    picOptOne(0).Cls
    picOptOne(1).Cls
    picOptOne(2).Cls
    picOptOne(3).Cls
    picQuestion.AutoRedraw = True
    picQuestion.Width = Len(main(4, 16))
    picQuestion.Print main(4, 16)
    picOptOne(0).Print main(4, 17)
    picOptOne(1).Print main(4, 18)
    picOptOne(2).Print main(4, 19)
    picOptOne(3).Print main(4, 20)
    Answer = 3
    Value = 800
    cmd800D.Enabled = False
End Sub

Private Sub cmd800E_Click()
    cmdSubmitAnswer.Enabled = True
    cmdSubmitAnswer.ToolTipText = StrReverse(main(5, 19))
    Option1(0) = False
    Option1(1) = False
    Option1(2) = False
    Option1(3) = False
    picQuestion.Cls
    picOptOne(0).Cls
    picOptOne(1).Cls
    picOptOne(2).Cls
    picOptOne(3).Cls
    picQuestion.AutoRedraw = True
    picQuestion.Width = Len(main(5, 16))
    picQuestion.Print main(5, 16)
    picOptOne(0).Print main(5, 17)
    picOptOne(1).Print main(5, 18)
    picOptOne(2).Print main(5, 19)
    picOptOne(3).Print main(5, 20)
    Answer = 2
    Value = 800
    cmd800E.Enabled = False
End Sub

Private Sub cmdLoad_Click()
    'declare subroutine variables
    Dim row As Integer, col As Integer
    
    'open all data
    Open App.Path & "\main.txt" For Input As #1
    For row = 1 To 5
        For col = 1 To 25
            Input #1, main(row, col)
        Next col
    Next row
    Close #1
        
    'prompt user to enter name and dispaly results
    Player = InputBox("What is your name?", "What is your name?")
    picName.Print Player
    
    'initialize player score to $0 and display
    Score = 0
    picScore.Print Tab(10); FormatCurrency(Score, 0)
    
    'hide name entry command button
    cmdLoad.Visible = False
        
    'display message box that informs user he or she is ready to play
    MsgBox "Please click on the value of the question you wish to answer to begin the game", , "You are ready to play!"

End Sub

Private Sub cmdRank_Click()
    'compute and display ranking
    Select Case Score
        Case 14000 To 16000
            MsgBox "Congratulations!  You earned " & FormatCurrency(Score, 0) & ".  You are the master of Children's Jeopardy!", , "Children's Jeopardy Master"
        Case 10000 To 13999
            MsgBox "Nice work!  You earned " & FormatCurrency(Score, 0) & ".  You are a Jeopardy master-in-training!", , "Children's Jeopardy Master-in-Training"
        Case 5000 To 9999
            MsgBox "Not too bad!  You earned " & FormatCurrency(Score, 0) & ".  With a little more practice, you could be really good!", , "Children's Jeopardy Average Player"
        Case 1 To 4999
            MsgBox "That was terrible!  You earned " & FormatCurrency(Score, 0) & ".  Just remember, practice makes perfect!", , "Children's Jeopardy Rookie"
        Case Else
            MsgBox "Wow!  What can a guy as smart as Alex Trebek say...  With an earnings total of " & FormatCurrency(Score, 0) & ", you need some serious training!", , "Try Jeopardy for Babies"
    End Select
    
    'thank user for playing
    MsgBox "Thank you very much for playing Children's Jeopardy.  Be sure to check the high score list to see how you rank among other players.  Until next time, goodbye.", , "Thank you"
    
    'enable high score button on welcome form
    frmWelcome.cmdHighScores.Enabled = True
    
    'disable play button on welcome form
    frmWelcome.cmdPlay.Enabled = False
    
    'hide gameboard and show welcome
    frmGameboard.Hide
    frmWelcome.Show
End Sub

Private Sub cmdSubmitAnswer_Click()
    'identify correct or incorrect answer and compute new score
    cmdSubmitAnswer.Enabled = False
    For Increment = 0 To 3
        If Option1(Increment) = True Then
            If Increment = Answer Then
                Score = Score + Value
                picScore.Cls
                picScore.Print Tab(10); FormatCurrency(Score, 0)
                MsgBox "Correct!  Please choose again.", vbInformation, "Correct Answer"
            ElseIf Increment <> Answer Then
                Score = Score - Value
                picScore.Cls
                picScore.Print Tab(10); FormatCurrency(Score, 0)
                MsgBox "Sorry, that is incorrect!  My name is Alex Trebek, I know everything!", vbCritical, "Incorrect Answer"
            End If
        End If
    Next Increment
    
    'determine when all questions have been answered
    ButtonsClicked = ButtonsClicked + 1
    If ButtonsClicked = 25 Then
        MsgBox "You have answered all of the questions.  Please click Finish and See Ranking to Exit", , "All questions answered"
    End If
End Sub

