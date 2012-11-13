VERSION 5.00
Begin VB.Form Hangman2 
   Caption         =   "Form1"
   ClientHeight    =   9735
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12930
   LinkTopic       =   "Form1"
   Picture         =   "formhangmaninfluenza.frx":0000
   ScaleHeight     =   9735
   ScaleWidth      =   12930
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton Command2 
      BackColor       =   &H000080FF&
      Caption         =   "If You Lose, Click Here"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4560
      TabIndex        =   34
      Top             =   6960
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go Back to Hangman Page"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   8760
      TabIndex        =   33
      Top             =   6840
      Width           =   3135
   End
   Begin VB.PictureBox Picresults 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   3240
      ScaleHeight     =   1515
      ScaleWidth      =   6195
      TabIndex        =   28
      Top             =   1440
      Width           =   6255
   End
   Begin VB.PictureBox Picfails 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   480
      ScaleHeight     =   3915
      ScaleWidth      =   1755
      TabIndex        =   27
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CheckBox chka 
      BackColor       =   &H0000FFFF&
      Caption         =   "A"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   3120
      TabIndex        =   26
      Top             =   3000
      Width           =   615
   End
   Begin VB.CheckBox chkb 
      BackColor       =   &H0000FFFF&
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   4080
      TabIndex        =   25
      Top             =   3000
      Width           =   615
   End
   Begin VB.CheckBox chkc 
      BackColor       =   &H0000FFFF&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   5040
      TabIndex        =   24
      Top             =   3000
      Width           =   615
   End
   Begin VB.CheckBox chkd 
      BackColor       =   &H0000FFFF&
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   6000
      TabIndex        =   23
      Top             =   3000
      Width           =   615
   End
   Begin VB.CheckBox chkg 
      BackColor       =   &H0000FFFF&
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   8880
      TabIndex        =   22
      Top             =   3000
      Width           =   615
   End
   Begin VB.CheckBox chkf 
      BackColor       =   &H0000FFFF&
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   7920
      TabIndex        =   21
      Top             =   3000
      Width           =   615
   End
   Begin VB.CheckBox chke 
      BackColor       =   &H0000FFFF&
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   6960
      TabIndex        =   20
      Top             =   3000
      Width           =   615
   End
   Begin VB.CheckBox chkt 
      BackColor       =   &H0000FFFF&
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   7920
      TabIndex        =   19
      Top             =   4920
      Width           =   615
   End
   Begin VB.CheckBox chko 
      BackColor       =   &H0000FFFF&
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   3120
      TabIndex        =   18
      Top             =   4920
      Width           =   615
   End
   Begin VB.CheckBox chkr 
      BackColor       =   &H0000FFFF&
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   6000
      TabIndex        =   17
      Top             =   4920
      Width           =   615
   End
   Begin VB.CheckBox chki 
      BackColor       =   &H0000FFFF&
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   4080
      TabIndex        =   16
      Top             =   3960
      Width           =   615
   End
   Begin VB.CheckBox chkn 
      BackColor       =   &H0000FFFF&
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   8880
      TabIndex        =   15
      Top             =   3960
      Width           =   615
   End
   Begin VB.CheckBox chkm 
      BackColor       =   &H0000FFFF&
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   7920
      TabIndex        =   14
      Top             =   3960
      Width           =   615
   End
   Begin VB.CheckBox chkl 
      BackColor       =   &H0000FFFF&
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   6960
      TabIndex        =   13
      Top             =   3960
      Width           =   615
   End
   Begin VB.CheckBox chkh 
      BackColor       =   &H0000FFFF&
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   3120
      TabIndex        =   12
      Top             =   3960
      Width           =   615
   End
   Begin VB.PictureBox Picbroth 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   9840
      ScaleHeight     =   3075
      ScaleWidth      =   2355
      TabIndex        =   11
      Top             =   2760
      Width           =   2415
   End
   Begin VB.CheckBox chkj 
      BackColor       =   &H0000FFFF&
      Caption         =   "J"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   5040
      TabIndex        =   10
      Top             =   3960
      Width           =   615
   End
   Begin VB.CheckBox chkk 
      BackColor       =   &H0000FFFF&
      Caption         =   "K"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   6000
      TabIndex        =   9
      Top             =   3960
      Width           =   615
   End
   Begin VB.CheckBox chkx 
      BackColor       =   &H0000FFFF&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   6000
      TabIndex        =   8
      Top             =   5880
      Width           =   615
   End
   Begin VB.CheckBox chku 
      BackColor       =   &H0000FFFF&
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   8880
      TabIndex        =   7
      Top             =   4920
      Width           =   615
   End
   Begin VB.CheckBox chkv 
      BackColor       =   &H0000FFFF&
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   4080
      TabIndex        =   6
      Top             =   5880
      Width           =   615
   End
   Begin VB.CheckBox chkp 
      BackColor       =   &H0000FFFF&
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   4080
      TabIndex        =   5
      Top             =   4920
      Width           =   615
   End
   Begin VB.CheckBox chkw 
      BackColor       =   &H0000FFFF&
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   5040
      TabIndex        =   4
      Top             =   5880
      Width           =   615
   End
   Begin VB.CheckBox chkq 
      BackColor       =   &H0000FFFF&
      Caption         =   "Q"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   5040
      TabIndex        =   3
      Top             =   4920
      Width           =   615
   End
   Begin VB.CheckBox chks 
      BackColor       =   &H0000FFFF&
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   6960
      TabIndex        =   2
      Top             =   4920
      Width           =   615
   End
   Begin VB.CheckBox chky 
      BackColor       =   &H0000FFFF&
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   6960
      TabIndex        =   1
      Top             =   5880
      Width           =   615
   End
   Begin VB.CheckBox chkz 
      BackColor       =   &H0000FFFF&
      Caption         =   "Z"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   7920
      TabIndex        =   0
      Top             =   5880
      Width           =   615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   1200
      X2              =   1200
      Y1              =   8880
      Y2              =   6600
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   1200
      X2              =   2400
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   2400
      X2              =   2400
      Y1              =   6600
      Y2              =   6960
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   2400
      X2              =   2400
      Y1              =   7440
      Y2              =   8040
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   2400
      X2              =   3000
      Y1              =   8040
      Y2              =   8520
   End
   Begin VB.Line Line6 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   2400
      X2              =   2040
      Y1              =   8040
      Y2              =   8640
   End
   Begin VB.Line Line7 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   2400
      X2              =   3120
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Line Line8 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   2640
      X2              =   1800
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Line Line9 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   600
      X2              =   1800
      Y1              =   8880
      Y2              =   8880
   End
   Begin VB.Line Line11 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   2040
      X2              =   2760
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Line Line12 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   2760
      X2              =   2760
      Y1              =   6960
      Y2              =   7440
   End
   Begin VB.Line Line13 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   2040
      X2              =   2040
      Y1              =   6960
      Y2              =   7440
   End
   Begin VB.Line Line14 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   2040
      X2              =   2760
      Y1              =   7440
      Y2              =   7440
   End
   Begin VB.Image imghangman 
      Height          =   3765
      Left            =   360
      Picture         =   "formhangmaninfluenza.frx":1CF50E
      Top             =   5880
      Width           =   3360
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000FFFF&
      Caption         =   "Am I Correct?"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   32
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FFFF&
      Caption         =   "Correct Letters in Appropriate Order"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10200
      TabIndex        =   31
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Fails (8 to lose)"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   30
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Your Word:    ______  ______  ______  ______  ______  ______  ______  ______  ______   Hint-> It is an illness!"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      TabIndex        =   29
      Top             =   120
      Width           =   9375
   End
End
Attribute VB_Name = "Hangman2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Healthy Living
'Form Hangman2
'Joel Coleman
'March 29, 2008
'To write a game which deals with words associated with health
'I used many true/false statements in order to complete my task

Option Explicit
'Declaring variables for entire form
Dim CTR As Single, chkivalue As Boolean, chknvalue As Boolean, chkfvalue As Boolean
Dim chklvalue As Boolean, chkuvalue As Boolean, chkevalue As Boolean, chknzvalue As Boolean
Dim chkzvalue As Boolean, chkavalue As Boolean
Private Sub chkb_click(Index As Integer)
'Adds up how many wrong answers and prints, if more than 7 then they lose
CTR = CTR + 1
'Select case with a true and false visibility to draw hangman as game progresses
Select Case CTR
Case Is = 1
    Line9 = True
Case Is = 2
    Line1 = True
    Line2 = True
    Line3 = True
Case Is = 3
    Line11 = True
    Line12 = True
    Line13 = True
    Line14 = True
Case Is = 4
    Line4 = True
Case Is = 5
    Line5 = True
Case Is = 6
    Line6 = True
Case Is = 7
    Line7 = True
Case Is = 8
    Line8 = True
End Select
'Prints how many wrongs and if over 7 times, then game is over
Picfails.Print CTR
If CTR > 7 Then Picfails.Print "You Lose"
If CTR > 7 Then MsgBox "You Lose!", , "You Lose"
End Sub

Private Sub chkh_Click(Index As Integer)
'Adds up how many wrong answers and prints, if more than 7 then they lose
CTR = CTR + 1
'Select case with a true and false visibility to draw hangman as game progresses
Select Case CTR
Case Is = 1
    Line9 = True
Case Is = 2
    Line1 = True
    Line2 = True
    Line3 = True
Case Is = 3
    Line11 = True
    Line12 = True
    Line13 = True
    Line14 = True
Case Is = 4
    Line4 = True
Case Is = 5
    Line5 = True
Case Is = 6
    Line6 = True
Case Is = 7
    Line7 = True
Case Is = 8
    Line8 = True
End Select
'Prints how many wrongs and if over 7 times, then game is over
Picfails.Print CTR
If CTR > 7 Then Picfails.Print "You Lose"
If CTR > 7 Then MsgBox "You Lose!", , "You Lose"
End Sub

Private Sub chko_Click(Index As Integer)
'Adds up how many wrong answers and prints, if more than 7 then they lose
CTR = CTR + 1
'Select case with a true and false visibility to draw hangman as game progresses
Select Case CTR
Case Is = 1
    Line9 = True
Case Is = 2
    Line1 = True
    Line2 = True
    Line3 = True
Case Is = 3
    Line11 = True
    Line12 = True
    Line13 = True
    Line14 = True
Case Is = 4
    Line4 = True
Case Is = 5
    Line5 = True
Case Is = 6
    Line6 = True
Case Is = 7
    Line7 = True
Case Is = 8
    Line8 = True
End Select
'Prints how many wrongs and if over 7 times, then game is over
Picfails.Print CTR
If CTR > 7 Then Picfails.Print "You Lose"
If CTR > 7 Then MsgBox "You Lose!", , "You Lose"
End Sub

Private Sub chkr_Click(Index As Integer)
'Adds up how many wrong answers and prints, if more than 7 then they lose
CTR = CTR + 1
'Select case with a true and false visibility to draw hangman as game progresses
Select Case CTR
Case Is = 1
    Line9 = True
Case Is = 2
    Line1 = True
    Line2 = True
    Line3 = True
Case Is = 3
    Line11 = True
    Line12 = True
    Line13 = True
    Line14 = True
Case Is = 4
    Line4 = True
Case Is = 5
    Line5 = True
Case Is = 6
    Line6 = True
Case Is = 7
    Line7 = True
Case Is = 8
    Line8 = True
End Select
'Prints how many wrongs and if over 7 times, then game is over
Picfails.Print CTR
If CTR > 7 Then Picfails.Print "You Lose"
If CTR > 7 Then MsgBox "You Lose!", , "You Lose"
End Sub

Private Sub chkt_Click(Index As Integer)
'Adds up how many wrong answers and prints, if more than 7 then they lose
CTR = CTR + 1
'Select case with a true and false visibility to draw hangman as game progresses
Select Case CTR
Case Is = 1
    Line9 = True
Case Is = 2
    Line1 = True
    Line2 = True
    Line3 = True
Case Is = 3
    Line11 = True
    Line12 = True
    Line13 = True
    Line14 = True
Case Is = 4
    Line4 = True
Case Is = 5
    Line5 = True
Case Is = 6
    Line6 = True
Case Is = 7
    Line7 = True
Case Is = 8
    Line8 = True
End Select
'Prints how many wrongs and if over 7 times, then game is over
Picfails.Print CTR
If CTR > 7 Then Picfails.Print "You Lose"
If CTR > 7 Then MsgBox "You Lose!", , "You Lose"
End Sub
Private Sub chka_Click(Index As Integer)
'Stores in program that letter is selected
chkavalue = True
'If all correct letters are checked, it will print the winning answer, otherwise it will print the letter selected in its appropriate spot
If chkivalue = True And chknzvalue = True And chkfvalue = True And chklvalue = True And chkuvalue = True And chkevalue = True And chkzvalue = True And chkavalue = True Then
picResults.Print "INFLUENZA is Correct!!", "YOU WIN!!"
End If
Picbroth.Print "_ _ _ _ _ _ _ _ A "
End Sub
Private Sub chkc_Click(Index As Integer)
'Adds up how many wrong answers and prints, if more than 7 then they lose
CTR = CTR + 1
'Select case with a true and false visibility to draw hangman as game progresses
Select Case CTR
Case Is = 1
    Line9 = True
Case Is = 2
    Line1 = True
    Line2 = True
    Line3 = True
Case Is = 3
    Line11 = True
    Line12 = True
    Line13 = True
    Line14 = True
Case Is = 4
    Line4 = True
Case Is = 5
    Line5 = True
Case Is = 6
    Line6 = True
Case Is = 7
    Line7 = True
Case Is = 8
    Line8 = True
End Select
'Prints how many wrongs and if over 7 times, then game is over
Picfails.Print CTR
If CTR > 7 Then Picfails.Print "You Lose"
If CTR > 7 Then MsgBox "You Lose!", , "You Lose"
End Sub
Private Sub chkd_Click(Index As Integer)
'Adds up how many wrong answers and prints, if more than 7 then they lose
CTR = CTR + 1
'Select case with a true and false visibility to draw hangman as game progresses
Select Case CTR
Case Is = 1
    Line9 = True
Case Is = 2
    Line1 = True
    Line2 = True
    Line3 = True
Case Is = 3
    Line11 = True
    Line12 = True
    Line13 = True
    Line14 = True
Case Is = 4
    Line4 = True
Case Is = 5
    Line5 = True
Case Is = 6
    Line6 = True
Case Is = 7
    Line7 = True
Case Is = 8
    Line8 = True
End Select
'Prints how many wrongs and if over 7 times, then game is over
Picfails.Print CTR
If CTR > 7 Then Picfails.Print "You Lose"
If CTR > 7 Then MsgBox "You Lose!", , "You Lose"
End Sub
Private Sub chke_Click(Index As Integer)
'Stores in program that letter is selected
chkevalue = True
'If all correct letters are checked, it will print the winning answer, otherwise it will print the letter selected in its appropriate spot
If chkivalue = True And chknzvalue = True And chkfvalue = True And chklvalue = True And chkuvalue = True And chkevalue = True And chkzvalue = True And chkavalue = True Then
picResults.Print "INFLUENZA is Correct!!", "YOU WIN!!"
End If
Picbroth.Print "_ _ _ _ _ E _ _ _"
End Sub
Private Sub chkf_Click(Index As Integer)
'Stores in program that letter is selected
chkfvalue = True
'If all correct letters are checked, it will print the winning answer, otherwise it will print the letter selected in its appropriate spot
If chkivalue = True And chknzvalue = True And chkfvalue = True And chklvalue = True And chkuvalue = True And chkevalue = True And chkzvalue = True And chkavalue = True Then
picResults.Print "INFLUENZA is Correct!!", "YOU WIN!!"
End If
Picbroth.Print "_ _ F _ _ _ _ _ _"
End Sub
Private Sub chkg_Click(Index As Integer)
'Adds up how many wrong answers and prints, if more than 7 then they lose
CTR = CTR + 1
'Select case with a true and false visibility to draw hangman as game progresses
Select Case CTR
Case Is = 1
    Line9 = True
Case Is = 2
    Line1 = True
    Line2 = True
    Line3 = True
Case Is = 3
    Line11 = True
    Line12 = True
    Line13 = True
    Line14 = True
Case Is = 4
    Line4 = True
Case Is = 5
    Line5 = True
Case Is = 6
    Line6 = True
Case Is = 7
    Line7 = True
Case Is = 8
    Line8 = True
End Select
'Prints how many wrongs and if over 7 times, then game is over
Picfails.Print CTR
If CTR > 7 Then Picfails.Print "You Lose"
If CTR > 7 Then MsgBox "You Lose!", , "You Lose"
End Sub
Private Sub chki_Click(Index As Integer)
'Stores in program that letter is selected
chkivalue = True
'If all correct letters are checked, it will print the winning answer, otherwise it will print the letter selected in its appropriate spot
If chkivalue = True And chknzvalue = True And chkfvalue = True And chklvalue = True And chkuvalue = True And chkevalue = True And chkzvalue = True And chkavalue = True Then
picResults.Print "INFLUENZA is Correct!!", "YOU WIN!!"
End If
Picbroth.Print "I _ _ _ _ _ _ _ _"
End Sub
Private Sub chkj_Click(Index As Integer)
'Adds up how many wrong answers and prints, if more than 7 then they lose
CTR = CTR + 1
'Select case with a true and false visibility to draw hangman as game progresses
Select Case CTR
Case Is = 1
    Line9 = True
Case Is = 2
    Line1 = True
    Line2 = True
    Line3 = True
Case Is = 3
    Line11 = True
    Line12 = True
    Line13 = True
    Line14 = True
Case Is = 4
    Line4 = True
Case Is = 5
    Line5 = True
Case Is = 6
    Line6 = True
Case Is = 7
    Line7 = True
Case Is = 8
    Line8 = True
End Select
'Prints how many wrongs and if over 7 times, then game is over
Picfails.Print CTR
If CTR > 7 Then Picfails.Print "You Lose"
If CTR > 7 Then MsgBox "You Lose!", , "You Lose"
End Sub
Private Sub chkk_Click(Index As Integer)
'Adds up how many wrong answers and prints, if more than 7 then they lose
CTR = CTR + 1
'Select case with a true and false visibility to draw hangman as game progresses
Select Case CTR
Case Is = 1
    Line9 = True
Case Is = 2
    Line1 = True
    Line2 = True
    Line3 = True
Case Is = 3
    Line11 = True
    Line12 = True
    Line13 = True
    Line14 = True
Case Is = 4
    Line4 = True
Case Is = 5
    Line5 = True
Case Is = 6
    Line6 = True
Case Is = 7
    Line7 = True
Case Is = 8
    Line8 = True
End Select
'Prints how many wrongs and if over 7 times, then game is over
Picfails.Print CTR
If CTR > 7 Then Picfails.Print "You Lose"
If CTR > 7 Then MsgBox "You Lose!", , "You Lose"
End Sub
Private Sub chkl_Click(Index As Integer)
'Stores in program that letter is selected
chklvalue = True
'If all correct letters are checked, it will print the winning answer, otherwise it will print the letter selected in its appropriate spot
If chkivalue = True And chknzvalue = True And chkfvalue = True And chklvalue = True And chkuvalue = True And chkevalue = True And chkzvalue = True And chkavalue = True Then
picResults.Print "INFLUENZA is Correct!!", "YOU WIN!!"
End If
Picbroth.Print "_ _ _ L _ _ _ _ _"
End Sub
Private Sub chkm_Click(Index As Integer)
'Adds up how many wrong answers and prints, if more than 7 then they lose
CTR = CTR + 1
'Select case with a true and false visibility to draw hangman as game progresses
Select Case CTR
Case Is = 1
    Line9 = True
Case Is = 2
    Line1 = True
    Line2 = True
    Line3 = True
Case Is = 3
    Line11 = True
    Line12 = True
    Line13 = True
    Line14 = True
Case Is = 4
    Line4 = True
Case Is = 5
    Line5 = True
Case Is = 6
    Line6 = True
Case Is = 7
    Line7 = True
Case Is = 8
    Line8 = True
End Select
'Prints how many wrongs and if over 7 times, then game is over
Picfails.Print CTR
If CTR > 7 Then Picfails.Print "You Lose"
If CTR > 7 Then MsgBox "You Lose!", , "You Lose"
End Sub
Private Sub chkn_Click(Index As Integer)
'Stores in program that letter is selected
chknzvalue = True
'If all correct letters are checked, it will print the winning answer, otherwise it will print the letter selected in its appropriate spot
If chkivalue = True And chknzvalue = True And chkfvalue = True And chklvalue = True And chkuvalue = True And chkevalue = True And chkzvalue = True And chkavalue = True Then
picResults.Print "INFLUENZA is Correct!!", "YOU WIN!!"
End If
Picbroth.Print "_ N _ _ _ _ N _ _"
End Sub
Private Sub chkp_Click(Index As Integer)
'Adds up how many wrong answers and prints, if more than 7 then they lose
CTR = CTR + 1
'Select case with a true and false visibility to draw hangman as game progresses
Select Case CTR
Case Is = 1
    Line9 = True
Case Is = 2
    Line1 = True
    Line2 = True
    Line3 = True
Case Is = 3
    Line11 = True
    Line12 = True
    Line13 = True
    Line14 = True
Case Is = 4
    Line4 = True
Case Is = 5
    Line5 = True
Case Is = 6
    Line6 = True
Case Is = 7
    Line7 = True
Case Is = 8
    Line8 = True
End Select
'Prints how many wrongs and if over 7 times, then game is over
Picfails.Print CTR
If CTR > 7 Then Picfails.Print "You Lose"
If CTR > 7 Then MsgBox "You Lose!", , "You Lose"
End Sub
Private Sub chkq_Click(Index As Integer)
'Adds up how many wrong answers and prints, if more than 7 then they lose
CTR = CTR + 1
'Select case with a true and false visibility to draw hangman as game progresses
Select Case CTR
Case Is = 1
    Line9 = True
Case Is = 2
    Line1 = True
    Line2 = True
    Line3 = True
Case Is = 3
    Line11 = True
    Line12 = True
    Line13 = True
    Line14 = True
Case Is = 4
    Line4 = True
Case Is = 5
    Line5 = True
Case Is = 6
    Line6 = True
Case Is = 7
    Line7 = True
Case Is = 8
    Line8 = True
End Select
'Prints how many wrongs and if over 7 times, then game is over
Picfails.Print CTR
If CTR > 7 Then Picfails.Print "You Lose"
If CTR > 7 Then MsgBox "You Lose!", , "You Lose"
End Sub
Private Sub chks_Click(Index As Integer)
'Adds up how many wrong answers and prints, if more than 7 then they lose
CTR = CTR + 1
'Select case with a true and false visibility to draw hangman as game progresses
Select Case CTR
Case Is = 1
    Line9 = True
Case Is = 2
    Line1 = True
    Line2 = True
    Line3 = True
Case Is = 3
    Line11 = True
    Line12 = True
    Line13 = True
    Line14 = True
Case Is = 4
    Line4 = True
Case Is = 5
    Line5 = True
Case Is = 6
    Line6 = True
Case Is = 7
    Line7 = True
Case Is = 8
    Line8 = True
End Select
'Prints how many wrongs and if over 7 times, then game is over
Picfails.Print CTR
If CTR > 7 Then Picfails.Print "You Lose"
If CTR > 7 Then MsgBox "You Lose!", , "You Lose"
End Sub
Private Sub chku_Click(Index As Integer)
'Stores in program that letter is selected
chkuvalue = True
'If all correct letters are checked, it will print the winning answer, otherwise it will print the letter selected in its appropriate spot
If chkivalue = True And chknzvalue = True And chkfvalue = True And chklvalue = True And chkuvalue = True And chkevalue = True And chkzvalue = True And chkavalue = True Then
picResults.Print "INFLUENZA is Correct!!", "YOU WIN!!"
End If
Picbroth.Print "_ _ _ _ U _ _ _ _"
End Sub
Private Sub chkv_Click(Index As Integer)
'Adds up how many wrong answers and prints, if more than 7 then they lose
CTR = CTR + 1
'Select case with a true and false visibility to draw hangman as game progresses
Select Case CTR
Case Is = 1
    Line9 = True
Case Is = 2
    Line1 = True
    Line2 = True
    Line3 = True
Case Is = 3
    Line11 = True
    Line12 = True
    Line13 = True
    Line14 = True
Case Is = 4
    Line4 = True
Case Is = 5
    Line5 = True
Case Is = 6
    Line6 = True
Case Is = 7
    Line7 = True
Case Is = 8
    Line8 = True
End Select
'Prints how many wrongs and if over 7 times, then game is over
Picfails.Print CTR
If CTR > 7 Then Picfails.Print "You Lose"
If CTR > 7 Then MsgBox "You Lose!", , "You Lose"
End Sub
Private Sub chkw_Click(Index As Integer)
'Adds up how many wrong answers and prints, if more than 7 then they lose
CTR = CTR + 1
'Select case with a true and false visibility to draw hangman as game progresses
Select Case CTR
Case Is = 1
    Line9 = True
Case Is = 2
    Line1 = True
    Line2 = True
    Line3 = True
Case Is = 3
    Line11 = True
    Line12 = True
    Line13 = True
    Line14 = True
Case Is = 4
    Line4 = True
Case Is = 5
    Line5 = True
Case Is = 6
    Line6 = True
Case Is = 7
    Line7 = True
Case Is = 8
    Line8 = True
End Select
'Prints how many wrongs and if over 7 times, then game is over
Picfails.Print CTR
If CTR > 7 Then Picfails.Print "You Lose"
If CTR > 7 Then MsgBox "You Lose!", , "You Lose"
End Sub
Private Sub chkx_Click(Index As Integer)
'Adds up how many wrong answers and prints, if more than 7 then they lose
CTR = CTR + 1
'Select case with a true and false visibility to draw hangman as game progresses
Select Case CTR
Case Is = 1
    Line9 = True
Case Is = 2
    Line1 = True
    Line2 = True
    Line3 = True
Case Is = 3
    Line11 = True
    Line12 = True
    Line13 = True
    Line14 = True
Case Is = 4
    Line4 = True
Case Is = 5
    Line5 = True
Case Is = 6
    Line6 = True
Case Is = 7
    Line7 = True
Case Is = 8
    Line8 = True
End Select
'Prints how many wrongs and if over 7 times, then game is over
Picfails.Print CTR
If CTR > 7 Then Picfails.Print "You Lose"
If CTR > 7 Then MsgBox "You Lose!", , "You Lose"
End Sub
Private Sub chky_Click(Index As Integer)
'Adds up how many wrong answers and prints, if more than 7 then they lose
CTR = CTR + 1
'Select case with a true and false visibility to draw hangman as game progresses
Select Case CTR
Case Is = 1
    Line9 = True
Case Is = 2
    Line1 = True
    Line2 = True
    Line3 = True
Case Is = 3
    Line11 = True
    Line12 = True
    Line13 = True
    Line14 = True
Case Is = 4
    Line4 = True
Case Is = 5
    Line5 = True
Case Is = 6
    Line6 = True
Case Is = 7
    Line7 = True
Case Is = 8
    Line8 = True
End Select
'Prints how many wrongs and if over 7 times, then game is over
Picfails.Print CTR
If CTR > 7 Then Picfails.Print "You Lose"
If CTR > 7 Then MsgBox "You Lose!", , "You Lose"
End Sub
Private Sub chkz_Click(Index As Integer)
'Stores in program that letter is selected
chkzvalue = True
'If all correct letters are checked, it will print the winning answer, otherwise it will print the letter selected in its appropriate spot
If chkivalue = True And chknzvalue = True And chkfvalue = True And chklvalue = True And chkuvalue = True And chkevalue = True And chkzvalue = True And chkavalue = True Then
picResults.Print "INFLUENZA is Correct!!", "YOU WIN!!"
End If
Picbroth.Print "_ _ _ _ _ _ _ Z _"
End Sub

Private Sub Command1_Click()
Hangman.Show
Hangman2.Hide
End Sub

Private Sub Command2_Click()
Hangman.Hide
Swing.Show
End Sub
