VERSION 5.00
Begin VB.Form Hangman1 
   Caption         =   "Form1"
   ClientHeight    =   9735
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14100
   LinkTopic       =   "Form1"
   Picture         =   "formhangman.frx":0000
   ScaleHeight     =   9735
   ScaleWidth      =   14100
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton Command2 
      BackColor       =   &H000080FF&
      Caption         =   "If You Lose, Click Here"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   34
      Top             =   4320
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "Go Back to Hangman Page"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   6120
      TabIndex        =   33
      Top             =   6120
      Width           =   4935
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
      Left            =   11280
      TabIndex        =   29
      Top             =   5040
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
      Left            =   10440
      TabIndex        =   28
      Top             =   5040
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
      Left            =   4680
      TabIndex        =   27
      Top             =   5040
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
      Left            =   2760
      TabIndex        =   26
      Top             =   5040
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
      Left            =   8520
      TabIndex        =   25
      Top             =   5040
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
      Left            =   9240
      TabIndex        =   24
      Top             =   4200
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
      Left            =   7560
      TabIndex        =   23
      Top             =   5040
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
      Left            =   6600
      TabIndex        =   22
      Top             =   5040
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
      Left            =   9480
      TabIndex        =   21
      Top             =   5040
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
      Left            =   4440
      TabIndex        =   20
      Top             =   4200
      Width           =   615
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
      Left            =   3480
      TabIndex        =   19
      Top             =   4200
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
      Left            =   9600
      ScaleHeight     =   3075
      ScaleWidth      =   2355
      TabIndex        =   18
      Top             =   840
      Width           =   2415
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
      Left            =   9000
      TabIndex        =   17
      Top             =   3360
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
      Left            =   5400
      TabIndex        =   16
      Top             =   4200
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
      Left            =   6360
      TabIndex        =   15
      Top             =   4200
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
      Left            =   7320
      TabIndex        =   14
      Top             =   4200
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
      Left            =   2520
      TabIndex        =   13
      Top             =   4200
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
      Left            =   3720
      TabIndex        =   12
      Top             =   5040
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
      Left            =   8280
      TabIndex        =   11
      Top             =   4200
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
      Left            =   5640
      TabIndex        =   10
      Top             =   5040
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
      Left            =   6120
      TabIndex        =   9
      Top             =   3360
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
      Left            =   7080
      TabIndex        =   8
      Top             =   3360
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
      Left            =   8040
      TabIndex        =   7
      Top             =   3360
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
      Left            =   5160
      TabIndex        =   6
      Top             =   3360
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
      Left            =   4200
      TabIndex        =   5
      Top             =   3360
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
      Left            =   3240
      TabIndex        =   4
      Top             =   3360
      Width           =   615
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
      Left            =   2280
      TabIndex        =   3
      Top             =   3360
      Width           =   615
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
      Height          =   3375
      Left            =   120
      ScaleHeight     =   3315
      ScaleWidth      =   1755
      TabIndex        =   1
      Top             =   840
      Width           =   1815
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
      Height          =   1095
      Left            =   3000
      ScaleHeight     =   1035
      ScaleWidth      =   4875
      TabIndex        =   0
      Top             =   1920
      Width           =   4935
   End
   Begin VB.Line Line14 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   1560
      X2              =   2280
      Y1              =   7320
      Y2              =   7320
   End
   Begin VB.Line Line13 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   1560
      X2              =   1560
      Y1              =   6840
      Y2              =   7320
   End
   Begin VB.Line Line12 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   2280
      X2              =   2280
      Y1              =   6840
      Y2              =   7320
   End
   Begin VB.Line Line11 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   1560
      X2              =   2280
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Line Line9 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   120
      X2              =   1320
      Y1              =   8760
      Y2              =   8760
   End
   Begin VB.Line Line8 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   2160
      X2              =   1320
      Y1              =   7560
      Y2              =   7560
   End
   Begin VB.Line Line7 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   1920
      X2              =   2640
      Y1              =   7560
      Y2              =   7560
   End
   Begin VB.Line Line6 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   1920
      X2              =   1560
      Y1              =   7920
      Y2              =   8520
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   1920
      X2              =   2520
      Y1              =   7920
      Y2              =   8400
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   1920
      X2              =   1920
      Y1              =   7320
      Y2              =   7920
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   1920
      X2              =   1920
      Y1              =   6480
      Y2              =   6840
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   720
      X2              =   1920
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   720
      X2              =   720
      Y1              =   8760
      Y2              =   6480
   End
   Begin VB.Image imghangman 
      Height          =   3765
      Left            =   120
      Picture         =   "formhangman.frx":1CF50E
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
      Left            =   4560
      TabIndex        =   32
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FFFF&
      Caption         =   "Correct Letters in appropriate order"
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
      Left            =   9720
      TabIndex        =   31
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Your Word:     ______  ______  ______  ______  ______          Hint-> It is drinkable!"
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
      Left            =   2040
      TabIndex        =   30
      Top             =   360
      Width           =   7095
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
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "Hangman1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Healthy Living
'Form Hangman1
'Joel Coleman
'March 29, 2008
'To write a game which deals with words associated with health
'I used many true/false statements in order to complete my task
Option Explicit
'Declaring variables for entire form
Dim CTR As Single, chkbvalue As Boolean, chkrvalue As Boolean, chkovalue As Boolean
Dim chktvalue As Boolean, chkhvalue As Boolean

Private Sub chkb_click(Index As Integer)
'Stores in program that letter is selected
chkbvalue = True
'If all correct letters are checked, it will print the winning answer, otherwise it will print the letter selected in its appropriate spot
If chkrvalue = True And chkovalue = True And chktvalue = True And chkhvalue = True And chkbvalue = True Then
picResults.Print "BROTH", "YOU WIN!!"
End If
Picbroth.Print "B _ _ _ _"
End Sub

Private Sub chkh_Click(Index As Integer)
'Stores in program that letter is selected
chkhvalue = True
'If all correct letters are checked, it will print the winning answer, otherwise it will print the letter selected in its appropriate spot
If chkrvalue = True And chkovalue = True And chktvalue = True And chkhvalue = True And chkbvalue = True Then
picResults.Print "BROTH", "YOU WIN!!"
End If
Picbroth.Print "_ _ _ _ H"
End Sub

Private Sub chko_Click(Index As Integer)
'Stores in program that letter is selected
chkovalue = True
'If all correct letters are checked, it will print the winning answer, otherwise it will print the letter selected in its appropriate spot
If chkrvalue = True And chkovalue = True And chktvalue = True And chkhvalue = True And chkbvalue = True Then
picResults.Print "BROTH", "YOU WIN!!"
End If
Picbroth.Print "_ _ O _ _"
End Sub

Private Sub chkr_Click(Index As Integer)
'Stores in program that letter is selected
chkrvalue = True
'If all correct letters are checked, it will print the winning answer, otherwise it will print the letter selected in its appropriate spot
If chkrvalue = True And chkovalue = True And chktvalue = True And chkhvalue = True And chkbvalue = True Then
picResults.Print "BROTH", "YOU WIN!!"
End If
Picbroth.Print "_ R _ _ _"
End Sub

Private Sub chkt_Click(Index As Integer)
'Stores in program that letter is selected
chktvalue = True
'If all correct letters are checked, it will print the winning answer, otherwise it will print the letter selected in its appropriate spot
If chkrvalue = True And chkovalue = True And chktvalue = True And chkhvalue = True And chkbvalue = True Then
picResults.Print "BROTH", "YOU WIN!!"
End If
Picbroth.Print "_ _ _ T _"
End Sub
Private Sub chka_Click(Index As Integer)
'Adds up how many wrong answers and prints
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
Private Sub chkc_Click(Index As Integer)
'Adds up how many wrong answers and prints
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
'Adds up how many wrong answers and prints
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
'Adds up how many wrong answers and prints
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
Private Sub chkf_Click(Index As Integer)
'Adds up how many wrong answers and prints
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
Private Sub chkg_Click(Index As Integer)
'Adds up how many wrong answers and prints
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
'Adds up how many wrong answers and prints
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
Private Sub chkj_Click(Index As Integer)
'Adds up how many wrong answers and prints
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
'Adds up how many wrong answers and prints
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
'Adds up how many wrong answers and prints
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
Private Sub chkm_Click(Index As Integer)
'Adds up how many wrong answers and prints
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
'Adds up how many wrong answers and prints
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
Private Sub chkp_Click(Index As Integer)
'Adds up how many wrong answers and prints
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
'Adds up how many wrong answers and prints
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
'Adds up how many wrong answers and prints
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
'Adds up how many wrong answers and prints
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
Private Sub chkv_Click(Index As Integer)
'Adds up how many wrong answers and prints
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
'Adds up how many wrong answers and prints
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
'Adds up how many wrong answers and prints
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
'Adds up how many wrong answers and prints
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
'Adds up how many wrong answers and prints
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

Private Sub Command1_Click()
Hangman1.Hide
Hangman.Show
End Sub
Private Sub Command2_Click()
Hangman1.Hide
Swing.Show
End Sub
