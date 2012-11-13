VERSION 5.00
Begin VB.Form frmScreen2 
   Caption         =   "Puzzle"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7500
   LinkTopic       =   "Form2"
   Picture         =   "WheelFortune.frx":0000
   ScaleHeight     =   5595
   ScaleWidth      =   7500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H80000009&
      Caption         =   "Add Up Total"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton cmdSolve 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Ready to Solve?"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton cmdEnd 
      BackColor       =   &H00FF8080&
      Caption         =   "End Game"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   0
      Width           =   735
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5520
      ScaleHeight     =   435
      ScaleWidth      =   1755
      TabIndex        =   14
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton cmdInPuzzle 
      BackColor       =   &H0080FFFF&
      Caption         =   "Is Your Letter In The Puzzle?"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CommandButton cmdVowel 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Or Buy A Vowel!"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton cmdSpin 
      BackColor       =   &H00FFFF00&
      Caption         =   """Spin""!"
      BeginProperty Font 
         Name            =   "JazzText"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox txtTopic 
      BackColor       =   &H00800080&
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2040
      TabIndex        =   10
      Text            =   "FICTIONAL CHARACTER"
      Top             =   360
      Width           =   3255
   End
   Begin VB.CommandButton cmdD2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "9"
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton cmdD 
      BackColor       =   &H00FFFFFF&
      Caption         =   "8"
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton cmdU 
      BackColor       =   &H00FFFFFF&
      Caption         =   "7"
      Height          =   375
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton cmdE2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "4"
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton cmdR 
      BackColor       =   &H00FFFFFF&
      Caption         =   "5"
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton cmdM 
      BackColor       =   &H00FFFFFF&
      Caption         =   "3"
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton cmdF 
      BackColor       =   &H00FFFFFF&
      Caption         =   "6"
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton cmdL 
      BackColor       =   &H00FFFFFF&
      Caption         =   "2"
      Height          =   375
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton cmdE 
      BackColor       =   &H00FFFFFF&
      Caption         =   "1"
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1440
      Width           =   255
   End
   Begin VB.PictureBox picName 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3000
      ScaleHeight     =   195
      ScaleWidth      =   1515
      TabIndex        =   0
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Image Image4 
      Height          =   435
      Left            =   3960
      Picture         =   "WheelFortune.frx":9DB1
      Top             =   1800
      Width           =   1635
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   3960
      Top             =   1440
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   420
      Left            =   2520
      Picture         =   "WheelFortune.frx":1009E
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   2520
      Picture         =   "WheelFortune.frx":12408
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Image Image7 
      Height          =   435
      Left            =   1800
      Picture         =   "WheelFortune.frx":1476F
      Top             =   1800
      Width           =   690
   End
   Begin VB.Image Image6 
      Height          =   435
      Left            =   1800
      Picture         =   "WheelFortune.frx":1A2B6
      Top             =   1440
      Width           =   690
   End
   Begin VB.Image Image5 
      Height          =   435
      Left            =   3960
      Picture         =   "WheelFortune.frx":1FDFD
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   375
   End
End
Attribute VB_Name = "frmScreen2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Wheel of Fortune!(WheelofFortune.vbp)
'Form name: frmScreen2(WheelFortune.frm); Form caption: Puzzle
'Author: Maria Zipp
'Date written: 1st November, 2006
'Form Objective: this is the main form that gives different options:
'               "spin", buy a vowel, shows results of user's guess
'               and shows the player's totals. When the user guesses
'               a letter (in frmWheel), this form takes that letter
'               and searches for it in the answer (String function: InStr)
'               then displays which button to puch to reveal the letter guessed.

Private Sub cmdAdd_Click()
    'adds up total from spins and subtracts from buying vowels
    picName.Cls
    picName.Print nombre; Tab(10); FormatCurrency(total, 0)
End Sub

Private Sub cmdD_Click()
    'hides button when clicked
    cmdD.Visible = False
End Sub

Private Sub cmdD2_Click()
    cmdD2.Visible = False
End Sub

Private Sub cmdE_Click()
    cmdE.Visible = False
End Sub

Private Sub cmdE2_Click()
    cmdE2.Visible = False
End Sub

Private Sub cmdEnd_Click()
    End
End Sub

Private Sub cmdF_Click()
    cmdF.Visible = False
End Sub

Private Sub cmdInPuzzle_Click()
    'clears results picbox to display position of letter in puzzle
    picResults.Cls
    position = InStr("ELMERFUDD", cletter)
    If position > 0 Then
        picResults.Print "YES!!! Click on button"
        picResults.Print "number " & position
    Else
        picResults.Print "Sorry, your letter is"
        picResults.Print "not in the puzzle!"
    End If
End Sub

Private Sub cmdL_Click()
    cmdL.Visible = False
End Sub

Private Sub cmdM_Click()
    cmdM.Visible = False
End Sub

Private Sub cmdR_Click()
    cmdR.Visible = False
End Sub

Private Sub cmdSolve_Click()
    Dim solved As String
    'this button gets the guess from the user via inputbox.
    'the answer must be in caps lock, then a message will display
    'telling the user if he/she guessed right or not.
    solved = InputBox("Type guess in CAPS", "Drum Roll...")
    If solved = "ELMER FUDD" Then
        MsgBox "Congratulations!!! You Solved the Puzzle!!", , "Yea!"
        cmdE.Visible = False
        cmdL.Visible = False
        cmdM.Visible = False
        cmdE2.Visible = False
        cmdR.Visible = False
        cmdF.Visible = False
        cmdU.Visible = False
        cmdD.Visible = False
        cmdD2.Visible = False
    Else
        MsgBox "Sorry, try again!", , "Oops!"
    End If
        
End Sub

Private Sub cmdSpin_Click()
    'this just hides the main form and shows the wheel
    
    frmScreen2.Visible = False
    frmWheel.Visible = True
End Sub

Private Sub cmdU_Click()
    cmdU.Visible = False
End Sub

Private Sub cmdVowel_Click()
    picName.Print nombre; Tab(10); FormatCurrency(total, 0)
    found = False
    counter = 0
    If total > 150 Then
        cvowel = InputBox("Enter a vowel in caps", "Enter")
        'searches alphabet to verify vowel
        Do Until found = True Or counter > 5
            counter = counter + 1
            If vowelArray(counter) = cvowel Then
                found = True
            End If
        Loop
        'if vowel is verified, then subtract $150 from users total
        'and shows user what position(which button on puzzle) to
        'push in the picturebox
        If found = True Then
            total = total - 150
            position = InStr("ELMERFUDD", cvowel)
            picResults.Cls
            picResults.Print "YES!!! Click on button"
            picResults.Print "number(s) " & position
        Else
            picResults.Print "Your vowel is not in the puzzle"
        End If
    Else
        MsgBox "You need more money! Spin!", , "Uh-Oh"
    End If
End Sub


