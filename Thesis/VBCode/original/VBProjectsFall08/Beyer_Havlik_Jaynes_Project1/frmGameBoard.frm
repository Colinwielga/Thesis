VERSION 5.00
Begin VB.Form frmGameBoard 
   BackColor       =   &H8000000D&
   Caption         =   "Game Board"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12300
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   12300
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdC300 
      BackColor       =   &H0000FFFF&
      Caption         =   "$300"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5880
      Width           =   1695
   End
   Begin VB.CommandButton cmdB100 
      BackColor       =   &H000080FF&
      Caption         =   "$100"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton cmdA100 
      BackColor       =   &H000000FF&
      Caption         =   "$100"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton cmdDouble 
      Caption         =   "Go to the Double Jeopardy! Round"
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8760
      TabIndex        =   16
      Top             =   7440
      Width           =   3255
   End
   Begin VB.PictureBox picPicture 
      Height          =   3615
      Left            =   240
      Picture         =   "frmGameBoard.frx":0000
      ScaleHeight     =   3555
      ScaleWidth      =   3675
      TabIndex        =   15
      Top             =   3240
      Width           =   3735
   End
   Begin VB.CommandButton cmdMainMenu 
      Caption         =   "Go back to Main Menu without saving"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   14
      Top             =   7080
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit game without saving"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2280
      TabIndex        =   13
      Top             =   7080
      Width           =   1695
   End
   Begin VB.PictureBox picTotal 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      ScaleHeight     =   435
      ScaleWidth      =   1035
      TabIndex        =   12
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdC200 
      BackColor       =   &H0000FFFF&
      Caption         =   "$200"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4200
      Width           =   1695
   End
   Begin VB.CommandButton cmdC100 
      BackColor       =   &H0000FFFF&
      Caption         =   "$100"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton cmdB300 
      BackColor       =   &H000080FF&
      Caption         =   "$300"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5880
      Width           =   1695
   End
   Begin VB.CommandButton cmdB200 
      BackColor       =   &H000080FF&
      Caption         =   "$200"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4200
      Width           =   1695
   End
   Begin VB.CommandButton cmdA300 
      BackColor       =   &H000000FF&
      Caption         =   "$300"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5880
      Width           =   1695
   End
   Begin VB.CommandButton cmdA200 
      BackColor       =   &H000000FF&
      Caption         =   "$200"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4200
      Width           =   1695
   End
   Begin VB.PictureBox picContestant 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      ScaleHeight     =   435
      ScaleWidth      =   2835
      TabIndex        =   4
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label lblTotal 
      BackColor       =   &H00FF00FF&
      Caption         =   "    Total:"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   11
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label lbl5 
      BackColor       =   &H00FF00FF&
      Caption         =   "Contestant:"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lbl4 
      BackColor       =   &H000080FF&
      Caption         =   "          A                 "" Major""          Decision"
      BeginProperty Font 
         Name            =   "Jokerman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7200
      TabIndex        =   2
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label lbl2 
      BackColor       =   &H000000FF&
      Caption         =   "                               Abbrev's"
      BeginProperty Font 
         Name            =   "Jokerman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4800
      TabIndex        =   1
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label lbl1 
      BackColor       =   &H0000FFFF&
      Caption         =   " Ballin' With     the Blazers"
      BeginProperty Font 
         Name            =   "Jokerman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   9600
      TabIndex        =   0
      Top             =   720
      Width           =   1935
   End
End
Attribute VB_Name = "frmGameBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: CSB/SJU Jeopardy
'Form Name: frmGameBoard
'Authors: Emma Jaynes, Lindsay Havlik, Brooke Beyer
'Date Written: 11/02/08
'Objective: This form displays our 1st round game board with 9 questions.  Users click on buttons and answer corresponding quiestions to increase or decrease their scores.
'Comments: The contestant name is entered here and is used in each of the forms following the 1st. It has links to main menu, a quit option, and, upon completion of all questions, a link to the second game board form.

Option Explicit
Dim CTR As Integer

Private Sub cmdA100_Click()
'asks question, if correct adds labeled points to total, if incorrect, deducts points from total
'counts question as complete and disables button from being clicked again

CTR = CTR + 1

cmdA100.Enabled = False

picTotal.Cls

X = InputBox("HCC: What is the...")

If LCase(X) = "haehn campus center" Then
    MsgBox ("That is correct!")
    Total = Total + 100
Else: MsgBox ("Sorry, that is incorrect. The correct answer is Haehn Campus Center")
    Total = Total - 100
End If

picTotal.Print FormatCurrency(Total, 0)

If CTR = 9 Then
    cmdDouble.Enabled = True
End If

End Sub


Private Sub cmdA200_Click()
'asks question, if correct adds labeled points to total, if incorrect, deducts points from total
'counts question as complete and disables button from being clicked again
CTR = CTR + 1

cmdA200.Enabled = False

picTotal.Cls

X = InputBox("Pengl: What is the...")

If LCase(X) = "peter engel science center" Then
    MsgBox ("That is correct!")
    Total = Total + 200
Else: MsgBox ("Sorry, that is incorrect. The correct answer is Peter Engel Science Center.")
    Total = Total - 200
End If

picTotal.Print FormatCurrency(Total, 0)

If CTR = 9 Then
    cmdDouble.Enabled = True
End If

End Sub

Private Sub cmdA300_Click()
'asks question, if correct adds labeled points to total, if incorrect, deducts points from total
'counts question as complete and disables button from being clicked again
CTR = CTR + 1

picTotal.Cls

cmdA300.Enabled = False

X = InputBox("HAB: What is the...")

If LCase(X) = "henrietta academic building" Then
    MsgBox ("That is correct!")
    Total = Total + 300
Else: MsgBox ("Sorry, that is incorrect. The correct answer is Henrieta Academic Building.")
    Total = Total - 300
End If

picTotal.Print FormatCurrency(Total, 0)

If CTR = 9 Then
    cmdDouble.Enabled = True
End If

End Sub

Private Sub cmdB100_Click()
'asks question, if correct adds labeled points to total, if incorrect, deducts points from total
'counts question as complete and disables button from being clicked again
CTR = CTR + 1

cmdB100.Enabled = False

picTotal.Cls

X = InputBox("If you like acting, perhaps this should be your major! What is...")

If LCase(X) = "theater" Then
    MsgBox ("That is correct!")
    Total = Total + 100
Else: MsgBox ("Sorry, that is incorrect. The correct answer is Theater.")
    Total = Total - 100
End If

picTotal.Print FormatCurrency(Total, 0)

If CTR = 9 Then
    cmdDouble.Enabled = True
End If

End Sub

Private Sub cmdB200_Click()
'asks question, if correct adds labeled points to total, if incorrect, deducts points from total
'counts question as complete and disables button from being clicked again
CTR = CTR + 1

cmdB200.Enabled = False

picTotal.Cls

MsgBox ("You have chosen the Daily Double! This question is now worth $400!")

X = InputBox("This major has 3 different concentrations for its students to choose from: the Finance, Traditional, or CPA. What is...")

If LCase(X) = "accounting" Then
    MsgBox ("That is correct!")
    Total = Total + 400
Else: MsgBox ("Sorry, that is incorrect. The correct answer is Accounting.")
    Total = Total - 400
End If

picTotal.Print FormatCurrency(Total, 0)

If CTR = 9 Then
    cmdDouble.Enabled = True
End If

End Sub

Private Sub cmdB300_Click()
'asks question, if correct adds labeled points to total, if incorrect, deducts points from total
'counts question as complete and disables button from being clicked again
CTR = CTR + 1

cmdB300.Enabled = False

picTotal.Cls

X = InputBox("If you like the game you're playing right now and would like to learn how to create it, this is the major for you! What is...")

If LCase(X) = "computer science" Then
    MsgBox ("That is correct!")
    Total = Total + 300
Else: MsgBox ("Sorry, that is incorrect. The correct answer is Computer Science.")
    Total = Total - 300
End If

picTotal.Print FormatCurrency(Total, 0)

If CTR = 9 Then
    cmdDouble.Enabled = True
End If

End Sub

Private Sub cmdC100_Click()
'asks question, if correct adds labeled points to total, if incorrect, deducts points from total
'counts question as complete and disables button from being clicked again
CTR = CTR + 1

cmdC100.Enabled = False

picTotal.Cls

X = InputBox("This Blazer sports team has won the MIAC for the past 3 seasons. What is...")

If LCase(X) = "basketball" Then
    MsgBox ("That is correct!")
    Total = Total + 100
Else: MsgBox ("Sorry, that is incorrect. The correct answer is Basketball.")
    Total = Total - 100
End If

picTotal.Print FormatCurrency(Total, 0)

If CTR = 9 Then
    cmdDouble.Enabled = True
End If

End Sub

Private Sub cmdC200_Click()
'asks question, if correct adds labeled points to total, if incorrect, deducts points from total
'counts question as complete and disables button from being clicked again
CTR = CTR + 1

cmdC200.Enabled = False

picTotal.Cls

X = InputBox("The gymnasium where the Blazer Basketball and Volleyball teams play home games. What is...")

If LCase(X) = "claire lynch hall" Then
    MsgBox ("That is correct!")
    Total = Total + 200
Else: MsgBox ("Sorry, that is incorrect. The correct answer is Claire Lynch Hall.")
    Total = Total - 200
End If

picTotal.Print FormatCurrency(Total, 0)

If CTR = 9 Then
    cmdDouble.Enabled = True
End If

End Sub

Private Sub cmdC300_Click()
'asks question, if correct adds labeled points to total, if incorrect, deducts points from total
'counts question as complete and disables button from being clicked again
CTR = CTR + 1

cmdC300.Enabled = False

picTotal.Cls

X = InputBox("In 2007, she broke the St. Ben's set assist record for Volleyball. Who is...")

If LCase(X) = "beth hanson" Then
    MsgBox ("That is correct!")
    Total = Total + 300
Else: MsgBox ("Sorry, that is incorrect. The correct answer is Beth Hanson.")
    Total = Total - 300
End If

picTotal.Print FormatCurrency(Total, 0)

If CTR = 9 Then
    cmdDouble.Enabled = True
End If
End Sub

Private Sub cmdDouble_Click()
'hides the first game board and brings us to the second round
frmGameBoard.Hide
frmGameBoard2.Show

'carries the contestant name and total score into the second round of the game and disables the button leading us to the final jeopardy round
frmGameBoard2.picContestant.Print Contestant
frmGameBoard2.picTotal.Print FormatCurrency(Total, 0)
frmGameBoard2.cmdFinal.Enabled = False

frmFinal.cmdClue.Enabled = False

End Sub

Private Sub cmdMainMenu_Click()
'hides the game board, brings us back to the main menu while clearing the contestant name and total score
frmGameBoard.Hide
frmMainMenu.Show

picContestant.Cls
picTotal.Cls

End Sub

Private Sub cmdQuit_Click()
'quits the game
End

End Sub


