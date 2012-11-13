VERSION 5.00
Begin VB.Form frmGameBoard2 
   BackColor       =   &H8000000D&
   Caption         =   "Double Jeopardy Round"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   ScaleHeight     =   8505
   ScaleWidth      =   11070
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdC600 
      BackColor       =   &H00800080&
      Caption         =   "$600"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton cmdB600 
      BackColor       =   &H00C0C000&
      Caption         =   "$600"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton cmdA600 
      BackColor       =   &H0000FF00&
      Caption         =   "$600"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton cmdC400 
      BackColor       =   &H00800080&
      Caption         =   "$400"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton cmdB400 
      BackColor       =   &H00C0C000&
      Caption         =   "$400"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton cmdA400 
      BackColor       =   &H0000FF00&
      Caption         =   "$400"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton cmdC200 
      BackColor       =   &H00800080&
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
      Height          =   975
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton cmdB200 
      BackColor       =   &H00C0C000&
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
      Height          =   975
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit Game without Saving"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2160
      TabIndex        =   11
      Top             =   6960
      Width           =   1695
   End
   Begin VB.PictureBox picPicture 
      Height          =   3495
      Left            =   240
      Picture         =   "frmGameBoard2.frx":0000
      ScaleHeight     =   3435
      ScaleWidth      =   3555
      TabIndex        =   10
      Top             =   3240
      Width           =   3615
   End
   Begin VB.CommandButton cmdFinal 
      Caption         =   "Go to the Final Jeopardy! Round"
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
      Left            =   7200
      TabIndex        =   9
      Top             =   7200
      Width           =   3135
   End
   Begin VB.CommandButton cmdA200 
      BackColor       =   &H0000FF00&
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
      Height          =   975
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Main Menu without Saving"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   4
      Top             =   6960
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
      Left            =   1080
      ScaleHeight     =   435
      ScaleWidth      =   1275
      TabIndex        =   3
      Top             =   2520
      Width           =   1335
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
      Left            =   600
      ScaleHeight     =   435
      ScaleWidth      =   2355
      TabIndex        =   1
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C000&
      Caption         =   "                        Tough Trivia"
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
      Left            =   6360
      TabIndex        =   7
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H00800080&
      Caption         =   "      Sports                 at                   SJU"
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
      Left            =   8640
      TabIndex        =   6
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FF00&
      Caption         =   "Where you           at on           campus?!"
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
      Left            =   4080
      TabIndex        =   5
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label2 
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
      Left            =   960
      TabIndex        =   2
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label1 
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
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
End
Attribute VB_Name = "frmGameBoard2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: CSJ/SJU Jeopardy
'Form Name: frmGameBoard2
'Authors: Emma Jaynes, Linsday Havlik, Brooke Beyer
'Date Written: 11/02/08
'Objective: This form is similar to the 1st game board but has more difficult questions worth more points.  Has links to main menu,
'   a quit option, and, upon completion of all questions, a link to the final Jeopardy form
'Other Comments: The total and contestant names are carried into this form from previous form. It has links to main menu, a quit option,
'   and, upon completion of all questions, a link to the final Jeopardy form.

Option Explicit
Dim CTR As Integer

Private Sub cmdA200_Click()
'asks question, if correct adds labeled points to total, if incorrect, deducts points from total
'counts question as complete and disables button from being clicked again
CTR = CTR + 1
cmdA200.Enabled = False

picTotal.Cls

X = InputBox("The SJU bookstore is located in this building. What is...")

If LCase(X) = "sexton" Then
    MsgBox ("That is correct!")
    Total = Total + 200
Else: MsgBox ("Sorry, that is incorrect. The correct answer is Sexton")
    Total = Total - 200
End If

picTotal.Print FormatCurrency(Total, 0)

If CTR = 9 Then
    cmdFinal.Enabled = True
End If
    
End Sub

Private Sub cmdA400_Click()
'asks question, if correct adds labeled points to total, if incorrect, deducts points from total
'counts question as complete and disables button from being clicked again
CTR = CTR + 1

cmdA400.Enabled = False

picTotal.Cls

X = InputBox("The Bennies get their mail in this building. What is...")

If LCase(X) = "mary hall commons" Then
    MsgBox ("That is correct!")
    Total = Total + 400
Else: MsgBox ("Sorry, that is incorrect. The correct answer is Mary Hall Commons")
    Total = Total - 400
End If

picTotal.Print FormatCurrency(Total, 0)

If CTR = 9 Then
    cmdFinal.Enabled = True
End If
End Sub

Private Sub cmdA600_Click()
'asks question, if correct adds labeled points to total, if incorrect, deducts points from total
'counts question as complete and disables button from being clicked again
CTR = CTR + 1

cmdA600.Enabled = False

picTotal.Cls

X = InputBox("This is the name of the SJU library. What is...")

If LCase(X) = "alcuin" Then
    MsgBox ("That is correct!")
    Total = Total + 600
Else: MsgBox ("Sorry, that is incorrect. The correct answer is Alcuin.")
    Total = Total - 600
End If

picTotal.Print FormatCurrency(Total, 0)

If CTR = 9 Then
    cmdFinal.Enabled = True
End If
End Sub

Private Sub cmdB200_Click()
'asks question, if correct adds labeled points to total, if incorrect, deducts points from total
'counts question as complete and disables button from being clicked again
CTR = CTR + 1

cmdB200.Enabled = False


picTotal.Cls

X = InputBox("This is the number of states represented by the student body. What is...")

If X = "42" Then
    MsgBox ("That is correct!")
    Total = Total + 200
Else: MsgBox ("Sorry, that is incorrect. The correct answer is 42.")
    Total = Total - 200
End If

picTotal.Print FormatCurrency(Total, 0)

If CTR = 9 Then
    cmdFinal.Enabled = True
End If

End Sub

Private Sub cmdB400_Click()
'asks question, if correct adds labeled points to total, if incorrect, deducts points from total
'counts question as complete and disables button from being clicked again
CTR = CTR + 1

cmdB400.Enabled = False

picTotal.Cls

X = InputBox("This is the number of acres the share campuses share. What is...")

If X = "3200" Then
    MsgBox ("That is correct!")
    Total = Total + 400
Else: MsgBox ("Sorry, that is incorrect. The correct answer is 3200.")
    Total = Total - 400
End If

picTotal.Print FormatCurrency(Total, 0)

If CTR = 9 Then
    cmdFinal.Enabled = True
End If

End Sub

Private Sub cmdB600_Click()
'asks question, if correct adds labeled points to total, if incorrect, deducts points from total
'counts question as complete and disables button from being clicked again
CTR = CTR + 1

cmdB600.Enabled = False

picTotal.Cls

X = InputBox("Replacing Brother Dietrich, this is the man who was named interim President of SJU. Who is...")

If LCase(X) = "dan whalen" Then
    MsgBox ("That is correct!")
    Total = Total + 600
Else: MsgBox ("Sorry, that is incorrect. The correct answer is Dan Whalen.")
    Total = Total - 600
End If

picTotal.Print FormatCurrency(Total, 0)

If CTR = 9 Then
    cmdFinal.Enabled = True
End If

End Sub

Private Sub cmdBack_Click()
'hides the second game board, brings us back to the main menu while clearing the contestant name and total score
    frmGameBoard2.Hide
    frmMainMenu.Show
    picContestant.Cls
    picTotal.Cls
    
End Sub

Private Sub cmdC200_Click()
'asks question, if correct adds labeled points to total, if incorrect, deducts points from total
'counts question as complete and disables button from being clicked again
CTR = CTR + 1

cmdC200.Enabled = False

picTotal.Cls

X = InputBox("The number of club sports SJU has. What is...")

If X = "17" Then
    MsgBox ("That is correct!")
    Total = Total + 200
Else: MsgBox ("Sorry, that is incorrect. The correct answer is 17.")
    Total = Total - 200
End If

picTotal.Print FormatCurrency(Total, 0)

If CTR = 9 Then
    cmdFinal.Enabled = True
End If

End Sub

Private Sub cmdC400_Click()
'asks question, if correct adds labeled points to total, if incorrect, deducts points from total
'counts question as complete and disables button from being clicked again
CTR = CTR + 1

cmdC400.Enabled = False

picTotal.Cls

X = InputBox("The number of seasons John Gagliardi has coached the Johnnie football team. What is...")

If X = "55" Then
    MsgBox ("That is correct!")
    Total = Total + 400
Else: MsgBox ("Sorry, that is incorrect. The correct answer is 55.")
    Total = Total - 400
End If

picTotal.Print FormatCurrency(Total, 0)

If CTR = 9 Then
    cmdFinal.Enabled = True
End If

End Sub

Private Sub cmdC600_Click()
'asks question, if correct adds labeled points to total, if incorrect, deducts points from total
'counts question as complete and disables button from being clicked again
CTR = CTR + 1

cmdC600.Enabled = False

picTotal.Cls

MsgBox ("You have chosen a Daily Double. This question is now worth $1200!")

X = InputBox("The number of total MIAC championships SJU has won. What is...")

If X = "120" Then
    MsgBox ("That is correct!")
    Total = Total + 600
Else: MsgBox ("Sorry, that is incorrect. The correct answer is 120.")
    Total = Total - 600
End If

picTotal.Print FormatCurrency(Total, 0)

If CTR = 9 Then
    cmdFinal.Enabled = True
End If

End Sub

Private Sub cmdFinal_Click()
'if the total score isn't <= 0 this takes us to the final jeopardy form and carries over the contestant name and score from the previous 2 rounds

If Total <= 0 Then
    MsgBox ("Sorry, you do not have a positive score, the game is over for you!")
    frmMainMenu.Show
    frmGameBoard2.Hide
Else
    frmFinal.Show
    frmGameBoard2.Hide
End If

frmFinal.picContestant.Print Contestant
frmFinal.picTotal.Print FormatCurrency(Total, 0)

End Sub

Private Sub cmdQuit_Click()
'quits the project
    End
End Sub


