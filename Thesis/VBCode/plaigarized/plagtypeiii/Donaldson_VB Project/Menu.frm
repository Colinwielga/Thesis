VERSION 5.00
Begin VB.Form frmMain
   BackColor       =   &H00400040&
   Caption         =   "Main"
   ClientHeight    =   12390
   ClientLeft      =   5280
   ClientTop       =   885
   ClientWidth     =   13560
   LinkTopic       =   "Form1"
   Picture         =   "Menu.frx":0000
   ScaleHeight     =   12390
   ScaleWidth      =   13560
   Begin VB.CommandButton cmdStart
      BackColor       =   &H00C000C0&
      Caption         =   "Start!"
      BeginProperty Font
         Name            =   "Franklin Gothic Demi"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9600
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1080
      Width           =   2655
   End
   Begin VB.CommandButton cmdGame
      BackColor       =   &H0000FFFF&
      Caption         =   "The Favre Game"
      BeginProperty Font
         Name            =   "Franklin Gothic Demi"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9600
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4080
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton cmdStats
      BackColor       =   &H00C000C0&
      Caption         =   "Favre's Career Stats"
      BeginProperty Font
         Name            =   "Franklin Gothic Demi"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5880
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton cmdTrivia
      BackColor       =   &H0000FFFF&
      Caption         =   "Favre Trivia Challenge"
      BeginProperty Font
         Name            =   "Franklin Gothic Demi"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7800
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton cmdQuit
      BackColor       =   &H0000FFFF&
      Caption         =   "Retire"
      BeginProperty Font
         Name            =   "Franklin Gothic Demi"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   10680
      Width           =   2655
   End
   Begin VB.Label lblMenu
      BackStyle       =   0  'Transparent
      Caption         =   "Main Menu"
      BeginProperty Font
         Name            =   "Franklin Gothic Demi"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   975
      Left            =   9960
      TabIndex        =   6
      Top             =   2280
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label lblTitle
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "The Brett Favre Experience"
      BeginProperty Font
         Name            =   "Book Antiqua"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2055
      Left            =   360
      TabIndex        =   4
      Top             =   360
      Width           =   8175
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'The Brett Favre Experience
'frmMain
'Doug Donaldson
'2/24/10

'This is the menu and title page of the program. There are several options available for
'the user, including taking a trivia test, playing a decision-based game, and viewing
'career stats of Favre. The program incorporates command buttons, multiple forms,
'input and message boxes, and plenty of fun!



'The start button loads the career statistical data of Brett Favre from the FavreStats.txt
'file. All buttons but start and Retire(Quit) are disabled until the Start button is
'pressed.
Private Sub cmdStart_Click()

'welcome user to program and explain the layout/procedure/instructions, etc.

MsgBox "Welcome to The Brett Favre Experience! Use the menu to guide you through" & Chr(13) & "the program. You can play the Brett Favre Game, view his Career Statistics" & Chr(13) & "answer Trivia, and even calculate your Brett Favre score! Enjoy!", , ""


'open stats file
Open App.Path & "\FavreStats.txt" For Input As #1

'set counter variable values
CTR = 0

'perform file read
While Not (EOF(1) Or CTR >= 20)
    CTR = CTR + 1
    Input #1, Year(CTR), Games(CTR), Attempts(CTR), Completions(CTR), CompPercent(CTR), Yards(CTR), Touchdowns(CTR), Interceptions(CTR), PassRating(CTR)
End While
Close #1

'enable/disable buttons to guide user through interface
cmdGame.Visible = True
cmdStats.Visible = True
cmdTrivia.Visible = True
cmdStart.Visible = False
lblMenu.Visible = True
End Sub
Private Sub cmdStats_Click()
frmMain.Hide
frmStats.Show
End Sub


Private Sub cmdTrivia_Click()
frmMain.Hide
frmTrivia.Show
End Sub

Private Sub cmdGame_Click()
frmMain.Hide
frmFavreGame.Show
End Sub

Private Sub cmdQuit_Click()

'make a warning message to make sure user wants to retire (quit)
Dim Retire As Integer

'use input box to inquire user intent
Retire = InputBox("Are you sure you want to retire? Enter 1 if yes, 2 if no.", "")

'if statement for conditions of input box
Select Case Retire
Case 1

    'I had to look up how to perform this particular end function online at http://www.vb6.us/tutorials/understanding-msgbox-command-visual-basic
    Retire = MsgBox("Are you sure you're sure you want to retire?", vbYesNo + vbQuestion, "")
        If Retire = vbYes Then
            End
        End If

    'if statement used as a cancel of the retirement
    Case 2
        MsgBox "Congratulations! You have decided to come back for another season!", , ""
    Case Else
        MsgBox "Sorry, you have waffled incorrectly. Please answer either yes(1) or no(2).", , ""
End Select
End Sub
