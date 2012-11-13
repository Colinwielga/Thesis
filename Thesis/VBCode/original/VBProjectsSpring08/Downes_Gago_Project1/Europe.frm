VERSION 5.00
Begin VB.Form Europe 
   BackColor       =   &H000040C0&
   Caption         =   "Europe"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12960
   LinkTopic       =   "Form1"
   ScaleHeight     =   8055
   ScaleWidth      =   12960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuiz 
      BackColor       =   &H00C0E0FF&
      Caption         =   "SO YOU THINK YOU KNOW EUROPEAN CAPITALS?"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   3855
   End
   Begin VB.TextBox txtTitle 
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Text            =   "                                        Welcome to the Old Continent!"
      Top             =   120
      Width           =   12975
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00E0E0E0&
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6720
      Width           =   2175
   End
   Begin VB.CommandButton cmdMatch 
      BackColor       =   &H00C0E0FF&
      Caption         =   "GAME OF CONNOTATIONS"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   3855
   End
   Begin VB.Label lblPlay 
      BackColor       =   &H00C0E0FF&
      Caption         =   "      PLAY EUROPEAN MUSIC                 (CLICK TWICE BELOW)"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   4200
      Width           =   3855
   End
   Begin VB.OLE OLE1 
      BackColor       =   &H000040C0&
      BorderStyle     =   0  'None
      Class           =   "Package"
      Height          =   975
      Left            =   120
      OleObjectBlob   =   "Europe.frx":0000
      SourceDoc       =   "M:\CS130\PROJECT\Beethoven's Symphony No. 9 (Scherzo).wma"
      TabIndex        =   4
      Top             =   5040
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   6165
      Left            =   4680
      Picture         =   "Europe.frx":98A18
      Stretch         =   -1  'True
      Top             =   960
      Width           =   8175
   End
End
Attribute VB_Name = "Europe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: The Globe Trotter Experience
'Form name: Europe.frm
'Author: Marta Gago & Brian Downes
'Date Written: Thursday March 27th, 2008
'Objective of form:  The objective of this form is to play a music file,
'and two different games to test the knowledge of the user
Option Explicit
'Hides the Europe Form and Shows the Main Form
Private Sub cmdBack_Click()
Europe.Hide
Main.Show
End Sub
'Hides the Europe Form and Shows the EuropeanGame Form
Private Sub cmdMatch_Click()
Europe.Hide
EuropeanGame.Show
End Sub
'Is used to play music
Private Sub CmdMusic_Click()
OLE1.Enabled = True         'a music file that can be played
End Sub
'a game to test the knowledge of the user
Private Sub cmdQuiz_Click()
Dim country(1 To 100) As String
Dim capital(1 To 100) As String     'Dim variables
Dim cap As String


Open App.Path & "\CapitalCities.txt" For Input As #1    'open data file of capital cities into array

ctr = 0             'Start ctr off at zero

    Do While Not EOF(1)         'Using a Do While Loop, the data is put into the array
        ctr = ctr + 1
        Input #1, country(ctr), capital(ctr)
    Loop

For j = 1 To ctr And cap <> "Finish"

'Using the exhaustive search, the cap = the input from the user for each question until 'Finish' is typed in to quit, or the user finishes the game

cap = InputBox("What is the capital city of " & country(j) & "? Type 'Finish' to Quit")
        If cap = capital(j) Then
            MsgBox "Bravo!"
        ElseIf cap = "Finish" Then
            GoTo endloop            'Brings the user out of the loop, and ends the game
        Else:
            MsgBox "Incorrect Capital! The capital we are looking for is " & capital(j)
        End If
Next j
endloop:    'Out of the loop, the game ends
    
Close       'Closes the array

    
End Sub


