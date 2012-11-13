VERSION 5.00
Begin VB.Form frmDunkContest 
   Caption         =   "Timberwolves Dunk Contest"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   ScaleHeight     =   8445
   ScaleWidth      =   10770
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return to Main Page"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   6
      Top             =   7200
      Width           =   2535
   End
   Begin VB.TextBox picResult 
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   8280
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CommandButton cmdResults 
      Caption         =   "See The Results"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8400
      TabIndex        =   4
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton cmdAmir 
      Caption         =   "Amir"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   3
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdNate 
      Caption         =   "Nate"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   2
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdRudy 
      Caption         =   "Rudy"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdCarmello 
      Caption         =   "Carmello"
      BeginProperty Font 
         Name            =   "Tw Cen MT Condensed Extra Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin VB.Image Image5 
      Height          =   3795
      Left            =   840
      Picture         =   "frmDunkContest.frx":0000
      Top             =   3120
      Width           =   2850
   End
   Begin VB.Image Image4 
      Height          =   2775
      Left            =   3960
      Picture         =   "frmDunkContest.frx":177DA
      Top             =   4440
      Width           =   3735
   End
   Begin VB.Image Image3 
      Height          =   3435
      Left            =   6120
      Picture         =   "frmDunkContest.frx":2E17C
      Top             =   120
      Width           =   4500
   End
   Begin VB.Image Image2 
      Height          =   2400
      Left            =   2520
      Picture         =   "frmDunkContest.frx":4FA82
      Top             =   120
      Width           =   1965
   End
   Begin VB.Image Image1 
      Height          =   8640
      Left            =   0
      Picture         =   "frmDunkContest.frx":59FD0
      Top             =   0
      Width           =   10800
   End
End
Attribute VB_Name = "frmDunkContest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This form is a dunk contest
'The user can vote for which dunk they like the best by clicking on the command button which corespondes to the dunk of their choice.
'A running count is kept of how many times each command(or vote) button is clicked and when the result command button is used it will display the tally of votes for each dunker
'These variables must be accesible on the form level so they're declared in option explicit
Dim Amir As Integer
Dim Carmello As Integer
Dim Nate As Integer
Dim Rudy As Integer










Private Sub cmdAmir_Click()
'This is a counter like variable which keeps a tally of how many times the specific command button is pressed
Amir = Amir + 1

End Sub

Private Sub cmdCarmello_Click()
'This is a counter like variable which keeps a tally of how many times the specific command button is pressed
Carmello = Carmello + 1
End Sub


Private Sub cmdNate_Click()
'This is a counter like variable which keeps a tally of how many times the specific command button is pressed

Nate = Nate + 1

End Sub

Private Sub cmdResults_Click()
'This button displays the results of the dunk contest by showing how many votes each dunk has received
'Take notice that I used the untaught method of printing in text boxes
picResult.Text = "Carmello " & Carmello
picResult.Text = picResult.Text & vbNewLine & "Rudy " & Rudy
picResult.Text = picResult.Text & vbNewLine & "Nate " & Nate
picResult.Text = picResult.Text & vbNewLine & "Amir " & Amir
End Sub

Private Sub cmdreturn_Click()
'This returns the user to the main page form and away from the dunk contest form
frmDunkContest.Visible = False
frmMainPage.Visible = True

End Sub

Private Sub cmdRudy_Click()
'This is a counter like variable which keeps a tally of how many times the specific command button is pressed
Rudy = Rudy + 1

End Sub
