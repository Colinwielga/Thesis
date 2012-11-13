VERSION 5.00
Begin VB.Form frm1 
   Caption         =   "Minnesota Timberwolves"
   ClientHeight    =   8730
   ClientLeft      =   1965
   ClientTop       =   2040
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   ScaleHeight     =   8730
   ScaleWidth      =   10785
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H8000000C&
      Caption         =   "Search by Player Name"
      Height          =   855
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Press ""View Players"" button first, then enter the exact players name with capitalization."
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   3720
      Top             =   600
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   3240
      Top             =   600
   End
   Begin VB.PictureBox picLogo 
      Height          =   1695
      Left            =   3840
      Picture         =   "frm1.frx":0000
      ScaleHeight     =   1635
      ScaleWidth      =   2475
      TabIndex        =   2
      Top             =   3360
      Width           =   2535
   End
   Begin VB.CommandButton cmdEnter 
      BackColor       =   &H00FF0000&
      Caption         =   "ENTER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox txtInput 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4320
      TabIndex        =   10
      ToolTipText     =   "Press Enter After Entering Name"
      Top             =   0
      Width           =   3015
   End
   Begin VB.CommandButton cmdfun 
      BackColor       =   &H8000000C&
      Caption         =   "Fun Section!"
      BeginProperty Font 
         Name            =   "Myriad Condensed Web"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   7200
      MaskColor       =   &H008080FF&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1080
      UseMaskColor    =   -1  'True
      Width           =   2895
   End
   Begin VB.CommandButton cmdTakePoll 
      BackColor       =   &H8000000D&
      Caption         =   "             Take A Poll:                     What Are The Chances The Wolves Make The Playoffs?!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2760
      Width           =   2895
   End
   Begin VB.CommandButton CmdSignUp 
      BackColor       =   &H8000000C&
      Caption         =   "Sign Up For Your Chance To Win Free Wolves Tickets!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4440
      Width           =   2895
   End
   Begin VB.CommandButton cmdTickets 
      BackColor       =   &H8000000D&
      Caption         =   "Buy Tickets Here"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   7320
      Picture         =   "frm1.frx":1AE3
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6360
      Width           =   2775
   End
   Begin VB.CommandButton CmdNumber 
      BackColor       =   &H8000000D&
      Caption         =   "Enter The Player's Number to get Stats"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "You must push ""View Players"" button first, then enter the exact number of the player before you get stats."
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmdStandings 
      BackColor       =   &H8000000D&
      Caption         =   "Standings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6840
      UseMaskColor    =   -1  'True
      Width           =   2895
   End
   Begin VB.CommandButton cmdRecord 
      BackColor       =   &H8000000C&
      Caption         =   "View Current Record  "
      BeginProperty Font 
         Name            =   "Berlin Sans FB"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4920
      Width           =   2895
   End
   Begin VB.CommandButton cmdPlayers 
      BackColor       =   &H8000000D&
      Caption         =   "View Players"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   1935
      Left            =   3720
      TabIndex        =   14
      Top             =   3240
      Width           =   2775
      Begin VB.Image Image15 
         Height          =   135
         Left            =   1320
         Top             =   1800
         Width           =   15
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   135
      Left            =   1320
      TabIndex        =   15
      Top             =   2640
      Width           =   15
   End
   Begin VB.Image Image30 
      Height          =   1680
      Left            =   9480
      Picture         =   "frm1.frx":35C6
      Top             =   6720
      Width           =   2475
   End
   Begin VB.Image Image29 
      Height          =   1680
      Left            =   9600
      Picture         =   "frm1.frx":50A9
      Top             =   5040
      Width           =   2475
   End
   Begin VB.Image Image28 
      Height          =   1680
      Left            =   9600
      Picture         =   "frm1.frx":6B8C
      Top             =   3360
      Width           =   2475
   End
   Begin VB.Image Image26 
      Height          =   1815
      Left            =   9720
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Image Image25 
      Height          =   1680
      Left            =   9720
      Picture         =   "frm1.frx":866F
      Top             =   1680
      Width           =   2475
   End
   Begin VB.Image Image24 
      Height          =   1680
      Left            =   9600
      Picture         =   "frm1.frx":A152
      Top             =   0
      Width           =   2475
   End
   Begin VB.Image Image23 
      Height          =   1680
      Left            =   7200
      Picture         =   "frm1.frx":BC35
      Top             =   6720
      Width           =   2475
   End
   Begin VB.Image Image22 
      Height          =   1680
      Left            =   7200
      Picture         =   "frm1.frx":D718
      Top             =   4920
      Width           =   2475
   End
   Begin VB.Image Image21 
      Height          =   1680
      Left            =   7320
      Picture         =   "frm1.frx":F1FB
      Top             =   3240
      Width           =   2475
   End
   Begin VB.Image Image20 
      Height          =   1680
      Left            =   7320
      Picture         =   "frm1.frx":10CDE
      Top             =   1560
      Width           =   2475
   End
   Begin VB.Image Image19 
      Height          =   1680
      Left            =   7200
      Picture         =   "frm1.frx":127C1
      Top             =   0
      Width           =   2475
   End
   Begin VB.Image Image18 
      Height          =   1680
      Left            =   4800
      Picture         =   "frm1.frx":142A4
      Top             =   6720
      Width           =   2475
   End
   Begin VB.Image Image17 
      Height          =   1680
      Left            =   4800
      Picture         =   "frm1.frx":15D87
      Top             =   5040
      Width           =   2475
   End
   Begin VB.Image Image16 
      Height          =   1680
      Left            =   4920
      Picture         =   "frm1.frx":1786A
      Top             =   3360
      Width           =   2475
   End
   Begin VB.Image Image14 
      Height          =   1680
      Left            =   4800
      Picture         =   "frm1.frx":1934D
      Top             =   1680
      Width           =   2475
   End
   Begin VB.Image Image13 
      Height          =   255
      Left            =   5520
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label lblYourName 
      BackColor       =   &H00FF0000&
      Caption         =   "Input Your Name Here!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label lblSignIn 
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Please Sign In First!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   4695
   End
   Begin VB.Image Image12 
      Height          =   1680
      Left            =   4800
      Picture         =   "frm1.frx":1AE30
      Top             =   0
      Width           =   2475
   End
   Begin VB.Image Image11 
      Height          =   1680
      Left            =   2400
      Picture         =   "frm1.frx":1C913
      Top             =   6720
      Width           =   2475
   End
   Begin VB.Image Image10 
      Height          =   1680
      Left            =   0
      Picture         =   "frm1.frx":1E3F6
      Top             =   6720
      Width           =   2475
   End
   Begin VB.Image Image9 
      Height          =   1680
      Left            =   2400
      Picture         =   "frm1.frx":1FED9
      Top             =   5040
      Width           =   2475
   End
   Begin VB.Image Image8 
      Height          =   1680
      Left            =   0
      Picture         =   "frm1.frx":219BC
      Top             =   5040
      Width           =   2475
   End
   Begin VB.Image Image7 
      Height          =   1680
      Left            =   2400
      Picture         =   "frm1.frx":2349F
      Top             =   3360
      Width           =   2475
   End
   Begin VB.Image Image6 
      Height          =   1680
      Left            =   2400
      Picture         =   "frm1.frx":24F82
      Top             =   1680
      Width           =   2475
   End
   Begin VB.Image Image5 
      Height          =   1680
      Left            =   0
      Picture         =   "frm1.frx":26A65
      Top             =   3360
      Width           =   2475
   End
   Begin VB.Image Image4 
      Height          =   1680
      Left            =   0
      Picture         =   "frm1.frx":28548
      Top             =   1680
      Width           =   2475
   End
   Begin VB.Image Image3 
      Height          =   1455
      Left            =   0
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Image Image2 
      Height          =   1680
      Left            =   2400
      Picture         =   "frm1.frx":2A02B
      Top             =   0
      Width           =   2475
   End
   Begin VB.Image Image1 
      Height          =   1680
      Left            =   0
      Picture         =   "frm1.frx":2BB0E
      Top             =   0
      Width           =   2475
   End
   Begin VB.Label lblName 
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "By: Chad Henfling"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10455
   End
   Begin VB.Image Image27 
      Height          =   1680
      Left            =   0
      Picture         =   "frm1.frx":2D5F1
      Top             =   0
      Width           =   2475
   End
End
Attribute VB_Name = "frm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Minnesota Timberwolves Center (MinnesotaTimberwovlesbyChadHenfling.vbp)
'Main Form (frm1.frm)
'Chad Henfling
'Created March 23, 2006
'This is the main form and has all of the command buttons that will connect to other forms
'The overall purpose of this project is for Wolves fans to have fun, learn about the team and players, have some basketball interaction tools, and learn about the current status of the NBA.
Option Explicit
Private Sub cmdCoaches_Click()
    'jumping from one form to the other
    frm1.Visible = False
    frmCoaches.Visible = True
End Sub
Private Sub cmdEnter_Click()
    'input your name through text box and storing it as public
    YourName = txtInput.Text
End Sub
Private Sub cmdfun_Click()
    'goes to fun form
    
    frm1.Visible = False
    frmFun.Visible = True
End Sub
Private Sub CmdNumber_Click()
    Dim Number, counter As Integer
    counter = 0
    'inputing a players number to search for that player
    Number = InputBox("Enter thePlayer's Number whos profile you wish to see", "Player Number")
    For counter = 1 To size
        If Number = Numbers(counter) Then
            MsgBox "Number " & Number & " is " & Names(counter) & " He scores " & PPG(counter) & " points per game", , "To view more on" & Names(counter) & " Go to the players section"
        End If
    Next counter
End Sub
Private Sub cmdPlayers_Click()
    Dim counter As Integer
    'Opening file and reading information
    Open App.Path & "\NumbersNames.txt" For Input As #1
    counter = 0
    Do Until EOF(1)
        counter = counter + 1
        Input #1, Numbers(counter), Names(counter), PPG(counter)
    Loop
    Close #1
    size = counter
    frm1.Visible = False
    frmPlayers.Visible = True
End Sub
Private Sub cmdRecord_Click()
    'goes to other form
    frm1.Visible = False
    frmRecord.Visible = True
End Sub
Private Sub cmdSearch_Click()
    'insert a name via inputbox, then it searches by the string name through the stored arrays and shows via messagebox the results
    Dim guy As String
    Dim pos As Integer
    guy = InputBox("Enter the players first name", "enter a name")
    For pos = 1 To size
        If guy = Names(pos) Then
        MsgBox "Your Desired Results Are " & Names(pos) & " # " & Numbers(pos) & " ,ppg " & PPG(pos)
        End If
    Next pos
End Sub
Private Sub CmdSignUp_Click()
    Dim Number As String
    'inputing a phone number and entering it in a drawing
    Number = InputBox("Enter Your Contact Number", "Telephone Number")
    MsgBox "" & YourName & " ,Your Name will be put into a drawing for your chance to win Wolves tickets!", , "Congratulations!"
End Sub
Private Sub cmdStandings_Click()
    'going to a different form
    frm1.Visible = False
    frmStandings.Visible = True
End Sub
Private Sub cmdTakePoll_Click()
    'going to a different form
    frm1.Visible = False
    frmPoll.Visible = True
End Sub
Private Sub cmdTickets_Click()
    Dim Tickets As String
    Dim Price As Single
    'entering a ticket price and and a team or date you want to see them on and displaying that there are no tickets left for that game.
    Tickets = InputBox("Please pick a team or date for which you wish to attend a game!", "Tickets")
    Price = InputBox("Enter your desired ticket price range.", , "Price")
    MsgBox "We are sorry, " & YourName & " but we are currently out of tickets for the " & Tickets & " Game in the " & FormatCurrency(Price) & " Price range.", , "Sorry"
End Sub
Private Sub Timer2_Timer()
    'These Timers make the sign in box flash so it draws your attention
    Dim time As Single
    For time = 1 To 30
    Timer2 = True
    lblYourName = False
    lblYourName = "ENTER YOUR NAME"
    Next time
End Sub
Private Sub Timer3_Timer()
    Dim time As Single
    For time = 1 To 30
    Timer3 = True
    lblYourName = True
    lblYourName = "A NAME PLEASE"
    Next time
End Sub

