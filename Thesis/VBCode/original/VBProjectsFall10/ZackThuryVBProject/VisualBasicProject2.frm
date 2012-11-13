VERSION 5.00
Begin VB.Form FrmTopSalaries 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   10365
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12465
   LinkTopic       =   "Form1"
   ScaleHeight     =   10365
   ScaleWidth      =   12465
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picAverageIncome 
      BackColor       =   &H00FFFFFF&
      Height          =   2535
      Left            =   4320
      ScaleHeight     =   2475
      ScaleWidth      =   7395
      TabIndex        =   5
      Top             =   6240
      Width           =   7455
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate Average Income Per Event"
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      TabIndex        =   4
      Top             =   7080
      Width           =   2415
   End
   Begin VB.CommandButton cmdBackToHome 
      BackColor       =   &H00808000&
      Caption         =   "Back To Home Screen"
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9000
      Width           =   2415
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000080FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9000
      Width           =   2415
   End
   Begin VB.CommandButton cmdShowTop15 
      Caption         =   "Show 10 Top Paid Professional Golfers"
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      TabIndex        =   1
      Top             =   2520
      Width           =   2415
   End
   Begin VB.PictureBox picTopSalaries 
      BackColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   3960
      ScaleHeight     =   2595
      ScaleWidth      =   8115
      TabIndex        =   0
      Top             =   1800
      Width           =   8175
   End
   Begin VB.Label lblthetops 
      BackColor       =   &H00FFFFFF&
      Caption         =   "The Top Ten Highest Paid Professional Golfers Including Endorsements:"
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6240
      TabIndex        =   7
      Top             =   840
      Width           =   3735
   End
   Begin VB.Label lblcalculatedaveragesalaries 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Average Amount of Money Made by Each Golfer Per Event"
      BeginProperty Font 
         Name            =   "New Athena Unicode"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5880
      TabIndex        =   6
      Top             =   5160
      Width           =   4215
   End
End
Attribute VB_Name = "FrmTopSalaries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this form shows the top paid professional golfers in descending order by the money they made in 2010
'Declare variables
Dim Ctr As Integer
Dim Order(1 To 10) As Integer
Dim GolfersName(1 To 10) As String
Dim Salary(1 To 10) As Single
Dim EventsPlayed(1 To 10) As Integer
Dim AvgIncome(1 To 10) As Single

'hides the top salaries form and shows the title form
Private Sub cmdBackToHome_Click()
    FrmTopSalaries.Hide
    FrmTitle.Show
End Sub

Private Sub cmdCalculate_Click()
'clears picture box and prints out the labels for the table
    picAverageIncome.Cls
    picAverageIncome.Print Tab(3); "Rank"; Tab(23); "Golfer Name"; Tab(53); "Average Amount Earned Per Event"
    picAverageIncome.Print "__________________________________________________________________________"
    'declares variable
    Dim k As Integer
    'goes through the array and calculates the average income per event played by the golfer
    'prints the rank golfers name and average income per event
    For k = 1 To Ctr
        AvgIncome(k) = Salary(k) / EventsPlayed(k)
        picAverageIncome.Print Tab(3); Order(k); Tab(23); GolfersName(k); Tab(62); FormatCurrency(AvgIncome(k))
    Next k
End Sub

'quits the program
Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdShowTop15_Click()
    'clears the picture box and prints the heading for the table
    picTopSalaries.Cls
    picTopSalaries.Print Tab(3); "Rank"; Tab(13); "Golfer Name"; Tab(50); "Total Earnings in 2010"; Tab(80); "Events Played in 2010"
    picTopSalaries.Print "_____________________________________________________________________________________"
    'opens the toppaidgolfers file and saves the contents to arrays
    Open App.Path & "\TopPaidGolfers.txt" For Input As #2
    Ctr = 0
    Do While Not EOF(2)
        Ctr = Ctr + 1
        Input #2, Order(Ctr), GolfersName(Ctr), Salary(Ctr), EventsPlayed(Ctr)
        'prints the contents of the file
        picTopSalaries.Print Tab(3); Order(Ctr); Tab(13); GolfersName(Ctr); Tab(53); FormatCurrency(Salary(Ctr)); Tab(86); EventsPlayed(Ctr)
    Loop
    'closes the file
    Close #2
End Sub
