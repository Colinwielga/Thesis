VERSION 5.00
Begin VB.Form frmCalcOwnStats 
   Caption         =   "Calculate Your Own Stats"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   Picture         =   "frmCalcOwnStats.frx":0000
   ScaleHeight     =   6690
   ScaleWidth      =   9780
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Tekton Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8400
      TabIndex        =   7
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton cmdGoToSort 
      Caption         =   "Go To Sort Option"
      BeginProperty Font 
         Name            =   "Tekton Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6360
      TabIndex        =   6
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton cmdGoToPitching 
      Caption         =   "Go To Pitching Statistics"
      BeginProperty Font 
         Name            =   "Tekton Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      TabIndex        =   5
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton cmdGoToHitting 
      Caption         =   "Go To Hitting Statistics"
      BeginProperty Font 
         Name            =   "Tekton Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      TabIndex        =   4
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton cmdOnBase 
      Caption         =   "On Base %"
      BeginProperty Font 
         Name            =   "Tekton Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   2055
   End
   Begin VB.CommandButton cmdERA 
      Caption         =   "ERA"
      BeginProperty Font 
         Name            =   "Tekton Pro"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   4800
      Width           =   2055
   End
   Begin VB.CommandButton cmdSlugging 
      Caption         =   "Slugging %"
      BeginProperty Font 
         Name            =   "Tekton Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   2055
   End
   Begin VB.CommandButton cmdAvg 
      Caption         =   "Batting Average"
      BeginProperty Font 
         Name            =   "Tekton Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
End
Attribute VB_Name = "frmCalcOwnStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Avg As Single
Dim Slug As Single
Dim OBP As Single
Dim ERA As Single

''Twins Statistics, frmCalcOwnStats, By Mike Patnode, Written Nov.2, 2006, Objective to diplay Twins Stats

Private Sub cmdAvg_Click()
    'Declare Hits and AtBats as variable to use in formula
    Dim Hits As Integer
    Dim AtBats As Integer
    AtBats = InputBox("Insert Number of At Bats", "At Bats")
    Hits = InputBox("Insert Number of Hits", "Hits")
    Avg = Hits / AtBats
    MsgBox "You have a Batting Average of " & FormatNumber(Avg, 3), , "Batting Average"
    'Displays batting average
End Sub

Private Sub cmdERA_Click()
'(Earned Runs/Innings Pitched) x 9,Declare variable
    Dim ER As Integer
    Dim IP As Integer
    Dim ERA As Single
    IP = InputBox("Enter how many innings you pitched", "Innings Pitched")
    ER = InputBox("Enter how many earned runs you gave up", "Earned Runs")
    ERA = (ER / IP) * 9
    MsgBox "Your ERA is " & FormatNumber(ERA), , "ERA"
    'Displays ERA
End Sub

Private Sub cmdGoToHitting_Click()
    'switching forms
    frmHittingStats.Show
    frmPitchingStats.Hide
    frmCalcOwnStats.Hide
    frmSortTwins.Hide
End Sub

Private Sub cmdGoToPitching_Click()
    'switching forms
    frmHittingStats.Hide
    frmPitchingStats.Show
    frmCalcOwnStats.Hide
    frmSortTwins.Hide
End Sub

Private Sub cmdGoToSort_Click()
    'switching forms
    frmHittingStats.Hide
    frmPitchingStats.Hide
    frmCalcOwnStats.Hide
    frmSortTwins.Show
End Sub

Private Sub cmdOnBase_Click()
'OBP = (H+BB+HBP)/(AB+BB+HBP+SF), Delcare all variable
    Dim OBP As Single
    Dim Hits As Integer
    Dim Walks As Integer
    Dim HBP As Integer
    Dim SF As Integer
    Dim AB As Integer
    Hits = InputBox("How many hits did you have?", "Hits") 'Input Hits
    Walks = InputBox("How many times did you walk?", "Walks") 'Input Walks
    HBP = InputBox("How many times were you hit by a pitch?", "HBP") 'Input Hit by Pitch
    SF = InputBox("How many sacrifice flies did you have?", "Sac Flies") 'Input Sac Flies
    AB = InputBox("How many total at bats did you have?", "At Bats") 'Input At Bats
    OBP = (Hits + Walks + HBP) / (AB + Walks + HBP + SF) 'Formula for OBP
    MsgBox "You On Base Percentage is " & FormatNumber(OBP, 3), , "On Base Percentage"
    'Display OBP
End Sub


Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdSlugging_Click()
    'Declare variables
    Dim Singles As Integer
    Dim Doubles As Integer
    Dim Triples As Integer
    Dim HomeRuns As Integer
    Dim Bats As Integer
    Singles = InputBox("Enter how many singles you hit", "Singles") 'input singles
    Doubles = InputBox("Enter how many doubles you hit", "Doubles") 'inputs doubles
    Triples = InputBox("Enter how many triples you hit", "Triples") 'inputs triples
    HomeRuns = InputBox("Enter how many home runs you hit", "Home Runs") 'input homeruns
    Bats = InputBox("Enter how many total At Bats", "At Bats") 'input at bats
    If Bats < (Singles + Doubles + Triples + HomeRuns) Then 'Error message if number of hits are more than at bats
        MsgBox "Not Enough At Bats", , "Error"
        MsgBox "Start Over", , "Error"
    End If
    Slug = (Singles + (2 * Doubles) + (3 * Triples) + (4 * HomeRuns)) / Bats 'formula for slugging percentage
    MsgBox "Your Slugging Percentage is " & FormatNumber(Slug, 3), , "Slugging Percentage"
    'display slugging percentage
End Sub
