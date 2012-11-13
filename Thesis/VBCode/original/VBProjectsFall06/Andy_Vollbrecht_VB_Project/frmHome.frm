VERSION 5.00
Begin VB.Form frmHome 
   Caption         =   "Hall of Fame Scores"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5355
   ForeColor       =   &H80000007&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   5355
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRead 
      Caption         =   "Load Player Data"
      Height          =   735
      Left            =   1560
      TabIndex        =   6
      Top             =   2880
      Width           =   2055
   End
   Begin VB.PictureBox picBaseball 
      Height          =   1695
      Left            =   1920
      Picture         =   "frmHome.frx":0000
      ScaleHeight     =   1635
      ScaleWidth      =   1515
      TabIndex        =   5
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   855
      Left            =   3000
      TabIndex        =   3
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "Read about the method behind this system"
      Height          =   855
      Left            =   360
      TabIndex        =   2
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search for a Player"
      Enabled         =   0   'False
      Height          =   855
      Left            =   3000
      TabIndex        =   1
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View Rankings"
      Enabled         =   0   'False
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   $"frmHome.frx":0FD3
      Height          =   735
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This is a program that determines if a baseball player deserves to be elected to the hall of fame


Private Sub cmdInfo_Click()
    'Makes explanation page appear, and all other forms disappear
    frmExplanation.Visible = True
    frmHome.Visible = False
    frmRankings.Visible = False
    frmSearch.Visible = False
    
End Sub

Private Sub cmdQuit_Click()
    End 'Closes the program
End Sub

Private Sub cmdRead_Click()
    'Loads data from files into arrays
    'Declaring variables
    Dim pname As String, pcareer As Integer, prate As Integer, pdefense As Integer
    Dim bname As String, bcareer As Integer, brate As Integer, bdefense As Integer
    Dim PitchPos As Single, BatPos As Single
    
    bcounter = 0
    pcounter = 0
    'Opens and loads pitcher data
    Open App.Path & "\pitchers.txt" For Input As #1
    Do Until EOF(1)
        Input #1, pname, PitchPos, pcareer, prate, pdefense
        pcounter = pcounter + 1
        Pitchers(pcounter) = pname
        PitcherPos(pcounter) = PitchPos
        PCareers(pcounter) = pcareer
        PRates(pcounter) = prate
        PDefenses(pcounter) = pdefense
        PitcherTotals(pcounter) = pcareer + prate + pdefense
    Loop
    Close #1
    'Opens and loads batter data
    Open App.Path & "\positionplayers.txt" For Input As #2
    Do Until EOF(2)
        Input #2, bname, BatPos, bcareer, brate, bdefense
        bcounter = bcounter + 1
        BatterPos(bcounter) = BatPos
        Batters(bcounter) = bname
        BCareers(bcounter) = bcareer
        BRates(bcounter) = brate
        BDefenses(bcounter) = bdefense
        BatterTotals(bcounter) = bcareer + brate + bdefense
    Loop
    Close #2
    'Enables the buttons to search and view rankings
    cmdSearch.Enabled = True
    cmdView.Enabled = True
    cmdRead.Enabled = False
    
    
End Sub

Private Sub cmdSearch_Click()
    'Makes search form appear
    frmSearch.Visible = True
    frmHome.Visible = False
    frmRankings.Visible = False
    frmExplanation.Visible = False
    
End Sub

Private Sub cmdView_Click()
    'Makes rankings form appear
    frmRankings.Visible = True
    frmHome.Visible = False
    frmSearch.Visible = False
    frmExplanation.Visible = False
    
End Sub
