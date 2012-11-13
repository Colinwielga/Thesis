VERSION 5.00
Begin VB.Form frmReadTwinsStats 
   BackColor       =   &H8000000D&
   Caption         =   "View Twins Statistics"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14490
   LinkTopic       =   "Form2"
   Picture         =   "frmReadTwinsStats.frx":0000
   ScaleHeight     =   8055
   ScaleWidth      =   14490
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   855
      Left            =   720
      TabIndex        =   4
      Top             =   5880
      Width           =   2055
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Main Menu"
      Height          =   975
      Left            =   720
      TabIndex        =   3
      Top             =   4560
      Width           =   2055
   End
   Begin VB.CommandButton cmdFindStats 
      Caption         =   "After viewing, now you can figure out their statistics!    (BA, OBP, SLG, OPS)"
      Height          =   1215
      Left            =   720
      TabIndex        =   2
      Top             =   3000
      Width           =   2055
   End
   Begin VB.PictureBox picResults 
      Height          =   6735
      Left            =   3360
      ScaleHeight     =   6675
      ScaleWidth      =   7275
      TabIndex        =   1
      Top             =   480
      Width           =   7335
   End
   Begin VB.CommandButton cmdReadTwins 
      Caption         =   "Click here to view 2008 Twins Stats"
      Height          =   1095
      Left            =   720
      TabIndex        =   0
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblSource 
      BackColor       =   &H8000000D&
      Caption         =   "statistics courtesy of twinsbaseball.com"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   7080
      Width           =   2895
   End
End
Attribute VB_Name = "frmReadTwinsStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Ctr As Integer

'Baseball Batting Statistics
'frmReadTwinsStats
'Aaron Walsh
'March 24, 2009
'This program will figure out various batting statistics like BA, OPS, OBP, and SLG
'for Twins players based on 2008 numbers in certain batting catagories

Private Sub cmdBack_Click()
    frmReadTwinsStats.Hide
    frmInitialform.Show
End Sub

Private Sub cmdFindStats_Click()
    frmReadTwinsStats.Hide
    frmFigureTwins.Show
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdReadTwins_Click()
'this opens the file with the Twins 2008 batting numbers and reads them into arrays
    Open App.Path & "\Twins2008battingstats.txt" For Input As #1
    Ctr = 0
    picResults.Cls
    picResults.Print "Player", "At Bats", "Hits", "Home Runs", "Total Bases", "Walks", "Hit-by-Pitch", "Sac Flies"
    picResults.Print "*************************************************************************************************************************************"
    Do Until EOF(1)
        Ctr = Ctr + 1
        Input #1, TwinsNames(Ctr), AB(Ctr), H(Ctr), HR(Ctr), TB(Ctr), BB(Ctr), HBP(Ctr), SF(Ctr)
        picResults.Print TwinsNames(Ctr), AB(Ctr), H(Ctr), HR(Ctr), TB(Ctr), BB(Ctr), HBP(Ctr), SF(Ctr)
    Loop
    Close #1
End Sub
