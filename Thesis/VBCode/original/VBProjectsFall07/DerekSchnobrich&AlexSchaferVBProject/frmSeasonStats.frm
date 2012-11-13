VERSION 5.00
Begin VB.Form frmSeasonStats 
   BackColor       =   &H000000FF&
   Caption         =   "Season Stats"
   ClientHeight    =   5655
   ClientLeft      =   2985
   ClientTop       =   3630
   ClientWidth     =   6990
   FillColor       =   &H000000FF&
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   6990
   Begin VB.CommandButton cmdA 
      Caption         =   "Return to Main Menu"
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      Height          =   1335
      Left            =   120
      MaskColor       =   &H000000FF&
      TabIndex        =   2
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Enter Player to see Stats"
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.PictureBox picResults 
      Height          =   4935
      Left            =   1440
      ScaleHeight     =   4875
      ScaleWidth      =   4635
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmSeasonStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdA_Click()
 frmSeasonStats.Hide
 frmHome.Show
End Sub

Private Sub cmdDisplay_Click()
'Declares Variables
Dim Found As Boolean
Dim SearchName As String
Dim ctr As Integer
Dim counter As Integer
Dim Name(1 To 100) As String
Dim RushAtt(1 To 100) As Integer
Dim RushYards(1 To 100) As Integer
Dim RushTD(1 To 100) As Integer
Dim PassATT(1 To 100) As Integer
Dim PassComp(1 To 100) As Integer
Dim Passyds(1 To 100) As Integer
Dim PassTD(1 To 100) As Integer
Dim PassInt(1 To 100) As Integer
Dim Rec(1 To 100) As Integer
Dim RecYds(1 To 100) As Integer
Dim RecTD(1 To 100) As Integer
Found = False
Dim RushAvg As Double
Dim PassAvg As Double
Dim RecAvg As Double

'Fills the array with stats
Open App.Path & "\OffensiveStats.txt" For Input As #1
Do While Not EOF(1)
    ctr = ctr + 1
    Input #1, Name(ctr), RushAtt(ctr), RushYards(ctr), RushTD(ctr), PassATT(ctr), PassComp(ctr), Passyds(ctr), PassTD(ctr), PassInt(ctr), Rec(ctr), RecYds(ctr), RecTD(ctr)
Loop
Close #1

'Finds Player to display stats for
MsgBox ("Enter a name to enter stats for, formated FIRSTNAME LASTNAME.")
SearchName = InputBox("Player Name")

'searches for player
Do While counter < ctr And Not Found
    counter = counter + 1
    If Name(counter) = SearchName Then
    Found = True
    End If
Loop


'Computes Averages
RushAvg = RushYards(ctr) / RushAtt(ctr)
PassAvg = Passyds(ctr) / PassATT(ctr)
RecAvg = RecYds(ctr) / Rec(ctr)

'Prints Player's Stat Totals
    picResults.Print "Name", Name(counter); Chr(10); "Rush Att", RushAtt(counter); Chr(10); "Rushing Yrds", RushYards(counter); Chr(10); "Rushing TDs", RushTD(counter); Chr(10); "Pass Att", PassATT(counter); Chr(10); "Pass Comp", PassComp(counter); Chr(10); "Pass Yds", Passyds(counter); Chr(10); "Pass TDs", PassTD(counter); Chr(10); "Pass Int", PassInt(counter); Chr(10); "Rec", Rec(counter); Chr(10); "RecYds", RecYds(counter); Chr(10); "RecTDs", RecTD(counter); Chr(10); "Rush Avg", RushAvg; Chr(10); "Pass Avg", PassAvg; Chr(10); "Rec Avg", RecAvg; Chr(10)


End Sub
'Quits the program
Private Sub cmdQuit_Click()
 End
End Sub
