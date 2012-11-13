VERSION 5.00
Begin VB.Form frmschedule 
   BackColor       =   &H0000FFFF&
   Caption         =   "Form1"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   7710
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   40000
      Left            =   240
      Top             =   120
   End
   Begin VB.PictureBox picResults2 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   4800
      ScaleHeight     =   795
      ScaleWidth      =   2715
      TabIndex        =   5
      Top             =   3960
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Return to Front Page"
      Height          =   855
      Left            =   4920
      TabIndex        =   4
      Top             =   240
      Width           =   2535
   End
   Begin VB.CommandButton cmdbestrecord 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sort by Best 2005 Record"
      Height          =   735
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4920
      Width           =   2295
   End
   Begin VB.CommandButton cmdalphabetical 
      BackColor       =   &H00800080&
      Caption         =   "Sort Alphabetically"
      Height          =   735
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4920
      Width           =   2295
   End
   Begin VB.CommandButton cmdweek 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sort Chronologically"
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4920
      Width           =   2295
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00FFFFFF&
      Height          =   4575
      Left            =   120
      ScaleHeight     =   4515
      ScaleWidth      =   4515
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
   Begin VB.Image Image1 
      Height          =   2670
      Left            =   4800
      Picture         =   "frmschedule.frx":0000
      Top             =   1200
      Width           =   2850
   End
End
Attribute VB_Name = "frmschedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Title: Minnesota Vikings Fan Page

'Form Name: Schedule

'Written by Jim Zrust

'Date: November 1, 2006

'Form objective:the purpose of this form was to allow the user to see the teams that the Vikings will be playing this year
'and to sort these teams depending on three criteria (the data was inputted chronologically to eliminate the
'need for a third sort)

Option Explicit
Dim Teams(1 To 16) As String, Wins(1 To 16) As Single 'declared the variables form wide so i wouldn't have to do it for every command (planned on using them for multiple commands)
Dim TempTeams As String, TempWins As Integer
Private Sub cmdalphabetical_Click()
picResults.Cls 'clear past results
Open App.Path & "\schedule.txt" For Input As #1 'open schedule
I = 0
For I = 1 To 16
    Input #1, Teams(I), Wins(I) 'input array
Next I
For Pass = 1 To 15 'sort the results alphabetically
    For I = 1 To 15 - Pass
        If Teams(I) > Teams(I + 1) Then
            TempTeams = Teams(I)
            Teams(I) = Teams(I + 1)
            Teams(I + 1) = TempTeams
        End If
    Next I
Next Pass
For I = 1 To 16
    picResults.Print Teams(I) 'print the sorted list
Next I
Close #1 'make sure to close file
End Sub

Private Sub cmdbestrecord_Click()
Dim TotalWins As Single
Dim avg As Integer
picResults.Cls
Open App.Path & "\schedule.txt" For Input As #1
I = 0
For I = 1 To 16
    Input #1, Teams(I), Wins(I) 'input array again(had to input for every button because i didn't know which the user would click first
Next I
Close #1
For Pass = 1 To 15 'sort numerically (ascending)
    For I = 1 To 15 - Pass
        If Wins(I) > Wins(I + 1) Then
            TempWins = Wins(I)
            TempTeams = Teams(I)
            Wins(I) = Wins(I + 1)
            Teams(I) = Teams(I + 1)
            Wins(I + 1) = TempWins
            Teams(I + 1) = TempTeams
        End If
    Next I
Next Pass
For I = 1 To 16
    picResults.Print Teams(I); Tab(30); "Wins "; FormatNumber(Wins(I), 0)
Next I
For I = 1 To 16
    TotalWins = TotalWins + Wins(I)
Next I
avg = TotalWins / 16
picResults2.Print "Opponets Avg Wins "; avg
End Sub

Private Sub cmdweek_Click()
picResults.Cls
Open App.Path & "\schedule.txt" For Input As #1 'open the created notepad file
I = 0
For I = 1 To 16
    Input #1, Teams(I), Wins(I) 'input array
Next I
For I = 1 To 16
    picResults.Print Teams(I) 'print array in a column
Next I
Close #1
End Sub

Private Sub Command1_Click() ' allows the user to return to the front page
frmhome.Show
frmschedule.Hide
End Sub


Private Sub Timer1_Timer()
frmschedule.Hide
End Sub
