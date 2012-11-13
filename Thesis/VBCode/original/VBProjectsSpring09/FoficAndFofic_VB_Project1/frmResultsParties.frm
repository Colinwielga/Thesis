VERSION 5.00
Begin VB.Form frmResultsParties 
   BackColor       =   &H00808000&
   Caption         =   "Results Parties"
   ClientHeight    =   8535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   Picture         =   "frmResultsParties.frx":0000
   ScaleHeight     =   8535
   ScaleWidth      =   10470
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Quit"
      Height          =   855
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7560
      Width           =   1575
   End
   Begin VB.CommandButton cmdChart 
      BackColor       =   &H00C0E0FF&
      Caption         =   "View Chart"
      Height          =   1455
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox txtSeats 
      BackColor       =   &H00C0FFC0&
      Height          =   735
      Left            =   3480
      TabIndex        =   5
      Top             =   6240
      Width           =   1935
   End
   Begin VB.PictureBox picResultsPP 
      Height          =   5175
      Left            =   2160
      Picture         =   "frmResultsParties.frx":20802
      ScaleHeight     =   5115
      ScaleWidth      =   7275
      TabIndex        =   4
      Top             =   600
      Width           =   7335
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Back"
      Height          =   855
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7560
      Width           =   1695
   End
   Begin VB.CommandButton cmdbySeats 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Results by Number of Seats"
      Height          =   855
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton cmdbyVotes 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Results by Number of Votes"
      Height          =   1455
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Load"
      Height          =   1455
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lblSeats 
      BackColor       =   &H00C0FFC0&
      Caption         =   " Please, enter  number (1-21) to find  the  number of seats for the Parties                                                ==>"
      Height          =   735
      Left            =   360
      TabIndex        =   7
      Top             =   6240
      Width           =   2535
   End
End
Attribute VB_Name = "frmResultsParties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Local Election for Prozor-Rama
    'frmResultsParties
    'Josipa and Mario Fofic
    'Written 03/22/09


'the purpose of this form is to show the results of elections by number of votes,
'and by number of seats. This form searches through results of elections and sort and
'display parties from by number of votes. Also it allows user to search the number of
'seats for each political party, and display results in message box
'this forms loads the chart with general results

Option Explicit
Dim PoliticalParty(1 To 100) As String
Dim Votes(1 To 7000) As Integer
Dim Seat As Integer
Dim TempVotes As Integer
Dim TempPoliticalParty As String
Public Ctr As Integer
Public I As Integer
Dim Pos As Integer, Pass As Integer
Dim TempSeat As Integer
Dim J As Integer
Dim TempPercent As Single
Dim Percent(1 To 50) As Single


Private Sub cmdBack_Click()
frmResultsParties.Hide
frmResults.Show

End Sub

Private Sub cmdChart_Click()
picResultsPP.Cls

picResultsPP.Picture = LoadPicture(App.Path & "/ColumnChart.jpg")  'this line of code load the chart

End Sub



Private Sub cmdLoad_Click()
picResultsPP.Cls               'this line of code will clear the picturebox


Open App.Path & "\PartiesResults.txt" For Input As #1   'this line open the file
Ctr = 0

picResultsPP.Print Tab(2); "Name of Political Parties"; Tab(55); "Number of Votes"
picResultsPP.Print "*******************************************************************************************************************************"

Do While Not EOF(1)     'this line fill the array
    Ctr = Ctr + 1
    Input #1, PoliticalParty(Ctr), Votes(Ctr), Percent(Ctr)
    picResultsPP.Print Tab(2); PoliticalParty(Ctr); Tab(55); Votes(Ctr)
   
Loop
Close #1
End Sub

Private Sub cmdbyVotes_Click()
'this command button sorts the list of political parties by number of votes

picResultsPP.Cls

picResultsPP.Print    'prints a blank line
picResultsPP.Print Tab(2); "Name of Political Parties"; Tab(55); "Number of Votes"; Tab(80); "Percent"
picResultsPP.Print "******************************************************************************************************************************"
 
   For Pass = 1 To Ctr - 1                         'this part of code compare and arrange list
      For Pos = 1 To Ctr - Pass
         If Votes(Pos) < Votes(Pos + 1) Then
            TempVotes = Votes(Pos)
            Votes(Pos) = Votes(Pos + 1)
            Votes(Pos + 1) = TempVotes
            TempPoliticalParty = PoliticalParty(Pos)
            PoliticalParty(Pos) = PoliticalParty(Pos + 1)
            PoliticalParty(Pos + 1) = TempPoliticalParty
            TempPercent = Percent(Pos)
            Percent(Pos) = Percent(Pos + 1)
            Percent(Pos + 1) = TempPercent
         End If
      Next Pos
   Next Pass

    For I = 1 To Ctr
    
    picResultsPP.Print Tab(2); PoliticalParty(I); Tab(55); Votes(I); Tab(80); FormatPercent(Percent(I), 2)
   Next I
                                          'this line prints the results
            
End Sub
Private Sub cmdbySeats_Click()
'purpose of this command button is to select and display
'political parties by the number of seats


Seat = txtSeats.Text

Select Case Seat
    
    Case 8
        MsgBox "HDZ 1990 has " & Seat & " seats in the Assambly", , "Number of Seats"
    Case 6
       MsgBox "HDZ BIH has " & Seat & " seats in the Assambly", , "Number of Seats"
    Case 3
       MsgBox "SDA has " & Seat & " seats in the Assambly", , "Number of Seats"
    Case 2
       MsgBox "HSP has " & Seat & " seats in the Assambly", , "Number of Seats"
    Case 1
       MsgBox "SBIH and BPS, both have  " & Seat & " seat in the Assambly", , "Number of Seats"
    Case 0
       MsgBox "DNZ, HSS-NHI, SDP, and HKDU do not have any seat(s) in the Assambly", , "Number of Seats"
    Case Else
       MsgBox "There is no Party with " & Seat & " number of seats in the Assambly", , "Number of Seats"
End Select
       
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub Form_Load()
Top = Screen.Height / 2 - Height / 2
  Left = Screen.Width / 2 - Width / 2
End Sub

