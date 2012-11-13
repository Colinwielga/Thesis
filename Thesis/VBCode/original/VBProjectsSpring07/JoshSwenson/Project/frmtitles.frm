VERSION 5.00
Begin VB.Form frmtitles 
   BackColor       =   &H000000FF&
   Caption         =   "Tennis Grand Slam Titles"
   ClientHeight    =   9060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   ScaleHeight     =   9060
   ScaleWidth      =   10605
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdgo 
      Caption         =   "Go on to next page"
      Height          =   975
      Left            =   1560
      TabIndex        =   4
      Top             =   4440
      Width           =   2175
   End
   Begin VB.CommandButton cmdfind 
      Caption         =   "Find a Player"
      Height          =   1335
      Left            =   1440
      TabIndex        =   3
      Top             =   2640
      Width           =   2655
   End
   Begin VB.CommandButton cmdend 
      Caption         =   "Quit"
      Height          =   855
      Left            =   1560
      TabIndex        =   2
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CommandButton cmdshow 
      Caption         =   "Show Grand Slam Title Leaders"
      Height          =   1335
      Left            =   1320
      TabIndex        =   1
      Top             =   720
      Width           =   2775
   End
   Begin VB.PictureBox picdisplay 
      BackColor       =   &H00FFFF00&
      Height          =   10455
      Left            =   4560
      ScaleHeight     =   10395
      ScaleWidth      =   5595
      TabIndex        =   0
      Top             =   360
      Width           =   5655
   End
End
Attribute VB_Name = "frmtitles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This Form demonstrates using an array (Grand Slam Titles) and displaying them in picture boxs
'The use of Multiple Forms is also demonstrated
'The third page demonstrates how using text boxs you can input information
'and the program will read names input by the user and display the number of grand slam titles they have won
'and also which slam they won
'The pourpose of this page is to display players with the most grand slam titles current and past.






Private Sub cmdend_Click()
End
End Sub

Private Sub cmdfind_Click()
Dim player As String, found As Boolean, pos As Integer, find As Boolean, pos1 As Integer
'player inputs a name
player = InputBox("Enter a players name", , "Player")
pos = 0
found = False
'program searches for player
Do While found = False And pos <= ctr
pos = pos + 1
If names(pos) = player Then
found = True
End If
Loop
'if found then it will display the player and the number of titles he's won
If found = True Then
MsgBox "Your player " & names(pos) & " has won " & slams(pos) & " Grand Slams", , "Slams"
Else
'program will also search for players that have only won 1 title
Open App.Path & "\oneslam.txt" For Input As #2
ctr2 = 0
Do Until EOF(2)
ctr2 = ctr2 + 1
Input #2, name1(ctr2)
Loop
Close #2
pos1 = 0
find = False
Do Until find = True Or pos1 = ctr2
pos1 = pos1 + 1
If name1(pos1) = player Then
find = True
End If
Loop
If find = True Then
MsgBox "Your player " & name1(pos1) & " " & "has won one Grand Slam in his career", , "Slams"
Else
'informs user that the player they typed has not won a grand slam
MsgBox "Sorry your player has not won a Grand Slam in his career.", , "What a loser"
End If
End If



End Sub
'goes to next page
Private Sub cmdgo_Click()
frmtitles.Visible = False
frmpics.Visible = True

End Sub

Private Sub cmdshow_Click()
Dim pos As Integer
'loads information from text
Open App.Path & "\multislam.txt" For Input As #1
ctr = 0


Do Until EOF(1)

ctr = ctr + 1
Input #1, names(ctr), country(ctr), slams(ctr)

Loop

Close #1
'displays all players and number of slams they've won
pos = 0
picdisplay.Print "The following players have won more then 1 Grand Slam:"
picdisplay.Print "*********************************************************************"
Do Until pos = ctr
pos = pos + 1
picdisplay.Print names(pos); Tab(20); country(pos); Tab(40); slams(pos)
Loop

End Sub
