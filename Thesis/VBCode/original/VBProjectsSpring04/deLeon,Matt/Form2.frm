VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H0080FF80&
   Caption         =   "Form2"
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10740
   LinkTopic       =   "Form2"
   ScaleHeight     =   7290
   ScaleWidth      =   10740
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSortP 
      Caption         =   "Sort by Points"
      Height          =   855
      Left            =   600
      TabIndex        =   7
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton Cmdf3 
      Caption         =   "Go To Form 3"
      Height          =   855
      Left            =   8280
      TabIndex        =   6
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Go to Form1"
      Height          =   855
      Left            =   8280
      TabIndex        =   5
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton cmdplayer 
      Caption         =   "Get Player Statistics"
      Height          =   855
      Left            =   600
      TabIndex        =   4
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton cmdquit2 
      Caption         =   "QUIT"
      Height          =   735
      Left            =   600
      TabIndex        =   3
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton cmdclear2 
      Caption         =   "CLEAR"
      Height          =   735
      Left            =   600
      TabIndex        =   2
      Top             =   5280
      Width           =   1695
   End
   Begin VB.PictureBox picresults2 
      Height          =   4575
      Left            =   2880
      ScaleHeight     =   4515
      ScaleWidth      =   3555
      TabIndex        =   1
      Top             =   120
      Width           =   3615
   End
   Begin VB.CommandButton cmdread2 
      Caption         =   "See Statistics"
      Height          =   855
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin VB.Image Image2 
      Height          =   4620
      Left            =   6600
      Picture         =   "Form2.frx":0000
      Top             =   120
      Width           =   3810
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name:TeamStats.vbg
'Author Matt de Leon
'Form Name: Form 2 (Form2.frm)
'Written March 15 2004
'Purpose of form to allow user to see the statistics of the basketball team on form 1.
' Form also allows users to sort the players by their point averages.
'Form also allows user to find stats on  a specific player
Option Explicit
Dim AvgPoints(1 To 50) As Double
Dim Player(1 To 50) As String
Dim Rebounds(1 To 50) As Double
Dim Assists(1 To 50) As Double
Dim CTR As Integer
Dim app As Double
Dim runningtotal As Double
Dim slot As Integer
Dim Found As Boolean
Dim Path As String





Private Sub cmdback_Click()
Form1.Show
Form2.Hide
End Sub



Private Sub cmdclear2_Click()
picresults2.Cls
End Sub









Private Sub Cmdf3_Click()
Form3.Show
Form2.Hide
Form1.Hide
End Sub

Private Sub cmdplayer_Click()
Dim nplayer As String
nplayer = InputBox("Enter a name") 'Get name from User
slot = 0
Do While (Not Found) And (slot < CTR) 'Continue looking for player until found in the array
slot = slot + 1
If nplayer = Player(slot) Then
Found = True
picresults2.Print "Player", "Points", "REBS", "ASTS"
picresults2.Print Player(slot), AvgPoints(slot), Rebounds(slot), Assists(slot)
End If
Loop
If Found = False Then
picresults2.Print "player not found"
End If
End Sub





Private Sub cmdquit2_Click()
End
End Sub

Private Sub cmdread2_Click()
 'Start CTR at zero, to be used for position in the array
    CTR = 0
   
    'Open file to be used
    Open Path & "stats.txt" For Input As #1
    picresults2.Print "Player", "Points", "REBS", "ASTS"
    picresults2.Print "-------------------------------------------"
    Do While Not EOF(1)
        'Add one to CTR each time through the loop, to move to next spot in array
        CTR = CTR + 1
        
        'Put data into the array then print
        Input #1, Player(CTR), AvgPoints(CTR), Rebounds(CTR), Assists(CTR)
        picresults2.Print Player(CTR); Tab(17); AvgPoints(CTR); Tab(27); Rebounds(CTR); Tab(40); Assists(CTR); Tab(48)
        
        'add points to runningTotal
        runningtotal = runningtotal + AvgPoints(CTR)
    Loop
   'print the average total points for all players
    picresults2.Print "---------------------------------------------"

    picresults2.Print "Avg Total Points = "; Tab(20); runningtotal
   Close #1
    'Disable the command button for Reading the file
    cmdread2.Enabled = False
End Sub








Private Sub cmdSortP_Click()
Dim name As String
Dim points As Double
Dim PASS As Single, COMP As Double, J As Single

'Arrange players by points
For PASS = 1 To CTR - 1
For COMP = 1 To CTR - PASS
If AvgPoints(COMP) < AvgPoints(COMP + 1) Then
        
'get players in correct place
name = Player(COMP)
Player(COMP) = Player(COMP + 1)
Player(COMP + 1) = name
            
 'get players points in correct place
points = AvgPoints(COMP)
AvgPoints(COMP) = AvgPoints(COMP + 1)
AvgPoints(COMP + 1) = points
            
End If
Next COMP
Next PASS

picresults2.Print
picresults2.Print "------------Players by Points------------------"

'Put Players in Order by Points
For J = 1 To CTR
picresults2.Print Player(J), AvgPoints(J)
Next J
End Sub

Private Sub Form_Load()
Path = "N:\CS130\handin\deLeon,Matt\"
End Sub
