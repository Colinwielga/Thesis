VERSION 5.00
Begin VB.Form frmstats 
   BackColor       =   &H00FF0000&
   Caption         =   "2006 Comparisions"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   10995
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdind 
      Caption         =   "Click to view individual statistics"
      Height          =   975
      Left            =   240
      TabIndex        =   8
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Return to Main Menu"
      Height          =   975
      Left            =   2280
      TabIndex        =   6
      Top             =   5160
      Width           =   1815
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear"
      Height          =   975
      Left            =   2400
      TabIndex        =   5
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton cmdload 
      Caption         =   "Load Players and Data"
      Height          =   855
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   2415
   End
   Begin VB.CommandButton cmdjumps 
      Caption         =   "Top Vertical Jumps"
      Height          =   975
      Left            =   240
      TabIndex        =   3
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CommandButton cmdforty 
      Caption         =   "Top 40 yard sprints"
      Height          =   975
      Left            =   2400
      TabIndex        =   2
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton cmdbench 
      Caption         =   "Top Max Bench Presses"
      Height          =   975
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   1815
   End
   Begin VB.PictureBox picDisplay 
      Height          =   12255
      Left            =   4440
      ScaleHeight     =   12195
      ScaleWidth      =   6075
      TabIndex        =   0
      Top             =   120
      Width           =   6135
   End
   Begin VB.Label lblstats 
      BackColor       =   &H0000FFFF&
      Caption         =   "Search Top 40 in each Category"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label lblinfo 
      BackColor       =   &H0000FFFF&
      Caption         =   "Click the Clear Button, then reload Data to View next top 30 athletes"
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   3000
      Width           =   3855
   End
End
Attribute VB_Name = "frmstats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Size As Integer
Dim Pos As Integer
Dim Pass As Integer
Dim Temp1, Temp2, Temp3 As Single
'2006 NFL Draft Simulator (Draft.vbp)
'frmstats(frmstats.frm)
'Andy Lyons
'March 24, 2006
'The purpose of this form is to let the user look at the top 30 players statistics in the draft.  Through clicking on the buttons, it will arrange the players by name and their statistic.
'Clicking this button sorts the players name and their maximum bench press.  The user can see this in the picture box, from the best athlete to the least.
Private Sub cmdbench_Click()
For Pass = 1 To (60 - 1)
    For Pos = 1 To (60 - Pass)
        If Bench(Pos) < Bench(Pos + 1) Then
            Temp1 = Bench(Pos)
            Bench(Pos) = Bench(Pos + 1)
            Bench(Pos + 1) = Temp1
        
            Temp2 = Names(Pos)
            Names(Pos) = Names(Pos + 1)
            Names(Pos + 1) = Temp2
        End If
    Next Pos

Next Pass

For Pos = 1 To 30
    picDisplay.Print Names(Pos); Tab(30); Bench(Pos)
    Next Pos
End Sub
'Clears contents in picture box
Private Sub cmdclear_Click()
    picDisplay.Cls
End Sub

'Clicking this button allows the user to view what 30 players have the fastest forty yard sprint.  This button arranges them from the fastest to the slowest.

Private Sub cmdforty_Click()
For Pass = 1 To (60 - 1)
    For Pos = 1 To (60 - Pass)
        If Forty(Pos) > Forty(Pos + 1) Then
            Temp2 = Forty(Pos)
            Forty(Pos) = Forty(Pos + 1)
            Forty(Pos + 1) = Temp2
        
            Temp2 = Names(Pos)
            Names(Pos) = Names(Pos + 1)
            Names(Pos + 1) = Temp2
        End If
    Next Pos
Next Pass
For Pos = 1 To 30
    picDisplay.Print Names(Pos); Tab(30); Forty(Pos)
    Next Pos
End Sub

'This button allows the user to look at each individual athletes statistics with the click of the button"
Private Sub cmdind_Click()
    picDisplay.Print "Names"; Tab(25); "Bench"; Tab(35); "Forty Time"; Tab(47); "Jump"
For Pos = 1 To 60
     picDisplay.Print Names(Pos); Tab(25); Bench(Pos); Tab(35); Forty(Pos); Tab(47); Jump(Pos)
Next Pos
End Sub

'Clicking this button sorts the data from the best 30 player's vertical leap to the least among the players.
Private Sub cmdjumps_Click()
For Pass = 1 To (60 - 1)
    For Pos = 1 To (60 - Pass)
        If Jump(Pos) < Jump(Pos + 1) Then
            Temp3 = Jump(Pos)
            Jump(Pos) = Jump(Pos + 1)
            Jump(Pos + 1) = Temp3
            
            Temp2 = Names(Pos)
            Names(Pos) = Names(Pos + 1)
            Names(Pos + 1) = Temp2
        
        End If
    Next Pos
Next Pass
For Pos = 1 To 30
    picDisplay.Print Names(Pos); Tab(30); Jump(Pos)
    Next Pos
End Sub
'Clicking this button loads the data for the three categories(Bench, Forty dash, and Vertical Leap). This button loads the data for sorting.
Private Sub cmdload_Click()
    Pos = 0
    Open App.Path & "\draftnames.txt" For Input As #1
        
        Do Until EOF(1)
        Pos = Pos + 1
        Input #1, Names(Pos), Position(Pos), Bench(Pos), Forty(Pos), Jump(Pos)
   
    Loop
    Close #1

End Sub
'returns user to main menu
Private Sub Command1_Click()
    frmNFLDraft.Show
    frmstats.Hide
End Sub
