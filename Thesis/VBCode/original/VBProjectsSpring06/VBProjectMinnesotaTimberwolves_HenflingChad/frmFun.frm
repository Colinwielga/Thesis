VERSION 5.00
Begin VB.Form frmFun 
   BackColor       =   &H8000000D&
   Caption         =   "Fan Fun Section!"
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10635
   DrawStyle       =   3  'Dash-Dot
   FillColor       =   &H000000FF&
   ForeColor       =   &H80000002&
   LinkTopic       =   "Form1"
   Picture         =   "frmFun.frx":0000
   ScaleHeight     =   8730
   ScaleWidth      =   10635
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   1695
      Left            =   3480
      Picture         =   "frmFun.frx":11622
      ScaleHeight     =   1635
      ScaleWidth      =   2355
      TabIndex        =   8
      Top             =   840
      Width           =   2415
   End
   Begin VB.CommandButton cmdDisplay 
      BackColor       =   &H8000000C&
      Caption         =   "Display total Votes thus far."
      Height          =   1815
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6120
      Width           =   2535
   End
   Begin VB.PictureBox picOutput 
      BackColor       =   &H8000000A&
      Height          =   8295
      Left            =   6600
      ScaleHeight     =   8235
      ScaleWidth      =   3675
      TabIndex        =   6
      Top             =   0
      Width           =   3735
   End
   Begin VB.CommandButton cmdVote 
      BackColor       =   &H8000000C&
      Caption         =   "Vote For Your Favorite Timberwolves Player!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3720
      Width           =   2775
   End
   Begin VB.ComboBox cmbVote 
      BackColor       =   &H80000003&
      Height          =   315
      ItemData        =   "frmFun.frx":13105
      Left            =   3360
      List            =   "frmFun.frx":1311E
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   3360
      Width           =   2895
   End
   Begin VB.CommandButton cmdPPG 
      BackColor       =   &H8000000C&
      Caption         =   "Calculate Your Points Per Game"
      Height          =   1935
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3240
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000C&
      Caption         =   "Compute your shot Percentage!"
      Height          =   1935
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H8000000C&
      Caption         =   "Go Back"
      Height          =   1815
      Left            =   0
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6120
      Width           =   2535
   End
   Begin VB.Label lblName 
      BackColor       =   &H8000000D&
      Caption         =   "By: Chad Henfling"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "frmFun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Minnesota Timberwolves Center (MinnesotaTimberwovlesbyChadHenfling.vbp)
'Main Form (frmFun.frm)
'Chad Henfling
'Created March 23, 2006
'This form allows the user to calculate his or her free throw percentage and points per game.
Option Explicit
    Dim KG, Ricky, Mark, Rashad, Trenton, Troy, Eddie As String
    Dim tally1, t2, t3, t4, t5, t6, t7 As Integer
Private Sub cmdBack_Click()
    'going back to main form
    frmFun.Visible = False
    frm1.Visible = True
End Sub

Private Sub cmdDisplay_Click()
     'declaring variables
    Dim pos, big, counter As Integer
    Dim run(1 To 100) As Integer
    Dim person(1 To 100), guy(1 To 100) As String
    Dim tally(1 To 100) As Integer
    tally1 = 0 + tally1
    t2 = 0 + t2
    t3 = 0 + t3
    t4 = 0 + t4
    t5 = 0 + t5
    t6 = 0 + t6
    t7 = 0 + t7
    'Making sure everything is ready for output into text file
    KG = "Kevin Garnett"
    Ricky = "Ricky Davis"
    Mark = "Mark Madsen"
    Rashad = "Rashad McCants"
    Trenton = "Trenton Hassell"
    Troy = "Troy Hudson"
    Eddie = "Eddie Griffin"
    picOutput.Cls
    picOutput.Print "Total Votes Thus Far"
    picOutput.Print "------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Open App.Path & "/vote.txt" For Input As #4
        counter = 0
        Do Until EOF(4)
            counter = counter + 1
            Input #4, guy(counter), run(counter)
        Loop
        Close #4
    tally1 = run(1) + tally1
    t2 = run(2) + t2
    t3 = run(3) + t3
    t4 = run(4) + t4
    t5 = run(5) + t5
    t6 = run(6) + t6
    t7 = run(7) + t7
    'opening vote.txt for outputing the player and their tally of votes which is gained below from the list box
    Open App.Path & "\Vote.txt" For Output As #4
        Write #4, "Kevin Garnett  ", tally1
        Write #4, "Ricky Davis    ", t2
        Write #4, "Mark Madsen    ", t3
        Write #4, "Rashad McCants ", t4
        Write #4, "Trenton Hassell", t5
        Write #4, "Troy Hudson    ", t6
        Write #4, "Eddie Griffin     ", t7
    Close #4
    Open App.Path & "\Vote.txt" For Input As #4
    pos = 0
    'inputing file
        Do Until EOF(4)
            pos = pos + 1
            Input #4, person(pos), tally(pos)
        Loop
    Close #4
    big = pos
    'printing the results of the vote in picture box
    For pos = 1 To big
        picOutput.Print person(pos), tally(pos)
    Next pos
    tally1 = 0
    t2 = 0
    t3 = 0
    t4 = 0
    t5 = 0
    t6 = 0
    t7 = 0
End Sub

Private Sub cmdPPG_Click()
    Dim PPG, sum As Single
    Dim points, counter As Integer
    counter = 0
    points = 0
    'entering as many games and points scored as you want and finding the points per game average.
    Do While points <> -1
        counter = counter + 1
        points = InputBox("Enter your points game number " & counter & ".  Enter -1 if you are finished entering games", "Points")
        sum = sum + points
    Loop
    'calculating points per game
    sum = sum + 1
    counter = counter - 1
    PPG = sum / counter
    'displaying in message box the points per game
    MsgBox "Your Points Per Game =" & PPG, , "Points Per Game"
End Sub

Private Sub cmdVote_Click()
    picOutput.Cls
    KG = "Kevin Garnett"
    Ricky = "Ricky Davis"
    Mark = "Mark Madsen"
    Rashad = "Rashad McCants"
    Trenton = "Trenton Hassell"
    Troy = "Troy Hudson"
    Eddie = "Eddie Griffin"
    'after voting for a person this program goes through and tallies the number of votes for each person.
    If cmbVote = "Kevin Garnett" Then
        tally1 = tally1 + 1
        'displays player and vote
        picOutput.Print KG, , tally1
    End If
    If cmbVote = "Ricky Davis" Then
        t2 = t2 + 1
        picOutput.Print Ricky, , t2
    End If
    If cmbVote = "Mark Madsen" Then
        t3 = t3 + 1
        picOutput.Print Mark, , t3
    End If
    If cmbVote = "Rashad McCants" Then
        t4 = t4 + 1
        picOutput.Print Rashad, , t4
    End If
    If cmbVote = "Trenton Hassell" Then
        t5 = t5 + 1
        picOutput.Print Trenton, , t5
    End If
    If cmbVote = "Troy Hudson" Then
        t6 = t6 + 1
        picOutput.Print Troy, , t6
    End If
    If cmbVote = "Eddie Griffin" Then
        t7 = t7 + 1
        picOutput.Print Eddie, , t7
    End If
End Sub

Private Sub Command1_Click()
    Dim Percentage As Single
    Dim shots, made As Integer
    'finding percentage of shots made
    shots = InputBox("Enter your number of shots", "Shots")
    made = InputBox("enter the number of shots you made", "Shots Made")
    'displaying in message box the percentage or an impossible percentage
    If made > shots Then
        MsgBox "WOW!  YOU MADE MORE SHOTS THAN YOU ATTEMPTED!", , "IMPOSSIBLE"
    Else: Percentage = made / shots
    MsgBox "Your shooting percentage is " & Percentage, , "Shooting Percentage!"
    End If
End Sub

