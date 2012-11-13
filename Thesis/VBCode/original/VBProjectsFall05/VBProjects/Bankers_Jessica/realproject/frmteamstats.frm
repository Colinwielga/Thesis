VERSION 5.00
Begin VB.Form frmteamstats 
   BackColor       =   &H00000000&
   Caption         =   "Team Statistics"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   9465
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0000FFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton cmdTeam 
      BackColor       =   &H0000FFFF&
      Caption         =   "Read Team Stats"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdpoints 
      BackColor       =   &H0000FFFF&
      Caption         =   "Points"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdAssists 
      BackColor       =   &H0000FFFF&
      Caption         =   "Assists"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdRebound 
      BackColor       =   &H0000FFFF&
      Caption         =   "Rebounds"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4560
      Width           =   1215
   End
   Begin VB.PictureBox picbox 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   1680
      ScaleHeight     =   6075
      ScaleWidth      =   7515
      TabIndex        =   0
      Top             =   480
      Width           =   7575
      Begin VB.PictureBox picBuffs 
         Height          =   2055
         Left            =   4560
         Picture         =   "frmteamstats.frx":0000
         ScaleHeight     =   1995
         ScaleWidth      =   2835
         TabIndex        =   1
         Top             =   3960
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmteamstats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAssists_Click()
picbox.Cls

For Pass = 1 To 13 - 1
    For J = 1 To 13 - Pass
        If (Assists(J) > (Assists(J + 1))) Then
        tempname = PlayerName(J)
        PlayerName(J) = PlayerName(J + 1)
        PlayerName(J + 1) = tempname
        temptwopoints = TwoPoints(J)
        TwoPoints(J) = TwoPoints(J + 1)
        TwoPoints(J + 1) = temptwopoints
        tempthreepoints = ThreePoints(J)
        ThreePoints(J) = ThreePoints(J + 1)
        ThreePoints(J + 1) = tempthreepoints
        tempgames = Games(J)
        Games(J) = Games(J + 1)
        Games(J + 1) = tempgames
        temprebounds = Rebounds(J)
        Rebounds(J) = Rebounds(J + 1)
        Rebounds(J + 1) = temprebounds
        tempassists = Assists(J)
        Assists(J) = Assists(J + 1)
        Assists(J + 1) = tempassists
        End If
    Next J
Next Pass

Q = 1
picbox.Print "Player Name"; Tab(30); "Assists"
For Q = 1 To 13
    picbox.Print PlayerName(Q); Tab(30); Assists(Q)
Next Q

Close #1
End Sub

Private Sub cmdpoints_Click()
picbox.Cls

For Pass = 1 To 13 - 1
    For J = 1 To 13 - Pass
        If (TwoPoints(J) * 2 + ThreePoints(J) * 3) > (TwoPoints(J + 1) * 2 + ThreePoints(J + 1) * 3) Then
        tempname = PlayerName(J)
        PlayerName(J) = PlayerName(J + 1)
        PlayerName(J + 1) = tempname
        temptwopoints = TwoPoints(J)
        TwoPoints(J) = TwoPoints(J + 1)
        TwoPoints(J + 1) = temptwopoints
        tempthreepoints = ThreePoints(J)
        ThreePoints(J) = ThreePoints(J + 1)
        ThreePoints(J + 1) = tempthreepoints
        tempgames = Games(J)
        Games(J) = Games(J + 1)
        Games(J + 1) = tempgames
        temprebounds = Rebounds(J)
        Rebounds(J) = Rebounds(J + 1)
        Rebounds(J + 1) = temprebounds
        tempassists = Assists(J)
        Assists(J) = Assists(J + 1)
        Assists(J + 1) = tempassists
        

        
        End If
    Next J
Next Pass

Q = 1
picbox.Print "Player Name"; Tab(30); "Two Points", "Three Points"
For Q = 1 To 13
    
    picbox.Print PlayerName(Q); Tab(30); TwoPoints(Q), ThreePoints(Q)
    
Next Q

Close #1
End Sub


Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdRebound_Click()
picbox.Cls

For Pass = 1 To 13 - 1
    For J = 1 To 13 - Pass
        If (Rebounds(J) > (Rebounds(J + 1))) Then
        tempname = PlayerName(J)
        PlayerName(J) = PlayerName(J + 1)
        PlayerName(J + 1) = tempname
        temptwopoints = TwoPoints(J)
        TwoPoints(J) = TwoPoints(J + 1)
        TwoPoints(J + 1) = temptwopoints
        tempthreepoints = ThreePoints(J)
        ThreePoints(J) = ThreePoints(J + 1)
        ThreePoints(J + 1) = tempthreepoints
        tempgames = Games(J)
        Games(J) = Games(J + 1)
        Games(J + 1) = tempgames
        temprebounds = Rebounds(J)
        Rebounds(J) = Rebounds(J + 1)
        Rebounds(J + 1) = temprebounds
        tempassists = Assists(J)
        Assists(J) = Assists(J + 1)
        Assists(J + 1) = tempassists
        End If
    Next J
Next Pass

Q = 1
picbox.Print "Player Name"; Tab(30); "Rebounds"
For Q = 1 To 13
    picbox.Print PlayerName(Q); Tab(30); Rebounds(Q)
Next Q

Close #1
End Sub


Private Sub cmdTeam_Click()
Open App.Path & "\teamstats.txt" For Input As #1
   NumElements = 1
    For NumElements = 1 To 13
        Input #1, PlayerName(NumElements), Games(NumElements), Minutes(NumElements), TwoPoints(NumElements), ThreePoints(NumElements), Rebounds(NumElements), Assists(NumElements)
    Next NumElements
    
End Sub
