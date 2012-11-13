VERSION 5.00
Begin VB.Form frmBasketball 
   BackColor       =   &H00000000&
   Caption         =   "CU Buffs"
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10485
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   10485
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNavigate 
      BackColor       =   &H0000FFFF&
      Caption         =   "Navigate"
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
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5400
      Width           =   1695
   End
   Begin VB.PictureBox picbox 
      BackColor       =   &H00FFFFFF&
      Height          =   4935
      Left            =   4440
      ScaleHeight     =   4875
      ScaleWidth      =   5475
      TabIndex        =   9
      Top             =   1920
      Width           =   5535
      Begin VB.PictureBox picBuffs 
         Height          =   2175
         Left            =   2520
         Picture         =   "BASKET~1.frx":0000
         ScaleHeight     =   2115
         ScaleWidth      =   2835
         TabIndex        =   10
         Top             =   2640
         Width           =   2895
      End
   End
   Begin VB.TextBox txtSearch 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   4560
      TabIndex        =   8
      Top             =   360
      Width           =   5175
   End
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
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6480
      Width           =   1695
   End
   Begin VB.CommandButton cmdPurchase 
      BackColor       =   &H0000FFFF&
      Caption         =   "Buy Tickets"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton cmdAssists 
      BackColor       =   &H0000FFFF&
      Caption         =   "Most Assists"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton cmdRebound 
      BackColor       =   &H0000FFFF&
      Caption         =   "Most Rebounds"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton cmdTeam 
      BackColor       =   &H0000FFFF&
      Caption         =   "Team Average of Points Per Game"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton cmdHighscore 
      BackColor       =   &H0000FFFF&
      Caption         =   "Highest Scorer"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H0000FFFF&
      Caption         =   "Find Player"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton cmdRead 
      BackColor       =   &H0000FFFF&
      Caption         =   "Read Stats"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmBasketball"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: ((BBALL.vbp).vbp)
'Form Name: ((BBALL.frm).frm)
'Module Name: (CObuffs.bas)
'Author: Jessica Bankers'
'Date Written: 10/31/05'
'Purpose: This form sorts the colorado buffalos women's basketball team statistics into 7 arrays.
'Then searches, sorts, and or computes the players stats.'
'The purpose of this project is so that individuals can search for players and their personal statistics quickly. They can also compare statistics of players based on several factors; points, assists,  and rebounds.
'The module created is used to dim all variabls for all forms. This way anyone can view the variabls dimmed.

Option Explicit
Dim PlayerName(1 To 13) As String
Dim Games(1 To 13) As Integer
Dim Minutes(1 To 13) As Single
Dim TwoPoints(1 To 13) As Single
Dim ThreePoints(1 To 13) As Single
Dim Rebounds(1 To 13) As Single
Dim Assists(1 To 13) As Single
Dim LookupName As String
Dim J As Integer
Dim Q As Integer
Dim Counter, NumElements As Integer
Dim NotFound As Boolean
Dim Pass As Integer
Dim Comp As Integer
Dim Leader As Integer
Dim Winner As Integer

Dim TeamScore As Single

Private Sub cmdAssists_Click()
picbox.Cls
J = 1
Winner = Assists(1)
For Pass = 2 To 13
    If Winner < Assists(Pass) Then
            Winner = Assists(Pass)
            J = Pass
            End If
    Next Pass

picbox.Print PlayerName(J); " averages"; Round(Assists(J)); "Total assisits in a season"
End Sub

Private Sub cmdHighscore_Click()
picbox.Cls
J = 1
Leader = TwoPoints(1)
For Pass = 2 To 13
    If Leader < TwoPoints(Pass) Then
            Leader = TwoPoints(Pass)
            J = Pass
            End If
    Next Pass

picbox.Print PlayerName(J), "averages"; Round(TwoPoints(J) * 2); "points per season"

End Sub

Private Sub cmdNavigate_Click()
frmBasketball.Hide
frmpictures.Show

End Sub

Private Sub cmdPurchase_Click()
MsgBox " Go to http://www.cubuffs.com/ViewArticle.dbml?DB_OEM_ID=600&KEY=&ATCLID=69315", , "Colorado Buffalo Ticket Information"
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdRead_Click()
picbox.Cls
Open App.Path & "\statistics.txt" For Input As #1
   NumElements = 0
    For NumElements = 1 To 13
        Input #1, PlayerName(NumElements), Games(NumElements), Minutes(NumElements), TwoPoints(NumElements), ThreePoints(NumElements), Rebounds(NumElements), Assists(NumElements)
    Next NumElements
Close #1
End Sub

Private Sub cmdRebound_Click()
picbox.Cls
J = 1
Winner = Rebounds(1)
For Pass = 2 To 13
    If Winner < Rebounds(Pass) Then
            Winner = Rebounds(Pass)
            J = Pass
            End If
    Next Pass

picbox.Print PlayerName(J); " averages"; Round(Rebounds(J)); "Rebounds a season"
End Sub

Private Sub cmdSearch_Click()

   picbox.Cls
   LookupName = txtSearch.Text
    NotFound = True
    Counter = 1
    Do While NotFound And Counter <= 13
        If LookupName = PlayerName(Counter) Then
            picbox.Print "Player Name           ", PlayerName(Counter)
            picbox.Print "Games Played          ", Games(Counter)
            picbox.Print "Total Minutes Played  ", Minutes(Counter)
            picbox.Print "Two Pointers          ", TwoPoints(Counter)
            picbox.Print "Three Pointers        ", ThreePoints(Counter)
            picbox.Print "Total Rebounds        ", Rebounds(Counter)
            picbox.Print "Total Assists         ", Assists(Counter)
        
            NotFound = False
            Counter = Counter + 1
        Else
           Counter = Counter + 1
            End If
    Loop
    If NotFound Then
        picbox.Print LookupName; " not Found"
    End If
End Sub


Private Sub cmdTeam_Click()
picbox.Cls

    TeamScore = (TwoPoints(1) + TwoPoints(2) + TwoPoints(3) + TwoPoints(4) + TwoPoints(5) + TwoPoints(6) + TwoPoints(7) + TwoPoints(8) + TwoPoints(9) + TwoPoints(10) + TwoPoints(11) + TwoPoints(12) + TwoPoints(13)) / 13
    picbox.Print "The Team Point Average Per game is"; Round(TeamScore)
    
End Sub
