VERSION 5.00
Begin VB.Form frmReceivingStats 
   BackColor       =   &H00FF0000&
   Caption         =   "ReceivingStats"
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9075
   LinkTopic       =   "Form1"
   ScaleHeight     =   4395
   ScaleWidth      =   9075
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picJohnson 
      Height          =   2535
      Left            =   6840
      Picture         =   "frmReceivingStats.frx":0000
      ScaleHeight     =   2475
      ScaleWidth      =   2115
      TabIndex        =   7
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton cmdTouchdowns 
      Caption         =   "Touchdowns"
      Height          =   615
      Left            =   7440
      TabIndex        =   5
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdLongest 
      Caption         =   "Longest"
      Height          =   615
      Left            =   6120
      TabIndex        =   4
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmdAverage 
      Caption         =   "Average"
      Height          =   615
      Left            =   4680
      TabIndex        =   3
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdYards 
      Caption         =   "Yards"
      Height          =   615
      Left            =   3240
      TabIndex        =   2
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdCatches 
      Caption         =   "Catches"
      Height          =   615
      Left            =   1800
      TabIndex        =   1
      Top             =   3600
      Width           =   1095
   End
   Begin VB.PictureBox picReceivingStats 
      BackColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   240
      ScaleHeight     =   2955
      ScaleWidth      =   6435
      TabIndex        =   0
      Top             =   240
      Width           =   6495
   End
   Begin VB.Label lblJohnson 
      BackColor       =   &H00FF0000&
      Caption         =   "#19 Keyshawn      Johnson"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6840
      TabIndex        =   8
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label lblSortBy 
      BackColor       =   &H00FF0000&
      Caption         =   "Sort By:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   3600
      Width           =   1455
   End
End
Attribute VB_Name = "frmReceivingStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Integer
Dim Temp As Integer
Dim TempPlayer As String
Dim TempNumber As Double, TempYards As Double, TempAverage As Double, TempLongest As Double, TempTouchdowns As Double
Dim Pass As Integer
Dim Player(1 To 12) As String
Dim Number(1 To 12) As Double, Yards(1 To 12) As Double, Average(1 To 12) As Double, Longest(1 To 12) As Double, TouchDowns(1 To 12) As Double

Private Sub cmdAverage_Click() 'This button sorts the information by the average yards for the players'
    picReceivingStats.Cls
    picReceivingStats.Print "Player"; Tab(20); "Number", "Yards"; Tab(40); "Average", "Longest", "Touch Downs"
    picReceivingStats.Print "************************************************************************************************************************"
    Open App.Path & "\receivingstats.txt" For Input As #2
    For I = 1 To 12
        Input #2, Player(I), Number(I), Yards(I), Average(I), Longest(I), TouchDowns(I)
    Next I
    Close #2
    For Pass = 1 To 11
        For I = 1 To (12 - Pass)
            If Average(I) < Average(I + 1) Then
                TempNumber = Number(I)
                Number(I) = Number(I + 1)
                Number(I + 1) = TempNumber
                TempPlayer = Player(I)
                Player(I) = Player(I + 1)
                Player(I + 1) = TempPlayer
                TempYards = Yards(I)
                Yards(I) = Yards(I + 1)
                Yards(I + 1) = TempYards
                TempAverage = Average(I)
                Average(I) = Average(I + 1)
                Average(I + 1) = TempAverage
                TempLongest = Longest(I)
                Longest(I) = Longest(I + 1)
                Longest(I + 1) = TempLongest
                TempTouchdowns = TouchDowns(I)
                TouchDowns(I) = TouchDowns(I + 1)
                TouchDowns(I + 1) = TempTouchdowns
            End If
        Next I
    Next Pass
    For I = 1 To 12
        picReceivingStats.Print Player(I); Tab(20); Number(I), Yards(I); Tab(40); Average(I), Longest(I), TouchDowns(I)
    Next I
End Sub

Private Sub cmdCatches_Click() 'This button sorts the receiving stats by how many catches the players have'
    picReceivingStats.Cls
    picReceivingStats.Print "Player"; Tab(20); "Number", "Yards"; Tab(40); "Average", "Longest", "Touch Downs"
    picReceivingStats.Print "************************************************************************************************************************"
    Open App.Path & "\receivingstats.txt" For Input As #2
    For I = 1 To 12
        Input #2, Player(I), Number(I), Yards(I), Average(I), Longest(I), TouchDowns(I)
    Next I
    Close #2
    For Pass = 1 To 11
        For I = 1 To (12 - Pass)
            If Number(I) < Number(I + 1) Then
                TempNumber = Number(I)
                Number(I) = Number(I + 1)
                Number(I + 1) = TempNumber
                TempPlayer = Player(I)
                Player(I) = Player(I + 1)
                Player(I + 1) = TempPlayer
                TempYards = Yards(I)
                Yards(I) = Yards(I + 1)
                Yards(I + 1) = TempYards
                TempAverage = Average(I)
                Average(I) = Average(I + 1)
                Average(I + 1) = TempAverage
                TempLongest = Longest(I)
                Longest(I) = Longest(I + 1)
                Longest(I + 1) = TempLongest
                TempTouchdowns = TouchDowns(I)
                TouchDowns(I) = TouchDowns(I + 1)
                TouchDowns(I + 1) = TempTouchdowns
            End If
        Next I
    Next Pass
    For I = 1 To 12
        picReceivingStats.Print Player(I); Tab(20); Number(I), Yards(I); Tab(40); Average(I), Longest(I), TouchDowns(I)
    Next I
End Sub

Private Sub cmdLongest_Click() 'This button sorts the information by the longest catch for the players'
     picReceivingStats.Cls
    picReceivingStats.Print "Player"; Tab(20); "Number", "Yards"; Tab(40); "Average", "Longest", "Touch Downs"
    picReceivingStats.Print "************************************************************************************************************************"
    Open App.Path & "\receivingstats.txt" For Input As #2
    For I = 1 To 12
        Input #2, Player(I), Number(I), Yards(I), Average(I), Longest(I), TouchDowns(I)
    Next I
    Close #2
    For Pass = 1 To 11
        For I = 1 To (12 - Pass)
            If Longest(I) < Longest(I + 1) Then
                TempNumber = Number(I)
                Number(I) = Number(I + 1)
                Number(I + 1) = TempNumber
                TempPlayer = Player(I)
                Player(I) = Player(I + 1)
                Player(I + 1) = TempPlayer
                TempYards = Yards(I)
                Yards(I) = Yards(I + 1)
                Yards(I + 1) = TempYards
                TempAverage = Average(I)
                Average(I) = Average(I + 1)
                Average(I + 1) = TempAverage
                TempLongest = Longest(I)
                Longest(I) = Longest(I + 1)
                Longest(I + 1) = TempLongest
                TempTouchdowns = TouchDowns(I)
                TouchDowns(I) = TouchDowns(I + 1)
                TouchDowns(I + 1) = TempTouchdowns
            End If
        Next I
    Next Pass
    For I = 1 To 12
        picReceivingStats.Print Player(I); Tab(20); Number(I), Yards(I); Tab(40); Average(I), Longest(I), TouchDowns(I)
    Next I
End Sub

Private Sub cmdTouchdowns_Click() 'This button sorts the information by the number of touchdowns the players have'
     picReceivingStats.Cls
    picReceivingStats.Print "Player"; Tab(20); "Number", "Yards"; Tab(40); "Average", "Longest", "Touch Downs"
    picReceivingStats.Print "************************************************************************************************************************"
    Open App.Path & "\receivingstats.txt" For Input As #2
    For I = 1 To 12
        Input #2, Player(I), Number(I), Yards(I), Average(I), Longest(I), TouchDowns(I)
    Next I
    Close #2
    For Pass = 1 To 11
        For I = 1 To (12 - Pass)
            If TouchDowns(I) < TouchDowns(I + 1) Then
                TempNumber = Number(I)
                Number(I) = Number(I + 1)
                Number(I + 1) = TempNumber
                TempPlayer = Player(I)
                Player(I) = Player(I + 1)
                Player(I + 1) = TempPlayer
                TempYards = Yards(I)
                Yards(I) = Yards(I + 1)
                Yards(I + 1) = TempYards
                TempAverage = Average(I)
                Average(I) = Average(I + 1)
                Average(I + 1) = TempAverage
                TempLongest = Longest(I)
                Longest(I) = Longest(I + 1)
                Longest(I + 1) = TempLongest
                TempTouchdowns = TouchDowns(I)
                TouchDowns(I) = TouchDowns(I + 1)
                TouchDowns(I + 1) = TempTouchdowns
            End If
        Next I
    Next Pass
    For I = 1 To 12
        picReceivingStats.Print Player(I); Tab(20); Number(I), Yards(I); Tab(40); Average(I), Longest(I), TouchDowns(I)
    Next I
End Sub

Private Sub cmdYards_Click() 'This button sorts the information by number of yards for the players'
     picReceivingStats.Cls
    picReceivingStats.Print "Player"; Tab(20); "Number", "Yards"; Tab(40); "Average", "Longest", "Touch Downs"
    picReceivingStats.Print "************************************************************************************************************************"
    Open App.Path & "\receivingstats.txt" For Input As #2
    For I = 1 To 12
        Input #2, Player(I), Number(I), Yards(I), Average(I), Longest(I), TouchDowns(I)
    Next I
    Close #2
    For Pass = 1 To 11
        For I = 1 To (12 - Pass)
            If Yards(I) < Yards(I + 1) Then
                TempNumber = Number(I)
                Number(I) = Number(I + 1)
                Number(I + 1) = TempNumber
                TempPlayer = Player(I)
                Player(I) = Player(I + 1)
                Player(I + 1) = TempPlayer
                TempYards = Yards(I)
                Yards(I) = Yards(I + 1)
                Yards(I + 1) = TempYards
                TempAverage = Average(I)
                Average(I) = Average(I + 1)
                Average(I + 1) = TempAverage
                TempLongest = Longest(I)
                Longest(I) = Longest(I + 1)
                Longest(I + 1) = TempLongest
                TempTouchdowns = TouchDowns(I)
                TouchDowns(I) = TouchDowns(I + 1)
                TouchDowns(I + 1) = TempTouchdowns
            End If
        Next I
    Next Pass
    For I = 1 To 12
        picReceivingStats.Print Player(I); Tab(20); Number(I), Yards(I); Tab(40); Average(I), Longest(I), TouchDowns(I)
    Next I
End Sub

Private Sub picReceivingStats_Paint()
    
    Open App.Path & "\receivingstats.txt" For Input As #2
    picReceivingStats.Print "Player"; Tab(20); "Number", "Yards"; Tab(40); "Average", "Longest", "Touch Downs"
    picReceivingStats.Print "*********************************************************************************************************"
    For I = 1 To 12
        Input #2, Player(I), Number(I), Yards(I), Average(I), Longest(I), TouchDowns(I)
        picReceivingStats.Print Player(I); Tab(20); Number(I), Yards(I); Tab(40); Average(I), Longest(I), TouchDowns(I)
    Next I
    Close #2
End Sub
