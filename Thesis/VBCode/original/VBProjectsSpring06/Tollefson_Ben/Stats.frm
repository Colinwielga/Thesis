VERSION 5.00
Begin VB.Form frmStats 
   BackColor       =   &H8000000D&
   Caption         =   "Stats"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCalcAvg 
      Caption         =   "Calculate Batting Average"
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton cmdAvg 
      Caption         =   "Rank By Average"
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton cmdHits 
      Caption         =   "Rank By Number of Hits"
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmdHR 
      Caption         =   "Rank By Number of Home Runs"
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdAlphOrder 
      Caption         =   "Place In Alpabetical Order"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   1455
   End
   Begin VB.PictureBox picEnter 
      Height          =   4695
      Left            =   1920
      ScaleHeight     =   4635
      ScaleWidth      =   8595
      TabIndex        =   1
      Top             =   1080
      Width           =   8655
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search By Name"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   1455
   End
End
Attribute VB_Name = "frmStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Avg(1 To 20) As Single
Private Sub cmdAlphOrder_Click()
    'Ben Tollefson
    'March 22, 2006
    'This Button puts the players in alphabetical order
    
    picEnter.Cls
    Dim Pass As Integer
    Dim TempN As String
    Dim TempHR, TempH, TempAB As Integer
    Dim Pos As Integer
    picEnter.Print "Name"; Tab(20); "Home Runs"; Tab(35); "Hits"; Tab(45); "At-Bats"
    picEnter.Print
    For Pass = 1 To Size - 1
        For Pos = 1 To Size - Pass
            If Names(Pos) > Names(Pos + 1) Then
                TempHR = HR(Pos)
                HR(Pos) = HR(Pos + 1)
                HR(Pos + 1) = TempHR
                
                TempN = Names(Pos)
                Names(Pos) = Names(Pos + 1)
                Names(Pos + 1) = TempN
                
                TempH = Hits(Pos)
                Hits(Pos) = Hits(Pos + 1)
                Hits(Pos + 1) = TempH
                
                TempAB = AB(Pos)
                AB(Pos) = AB(Pos + 1)
                AB(Pos + 1) = TempAB
            End If
        Next Pos
    Next Pass
    For Pos = 1 To Size
    picEnter.Print Names(Pos); Tab(20); HR(Pos); Tab(35); Hits(Pos); Tab(45); AB(Pos)
    Next Pos
End Sub

Private Sub cmdAvg_Click()
    'Ben Tollefson
    'March 23, 2006
    'This Button is used to rank the players by highest Batting Average"
    
    Dim P, Pass As Integer
    Dim TempA As Single
    Dim TempNa As String
    picEnter.Cls
    picEnter.Print "Name"; Tab(20); "Batting Average"
    picEnter.Print
    For P = 1 To 10
            Avg(P) = Hits(P) / AB(P)
    Next P
    For Pass = 1 To Size - 1
        For P = 1 To Size - 1
            If Avg(P) < Avg(P + 1) Then
                TempA = Avg(P)
                Avg(P) = Avg(P + 1)
                Avg(P + 1) = TempA
                
                TempNa = Names(P)
                Names(P) = Names(P + 1)
                Names(P + 1) = TempNa
             End If
        Next P
    Next Pass
    For P = 1 To Size
        picEnter.Print Names(P); Tab(20); FormatNumber(Avg(P), 3)
    Next P
    picEnter.Print
    picEnter.Print Names(1); " was the team leader with an "; FormatNumber(Avg(1), 3); " Batting Average"
End Sub

Private Sub cmdCalcAvg_Click()
    'Ben Tollefson
    'March 22, 2006
    'This Button calculates the Batting Average for each player
    
    picEnter.Cls
    Dim P As Integer
    Dim AvgR As String
    picEnter.Print "Name"; Tab(20); "Batting Average"
    picEnter.Print
    For P = 1 To 10
        Avg(P) = Hits(P) / AB(P)
        Select Case Avg(P)
            Case Is >= 0.35
                AvgR = "All Star"
            Case 0.3 To 0.35
                AvgR = "Good Hitter"
            Case 0.25 To 0.3
                AvgR = "Average Hitter"
            Case Else
                AvgR = "Poor Hitter"
        End Select
        picEnter.Print Names(P); Tab(20); FormatNumber(Avg(P), 3); Tab(30); AvgR
    Next P
End Sub

Private Sub cmdHits_Click()
    'Ben Tollefson
    'March 22, 2006
    'This Button places the players in order by number of hits
    
    
    picEnter.Cls
    Dim Pass As Integer
    Dim TempN As String
    Dim TempHR, TempH, TempAB As Integer
    Dim Pos As Integer
    picEnter.Print "Name"; Tab(20); "Home Runs"; Tab(35); "Hits"; Tab(45); "At-Bats"
    picEnter.Print
    For Pass = 1 To Size - 1
        For Pos = 1 To Size - 1
            If Hits(Pos) < Hits(Pos + 1) Then
                TempHR = HR(Pos)
                HR(Pos) = HR(Pos + 1)
                HR(Pos + 1) = TempHR
                
                TempN = Names(Pos)
                Names(Pos) = Names(Pos + 1)
                Names(Pos + 1) = TempN
                
                TempH = Hits(Pos)
                Hits(Pos) = Hits(Pos + 1)
                Hits(Pos + 1) = TempH
                
                TempAB = AB(Pos)
                AB(Pos) = AB(Pos + 1)
                AB(Pos + 1) = TempAB
            End If
        Next Pos
    Next Pass
    For Pos = 1 To Size
    picEnter.Print Names(Pos); Tab(20); HR(Pos); Tab(35); Hits(Pos); Tab(45); AB(Pos)
    Next Pos
    picEnter.Print
    picEnter.Print Names(1); " was the team leader with "; Hits(1); " Hits"
End Sub

Private Sub cmdHR_Click()
    'Ben Tollefson
    'March 22, 2006
    'This Button places the players in order by number of Home Runs
    
    
    picEnter.Cls
    picEnter.Print "Name"; Tab(20); "Home Runs"; Tab(35); "Hits"; Tab(45); "At-Bats"
    picEnter.Print
    Dim Pass As Integer
    Dim TempN As String
    Dim TempHR, TempH, TempAB As Integer
    Dim Pos As Integer
    For Pass = 1 To Size - 1
        For Pos = 1 To Size - 1
            If HR(Pos) < HR(Pos + 1) Then
                TempHR = HR(Pos)
                HR(Pos) = HR(Pos + 1)
                HR(Pos + 1) = TempHR
                
                TempN = Names(Pos)
                Names(Pos) = Names(Pos + 1)
                Names(Pos + 1) = TempN
                
                TempH = Hits(Pos)
                Hits(Pos) = Hits(Pos + 1)
                Hits(Pos + 1) = TempH
                
                TempAB = AB(Pos)
                AB(Pos) = AB(Pos + 1)
                AB(Pos + 1) = TempAB
            End If
        Next Pos
    Next Pass
    For Pos = 1 To Size
    picEnter.Print Names(Pos); Tab(20); HR(Pos); Tab(35); Hits(Pos); Tab(45); AB(Pos)
    Next Pos
    picEnter.Print
    picEnter.Print Names(1); " was the team leader with "; HR(1); " Home Runs"
End Sub

Private Sub cmdSearch_Click()
    'Ben Tollefson
    'March 22, 2006
    'This Button is used to search for a specific player
    
    picEnter.Cls
    Dim Found As Boolean
    Dim Counter As Integer
    Dim SName As String
    Dim Space As Integer
    Dim LName, FName As String
    SName = InputBox("Enter LastName, First", "Search For Player")
    Space = InStr(SName, " ")
    LName = Left(SName, Space - 2)
    FName = Right(SName, Len(SName) - Space)
    Found = False
    Counter = 0
    Do While Found = False And Counter < Size
        Counter = Counter + 1
            If Names(Counter) = SName Then
                Found = True
            End If
    Loop
    If Found = True Then
        picEnter.Print FName; " "; LName; "   had "; HR(Counter); " Home Runs "; Hits(Counter); " Hits and "; AB(Counter); "At-Bats"
    Else
        MsgBox "Player Not Found, Make Sure Name Is Correct", , "Error!!!"
    End If
End Sub

