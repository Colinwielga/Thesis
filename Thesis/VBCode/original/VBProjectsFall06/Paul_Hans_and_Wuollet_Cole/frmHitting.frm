VERSION 5.00
Begin VB.Form frmHitting 
   BackColor       =   &H00FF0000&
   Caption         =   "Hitting Stats"
   ClientHeight    =   7200
   ClientLeft      =   2055
   ClientTop       =   3090
   ClientWidth     =   9570
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   9570
   Visible         =   0   'False
   Begin VB.CommandButton cmdreturns 
      Caption         =   "Return To Main Menu"
      Height          =   615
      Left            =   4800
      TabIndex        =   10
      Top             =   6480
      Width           =   3135
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   6135
      Left            =   3360
      ScaleHeight     =   6075
      ScaleWidth      =   5835
      TabIndex        =   9
      Top             =   120
      Width           =   5895
   End
   Begin VB.CommandButton cmdcs 
      Caption         =   "Sort By Caught Stealing"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   5880
      Width           =   2535
   End
   Begin VB.CommandButton cmdStolen 
      Caption         =   "Sort By Stolen Bases"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   5160
      Width           =   2535
   End
   Begin VB.CommandButton cmdBA 
      Caption         =   "Sort By Batting Average"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   4440
      Width           =   2535
   End
   Begin VB.CommandButton cmdRBIs 
      Caption         =   "Sort By Runs Batted In"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   2535
   End
   Begin VB.CommandButton cmdHR 
      Caption         =   "Sort By Home Runs"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   2535
   End
   Begin VB.CommandButton cmdHits 
      Caption         =   "Sort By Hits"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   2535
   End
   Begin VB.CommandButton cmdAB 
      Caption         =   "Sort By At Bats"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   2535
   End
   Begin VB.CommandButton CmdGames 
      Caption         =   "Sort By Games Played"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   2535
   End
   Begin VB.CommandButton cmdName 
      Caption         =   "Sort By Name"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmHitting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: 1987 World Series
'Form Name: frmHitting
'Authors: Hans Paul and Cole Wuollet
'Date Written: Wednesday November 1, 2006
'Objective: Load and Sort an array from a file, and display different results
Option Explicit

Private Sub cmdAB_Click()
    For Pass = 1 To (Size - 1)
        For Counter = 1 To (Size - Pass)
            If AtBs(Counter) < AtBs(Counter + 1) Then
                Tempn = Names(Counter)
                Names(Counter) = Names(Counter + 1)
                Names(Counter + 1) = Tempn

                Tempa = AtBs(Counter)
                AtBs(Counter) = AtBs(Counter + 1)
                AtBs(Counter + 1) = Tempa
            End If
        Next Counter
    Next Pass
    picResults.Cls
    picResults.Print "Name"; , , "At Bats"
    picResults.Print "***********************************************************"
    For Pos = 1 To Size
        picResults.Print Names(Pos); Tab; AtBs(Pos)
    Next Pos
End Sub

Private Sub cmdBA_Click()
    For Pass = 1 To (Size - 1)
        For Counter = 1 To (Size - Pass)
            If BAs(Counter) < BAs(Counter + 1) Then
                Tempn = Names(Counter)
                Names(Counter) = Names(Counter + 1)
                Names(Counter + 1) = Tempn
                
                Tempb = BAs(Counter)
                BAs(Counter) = BAs(Counter + 1)
                BAs(Counter + 1) = Tempb
            End If
        Next Counter
    Next Pass
    picResults.Cls
    picResults.Print "Name"; , , "Batting Average"
    picResults.Print "***********************************************************"
    For Pos = 1 To Size
        picResults.Print Names(Pos); Tab; BAs(Pos)
    Next Pos
End Sub

Private Sub cmdcs_Click()
    For Pass = 1 To (Size - 1)
        For Counter = 1 To (Size - Pass)
            If CSs(Counter) < CSs(Counter + 1) Then
                Tempn = Names(Counter)
                Names(Counter) = Names(Counter + 1)
                Names(Counter + 1) = Tempn

                Tempc = CSs(Counter)
                CSs(Counter) = CSs(Counter + 1)
                CSs(Counter + 1) = Tempc
            End If
        Next Counter
    Next Pass
    picResults.Cls
    picResults.Print "Name"; , , "Caught Stealing"
    picResults.Print "***********************************************************"
    For Pos = 1 To Size
        picResults.Print Names(Pos); Tab; CSs(Pos)
    Next Pos
End Sub

Private Sub CmdGames_Click()
    For Pass = 1 To (Size - 1)
        For Counter = 1 To (Size - Pass)
            If Games(Counter) < Games(Counter + 1) Then
                Tempn = Names(Counter)
                Names(Counter) = Names(Counter + 1)
                Names(Counter + 1) = Tempn
    
                Tempg = Games(Counter)
                Games(Counter) = Games(Counter + 1)
                Games(Counter + 1) = Tempg
            End If
        Next Counter
    Next Pass
    picResults.Cls
    picResults.Print "Name"; , , "Games Played"
    picResults.Print "***********************************************************"
    For Pos = 1 To Size
        picResults.Print Names(Pos); Tab; Games(Pos)
    Next Pos

End Sub

Private Sub cmdHits_Click()
    For Pass = 1 To (Size - 1)
        For Counter = 1 To (Size - Pass)
            If Hits(Counter) < Hits(Counter + 1) Then
                Tempn = Names(Counter)
                Names(Counter) = Names(Counter + 1)
                Names(Counter + 1) = Tempn
    
                Temph = Hits(Counter)
                Hits(Counter) = Hits(Counter + 1)
                Hits(Counter + 1) = Temph
            End If
        Next Counter
    Next Pass
    picResults.Cls
    picResults.Print "Name"; , , "Hits"
    picResults.Print "***********************************************************"
    For Pos = 1 To Size
        picResults.Print Names(Pos); Tab; Hits(Pos)
    Next Pos
End Sub

Private Sub cmdHR_Click()
For Pass = 1 To (Size - 1)
        For Counter = 1 To (Size - Pass)
            If HRs(Counter) < HRs(Counter + 1) Then
                Tempn = Names(Counter)
                Names(Counter) = Names(Counter + 1)
                Names(Counter + 1) = Tempn
    
                Tempr = HRs(Counter)
                HRs(Counter) = HRs(Counter + 1)
                HRs(Counter + 1) = Tempr
            End If
        Next Counter
    Next Pass
    picResults.Cls
    picResults.Print "Name"; , , "Homeruns"
    picResults.Print "***********************************************************"
    For Pos = 1 To Size
        picResults.Print Names(Pos); Tab; HRs(Pos)
    Next Pos
End Sub


Private Sub cmdName_Click()
    For Pass = 1 To (Size - 1)
        For Counter = 1 To (Size - Pass)
            If Names(Counter) > Names(Counter + 1) Then
                Tempn = Names(Counter)
                Names(Counter) = Names(Counter + 1)
                Names(Counter + 1) = Tempn
    
                Tempg = Games(Counter)
                Games(Counter) = Games(Counter + 1)
                Games(Counter + 1) = Tempg
            End If
        Next Counter
    Next Pass
    picResults.Cls
    picResults.Print "Name"; , , "Games Played"
    picResults.Print "***********************************************************"
    For Pos = 1 To Size
        picResults.Print Names(Pos); Tab; Games(Pos)
    Next Pos
End Sub

Private Sub cmdRBIs_Click()
    For Pass = 1 To (Size - 1)
        For Counter = 1 To (Size - Pass)
            If RBIs(Counter) < RBIs(Counter + 1) Then
                Tempn = Names(Counter)
                Names(Counter) = Names(Counter + 1)
                Names(Counter + 1) = Tempn
    
                Tempr = RBIs(Counter)
                RBIs(Counter) = RBIs(Counter + 1)
                RBIs(Counter + 1) = Tempr
            End If
        Next Counter
    Next Pass
    picResults.Cls
    picResults.Print "Name"; , , "Runs Batted In"
    picResults.Print "***********************************************************"
    For Pos = 1 To Size
        picResults.Print Names(Pos); Tab; RBIs(Pos)
    Next Pos
End Sub

Private Sub cmdreturns_Click()
    frmHitting.Hide
    frmTwins.Show
End Sub

Private Sub cmdStolen_Click()
    For Pass = 1 To (Size - 1)
        For Counter = 1 To (Size - Pass)
            If SBs(Counter) < SBs(Counter + 1) Then
                Tempn = Names(Counter)
                Names(Counter) = Names(Counter + 1)
                Names(Counter + 1) = Tempn
    
                Temps = SBs(Counter)
                SBs(Counter) = SBs(Counter + 1)
                SBs(Counter + 1) = Temps
            End If
        Next Counter
    Next Pass
    picResults.Cls
    picResults.Print "Name"; , , "Bases Stolen"
    picResults.Print "***********************************************************"
    For Pos = 1 To Size
        picResults.Print Names(Pos); Tab; SBs(Pos)
    Next Pos
End Sub

Private Sub Form_Load()
    Open App.Path & "\Batters.txt" For Input As #1
    Counter = 0
    Do Until EOF(1)
        Counter = Counter + 1
        Input #1, Names(Counter), Games(Counter), AtBs(Counter), Hits(Counter), HRs(Counter), RBIs(Counter), BAs(Counter), SBs(Counter), CSs(Counter)
    Loop
    Close #1
    Size = Counter
End Sub
