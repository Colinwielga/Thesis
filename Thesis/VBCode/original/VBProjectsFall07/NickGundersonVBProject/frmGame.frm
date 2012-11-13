VERSION 5.00
Begin VB.Form frmGame 
   Caption         =   "GameTime"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11355
   BeginProperty Font 
      Name            =   "JazzText"
      Size            =   15.75
      Charset         =   2
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "frmGame.frx":0000
   ScaleHeight     =   8445
   ScaleWidth      =   11355
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBox2 
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   5760
      ScaleHeight     =   6795
      ScaleWidth      =   3795
      TabIndex        =   4
      Top             =   240
      Width           =   3855
   End
   Begin VB.PictureBox picBox1 
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   1800
      ScaleHeight     =   6795
      ScaleWidth      =   3555
      TabIndex        =   3
      Top             =   240
      Width           =   3615
   End
   Begin VB.CommandButton CmdMove 
      Caption         =   "Move on to Choosing Starters for the week!"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1800
      TabIndex        =   2
      Top             =   7080
      Width           =   7935
   End
   Begin VB.CommandButton cmdGo1 
      Caption         =   "Show Team 2"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   9960
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdgo 
      Caption         =   "Show Team 1"
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Pass As Integer, Pass2 As Integer, Pass3 As Integer
Dim TempPos As String, TempPlay As String, TempTeam As String
'This screen will load the two teams the user has selected.
'I used for next statements to print the files after I loaded them

Private Sub cmdGo_Click()

Dim Pos As Integer
For Pos = 1 To CTR
    picBox1.Print Pos1(Pos), Player1(Pos)
    picBox1.Print Tab(15); PTeam1(Pos)
Next Pos
cmdGo1.Enabled = True
cmdGo.Enabled = False

End Sub

Private Sub cmdGo1_Click()

Dim Pos As Integer
For Pos = 1 To CTR
    picBox2.Print Pos2(Pos), Player2(Pos)
    picBox2.Print Tab(15); PTeam2(Pos)
Next Pos
CmdMove.Enabled = True
cmdGo1.Enabled = False

End Sub
'I used a message box here to inform the user how many players you can start at each position
Private Sub CmdMove_Click()
    
    MsgBox ("In this Fantasy Football league you can start 1 QB, 2 RB, 2 WR, 1 RB/WR, 1 K, and 1 DEF")
    frmGame.Visible = False
    frmStarters.Visible = True
    
End Sub


Private Sub Form_Load()
'this section is to load the teams into arrays
'each team here has a specific number attached to it 1-12
'also each file contains team name, Position, Player name, and team
'these file arrays will be used throughout the entire document
'each file load is the same except with a different number to keep track of which team is used
If Team1 = 1 And CTR = 0 Then
    
    Open App.Path & "\Schumacher.txt" For Input As #1
    
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Nteam1(CTR), Pos1(CTR), Player1(CTR), PTeam1(CTR)
    Loop
    Close #1
    
Else
    If Team2 = 1 And CTR1 = 0 Then
        Open App.Path & "\Schumacher.txt" For Input As #1
        Do Until EOF(1)
            CTR1 = CTR1 + 1
            Input #1, NTeam2(CTR1), Pos2(CTR1), Player2(CTR1), PTeam2(CTR1)
        Loop
        Close #1
    End If
End If

If Team1 = 2 And CTR = 0 Then
        Open App.Path & "\Gervias.txt" For Input As #1
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Nteam1(CTR), Pos1(CTR), Player1(CTR), PTeam1(CTR)
    Loop
    Close #1
Else
    If Team2 = 2 And CTR1 = 0 Then
        Open App.Path & "\Gervias.txt" For Input As #1
        Do Until EOF(1)
            CTR1 = CTR1 + 1
            Input #1, NTeam2(CTR1), Pos2(CTR1), Player2(CTR1), PTeam2(CTR1)
        Loop
        Close #1
    End If
End If

If Team1 = 3 And CTR = 0 Then
    Open App.Path & "\Cubby.txt" For Input As #1
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Nteam1(CTR), Pos1(CTR), Player1(CTR), PTeam1(CTR)
    Loop
    Close #1
Else
    If Team2 = 3 And CTR1 = 0 Then
        Open App.Path & "\Cubby.txt" For Input As #1
        Do Until EOF(1)
            CTR1 = CTR1 + 1
            Input #1, NTeam2(CTR1), Pos2(CTR1), Player2(CTR1), PTeam2(CTR1)
        Loop
        Close #1
    End If
End If

If Team1 = 4 And CTR = 0 Then
        Open App.Path & "\Pete.txt" For Input As #1
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Nteam1(CTR), Pos1(CTR), Player1(CTR), PTeam1(CTR)
    Loop
    Close #1
Else
    If Team2 = 4 And CTR1 = 0 Then
        Open App.Path & "\Pete.txt" For Input As #1
        Do Until EOF(1)
            CTR1 = CTR1 + 1
            Input #1, NTeam2(CTR1), Pos2(CTR1), Player2(CTR1), PTeam2(CTR1)
        Loop
        Close #1
    End If
End If

If Team1 = 5 And CTR = 0 Then
        Open App.Path & "\Paul.txt" For Input As #1
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Nteam1(CTR), Pos1(CTR), Player1(CTR), PTeam1(CTR)
    Loop
    Close #1
Else
    If Team2 = 5 And CTR1 = 0 Then
        Open App.Path & "\Paul.txt" For Input As #1
        Do Until EOF(1)
            CTR1 = CTR1 + 1
            Input #1, NTeam2(CTR1), Pos2(CTR1), Player2(CTR1), PTeam2(CTR1)
        Loop
        Close #1
    End If
End If

If Team1 = 6 And CTR = 0 Then
        Open App.Path & "\Gundy.txt" For Input As #1
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Nteam1(CTR), Pos1(CTR), Player1(CTR), PTeam1(CTR)
    Loop
    Close #1
Else
    If Team2 = 6 And CTR1 = 0 Then
        Open App.Path & "\Gundy.txt" For Input As #1
        Do Until EOF(1)
            CTR1 = CTR1 + 1
            Input #1, NTeam2(CTR1), Pos2(CTR1), Player2(CTR1), PTeam2(CTR1)
        Loop
        Close #1
    End If
End If

If Team1 = 7 And CTR = 0 Then
        Open App.Path & "\Andy.txt" For Input As #1
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Nteam1(CTR), Pos1(CTR), Player1(CTR), PTeam1(CTR)
    Loop
    Close #1
Else
    If Team2 = 7 And CTR1 = 0 Then
        Open App.Path & "\Andy.txt" For Input As #1
        Do Until EOF(1)
            CTR1 = CTR1 + 1
            Input #1, NTeam2(CTR1), Pos2(CTR1), Player2(CTR1), PTeam2(CTR1)
        Loop
        Close #1
    End If
End If

If Team1 = 8 And CTR = 0 Then
        Open App.Path & "\Case.txt" For Input As #1
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Nteam1(CTR), Pos1(CTR), Player1(CTR), PTeam1(CTR)
    Loop
    Close #1
Else
    If Team2 = 8 And CTR1 = 0 Then
        Open App.Path & "\Case.txt" For Input As #1
        Do Until EOF(1)
            CTR1 = CTR1 + 1
            Input #1, NTeam2(CTR1), Pos2(CTR1), Player2(CTR1), PTeam2(CTR1)
        Loop
        Close #1
    End If
End If

If Team1 = 9 And CTR = 0 Then
        Open App.Path & "\DK.txt" For Input As #1
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Nteam1(CTR), Pos1(CTR), Player1(CTR), PTeam1(CTR)
    Loop
    Close #1
Else
    If Team2 = 9 And CTR1 = 0 Then
        Open App.Path & "\DK.txt" For Input As #1
        Do Until EOF(1)
            CTR1 = CTR1 + 1
            Input #1, NTeam2(CTR1), Pos2(CTR1), Player2(CTR1), PTeam2(CTR1)
        Loop
        Close #1
    End If
End If

If Team1 = 10 And CTR = 0 Then
        Open App.Path & "\Jason.txt" For Input As #1
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Nteam1(CTR), Pos1(CTR), Player1(CTR), PTeam1(CTR)
    Loop
    Close #1
Else
    If Team2 = 10 And CTR1 = 0 Then
        Open App.Path & "\Jason.txt" For Input As #1
        Do Until EOF(1)
            CTR1 = CTR1 + 1
            Input #1, NTeam2(CTR1), Pos2(CTR1), Player2(CTR1), PTeam2(CTR1)
        Loop
        Close #1
    End If
End If

If Team1 = 11 And CTR = 0 Then
        Open App.Path & "\Inglis.txt" For Input As #1
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Nteam1(CTR), Pos1(CTR), Player1(CTR), PTeam1(CTR)
    Loop
    Close #1
Else
    If Team2 = 11 And CTR1 = 0 Then
        Open App.Path & "\Inglis.txt" For Input As #1
        Do Until EOF(1)
            CTR1 = CTR1 + 1
            Input #1, NTeam2(CTR1), Pos2(CTR1), Player2(CTR1), PTeam2(CTR1)
        Loop
        Close #1
    End If
End If

If Team1 = 12 And CTR = 0 Then
        Open App.Path & "\Buss.txt" For Input As #1
    Do Until EOF(1)
        CTR = CTR + 1
        Input #1, Nteam1(CTR), Pos1(CTR), Player1(CTR), PTeam1(CTR)
    Loop
    Close #1
Else
    If Team2 = 12 And CTR1 = 0 Then
        Open App.Path & "\Buss.txt" For Input As #1
        Do Until EOF(1)
            CTR1 = CTR1 + 1
            Input #1, NTeam2(CTR1), Pos2(CTR1), Player2(CTR1), PTeam2(CTR1)
        Loop
        Close #1
    End If
End If


End Sub
