VERSION 5.00
Begin VB.Form frmRead 
   BackColor       =   &H00FF0000&
   Caption         =   "Read Stats"
   ClientHeight    =   9105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11010
   LinkTopic       =   "Form1"
   ScaleHeight     =   9105
   ScaleWidth      =   11010
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   6375
      Left            =   2520
      Picture         =   "frmRead.frx":0000
      ScaleHeight     =   6315
      ScaleWidth      =   5595
      TabIndex        =   3
      Top             =   3000
      Width           =   5655
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "Begin the NFL Stats Experience"
      Height          =   1335
      Left            =   11160
      TabIndex        =   2
      Top             =   1680
      Width           =   2175
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   1455
      Index           =   1
      Left            =   11160
      TabIndex        =   1
      Top             =   7920
      Width           =   2175
   End
   Begin VB.Label lblStart 
      BackColor       =   &H00FF0000&
      Caption         =   "Click Here to Search for your NFL stats needs =====================================>"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   4560
      TabIndex        =   0
      Top             =   1800
      Width           =   3615
   End
End
Attribute VB_Name = "frmRead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project-NFL Stats
'Form-frmRead
'Written by Ryan Kempf and Ryan Dell
'3-22-09
'This form read the files and puts them into arrays
'The purpose of this project is to categorize NFL Statistics so they are can be easily accessed.
Private Sub cmdQuit_Click(Index As Integer)
    End
End Sub

Private Sub cmdRead_Click()
    Dim j As Integer
    Dim k As Integer
    Dim l As Integer
    Open App.Path & "\QB.txt" For Input As #1
    ctrQ = 0
    Do Until EOF(1)
        ctrQ = ctrQ + 1
        Input #1, FirstNameQB(ctrQ), LastNameQB(ctrQ), TeamQB(ctrQ), DivisionQB(ctrQ), Comp(ctrQ), AttQB(ctrQ), YardsQB(ctrQ), LongQB(ctrQ), PassTD(ctrQ), INTQB(ctrQ)
    Loop
    For j = 1 To ctrQ
        CompPct(j) = Comp(j) / AttQB(j)
        YdsAtt(j) = YardsQB(j) / AttQB(j)
        QBRating(j) = ((((CompPct(j) * 100) - 30) / 20) + (((PassTD(j) / AttQB(j)) * 100) / 5) + ((9.5 - ((INTQB(j) / AttQB(j)) * 100)) / 4) + (((YardsQB(j) / AttQB(j)) - 3) / 4)) / 0.06
    Next j
    Close #1
    Open App.Path & "\RB.txt" For Input As #2
    ctrR = 0
    Do Until EOF(2)
        ctrR = ctrR + 1
        Input #2, FirstNameRB(ctrR), LastNameRB(ctrR), TeamRB(ctrR), DivisionRB(ctrR), AttRB(ctrR), YardsRB(ctrR), LongRB(ctrR), TDRB(ctrR)
    Loop
    For k = 1 To ctrR
        YPCRB(k) = YardsRB(k) / AttRB(k)
    Next k
    Close #2
    Open App.Path & "\WR.txt" For Input As #3
    ctrW = 0
    Do Until EOF(3)
        ctrW = ctrW + 1
        Input #3, FirstNameWR(ctrW), LastNameWR(ctrW), TeamWR(ctrW), DivisionWR(ctrW), Receptions(ctrW), YardsWR(ctrW), LongWR(ctrW), TDWR(ctrW)
    Loop
    For l = 1 To ctrW
        YPRWR(l) = YardsWR(l) / Receptions(l)
    Next l
    Close #3
    frmRead.Hide
    frmStartup.Show
End Sub
