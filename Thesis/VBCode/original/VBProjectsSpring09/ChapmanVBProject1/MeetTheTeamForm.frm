VERSION 5.00
Begin VB.Form MeetTheTeam 
   Caption         =   "Form1"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Go Back"
      Height          =   1095
      Left            =   5040
      TabIndex        =   4
      Top             =   4560
      Width           =   2655
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Sort by name"
      Height          =   1215
      Left            =   5040
      TabIndex        =   3
      Top             =   3120
      Width           =   2535
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search for someone"
      Height          =   1215
      Left            =   5040
      TabIndex        =   2
      Top             =   1680
      Width           =   2535
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "Read Arrays"
      Height          =   1095
      Left            =   5040
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin VB.PictureBox picResults 
      Height          =   6495
      Left            =   480
      ScaleHeight     =   6435
      ScaleWidth      =   4155
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "MeetTheTeam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Names(1 To 15) As String, Goals(1 To 15) As Integer
Dim Ejections(1 To 15) As Integer, Ctr As Integer

Private Sub cmdRead_Click()
Dim TotalGoals As Integer, TotalEjections As Integer, N As Integer

Ctr = 0
N = 0

Open App.Path & "\TheTeam.txt" For Input As #1

picResults.Print Tab(10); "Name", Tab(30); "Goals", Tab(40); "Ejections"

Do While Not EOF(1)
    Ctr = Ctr + 1
    Input #1, Names(Ctr), Goals(Ctr), Ejections(Ctr)
    picResults.Print Names(Ctr), Goals(Ctr), Ejections(Ctr)
Loop

Do While N <= Ctr
    N = N + 1
    TotalGoals = TotalGoals + Goals(N)
    TotalEjections = TotalEjections + Ejections(N)
Loop

picResults.Print "********************************************************************"
picResults.Print "Total", Tab(30); TotalGoals, Tab(40); TotalEjections

Close #1

End Sub

Private Sub cmdSearch_Click()
Dim Search As String

Search = InputBox("Enter a name to find out year", "Search")

If Search = Names Then
    MsgBox "Sweet, " & Search & " is on the team", "Woot!"
    
    Else
        MsgBox "Sorry, but " & Search & " isn't on the team", , "Alert"
End If

End Sub

Private Sub cmdSort_Click()
Dim Pass As Integer, Pos As Integer, TempNames As String, i As Integer
Dim TempGoals As Integer, TempEjections As Integer

picResults.Cls
picResults.Print "Name", Tab(30); "Goals"; Tab(40); "Ejections"
picResults.Print "********************************************************"
For Pass = 1 To Ctr - 1
    For Pos = 1 To Ctr - Pass
        If Names(Pos) > Names(Pos + 1) Then
            TempNames = Names(Pos)
            Names(Pos) = Names(Pos + 1)
            Names(Pos + 1) = TempNames
            
            TempGoals = Goals(Pos)
            Goals(Pos) = Goals(Pos + 1)
            Goals(Pos + 1) = TempGoals
            
            TempEjections = Ejections(Pos)
            Ejections(Pos) = Ejections(Pos + 1)
            Ejections(Pos + 1) = TempEjections
        End If
    Next Pos
Next Pass

For i = 1 To Ctr
    picResults.Print Names(i), Tab(30); Goals(i); Tab(40); Ejections(i)
Next i
End Sub

Private Sub cmdReturn_Click()
MeetTheTeamForm.Hide
TeamInfoForm.Show
End Sub
