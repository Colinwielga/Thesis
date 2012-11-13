VERSION 5.00
Begin VB.Form frmIndivResults 
   BackColor       =   &H00FF0000&
   Caption         =   "Individual Results"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmddirect 
      Caption         =   "Back To Directory"
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   6960
      Width           =   2055
   End
   Begin VB.CommandButton cmdteam 
      Caption         =   "Team Results"
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   5640
      Width           =   2055
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search for Runner"
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   4200
      Width           =   2055
   End
   Begin VB.CommandButton cmdsort2 
      Caption         =   "Sort By School"
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton cmdSort1 
      Caption         =   "Sort By Name A - Z"
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   2055
   End
   Begin VB.PictureBox picOutput 
      Height          =   10935
      Left            =   5520
      ScaleHeight     =   10875
      ScaleWidth      =   8355
      TabIndex        =   1
      Top             =   0
      Width           =   8415
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Meet Results "
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FF0000&
      Caption         =   "By: Tyler Trettel and Josh Gunderson"
      Height          =   255
      Left            =   2520
      TabIndex        =   10
      Top             =   9960
      Width           =   2895
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF0000&
      Caption         =   "2008 MIAC Cross Country Project "
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   9
      Top             =   9960
      Width           =   2535
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FF0000&
      Caption         =   "Individual Results"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   10200
      Width           =   2535
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FF0000&
      Caption         =   "November 5, 2008"
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   10200
      Width           =   2895
   End
End
Attribute VB_Name = "frmIndivResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name: MIAC CC Project
'Form Name: frmIndivResults
'Authors: Josh Gunderson & Tyler Trettel
'Date: 5 November 2008
'Objective: The purpose of this form is for the user to see who places with in the top 50 in the meet.  They are also able to sort the data in multiple ways (By team, alphabetical name)
Private Sub cmddirect_Click()
frmIndivResults.Hide
frmdirectory.Show
End Sub

Private Sub cmdLoad_Click()
picOutput.Cls
picOutput.Print ""
picOutput.Print "Place", "Name"; Tab(45); "Team"; Tab(65); "Time", , "Pace"
picOutput.Print "=============================================================================="

Open App.Path & "\Runners.txt" For Input As #1
CTR = 0
Do Until EOF(1)
    CTR = CTR + 1
    Input #1, place(CTR), runner(CTR), School(CTR), Time(CTR), Pace(CTR)
Loop
Close #1

For N = 1 To CTR
    picOutput.Print place(N), runner(N), Tab(45); School(N), Tab(65); Time(N), Pace(N)
Next N




End Sub

Private Sub cmdSearch_Click()

picOutput.Cls

X = InputBox("Please enter a runners name that finished within the top 50.", "Runner Search")

CTR = 0
Found = False

Do Until Found = True Or CTR >= 50
    CTR = CTR + 1
    If X = runner(CTR) Then
        Found = True
    End If
Loop
    
If Found = True Then
    picOutput.Print ""
    picOutput.Print "Place", "Name"; Tab(45); "Team"; Tab(65); "Time", , "Pace"
    picOutput.Print "=============================================================================="
    picOutput.Print place(CTR), runner(CTR), Tab(45); School(CTR), Tab(65); Time(CTR), Pace(CTR)
Else
    MsgBox "There is no race participant with that name."
End If

End Sub

Private Sub cmdSort1_Click()

picOutput.Cls

Dim Pass As Integer, Pos As Integer, Temp As String, Temp2 As Integer, Temp3 As String, Temp4 As String, Temp5 As String

    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
            If runner(Pos) > runner(Pos + 1) Then
                Temp = runner(Pos)
                runner(Pos) = runner(Pos + 1)
                runner(Pos + 1) = Temp
                
                Temp2 = place(Pos)
                place(Pos) = place(Pos + 1)
                place(Pos + 1) = Temp2
                
                Temp3 = School(Pos)
                School(Pos) = School(Pos + 1)
                School(Pos + 1) = Temp3
                
                Temp4 = Time(Pos)
                Time(Pos) = Time(Pos + 1)
                Time(Pos + 1) = Temp4
                
                Temp5 = Pace(Pos)
                Pace(Pos) = Pace(Pos + 1)
                Pace(Pos + 1) = Temp5
            End If
        Next Pos
    Next Pass
picOutput.Print ""
picOutput.Print "Place", "Name"; Tab(45); "Team"; Tab(65); "Time", , "Pace"
picOutput.Print "=============================================================================="

For N = 1 To CTR
    picOutput.Print place(N), runner(N), Tab(45); School(N), Tab(65); Time(N), Pace(N)
   
Next N

End Sub

Private Sub cmdsort2_Click()
    
picOutput.Cls

Dim Pass As Integer, Pos As Integer, Temp As String, Temp2 As Integer, Temp3 As String, Temp4 As String, Temp5 As String

    For Pass = 1 To CTR - 1
        For Pos = 1 To CTR - Pass
            If School(Pos) > School(Pos + 1) Then
                Temp3 = School(Pos)
                School(Pos) = School(Pos + 1)
                School(Pos + 1) = Temp3
                
                Temp = runner(Pos)
                runner(Pos) = runner(Pos + 1)
                runner(Pos + 1) = Temp
                
                Temp2 = place(Pos)
                place(Pos) = place(Pos + 1)
                place(Pos + 1) = Temp2
                
                
                Temp4 = Time(Pos)
                Time(Pos) = Time(Pos + 1)
                Time(Pos + 1) = Temp4
                
                Temp5 = Pace(Pos)
                Pace(Pos) = Pace(Pos + 1)
                Pace(Pos + 1) = Temp5
            End If
        Next Pos
    Next Pass
picOutput.Print ""
picOutput.Print "Place", "Name"; Tab(45); "Team"; Tab(65); "Time", , "Pace"
picOutput.Print "=============================================================================="

For N = 1 To CTR
    picOutput.Print place(N), runner(N), Tab(45); School(N), Tab(65); Time(N), Pace(N)
Next N
Close #1
End Sub

Private Sub cmdTeam_Click()
frmIndivResults.Hide
frmTeamResults.Show

End Sub

