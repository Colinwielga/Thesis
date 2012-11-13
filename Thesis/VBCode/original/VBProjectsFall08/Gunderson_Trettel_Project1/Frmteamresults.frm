VERSION 5.00
Begin VB.Form frmTeamResults 
   BackColor       =   &H00000000&
   Caption         =   "Team Results"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleMode       =   0  'User
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Left            =   0
      Picture         =   "Frmteamresults.frx":0000
      ScaleHeight     =   1035
      ScaleWidth      =   14355
      TabIndex        =   6
      Top             =   9480
      Width           =   14415
   End
   Begin VB.PictureBox Picture10 
      Height          =   1095
      Left            =   14280
      Picture         =   "Frmteamresults.frx":34F86
      ScaleHeight     =   1035
      ScaleWidth      =   915
      TabIndex        =   5
      Top             =   9480
      Width           =   975
   End
   Begin VB.CommandButton cmddirectory 
      Caption         =   "Back To Directory"
      Height          =   1455
      Left            =   10680
      TabIndex        =   4
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton cmdTotalAvgs 
      Caption         =   "Team Total Times And Average Times "
      Height          =   615
      Left            =   5520
      TabIndex        =   3
      Top             =   1680
      Width           =   3615
   End
   Begin VB.CommandButton cmdTeam 
      Caption         =   "Which school would you like to see Stats for?"
      Height          =   615
      Left            =   5520
      TabIndex        =   2
      Top             =   600
      Width           =   3615
   End
   Begin VB.CommandButton cmdteamresults 
      Caption         =   "Load Team Results"
      Height          =   1455
      Left            =   1920
      Picture         =   "Frmteamresults.frx":3CBD0
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin VB.PictureBox picOutput2 
      Height          =   4935
      Left            =   4080
      ScaleHeight     =   4875
      ScaleWidth      =   6435
      TabIndex        =   0
      Top             =   2760
      Width           =   6495
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000012&
      Caption         =   "By: Tyler Trettel and Josh Gunderson"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2520
      TabIndex        =   10
      Top             =   9000
      Width           =   2895
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000012&
      Caption         =   "2008 MIAC Cross Country Project "
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   9
      Top             =   9000
      Width           =   2535
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000012&
      Caption         =   "What is Cross Country"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   9240
      Width           =   2535
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000012&
      Caption         =   "November 5, 2008"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   9240
      Width           =   2895
   End
End
Attribute VB_Name = "frmTeamResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim place(1 To 11) As Integer, points(1 To 11) As Integer, runner(1 To 7) As String, team(1 To 11) As String, T As String
Dim teamplace(1 To 7) As Integer, overallplace(1 To 7) As Integer, Time(1 To 7) As String
'Project Name: MIAC CC Project
'Form Name: frmTeamResults
'Authors: Josh Gunderson & Tyler Trettel
'Date: 5 November 2008
'Objective: The purpose of this form is for the user to view how each school preformed at the meet.  They are able to see the standings by overall points.  Then they are able to view a breakdown of each school.  The can view the top seven runners for each school along with the total time for those runners to complete along with the average time of the top 5 runners




Private Sub cmddirectory_Click()
frmTeamResults.Hide
frmdirectory.Show

End Sub

Private Sub cmdForm3_Click()

frmIndivResults.Hide
frmTeamResults.Hide
frmExplaination.Show

End Sub

Private Sub cmdTeam_Click()

picOutput2.Cls
picOutput2.Print "Team Place"; Tab(22); "Overall Place"; Tab(40); "Runner"; Tab(65); "Time"
picOutput2.Print "=============================================================="

T = InputBox("Enter a school.", "Team Scores")

Select Case LCase(T)
    
    Case Is = "st olaf"
                    Open App.Path & "\stolaf.txt" For Input As #1
                        Counter = 0
                     Do Until EOF(1)
                        Counter = Counter + 1
                         Input #1, teamplace(Counter), overallplace(Counter), runner(Counter), Time(Counter)
                      Loop
                      Close #1

                        For N = 1 To Counter
                            picOutput2.Print Tab(3); teamplace(N); Tab(25); overallplace(N); Tab(40); runner(N); Tab(65); Time(N)
                        Next N
    
    
    Case Is = "hamline"
                    Open App.Path & "\hamline.txt" For Input As #1
                        Counter = 0
                     Do Until EOF(1)
                        Counter = Counter + 1
                         Input #1, teamplace(Counter), overallplace(Counter), runner(Counter), Time(Counter)
                      Loop
                      Close #1

                        For N = 1 To Counter
                            picOutput2.Print Tab(3); teamplace(N); Tab(25); overallplace(N); Tab(40); runner(N); Tab(65); Time(N)
                        Next N
    
    
    Case Is = "st thomas"
                    Open App.Path & "\stthomas.txt" For Input As #1
                        Counter = 0
                     Do Until EOF(1)
                        Counter = Counter + 1
                         Input #1, teamplace(Counter), overallplace(Counter), runner(Counter), Time(Counter)
                      Loop
                      Close #1

                        For N = 1 To Counter
                            picOutput2.Print Tab(3); teamplace(N); Tab(25); overallplace(N); Tab(40); runner(N); Tab(65); Time(N)
                        Next N

    
    Case Is = "bethel"
                    Open App.Path & "\bethel.txt" For Input As #1
                        Counter = 0
                     Do Until EOF(1)
                        Counter = Counter + 1
                         Input #1, teamplace(Counter), overallplace(Counter), runner(Counter), Time(Counter)
                      Loop
                      Close #1

                        For N = 1 To Counter
                            picOutput2.Print Tab(3); teamplace(N); Tab(25); overallplace(N); Tab(40); runner(N); Tab(65); Time(N)
                        Next N
        
        
    Case Is = "st johns"
                    Open App.Path & "\stjohns.txt" For Input As #1
                        Counter = 0
                     Do Until EOF(1)
                        Counter = Counter + 1
                         Input #1, teamplace(Counter), overallplace(Counter), runner(Counter), Time(Counter)
                      Loop
                      Close #1

                        For N = 1 To Counter
                            picOutput2.Print Tab(3); teamplace(N); Tab(25); overallplace(N); Tab(40); runner(N); Tab(65); Time(N)
                        Next N
        
        
    Case Is = "gustavus"
                    Open App.Path & "\gustavus.txt" For Input As #1
                        Counter = 0
                     Do Until EOF(1)
                        Counter = Counter + 1
                         Input #1, teamplace(Counter), overallplace(Counter), runner(Counter), Time(Counter)
                      Loop
                      Close #1

                        For N = 1 To Counter
                            picOutput2.Print Tab(3); teamplace(N); Tab(25); overallplace(N); Tab(40); runner(N); Tab(65); Time(N)
                        Next N
    
    
    Case Is = "carleton"
                Open App.Path & "\carleton.txt" For Input As #1
                        Counter = 0
                     Do Until EOF(1)
                        Counter = Counter + 1
                         Input #1, teamplace(Counter), overallplace(Counter), runner(Counter), Time(Counter)
                      Loop
                      Close #1

                        For N = 1 To Counter
                            picOutput2.Print Tab(3); teamplace(N); Tab(25); overallplace(N); Tab(40); runner(N); Tab(65); Time(N)
                        Next N
     
     Case Is = "macalester"
                Open App.Path & "\macalester.txt" For Input As #1
                        Counter = 0
                     Do Until EOF(1)
                        Counter = Counter + 1
                         Input #1, teamplace(Counter), overallplace(Counter), runner(Counter), Time(Counter)
                      Loop
                      Close #1

                        For N = 1 To Counter
                            picOutput2.Print Tab(3); teamplace(N); Tab(25); overallplace(N); Tab(40); runner(N); Tab(65); Time(N)
                        Next N
         
     
     Case Is = "augsburg"
                     Open App.Path & "\augsburg.txt" For Input As #1
                        Counter = 0
                     Do Until EOF(1)
                        Counter = Counter + 1
                         Input #1, teamplace(Counter), overallplace(Counter), runner(Counter), Time(Counter)
                      Loop
                      Close #1

                        For N = 1 To Counter
                            picOutput2.Print Tab(3); teamplace(N); Tab(25); overallplace(N); Tab(40); runner(N); Tab(65); Time(N)
                        Next N
        
        
      Case Is = "st marys"
                    Open App.Path & "\stmarys.txt" For Input As #1
                        Counter = 0
                     Do Until EOF(1)
                        Counter = Counter + 1
                         Input #1, teamplace(Counter), overallplace(Counter), runner(Counter), Time(Counter)
                      Loop
                      Close #1

                        For N = 1 To Counter
                            picOutput2.Print Tab(3); teamplace(N); Tab(25); overallplace(N); Tab(40); runner(N); Tab(65); Time(N)
                        Next N
        
        
    Case Is = "concordia"
                    Open App.Path & "\concordia.txt" For Input As #1
                        Counter = 0
                     Do Until EOF(1)
                        Counter = Counter + 1
                         Input #1, teamplace(Counter), overallplace(Counter), runner(Counter), Time(Counter)
                      Loop
                      Close #1

                        For N = 1 To Counter
                            picOutput2.Print Tab(3); teamplace(N); Tab(25); overallplace(N); Tab(40); runner(N); Tab(65); Time(N)
                        Next N
        End Select
End Sub


Private Sub cmdteamresults_Click()
   picOutput2.Cls
   picOutput2.Print "Place", , "Points", , "Team"
   picOutput2.Print "========================================================"
   Open App.Path & "\Teams.txt" For Input As #1
Counter = 0
Do Until EOF(1)
    Counter = Counter + 1
    Input #1, place(Counter), points(Counter), team(Counter)
Loop
Close #1

For N = 1 To Counter
    picOutput2.Print place(N), , points(N), , team(N)
Next N
End Sub




Private Sub cmdTotalAvgs_Click()
    Dim team2(1 To 11) As String, totaltime(1 To 11) As String, averagetime(1 To 11) As String, Counter2 As Integer, place2(1 To 11) As Integer, points2(1 To 11) As Integer
    
    picOutput2.Cls
   
    picOutput2.Print Tab(3); "Place", Tab(15); "Points"; Tab(35); "Team"; Tab(53); "Average Time"; Tab(73); "Total Time"
    picOutput2.Print "============================================================================"
                    
    Open App.Path & "\totalsandaverages.txt" For Input As #1
                        
    Counter2 = 0
                     
            Do Until EOF(1)
                Counter2 = Counter2 + 1
                    Input #1, place2(Counter2), points2(Counter2), team2(Counter2), totaltime(Counter2), averagetime(Counter2)
            Loop
            Close #1

            For N = 1 To Counter2
                picOutput2.Print Tab(3); place2(N), Tab(15); points2(N), Tab(35); team2(N); Tab(53); totaltime(N); Tab(73); averagetime(N)
            Next N
End Sub
