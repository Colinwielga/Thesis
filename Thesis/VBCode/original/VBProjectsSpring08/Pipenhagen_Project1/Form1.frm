VERSION 5.00
Begin VB.Form frmpastscores 
   Caption         =   "Form1"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   7470
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "Back"
      Height          =   615
      Left            =   2040
      TabIndex        =   2
      Top             =   4200
      Width           =   3375
   End
   Begin VB.CommandButton cmdloadscores 
      Caption         =   "Load Scores"
      Height          =   615
      Left            =   2040
      TabIndex        =   1
      Top             =   3360
      Width           =   3375
   End
   Begin VB.PictureBox picresults 
      Height          =   2415
      Left            =   0
      ScaleHeight     =   2355
      ScaleWidth      =   7395
      TabIndex        =   0
      Top             =   600
      Width           =   7455
   End
End
Attribute VB_Name = "frmpastscores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: A Review of Theoretical Orientations in Clinical Psychology
'Form name: frmpastscores
'Author: Calvin Pipenhagen
'Date Written: March 27, 2008
'Objective: To let the user view their past scores

Option Explicit

Private Sub cmdloadscores_Click() ' loads an array of past scores
Dim ctr As Integer
Dim found As Boolean
Dim n As Integer
found = False
Open App.Path & "\scores.txt" For Input As #1 'loading former scores
    Do Until EOF(1)
        ctr = ctr + 1
    Input #1, namesarray(ctr), humanisticarray(ctr), cognitivearray(ctr), psychodynamicarray(ctr)
    Loop
Close #1

Open App.Path & "\scores.txt" For Output As #1 'adding the new scores or individual, assuming individual has taken all three quizes
    For n = 1 To ctr
        Write #1, namesarray(n), humanisticarray(n), cognitivearray(n), psychodynamicarray(n)
    Next n
        Write #1, names, humanisticquizsum, cognitivequizsum, psychodynamicquizsum
Close #1
Open App.Path & "\scores.txt" For Input As #1
   Do Until EOF(1)
    ctr = ctr + 1
    Input #1, namesarray(ctr), humanisticarray(ctr), cognitivearray(ctr), psychodynamicarray(ctr)
   Loop
Close #1
ctr = 0
Do While ctr < 100 And found = False
ctr = ctr + 1
If names = namesarray(ctr) Then ' determenting if the user is in the data file
            picresults.Cls
            picresults.Print "Name"; Tab(20); "Humanistic Score"; Tab(40); "Cognitive Score"; Tab(60); "Psychodynamic Score"
            picresults.Print "******************************************************************************************"
            picresults.Print namesarray(ctr), Tab(20); humanisticarray(ctr); Tab(40); cognitivearray(ctr); Tab(60); psychodynamicarray(ctr)
            found = True
End If
Loop
If found = False Then 'a message in case the user hasn't taken the quizes
    MsgBox "We have no record of you taking all of the quizes", , "error"
End If
End Sub

Private Sub cmdback_Click() ' back to the main menu
frmpastscores.Hide
frmselectschool.Show
End Sub


