VERSION 5.00
Begin VB.Form avgform 
   BackColor       =   &H00FF0000&
   Caption         =   "Form2"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13215
   LinkTopic       =   "Form2"
   Picture         =   "avgform.frx":0000
   ScaleHeight     =   7470
   ScaleWidth      =   13215
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtpro 
      Height          =   855
      Left            =   3240
      TabIndex        =   5
      Top             =   1440
      Width           =   2415
   End
   Begin VB.CommandButton cmdpro 
      Caption         =   "Find professional bowler's average"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6360
      TabIndex        =   4
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Back to the start"
      Height          =   615
      Left            =   8400
      TabIndex        =   3
      Top             =   6600
      Width           =   1935
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   10920
      TabIndex        =   2
      Top             =   6600
      Width           =   1575
   End
   Begin VB.PictureBox picresults 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   1680
      ScaleHeight     =   3555
      ScaleWidth      =   7395
      TabIndex        =   1
      Top             =   2640
      Width           =   7455
   End
   Begin VB.CommandButton cmdcompute 
      Caption         =   "Click here to enter scores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4560
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "avgform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'bowling prodject
'avgform
'Zach Neumann
'3/30/2008
'this form allows the user to enter in the scores of all the games they have bowled and finds their average
'this form also allows you to search the current professional bowlers and see their average
Option Explicit
Dim avg As Single

Private Sub cmdback_Click()
'back to the hame page
avgform.Hide
teamform.Hide
startform.Show

End Sub

Private Sub cmdcompute_Click()
Dim score As Integer, ctr As Integer, total As Integer
ctr = 0

picresults.Cls
Do While score <> -99
    'loop for user to enter their scores, must be above 0 and below 300
    score = InputBox("Please enter score number " & (ctr + 1) & " or -99 when done", "Scores")
    If score <> -99 Then
        ctr = ctr + 1
        total = score + total
        If score < 0 Or score > 300 Then
            MsgBox ("Impossible!")
            ctr = ctr - 1
            total = total - score
        End If
    End If
Loop

avg = total / ctr

picresults.Print "Your average for"; ctr; " games is"; avg
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdpro_Click()
Dim names(1 To 100) As String, avg(1 To 100) As Single, found As Boolean, ctr As Integer, pro As String, ctrtwo As Integer
pro = LCase(txtpro.Text)
ctr = 0
'opens a document with all of the professional bowlers and their averages
Open App.Path & "\probowlers.txt" For Input As #1
Do Until EOF(1)
    ctr = ctr + 1
    Input #1, names(ctr), avg(ctr)
Loop
Close #1

ctrtwo = 1
'Searches to find the bowler with the name entered
Do While (found = False) And ctrtwo < 100
    If names(ctrtwo) = pro Then
        found = True
    Else
        ctrtwo = ctrtwo + 1
    End If
Loop

If found = True Then
    picresults.Print StrConv(names(ctrtwo), vbProperCase); " has an average of: "; avg(ctrtwo)
Else
    MsgBox "There are no professional bowlers with that name", , "Sorry!"
End If

End Sub

Private Sub cmdquit_Click()
End

End Sub

