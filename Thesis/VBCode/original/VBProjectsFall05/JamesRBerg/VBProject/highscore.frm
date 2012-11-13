VERSION 5.00
Begin VB.Form frmHighscore 
   Caption         =   "Hall of Fame"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10755
   LinkTopic       =   "Form1"
   Picture         =   "highscore.frx":0000
   ScaleHeight     =   7230
   ScaleWidth      =   10755
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4800
      TabIndex        =   2
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "Display the High Scores"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4080
      TabIndex        =   1
      Top             =   4200
      Width           =   2775
   End
   Begin VB.PictureBox picresults 
      Height          =   3495
      Left            =   3120
      ScaleHeight     =   3435
      ScaleWidth      =   4515
      TabIndex        =   0
      Top             =   360
      Width           =   4575
   End
End
Attribute VB_Name = "frmhighscore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Simpsons TV Show test (final.vbp)
'main form (highscore.frm)
'Jim Berg
'October 30, 2005
'this form will display all of the scores from the quiz sorted from highest to lowest

Option Explicit

Dim CTR As Integer, player(1 To 10) As String, score(1 To 10) As Integer, N As Integer, tempx As String, tempy As String, Pass As Integer
Dim k As Integer

Private Sub cmdreturn_Click()
frmhighscore.Hide
frmmain.Show
frmCharacters.Hide

End Sub



Private Sub cmdView_Click()
    picresults.Cls
    picresults.Print "Name"; Tab(20); "Score"
    CTR = 0
    'opening file and reading information
    Open App.Path & "\highscore2.txt" For Input As #1
    CTR = 0
    Do While Not EOF(1)
        CTR = CTR + 1
        Input #1, player(CTR), score(CTR)
    Loop
        'bubble sort of the scores
    For Pass = 1 To CTR - 1
        For N = 1 To CTR - Pass
            If score(N) < score(N + 1) Then
                tempx = score(N)
                tempy = player(N)
                score(N) = score(N + 1)
                player(N) = player(N + 1)
                score(N + 1) = tempx
                player(N + 1) = tempy
                End If
        Next N
    Next Pass
    'printing the results
    For k = 1 To CTR
        picresults.Print player(k); Tab(20); score(k)
    Next k
    Close #1
End Sub
