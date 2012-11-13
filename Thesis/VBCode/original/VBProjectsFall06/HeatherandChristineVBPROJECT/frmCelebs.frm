VERSION 5.00
Begin VB.Form frmCelebs 
   BackColor       =   &H00FFFF80&
   Caption         =   "Form1"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FFFF00&
      Caption         =   "Return to Game"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   720
      Width           =   2655
   End
   Begin VB.CommandButton cmdBackto 
      Caption         =   "Back to Game Page"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5280
      TabIndex        =   10
      Top             =   11880
      Width           =   5655
   End
   Begin VB.PictureBox Picture6 
      Height          =   1935
      Left            =   12480
      Picture         =   "frmCelebs.frx":0000
      ScaleHeight     =   1875
      ScaleWidth      =   1395
      TabIndex        =   9
      Top             =   8040
      Width           =   1455
   End
   Begin VB.PictureBox Picture5 
      Height          =   1935
      Left            =   10200
      Picture         =   "frmCelebs.frx":0E08
      ScaleHeight     =   1875
      ScaleWidth      =   1515
      TabIndex        =   8
      Top             =   8040
      Width           =   1575
   End
   Begin VB.PictureBox Picture4 
      Height          =   1935
      Left            =   7920
      Picture         =   "frmCelebs.frx":1BB8
      ScaleHeight     =   1875
      ScaleWidth      =   1635
      TabIndex        =   7
      Top             =   8040
      Width           =   1695
   End
   Begin VB.PictureBox Picture3 
      Height          =   1935
      Left            =   12000
      Picture         =   "frmCelebs.frx":296D
      ScaleHeight     =   1875
      ScaleWidth      =   1755
      TabIndex        =   6
      Top             =   2400
      Width           =   1815
   End
   Begin VB.PictureBox Picture2 
      Height          =   1935
      Left            =   9960
      Picture         =   "frmCelebs.frx":3CA0
      ScaleHeight     =   1875
      ScaleWidth      =   1515
      TabIndex        =   5
      Top             =   2400
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   1935
      Left            =   7920
      Picture         =   "frmCelebs.frx":5074
      ScaleHeight     =   1875
      ScaleWidth      =   1755
      TabIndex        =   4
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton cmdCelebDates 
      BackColor       =   &H80000009&
      Caption         =   "Find Out When Celebrities Were on Jeopardy!"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4440
      Width           =   6255
   End
   Begin VB.Frame frmCelebTitle 
      BackColor       =   &H00FFFF80&
      Caption         =   "CELEBRITIES:"
      BeginProperty Font 
         Name            =   "Gloucester MT Extra Condensed"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   1575
      Left            =   3480
      TabIndex        =   2
      Top             =   480
      Width           =   6855
   End
   Begin VB.CommandButton cmdDisplayCelebs 
      BackColor       =   &H00FFFF00&
      Caption         =   "Show Celebrities Top Scores:"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1080
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2400
      Width           =   5775
   End
   Begin VB.PictureBox picCelebNames 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   840
      ScaleHeight     =   5955
      ScaleWidth      =   6195
      TabIndex        =   0
      Top             =   3960
      Width           =   6255
   End
End
Attribute VB_Name = "frmCelebs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Celebrity Form: On this form, the user has the ability to view
'                which celebrities have been on the show.This page
'                will display the the celebrities in order from the
'                lowest score to the highest score.From this page
'                the user will then be ableto type in a name to see
'                what date thecelebrity of their choice was on the show.
'                Following this page, theuser will go back to the first
'                for and then be able to begin the game.

Option Explicit
Dim Counter As Integer
Dim Pos As Integer
Dim InputCeleb As String
Dim Found As Boolean
Dim pass As Integer
Dim Comp As Integer
Dim TempName As String
Dim TempScore As Integer

Private Sub cmdBackto_Click()
    frmCelebs.Hide
    frmGamePage.Show
End Sub

Private Sub cmdCelebDates_Click()
'This button allows the player to type in the celebrity of their choice from the
'   list on the left to find out when their favorite celebrity was on the show!

    Counter = 0
    Open App.Path & "\CelebDates.txt" For Input As #1
    Do Until EOF(1)
        Input #1, CName, CDates
        Counter = Counter + 1
        Celebrities(Counter) = CName
        CelebDates(Counter) = CDates
    Loop
    Close #1
    InputCeleb = InputBox("Enter celebrity to find out when they were on Jeopardy!  Remember to spell the actor the exact way it is displayed to the right")
    Found = False
    Counter = 0
    Do While Found = False And Counter < 15
        Counter = Counter + 1
        If Celebrities(Counter) = InputCeleb Then
            Found = True
        End If
    Loop
    If Found = True Then
        MsgBox InputCeleb & " " & "Played on" & " " & CDates
    Else
        MsgBox "No Match Found"
    End If
    
    
End Sub

Private Sub cmdDisplayCelebs_Click()
'This button will print out the celebrities that were on the show. The celebrities
'   will print out from the celebrity with the lowest score first and move to the
'   highest score.

    picCelebNames.Cls
    Counter = 0
    Open App.Path & "\CelebScores.txt" For Input As #1
    Do Until EOF(1)
        Input #1, CName, CScores
        Counter = Counter + 1
        Celebrities(Counter) = CName
        Scores(Counter) = CScores
    Loop
    Close #1
    For pass = 1 To Counter - 1
        For Comp = 1 To Counter - pass
            If Scores(Comp) > Scores(Comp + 1) Then
                TempScore = Scores(Comp)
                Scores(Comp) = Scores(Comp + 1)
                Scores(Comp + 1) = TempScore
                TempName = Celebrities(Comp)
                Celebrities(Comp) = Celebrities(Comp + 1)
                Celebrities(Comp + 1) = TempName
                
            End If
        Next Comp
    Next pass
    picCelebNames.Print "Celebrities", "       ", "Score"
    picCelebNames.Print "___________________________________________________"
    For Pos = 1 To Counter
        picCelebNames.Print Celebrities(Pos), Scores(Pos)
    Next Pos
End Sub

Private Sub cmdReturn_Click()
' The Return to Game button will simply bring the player back to the first form so
'   the player may begin playing the game when he or she is ready.

    frmCelebs.Hide
    frmGamePage.Show
End Sub

