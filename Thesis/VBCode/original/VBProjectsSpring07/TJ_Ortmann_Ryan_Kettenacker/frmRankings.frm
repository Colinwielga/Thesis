VERSION 5.00
Begin VB.Form frmRankings 
   BackColor       =   &H00FF0000&
   Caption         =   "Rankings"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9960
   BeginProperty Font 
      Name            =   "Papyrus"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleWidth      =   9960
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   6600
      Picture         =   "frmRankings.frx":0000
      ScaleHeight     =   2355
      ScaleWidth      =   2355
      TabIndex        =   7
      Top             =   3120
      Width           =   2415
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   6240
      Picture         =   "frmRankings.frx":914D
      ScaleHeight     =   3075
      ScaleWidth      =   3075
      TabIndex        =   6
      Top             =   5640
      Width           =   3135
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   6000
      Picture         =   "frmRankings.frx":16D88
      ScaleHeight     =   2835
      ScaleWidth      =   3435
      TabIndex        =   5
      Top             =   120
      Width           =   3495
   End
   Begin VB.CommandButton cmdRankings 
      Caption         =   "CLICK HERE TO GET RANKINGS"
      Height          =   855
      Left            =   1200
      TabIndex        =   4
      Top             =   6480
      Width           =   3375
   End
   Begin VB.CommandButton cmdGoToTotals 
      Caption         =   "GO TO TOTALS PAGE"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   3
      Top             =   7440
      Width           =   2295
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "EXIT!"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   2
      Top             =   7440
      Width           =   2295
   End
   Begin VB.PictureBox picRankings 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   360
      ScaleHeight     =   4635
      ScaleWidth      =   5475
      TabIndex        =   1
      Top             =   1680
      Width           =   5535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "WHERE DO YOU STAND?"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   5535
   End
End
Attribute VB_Name = "frmRankings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' this form allows the user to see where his score compares to other scores
Private Sub cmdExit_Click()
    MsgBox "Thanks for participating in the MADNESS!", , "GOODBYE!"     'a message box to tell user thank you after ending program
    End                                                                 'end program
End Sub

Private Sub cmdGoToTotals_Click()
    frmRankings.Hide        'go from rankings page to totals page
    frmTotals.Show
End Sub

'this button will print rankings and where user stands compared to other scores
'this is done by using a notepad with the scores of other users and bubble sort by score and by comparing user score

Private Sub cmdRankings_Click()
    Dim Pass As Integer                                     'count number of times there is a pass
    Dim Pos As Integer                                      'count position
    Dim TempTotal As Single                                 'temporary total for bubble sort
    Dim TempName As String                                  'temorary name for bubble sort
    Dim Ctr As Integer                                      'ctr for arrray
    Dim RankingsNames(1 To 100) As String                   'names as a string array
    Dim RankingsTotal(1 To 100) As Integer                  'totals as a integer array
    
    
    picRankings.Cls                                         'clear picRankings before printing
    
    Open App.Path & "\LeaderBoard.txt" For Input As #1      'open notepade with scores
    
    Do Until EOF(1)                                         'loop to make arrays for names and scores from file
        Ctr = Ctr + 1
        Input #1, RankingsNames(Ctr), RankingsTotal(Ctr)
    Loop
    Close #1
                    
                    'this bubble sort is to take the user score and add it to the descending order so that the user is able to see where there score stands
    For Pos = 1 To Ctr
      For Pass = 1 To (Ctr - Pass)
        If RankingsTotal(Pos) <= OverallScore Then
                TempTotal = RankingsTotal(Pos)
                RankingsTotal(Pos) = OverallScore
                OverallScore = TempTotal
                
                TempName = RankingsNames(Pos)
                RankingsNames(Pos) = User
                User = TempName
        End If
      Next Pass
    Next Pos
      
    
                                 'this second bubble sort is taking all names and scores from file and sorting them from largest to smalles score (descending order)
    For Pass = 1 To (Ctr - 1)
        For Pos = 1 To (Ctr - Pass)
            If RankingsTotal(Pos) <= RankingsTotal(Pos + 1) Then
                TempTotal = RankingsTotal(Pos)
                RankingsTotal(Pos) = RankingsTotal(Pos + 1)
                RankingsTotal(Pos + 1) = TempTotal
                
                TempName = RankingsNames(Pos)
                RankingsNames(Pos) = RankingsNames(Pos + 1)
                RankingsNames(Pos + 1) = TempName
            End If
        Next Pos
    Next Pass
                                        

      picRankings.Print "Name"; Tab(30), "Score"; Tab(60)                           'print header for picRankings
    For Pos = 1 To Ctr
      picRankings.Print RankingsNames(Pos); Tab(30), RankingsTotal(Pos); Tab(60)    'will print all names and scores with scores being in descending order
    Next Pos
       
        
End Sub

