VERSION 5.00
Begin VB.Form frmFinals 
   BackColor       =   &H000080FF&
   Caption         =   "Finals !!!!!"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9765
   ForeColor       =   &H000080FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   9765
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFinalFourAppear 
      Caption         =   "CLICK TO HAVE FINAL FOUR APPEAR"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4560
      TabIndex        =   20
      Top             =   3600
      Width           =   1935
   End
   Begin VB.CommandButton cmdCompute 
      Caption         =   " Compute Section Total"
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
      Left            =   2520
      TabIndex        =   18
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton cmdGoToTotals 
      Caption         =   "Go To Totals"
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
      Left            =   4080
      TabIndex        =   11
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton cmdGoBackSouth 
      Caption         =   "Go Back To South"
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
      Left            =   2520
      TabIndex        =   10
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton cmdGoBackMidwest 
      Caption         =   "Go Back To Midwest"
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
      Left            =   2520
      TabIndex        =   9
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton cmdGoBackEast 
      Caption         =   "Go Back To East"
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
      Left            =   4080
      TabIndex        =   8
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton cmdGoBack 
      Caption         =   "Go Back To West"
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
      Left            =   4080
      TabIndex        =   7
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton cmdChampion 
      Height          =   375
      Left            =   5280
      TabIndex        =   6
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton cmdwinner2 
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmdwinner1 
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton cmdSouth 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton cmdEast 
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdWest 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton CmdMidwest 
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TO SEE YOUR SCORES CLICK!! ON SECTION TOTAL TO TRANSFER SCORE TO TOTALS PAGE AND SUBMIT BRACKET"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   19
      Top             =   4680
      Width           =   2175
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "CLICK ON THE BUTTON TO THE RIGHT TO HAVE YOUR TEAMS APPEAR"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   17
      Top             =   3840
      Width           =   4215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "South Champion"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "East Champion"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "West Champion"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Midwest Champion"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Image Image2 
      Height          =   2220
      Left            =   6960
      Picture         =   "frmFinals.frx":0000
      Top             =   120
      Width           =   2250
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "**** FINAL FOUR ****"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   6855
   End
   Begin VB.Image Image1 
      Height          =   2715
      Left            =   5880
      Picture         =   "frmFinals.frx":914D
      Top             =   4320
      Width           =   3450
   End
End
Attribute VB_Name = "frmFinals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'this form is for the user to pick the final four winners

Private Sub cmdCompute_Click()
    Dim Ctr As Integer, CTR2 As Integer, CTR3 As Integer        'set counters to winners
    Dim Final4(1 To 4) As String                                'array for names of fianl four winners
    Dim Final4Pos(1 To 4) As Integer                            'array of final four team rankings incase we want to points for upsets
    Dim Championship(1 To 2) As String                          'other arrays the same except for championship game and winner
    Dim ChampionshipPos(1 To 2) As Integer
    Dim Champs(1) As String
    Dim ChampsPos(1) As Integer
    
    
    Open App.Path & "\Final4.txt" For Input As #1       'notepad with final four winners
    Ctr = 0
    Do Until EOF(1)
        Ctr = Ctr + 1
        Input #1, Final4Pos(Ctr), Final4(Ctr)
    Loop
    Close #1
    Final4Sum = 0                                       'set sum to 0
    If CmdMidwest.Caption = Final4(1) Then              'if statements, when true will add to sum
        Final4Sum = Final4Sum + 8
    End If
    If cmdWest.Caption = Final4(2) Then
        Final4Sum = Final4Sum + 8
    End If
    If cmdEast.Caption = Final4(3) Then
        Final4Sum = Final4Sum + 8
    End If
    If cmdSouth.Caption = Final4(4) Then
        Final4Sum = Final4Sum + 8
    End If
    
    'same as above except used for championship and champion
    Open App.Path & "\Championship.txt" For Input As #2
    CTR2 = 0
    Do Until EOF(2)
        CTR2 = CTR2 + 1
        Input #2, ChampionshipPos(CTR2), Championship(CTR2)
    Loop
    Close #2
    ChampionshipSum = 0
    If cmdwinner1.Caption = Championship(1) Then
        ChampionshipSum = ChampionshipSum + 12
    End If
    If cmdwinner2.Caption = Championship(2) Then
        ChampionshipSum = ChampionshipSum + 12
    End If
    
    Open App.Path & "\Champion.txt" For Input As #3
    CTR3 = 0
    Do Until EOF(3)
        CTR3 = CTR3 + 1
        Input #3, ChampsPos(CTR3), Champs(CTR3)
    Loop
    Close #3
    
    ChampsSum = 0
    If cmdChampion.Caption = Champs(1) Then
        ChampsSum = ChampsSum + 20
    End If
       
End Sub

'will set captions from regional winners to captions of buttons on finals form
Private Sub cmdFinalFourAppear_Click()
    CmdMidwest.Caption = MidwestWinner
    cmdSouth.Caption = SouthWinner
    cmdWest.Caption = WestWinner
    cmdEast.Caption = EastWinner
End Sub

'will allow the user to freely move from finals form back to midwest, east, west, south and move ahead to rankings
Private Sub cmdGoBack_Click()
    frmFinals.Hide
    frmWest.Show
End Sub

Private Sub cmdGoBackEast_Click()
    frmFinals.Hide
    frmEast.Show
End Sub

Private Sub cmdGoBackMidwest_Click()
    frmFinals.Hide
    frmMidwest.Show
End Sub

Private Sub cmdGoBackSouth_Click()
    frmFinals.Hide
    frmSouth.Show
End Sub

Private Sub cmdGoToTotals_Click()
    frmFinals.Hide
    frmTotals.Show
End Sub


'will also transfer caption to button where winner will be displayed
Private Sub CmdMidwest_Click()
   
    cmdwinner1.Caption = CmdMidwest.Caption
End Sub

Private Sub cmdSouth_Click()
    
    cmdwinner2.Caption = cmdSouth.Caption
End Sub

Private Sub cmdWest_Click()
  
    cmdwinner1.Caption = cmdWest.Caption
End Sub
Private Sub cmdEast_Click()

    cmdwinner2.Caption = cmdEast.Caption
End Sub


Private Sub cmdwinner1_Click()
    cmdChampion.Caption = cmdwinner1.Caption
End Sub

Private Sub cmdwinner2_Click()
    cmdChampion.Caption = cmdwinner2.Caption
End Sub

