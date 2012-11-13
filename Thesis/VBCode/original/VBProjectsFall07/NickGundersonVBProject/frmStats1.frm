VERSION 5.00
Begin VB.Form frmStats1 
   Caption         =   "Form1"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10800
   LinkTopic       =   "Form1"
   Picture         =   "frmStats1.frx":0000
   ScaleHeight     =   8295
   ScaleWidth      =   10800
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDTD 
      Height          =   285
      Left            =   6120
      TabIndex        =   94
      Text            =   "0"
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox txtTO 
      Height          =   285
      Left            =   4920
      TabIndex        =   93
      Text            =   "0"
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox txtSacks 
      Height          =   285
      Left            =   3720
      TabIndex        =   92
      Text            =   "0"
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox txtPoints 
      Height          =   285
      Left            =   2880
      TabIndex        =   91
      Text            =   "0"
      Top             =   4680
      Width           =   615
   End
   Begin VB.TextBox txtSafety 
      Height          =   285
      Left            =   8280
      TabIndex        =   90
      Text            =   "0"
      Top             =   4680
      Width           =   615
   End
   Begin VB.PictureBox picK 
      Height          =   255
      Left            =   9480
      ScaleHeight     =   195
      ScaleWidth      =   915
      TabIndex        =   88
      Top             =   5640
      Width           =   975
   End
   Begin VB.TextBox txt40 
      Height          =   285
      Left            =   4080
      TabIndex        =   87
      Text            =   "0"
      Top             =   5640
      Width           =   615
   End
   Begin VB.TextBox txt50 
      Height          =   285
      Left            =   5160
      TabIndex        =   86
      Text            =   "0"
      Top             =   5640
      Width           =   615
   End
   Begin VB.TextBox txtMissedFG 
      Height          =   285
      Left            =   6240
      TabIndex        =   85
      Text            =   "0"
      Top             =   5640
      Width           =   615
   End
   Begin VB.TextBox txtPAT 
      Height          =   285
      Left            =   7320
      TabIndex        =   84
      Text            =   "0"
      Top             =   5640
      Width           =   735
   End
   Begin VB.TextBox txtPATM 
      Height          =   285
      Left            =   8520
      TabIndex        =   83
      Text            =   "0"
      Top             =   5640
      Width           =   735
   End
   Begin VB.TextBox txt30 
      Height          =   285
      Left            =   2760
      TabIndex        =   82
      Text            =   "0"
      Top             =   5640
      Width           =   735
   End
   Begin VB.TextBox txtRTD 
      Height          =   285
      Left            =   7320
      TabIndex        =   75
      Text            =   "0"
      Top             =   4680
      Width           =   735
   End
   Begin VB.PictureBox picDEF 
      Height          =   255
      Left            =   9480
      ScaleHeight     =   195
      ScaleWidth      =   915
      TabIndex        =   74
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox txtTEF 
      Height          =   285
      Left            =   8520
      TabIndex        =   68
      Text            =   "0"
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton cmdTeam2 
      Caption         =   "Choose team 2 Starters"
      Height          =   1095
      Left            =   2400
      TabIndex        =   67
      Top             =   6720
      Width           =   1815
   End
   Begin VB.TextBox txtTEINT 
      Height          =   285
      Left            =   7800
      TabIndex        =   66
      Text            =   "0"
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox txtWRRBF 
      Height          =   285
      Left            =   8520
      TabIndex        =   65
      Text            =   "0"
      Top             =   3240
      Width           =   375
   End
   Begin VB.PictureBox picTE 
      Height          =   255
      Left            =   9480
      ScaleHeight     =   195
      ScaleWidth      =   915
      TabIndex        =   64
      Top             =   3720
      Width           =   975
   End
   Begin VB.PictureBox picWRRB 
      Height          =   255
      Left            =   9480
      ScaleHeight     =   195
      ScaleWidth      =   915
      TabIndex        =   63
      Top             =   3240
      Width           =   975
   End
   Begin VB.PictureBox picWR2 
      Height          =   255
      Left            =   9480
      ScaleHeight     =   195
      ScaleWidth      =   915
      TabIndex        =   62
      Top             =   2760
      Width           =   975
   End
   Begin VB.PictureBox picWR1 
      Height          =   255
      Left            =   9480
      ScaleHeight     =   195
      ScaleWidth      =   915
      TabIndex        =   61
      Top             =   2280
      Width           =   975
   End
   Begin VB.PictureBox picRB2 
      Height          =   255
      Left            =   9480
      ScaleHeight     =   195
      ScaleWidth      =   915
      TabIndex        =   60
      Top             =   1800
      Width           =   975
   End
   Begin VB.PictureBox picRB1 
      Height          =   255
      Left            =   9480
      ScaleHeight     =   195
      ScaleWidth      =   915
      TabIndex        =   59
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox txtTEREY 
      Height          =   285
      Left            =   5520
      TabIndex        =   58
      Text            =   "0"
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox txtTERY 
      Height          =   285
      Left            =   4080
      TabIndex        =   57
      Text            =   "0"
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox txtTEP 
      Height          =   285
      Left            =   2760
      TabIndex        =   56
      Text            =   "0"
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox txtWR2REY 
      Height          =   285
      Left            =   5520
      TabIndex        =   55
      Text            =   "0"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtWR2RY 
      Height          =   285
      Left            =   4080
      TabIndex        =   54
      Text            =   "0"
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtWR2P 
      Height          =   285
      Left            =   2760
      TabIndex        =   53
      Text            =   "0"
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox txtTETD 
      Height          =   285
      Left            =   7080
      TabIndex        =   52
      Text            =   "0"
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox txtWR2TD 
      Height          =   285
      Left            =   7080
      TabIndex        =   51
      Text            =   "0"
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox txtWR2INT 
      Height          =   285
      Left            =   7800
      TabIndex        =   50
      Text            =   "0"
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox txtWR2F 
      Height          =   285
      Left            =   8520
      TabIndex        =   49
      Text            =   "0"
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox txtWRRBINT 
      Height          =   285
      Left            =   7800
      TabIndex        =   48
      Text            =   "0"
      Top             =   3240
      Width           =   375
   End
   Begin VB.TextBox txtWRRBTD 
      Height          =   285
      Left            =   7080
      TabIndex        =   47
      Text            =   "0"
      Top             =   3240
      Width           =   375
   End
   Begin VB.TextBox txtWRRBREY 
      Height          =   285
      Left            =   5520
      TabIndex        =   46
      Text            =   "0"
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txtWRRBRY 
      Height          =   285
      Left            =   4080
      TabIndex        =   45
      Text            =   "0"
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox txtWRRBP 
      Height          =   285
      Left            =   2760
      TabIndex        =   44
      Text            =   "0"
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox txtRB2RY 
      Height          =   285
      Left            =   4080
      TabIndex        =   43
      Text            =   "0"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtRB1TD 
      Height          =   285
      Left            =   7080
      TabIndex        =   42
      Text            =   "0"
      Top             =   1320
      Width           =   375
   End
   Begin VB.PictureBox picTotal 
      Height          =   975
      Left            =   8280
      ScaleHeight     =   915
      ScaleWidth      =   2235
      TabIndex        =   41
      Top             =   6600
      Width           =   2295
   End
   Begin VB.TextBox txtWR1F 
      Height          =   285
      Left            =   8520
      TabIndex        =   32
      Text            =   "0"
      Top             =   2280
      Width           =   375
   End
   Begin VB.TextBox txtWR1INT 
      Height          =   285
      Left            =   7800
      TabIndex        =   31
      Text            =   "0"
      Top             =   2280
      Width           =   375
   End
   Begin VB.TextBox txtWR1TD 
      Height          =   285
      Left            =   7080
      TabIndex        =   30
      Text            =   "0"
      Top             =   2280
      Width           =   375
   End
   Begin VB.TextBox txtWR1REY 
      Height          =   285
      Left            =   5520
      TabIndex        =   29
      Text            =   "0"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtWR1RY 
      Height          =   285
      Left            =   4080
      TabIndex        =   28
      Text            =   "0"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox txtWR1P 
      Height          =   285
      Left            =   2760
      TabIndex        =   27
      Text            =   "0"
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox txtRB2F 
      Height          =   285
      Left            =   8520
      TabIndex        =   26
      Text            =   "0"
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox txtRB2INT 
      Height          =   285
      Left            =   7800
      TabIndex        =   25
      Text            =   "0"
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox txtRB2TD 
      Height          =   285
      Left            =   7080
      TabIndex        =   24
      Text            =   "0"
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox txtRB2REY 
      Height          =   285
      Left            =   5520
      TabIndex        =   23
      Text            =   "0"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtRB2P 
      Height          =   285
      Left            =   2760
      TabIndex        =   22
      Text            =   "0"
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox txtRB1F 
      Height          =   285
      Left            =   8520
      TabIndex        =   21
      Text            =   "0"
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox txtRB1INT 
      Height          =   285
      Left            =   7800
      TabIndex        =   20
      Text            =   "0"
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox txtRB1REY 
      Height          =   285
      Left            =   5520
      TabIndex        =   19
      Text            =   "0"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtRB1RY 
      Height          =   285
      Left            =   4080
      TabIndex        =   18
      Text            =   "0"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txtRB1P 
      Height          =   285
      Left            =   2760
      TabIndex        =   17
      Text            =   "0"
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox txtQBRE 
      Height          =   285
      Left            =   5520
      TabIndex        =   16
      Text            =   "0"
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Find how many Points scored"
      Height          =   1095
      Left            =   4440
      TabIndex        =   13
      Top             =   6720
      Width           =   1815
   End
   Begin VB.TextBox txtQBF 
      Height          =   285
      Left            =   8520
      TabIndex        =   12
      Text            =   "0"
      Top             =   840
      Width           =   375
   End
   Begin VB.PictureBox picQB 
      Height          =   255
      Left            =   9480
      ScaleHeight     =   195
      ScaleWidth      =   915
      TabIndex        =   10
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox txtQBINT 
      Height          =   285
      Left            =   7800
      TabIndex        =   9
      Text            =   "0"
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox txtQBTD 
      Height          =   285
      Left            =   7080
      TabIndex        =   8
      Text            =   "0"
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox txtQBRU 
      Height          =   285
      Left            =   4080
      TabIndex        =   7
      Text            =   "0"
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtQBPY 
      Height          =   285
      Left            =   2760
      TabIndex        =   6
      Text            =   "0"
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblSafety 
      Caption         =   "Safety"
      Height          =   255
      Left            =   8280
      TabIndex        =   89
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label lblPatMissed 
      Caption         =   "PATs Missed"
      Height          =   255
      Left            =   8280
      TabIndex        =   81
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label lblPAT 
      Caption         =   "PAT Made"
      Height          =   255
      Left            =   7320
      TabIndex        =   80
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label lblFGmissed 
      Caption         =   "FGs missed<30"
      Height          =   255
      Left            =   6000
      TabIndex        =   79
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label lblFG50 
      Caption         =   "FGs made 50+"
      Height          =   255
      Left            =   4800
      TabIndex        =   78
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label lblFG40 
      Caption         =   "FGs made 40-49"
      Height          =   255
      Left            =   3480
      TabIndex        =   77
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label lblFG30 
      Caption         =   "FGs made<40"
      Height          =   255
      Left            =   2160
      TabIndex        =   76
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label LBLRTD 
      Caption         =   "Return Td"
      Height          =   255
      Left            =   7200
      TabIndex        =   73
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label lblDTD 
      Caption         =   "Defensive TD"
      Height          =   255
      Left            =   6000
      TabIndex        =   72
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label lblTO 
      Caption         =   "Turnovers"
      Height          =   255
      Left            =   4800
      TabIndex        =   71
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label lblSacks 
      Caption         =   "Sacks"
      Height          =   255
      Left            =   3720
      TabIndex        =   70
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label lblPoints 
      Alignment       =   2  'Center
      Caption         =   "Points Given Up"
      Height          =   255
      Left            =   2160
      TabIndex        =   69
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label lblTotal 
      Caption         =   "Total Points Scored"
      Height          =   255
      Left            =   6480
      TabIndex        =   40
      Top             =   7200
      Width           =   1455
   End
   Begin VB.Label lblDEF 
      Caption         =   "Label2"
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label lblK 
      Caption         =   "Label2"
      Height          =   255
      Left            =   0
      TabIndex        =   38
      Top             =   5640
      Width           =   2415
   End
   Begin VB.Label lblTE 
      Caption         =   "Label2"
      Height          =   255
      Left            =   120
      TabIndex        =   37
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Label lblWRRB 
      Caption         =   "Label2"
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Label lblWR2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   120
      TabIndex        =   35
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Label lblWR1 
      Caption         =   "Label2"
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Label lblRB2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label lblreceiving 
      Caption         =   "Receiving Yards"
      Height          =   255
      Left            =   5400
      TabIndex        =   15
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label lblRB1 
      Caption         =   "Label2"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label lblFumbles 
      Caption         =   "Fumbles"
      Height          =   255
      Left            =   8520
      TabIndex        =   11
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Points scored"
      Height          =   255
      Left            =   9480
      TabIndex        =   5
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblRushing 
      Caption         =   "Rushing Yards"
      Height          =   255
      Left            =   4080
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblINT 
      Caption         =   "INT"
      Height          =   255
      Left            =   7800
      TabIndex        =   3
      Top             =   240
      Width           =   375
   End
   Begin VB.Label lblTD 
      Caption         =   "TD"
      Height          =   255
      Left            =   7080
      TabIndex        =   2
      Top             =   240
      Width           =   375
   End
   Begin VB.Label lbl2 
      Caption         =   "Passing Yards"
      Height          =   255
      Left            =   2760
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblQB 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   2415
   End
End
Attribute VB_Name = "frmStats1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim QBP As Integer, RB1P As Integer, RB2P As Integer, WR1P As Integer, TEP As Integer
Dim WR2P As Integer, WRRBP As Integer, KP As Integer, DEFP As Integer
'this form will have point totals calculated from the textboxes on the following critera
'Passing Yards (15 yards per point)
'Passing Touchdowns(6)
'Interceptions (-1)
'Rushing Yards (10 yards per point)
'Rushing Touchdowns(6)
'Reception Yards (10 yards per points)
'Reception Touchdowns(6)
'Return Yards (30 yards per point)
'Return Touchdowns (10)
'Fumbles Lost(-3)
'Field Goals 0-19 Yards (3)
'Field Goals 20-29 Yards (3)
'Field Goals 30-39 Yards (3)
'Field Goals 40-49 Yards (4)
'Field Goals 50+ Yards (5)
'Field Goals Missed 0-29 Yards (-3)
'Point After Attempt Made (1)
'Point After Attempt Missed (-10)
'Sack (2)
'Interception (2)
'Fumble Recovery(2)
'Touchdown (6)
'Kickoff and Punt Return Touchdowns (6)
'Points Allowed 0 points (20)
'Points Allowed 1-6 points (10)
'Points Allowed 7-13 points (5)
'Points Allowed 14-20 points (3)
'Points Allowed 21-27 points (0)
'Points Allowed 28-34 points (-1)
'Points Allowed 35+ points (-4)

Private Sub cmdGo_Click()
    Dim DEFPoints As Integer
    'I subtracted by.5 to make sure that the computer will always round the orginal answer down
    'example 2.98 rounds usaully up to 3 but 2.98-.5 rounds down to two so the total is two
    
    QBP = txtQBPY / 15 + txtQBRU / 10 + txtQBRE / 10 + txtQBTD * 6 - txtQBINT - txtQBF * 3 - 0.5
    picQB.Print QBP
    RB1P = txtRB1P / 15 + txtRB1RY / 10 + txtRB1REY / 10 + txtRB1TD * 6 - txtRB1INT - txtRB1F * 3 - 0.5
    
    picRB1.Print RB1P
    RB2P = txtRB2P / 15 + txtRB2RY / 10 + txtRB2REY / 10 + txtRB2TD * 6 - txtRB2INT - txtRB2F * 3 - 0.5
    picRB2.Print RB2P
    WR1P = txtWR1P / 15 + txtWR1RY / 10 + txtWR1REY / 10 + txtWR1TD * 6 - txtWR1INT - txtWR1F * 3 - 0.5
    picWR1.Print WR1P
    WR2P = txtWR2P / 15 + txtWR2RY / 10 + txtWR2REY / 10 + txtWR2TD * 6 - txtWR2INT - txtWR2F * 3 - 0.5
    picWR2.Print WR2P
    WRRBP = txtWRRBP / 15 + txtWRRBRY / 10 + txtWRRBREY / 10 + txtWRRBTD * 6 - txtWRRBINT - txtWRRBF * 3 - 0.5
    picWRRB.Print WRRBP
    TEP = txtTEP / 15 + txtTERY / 10 + txtTEREY / 10 + txtTETD * 6 - txtTEINT - txtTEF - 0.5
    picTE.Print TEP
    Dim counts As Integer
    counts = txtPoints
    'a select case was used here in order to decide the point ranges
    Select Case counts
        Case 0
            DEFPoints = 20
        Case 1 To 6
            DEFPoints = 10
        Case 7 To 13
            DEFPoints = 5
        Case 14 To 20
            DEFPoints = 3
        Case 21 To 27
            DEFPoints = 0
        Case 28 To 34
            DEFPoints = -1
        Case Else
            DEFPoints = -3
    End Select
    
    DEFP = DEFPoints + txtTO * 2 + txtSacks * 2 + txtDTD * 6 + txtRTD * 10 + txtSafety * 3
    picDEF.Print DEFP
    KP = txt30 * 3 + txt40 * 4 + txt50 * 5 + txtMissedFG * -3 + txtPAT * 1 + txtPATM * -10
    picK.Print KP
    Team1Points = KP + DEFP + WRRBP + WR2P + WR1P + RB1P + RB2P + QBP
    picTotal.Print Team1Points
End Sub

Private Sub cmdRedo_Click()
    QBP = 0
    RB1P = 0
    RB2P = 0
    WR1P = 0
    WR2P = 0
    WRRBP = 0
    TEP = 0
    DEFP = 0
    KP = 0

End Sub

Private Sub cmdTeam2_Click()
frmStats1.Visible = False
frmTeam2.Visible = True
    QB = 0
    RB1 = 0
    RB2 = 0
    WRRB = 0
    WR1 = 0
    WR2 = 0
    TE = 0
    K = 0
    Def = 0
End Sub

Private Sub Form_Load()
'this will change all the labels to the actual players names
lblQB.Caption = Player1(QB)
lblRB1.Caption = Player1(RB1)
lblRB2.Caption = Player1(RB2)
lblWR1.Caption = Player1(WR1)
lblWR2.Caption = Player1(WR2)
lblWRRB.Caption = Player1(WRRB)
lblK.Caption = Player1(K)
lblDEF.Caption = Player1(Def)
lblTE.Caption = Player1(TE)

End Sub

