VERSION 5.00
Begin VB.Form frmHome 
   BackColor       =   &H00000040&
   Caption         =   "Home"
   ClientHeight    =   12270
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13980
   LinkTopic       =   "Form1"
   ScaleHeight     =   12270
   ScaleWidth      =   13980
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClearTotalPoints 
      Caption         =   "Clear Total Points"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1320
      TabIndex        =   7
      Top             =   8040
      Width           =   2895
   End
   Begin VB.CommandButton cmdTotalPoints 
      Caption         =   "Total Points "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5160
      TabIndex        =   6
      Top             =   6960
      Width           =   2895
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000C0&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   9600
      TabIndex        =   5
      Top             =   8040
      Width           =   2415
   End
   Begin VB.CommandButton cmdHvy 
      BackColor       =   &H000000FF&
      Caption         =   "Heavy Support"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   8400
      TabIndex        =   4
      Top             =   5160
      Width           =   3735
   End
   Begin VB.CommandButton cmdF 
      Caption         =   "Fast Attack"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   960
      TabIndex        =   3
      Top             =   5160
      Width           =   3975
   End
   Begin VB.CommandButton cmdT 
      BackColor       =   &H000000FF&
      Caption         =   "Troops"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   8400
      MaskColor       =   &H000000FF&
      TabIndex        =   2
      Top             =   3000
      Width           =   3735
   End
   Begin VB.CommandButton cmdE 
      Caption         =   "Elites "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   960
      TabIndex        =   1
      Top             =   3120
      Width           =   3975
   End
   Begin VB.CommandButton cmdHQ 
      Caption         =   "    HQ  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4920
      TabIndex        =   0
      Top             =   960
      Width           =   3735
   End
End
Attribute VB_Name = "frmHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Each button will Pull up the sub Grouping of unit Choices from that menu
'The Total points will add up the points the Total Points from each group and will
'compare it to the amount of points wanted stated from the beginning

Private Sub cmdClearTotalPoints_Click()
HQPoints = 0
ElitesPoints = 0
TroopsPoints = 0
FastAttackPoints = 0
HeavySupportPoints = 0


End Sub

Private Sub cmdE_Click()
frmHome.Hide
frmElites.Show

End Sub

Private Sub cmdF_Click()
frmHome.Hide
frmFastAttack.Show

End Sub

Private Sub cmdHQ_Click()
frmHome.Hide
frmHQ.Show

End Sub

Private Sub cmdHvy_Click()
frmHome.Hide
frmHeavySupport.Show

End Sub

Private Sub cmdQuit_Click()
End

End Sub

Private Sub cmdT_Click()
frmHome.Hide
frmTroops.Show

End Sub

Private Sub cmdTotalPoints_Click()
TotalPoints = HQPoints + ElitesPoints + TroopsPoints + FastAttackPoints + HeavySupportPoints
If TotalPoints = PointsWanted Then
    MsgBox "You are at " & TotalPoints & "pts. You wanted " & PointsWanted & " pts."
ElseIf TotalPoints > PointsWanted Then
    MsgBox "You wanted " & PointsWanted & "pts. You are at " & TotalPoints & " pts. You have too many pts."
ElseIf TotalPoints < PointsWanted Then
    MsgBox "You Wanted " & PointsWanted & "pts. You have " & TotalPoints & "pts. You need some more."
End If

End Sub

Private Sub Form_Load()
Open App.Path & "\Tyranids.txt" For Input As #1
    
CTR = 0

    Do While Not EOF(1)
       CTR = CTR + 1
       Input #1, names(CTR), WS(CTR), BS(CTR), S(CTR), T(CTR), W(CTR), I(CTR), A(CTR), Ld(CTR), Sv(CTR)
    Loop
    
Close #1



End Sub


