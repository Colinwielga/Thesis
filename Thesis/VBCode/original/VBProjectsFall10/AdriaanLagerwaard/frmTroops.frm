VERSION 5.00
Begin VB.Form frmTroops 
   BackColor       =   &H00000040&
   Caption         =   "Troops"
   ClientHeight    =   12420
   ClientLeft      =   3585
   ClientTop       =   2415
   ClientWidth     =   19860
   LinkTopic       =   "Form3"
   ScaleHeight     =   12420
   ScaleWidth      =   19860
   Begin VB.CommandButton cmdTroopsPoints 
      Caption         =   "Total Points Spent on Troops"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   12120
      TabIndex        =   6
      Top             =   5040
      Width           =   2895
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Back To Home"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   12120
      TabIndex        =   5
      Top             =   2880
      Width           =   2895
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Ripper Swarm Brood"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1800
      TabIndex        =   4
      Top             =   5400
      Width           =   3855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Hormagaunt Brood"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   6600
      TabIndex        =   3
      Top             =   3480
      Width           =   3615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Termagant Brood"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   6600
      TabIndex        =   2
      Top             =   1320
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Genestealer Brood"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1800
      TabIndex        =   1
      Top             =   3480
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Tyranid Warrior Brood"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   1800
      TabIndex        =   0
      Top             =   1320
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Up to 6 Troop Choices from the following :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      TabIndex        =   7
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "frmTroops"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Every button will pull up each individual unit and load the Data into a picture
'The Total points button will add up the total of each individual unit from their respective forms

Private Sub cmdTroopsPoints_Click()
TroopsPoints = WarriorTotal + GenestealerTotal + EquipGenestealer + EquipWarrior + EquipBroodlord + TermTotal + EquipTerm + HormTotal + EquipHorm + RipperSwarmTotal + EquipRipperSwarm

MsgBox "Your Total Points Spent on Troops is " & TroopsPoints
End Sub

Private Sub Command1_Click()
frmTroops.Hide
frmWarrior.Show
frmWarrior.loadData

End Sub

Private Sub Command2_Click()
frmTroops.Hide
frmGenestealer.Show
frmGenestealer.loadData

End Sub



Private Sub Command4_Click()
frmTroops.Hide
frmTermagant.Show
frmTermagant.loadData

End Sub

Private Sub Command5_Click()
frmTroops.Hide
frmHormagaunt.Show
frmHormagaunt.loadData

End Sub

Private Sub Command6_Click()
frmTroops.Hide
frmRipperSwarm.Show
frmRipperSwarm.loadData


End Sub

Private Sub Command7_Click()
frmTroops.Hide
frmHome.Show

End Sub
