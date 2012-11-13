VERSION 5.00
Begin VB.Form frmFastAttack 
   BackColor       =   &H00000040&
   Caption         =   "Fast Attack"
   ClientHeight    =   12075
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19050
   LinkTopic       =   "Form2"
   ScaleHeight     =   12075
   ScaleWidth      =   19050
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFastAttackPoints 
      Caption         =   "Total Points Spent on Fast Attack"
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
      Left            =   12360
      TabIndex        =   7
      Top             =   5520
      Width           =   2655
   End
   Begin VB.CommandButton cmdFastAttackBack 
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
      Height          =   1575
      Left            =   12360
      TabIndex        =   6
      Top             =   3600
      Width           =   2655
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Spore Mine Cluster"
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
      Left            =   6360
      TabIndex        =   5
      Top             =   5400
      Width           =   3375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Harpy"
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
      Left            =   6360
      TabIndex        =   4
      Top             =   3360
      Width           =   3375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Gargoyle Brood"
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
      Left            =   6360
      TabIndex        =   3
      Top             =   1320
      Width           =   3375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Sky-Slasher Swarm Brood"
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
      Left            =   1080
      TabIndex        =   2
      Top             =   5400
      Width           =   3735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ravener Brood"
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
      Left            =   1080
      TabIndex        =   1
      Top             =   3360
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Tyranid Shrike Brood"
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
      Left            =   1080
      TabIndex        =   0
      Top             =   1320
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Up to 3 Fast Attack Choices from the following:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   8
      Top             =   240
      Width           =   7575
   End
End
Attribute VB_Name = "frmFastAttack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Every button will pull up each individual unit and load the Data into a picture
'The Total points button will add up the total of each individual unit from their respective forms

Private Sub cmdFastAttackBack_Click()
frmFastAttack.Hide
frmHome.Show

End Sub

Private Sub cmdFastAttackPoints_Click()
FastAttackPoints = ShrikeTotal + EquipShrike + RavenerTotal + EquipRavener + SkySlasherTotal + EquipSkySlasher + GargTotal + EquipGarg + HarpyTotal + EquipHarpy + SporeTotal

MsgBox "Your Total Points Spent on FastAttack is " & FastAttackPoints

End Sub

Private Sub Command1_Click()
frmFastAttack.Hide
frmShrike.Show
frmShrike.loadData

End Sub

Private Sub Command2_Click()
frmFastAttack.Hide
frmRavener.Show
frmRavener.loadData

End Sub

Private Sub Command3_Click()
frmFastAttack.Hide
frmskyslasher.Show
frmskyslasher.loadData

End Sub

Private Sub Command4_Click()
frmFastAttack.Hide
frmGargoyle.Show
frmGargoyle.loadData

End Sub

Private Sub Command5_Click()
frmFastAttack.Hide
frmHarpy.Show
frmHarpy.loadData

End Sub

Private Sub Command6_Click()
frmFastAttack.Hide
frmSporeMineCluster.Show
frmSporeMineCluster.loadData

End Sub
