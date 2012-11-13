VERSION 5.00
Begin VB.Form frmHQ 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000040&
   Caption         =   "HQ"
   ClientHeight    =   12285
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13920
   LinkTopic       =   "Form2"
   ScaleHeight     =   12285
   ScaleWidth      =   13920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHQPoints 
      Caption         =   "Total HQ Points"
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
      Left            =   4200
      TabIndex        =   6
      Top             =   6600
      Width           =   2655
   End
   Begin VB.CommandButton cmdHQBack 
      Caption         =   "Go Back to Home"
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
      Left            =   7560
      TabIndex        =   5
      Top             =   6600
      Width           =   2655
   End
   Begin VB.CommandButton cmdParasiteOfMortrex 
      Caption         =   "The Parasite of Mortrex"
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
      Left            =   6120
      TabIndex        =   4
      Top             =   2760
      Width           =   3255
   End
   Begin VB.CommandButton cmdTyranidPrime 
      Caption         =   "Tyranid Prime"
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
      Left            =   6120
      TabIndex        =   3
      Top             =   1200
      Width           =   3255
   End
   Begin VB.CommandButton cmdTervigon 
      Caption         =   "Tervigon"
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
      Left            =   840
      TabIndex        =   2
      Top             =   4320
      Width           =   3255
   End
   Begin VB.CommandButton cmdSwarmLord 
      Caption         =   "The SwarmLord"
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
      Left            =   840
      TabIndex        =   1
      Top             =   2760
      Width           =   3255
   End
   Begin VB.CommandButton cmdHiveTyrant 
      Caption         =   "Hive Tyrant"
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
      Left            =   840
      TabIndex        =   0
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Up to Two Choices from the following"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   7
      Top             =   360
      Width           =   6255
   End
End
Attribute VB_Name = "frmHQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'Every button will pull up each individual unit and load the Data into a picture
'The Total points button will add up the total of each individual unit from their respective forms


Private Sub cmdHiveTyrant_Click()
frmHQ.Hide
frmHiveTyrant.Show
frmHiveTyrant.loadData

End Sub

Private Sub cmdHQBack_Click()
frmHQ.Hide
frmHome.Show

End Sub

Private Sub cmdHQPoints_Click()
HQPoints = HiveTyrantTotal + SwarmLordTotal + TervigonTotal + PrimeTotal + ParasiteOfMortrexTotal

MsgBox "Your Total Points Spent on HQs is " & HQPoints
End Sub

Private Sub cmdParasiteOfMortrex_Click()
frmHQ.Hide
frmParasiteOfMortrex.Show
frmParasiteOfMortrex.loadData

End Sub



Private Sub cmdSwarmLord_Click()
frmHQ.Hide
frmSwarmLord.Show
frmSwarmLord.loadData

End Sub

Private Sub cmdTervigon_Click()
frmHQ.Hide
frmTervigon.Show
frmTervigon.loadData

End Sub

Private Sub cmdTyranidPrime_Click()
frmHQ.Hide
frmPrime.Show
frmPrime.loadData

End Sub

