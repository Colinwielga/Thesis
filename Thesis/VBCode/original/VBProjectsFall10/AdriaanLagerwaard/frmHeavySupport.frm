VERSION 5.00
Begin VB.Form frmHeavySupport 
   BackColor       =   &H00000040&
   Caption         =   "Heavy Support"
   ClientHeight    =   12435
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20055
   LinkTopic       =   "Form1"
   ScaleHeight     =   12435
   ScaleWidth      =   20055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Total Points Spent on Heavy Support"
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
      Left            =   11160
      TabIndex        =   7
      Top             =   3960
      Width           =   2775
   End
   Begin VB.CommandButton cmdHeavySupportBack 
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
      Left            =   11160
      TabIndex        =   6
      Top             =   1920
      Width           =   2775
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Tyrannofex"
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
      Left            =   5760
      TabIndex        =   5
      Top             =   5640
      Width           =   3135
   End
   Begin VB.CommandButton Mawloc 
      Caption         =   "Mawloc"
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
      Left            =   5760
      TabIndex        =   4
      Top             =   3600
      Width           =   3135
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Trygon"
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
      Left            =   5760
      TabIndex        =   3
      Top             =   1560
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Biovore Brood"
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
      Left            =   720
      TabIndex        =   2
      Top             =   5640
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Old One Eye"
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
      Left            =   720
      TabIndex        =   1
      Top             =   3600
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Carnifex Brood"
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
      Left            =   720
      TabIndex        =   0
      Top             =   1560
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Up to 3 Heavy Support Choices from the following:"
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
      Left            =   840
      TabIndex        =   8
      Top             =   360
      Width           =   8175
   End
End
Attribute VB_Name = "frmHeavySupport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Every button will pull up each individual unit and load the Data into a picture
'The Total points button will add up the total of each individual unit from their respective forms

Private Sub cmdHeavySupportBack_Click()
frmHeavySupport.Hide
frmHome.Show

End Sub

Private Sub Command1_Click()
frmHeavySupport.Hide
frmCarnifex.Show
frmCarnifex.loadData

End Sub

Private Sub Command2_Click()
frmHeavySupport.Hide
frmOldOneEye.Show
frmOldOneEye.loadData

End Sub

Private Sub Command3_Click()
frmHeavySupport.Hide
frmBiovore.Show
frmBiovore.loadData

End Sub

Private Sub Command4_Click()
frmHeavySupport.Hide
frmTrygon.Show
frmTrygon.loadData


End Sub

Private Sub Command5_Click()
HeavySupportPoints = CarnifexTotal + EquipCarnifex + OldTotal + BiovoreTotal + TrygonTotal + EquipTrygon + MawlocTotal + EquipMawloc + TyranTotal + EquipTyran

MsgBox "Total Points Spent on Heavy Support Choices " & HeavySupportPoints
End Sub

Private Sub Command6_Click()
frmHeavySupport.Hide
frmTyrannofex.Show
frmTyrannofex.loadData

End Sub

Private Sub Mawloc_Click()
frmHeavySupport.Hide
frmMawloc.Show
frmMawloc.loadData

End Sub
