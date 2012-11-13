VERSION 5.00
Begin VB.Form frmElites 
   BackColor       =   &H00000040&
   Caption         =   "Elites"
   ClientHeight    =   12420
   ClientLeft      =   2430
   ClientTop       =   1275
   ClientWidth     =   18960
   FillColor       =   &H00000040&
   LinkTopic       =   "Form4"
   ScaleHeight     =   12420
   ScaleWidth      =   18960
   Begin VB.CommandButton cmdElitesPoints 
      Caption         =   "Total Elites Points"
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
      Left            =   10920
      TabIndex        =   9
      Top             =   4920
      Width           =   3015
   End
   Begin VB.CommandButton Command9 
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
      Height          =   1455
      Left            =   10920
      TabIndex        =   8
      Top             =   3120
      Width           =   3015
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Pyrovore Brood"
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
      Left            =   5760
      TabIndex        =   7
      Top             =   6840
      Width           =   3615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Venomthrope Brood"
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
      Left            =   1320
      TabIndex        =   6
      Top             =   6840
      Width           =   3615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Ymgarl Genestealer Brood"
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
      Left            =   5760
      TabIndex        =   5
      Top             =   5040
      Width           =   3615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "The Doom Of Malan'Tai"
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
      Left            =   5760
      TabIndex        =   4
      Top             =   3240
      Width           =   3615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Zoanthrope Brood"
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
      Left            =   5760
      TabIndex        =   3
      Top             =   1440
      Width           =   3615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Deathleaper"
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
      Left            =   1320
      TabIndex        =   2
      Top             =   5040
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Lictor Brood"
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
      Left            =   1320
      TabIndex        =   1
      Top             =   3240
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hive Guard Brood"
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
      Left            =   1320
      TabIndex        =   0
      Top             =   1440
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Up to Three Broods from Elites choices."
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
      Left            =   1920
      TabIndex        =   10
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "frmElites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Every button will pull up each individual unit and load the Data into a picture
'The Total points button will add up the total of each individual unit from their respective forms

Private Sub cmdElitesPoints_Click()
ElitesPoints = HiveGuardTotal + LictorTotal + DeathleaperTotal + VenomthropeTotal + ZoanthropeTotal + DoomOfMalantaiTotal + PyrovoreTotal + YmgarlGenestealersTotal

MsgBox "Your Total Points Spent on Elites is " & ElitesPoints
End Sub

Private Sub Command1_Click()
frmElites.Hide
frmHiveGuard.Show
frmHiveGuard.loadData
End Sub

Private Sub Command2_Click()
frmElites.Hide
frmLictor.Show
frmLictor.loadData

End Sub

Private Sub Command3_Click()
frmElites.Hide
frmDeathleaper.Show
frmDeathleaper.loadData

End Sub

Private Sub Command4_Click()
frmElites.Hide
frmZoanthrope.Show
frmZoanthrope.loadData

End Sub

Private Sub Command5_Click()
frmElites.Hide
frmDoomOfMalantai.Show
frmDoomOfMalantai.loadData

End Sub

Private Sub Command6_Click()
frmElites.Hide
frmYmgarlGenestealers.Show
frmYmgarlGenestealers.loadData

End Sub

Private Sub Command7_Click()
frmElites.Hide
frmVenomthrope.Show
frmVenomthrope.loadData

End Sub

Private Sub Command8_Click()
frmElites.Hide
frmPyrovore.Show
frmPyrovore.loadData

End Sub

Private Sub Command9_Click()
frmElites.Hide
frmHome.Show

End Sub
