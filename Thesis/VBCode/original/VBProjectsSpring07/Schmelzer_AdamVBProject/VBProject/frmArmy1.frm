VERSION 5.00
Begin VB.Form frmArmy1 
   Caption         =   "Raising an Army"
   ClientHeight    =   8625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   ScaleHeight     =   8625
   ScaleWidth      =   10995
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Capitulate"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8760
      TabIndex        =   5
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton cmdSengines 
      Caption         =   "Build 20 Siege Engines (2 months)"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8760
      TabIndex        =   4
      Top             =   2160
      Width           =   2055
   End
   Begin VB.CommandButton cmdKnights 
      Caption         =   "There are 100 Squires ready to be promoted to Knights at your calling (0)"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8760
      TabIndex        =   3
      Top             =   3120
      Width           =   2055
   End
   Begin VB.CommandButton cmdArchers 
      Caption         =   "Ready 250 Archers (2 months)"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8760
      TabIndex        =   2
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton cmdPikemen 
      Caption         =   "Ready 1000 pikemen (1 month)"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8760
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label lblInstructions 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmArmy1.frx":0000
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   8640
      Left            =   0
      Picture         =   "frmArmy1.frx":0199
      Top             =   0
      Width           =   10995
   End
End
Attribute VB_Name = "frmArmy1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Objective: to give the user different options through command buttonsm, which
'affect certain variables.  Namely, battlepoints, troop levels, and resources
'all of these are public variables
'each choice advances the user to a subsequent form


Private Sub cmdArchers_Click()
Battlepoints = Battlepoints + (250 * 4)
archers = archers + 250
Resources = Resources - 200
frmArmy1.Hide
frmElders11.Show

End Sub

Private Sub cmdKnights_Click()
Battlepoints = Battlepoints + (100 * 10)
knights = knights + 100
Resources = Resources - 0
frmArmy1.Hide
frmElders11.Show
End Sub

Private Sub cmdPikemen_Click()
Battlepoints = Battlepoints + (1000 * 1)
pikemen = pikemen + 500
Resources = Resources - 100
frmArmy1.Hide
frmElders11.Show
End Sub

Private Sub cmdSengines_Click()
Battlepoints = Battlepoints + (20 * 10)
siege = siege + 20
Resources = Resources - 200
frmArmy1.Hide
frmElders11.Show
End Sub


