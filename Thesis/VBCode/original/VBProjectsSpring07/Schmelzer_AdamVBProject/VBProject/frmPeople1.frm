VERSION 5.00
Begin VB.Form frmPeople1 
   Caption         =   "The People"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   Picture         =   "frmPeople1.frx":0000
   ScaleHeight     =   7215
   ScaleWidth      =   9600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Capitulate"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8160
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdFreebeer 
      Caption         =   "Alcohol will no longer be taxed."
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5520
      TabIndex        =   4
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton cmdRation2 
      Caption         =   "The people shall receive 5 free rations per week for the duration of the war."
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2880
      TabIndex        =   3
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton cmdNothing 
      Caption         =   "The people will have to make due for the time being."
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5520
      TabIndex        =   2
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CommandButton cmdFreerations 
      Caption         =   "Give the people 3 free rations per week for the duration of the war"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2880
      TabIndex        =   1
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label lblInstructions 
      Caption         =   $"frmPeople1.frx":3051D
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   7335
      Left            =   0
      Top             =   0
      Width           =   9975
   End
End
Attribute VB_Name = "frmPeople1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'to give the user different options through command buttons in order to deal with the
'fiction people and economy in the game
'certain choices affect the variety of variables (battlepoints, econpoints,
'and resources)

Private Sub cmdFreebeer_Click()
Econpoints = Econpoints + 0
'the men recruited will be more malleable and dumbness=bravery
Battlepoints = Battlepoints + 100
Resources = Resources - 0
Alcoholvariable = True
frmPeople1.Hide
frmCouncilors1.Show
End Sub

Private Sub cmdFreerations_Click()
Econpoints = Econpoints + 1
Resources = Resources - 100
frmPeople1.Hide
frmCouncilors1.Show
End Sub

Private Sub cmdNothing_Click()
Econpoints = Econpoints - 2
'resources will be saved
Resources = Resources + 100
frmPeople1.Hide
frmCouncilors1.Show
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdRation2_Click()
Econpoints = Econpoints + 2
Resources = Resources - 200
frmPeople1.Hide
frmCouncilors1.Show
End Sub
