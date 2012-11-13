VERSION 5.00
Begin VB.Form frmDeclaration 
   Caption         =   "Hail!"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   Picture         =   "frmDeclaration.frx":0000
   ScaleHeight     =   7140
   ScaleWidth      =   9465
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Capitulate?"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   6480
      Width           =   2775
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Begin anew?"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   6120
      Width           =   2415
   End
   Begin VB.CommandButton cmdBegin 
      Caption         =   "Forward!"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6720
      TabIndex        =   0
      Top             =   6120
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   5055
      Left            =   0
      Top             =   0
      Width           =   8415
   End
   Begin VB.Label lblInstructions 
      BackColor       =   &H0080FFFF&
      Caption         =   $"frmDeclaration.frx":327B5
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   600
      TabIndex        =   2
      Top             =   240
      Width           =   7095
   End
End
Attribute VB_Name = "frmDeclaration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Objective: to give the user further instructions, a chance to start the game over,
'or to quit the game
'restarting the program reinitializes variables that were altered in the first 'start'
'form
'this form also serves a large role in initializing many of the key variables
'that affect the rest of the game including alliances, and battle outcome variables in
'relation to the second battle form.  It also initializes the boolean
'variable stating whether or not a future character is alive.

Private Sub cmdBack_Click()
my1variable = False
my2variable = False
frmDeclaration.Hide
frmStart.Show
End Sub

Private Sub cmdBegin_Click()
'initializes later used variables
Battlepoints = Battlepoints + (3000 * 1) + (500 * 4) + (75 * 10)
BoltenArmy = 12000
Resources = 1000

LannisterAllianceP = False
LannisterAllianceN = False
BoltenAllianceP = False
BoltenAllianceN = False

LannisterLifeV = True

failsiegeV = False
successfulsiegeV = False
waitedV = False
blockadeV = False

frmDeclaration.Hide
frmMessage.Show
End Sub

Private Sub cmdQuit_Click()
End
End Sub
