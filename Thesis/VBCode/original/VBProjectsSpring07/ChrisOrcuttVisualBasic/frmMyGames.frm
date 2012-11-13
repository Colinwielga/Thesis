VERSION 5.00
Begin VB.Form frmMyGames 
   BackColor       =   &H80000007&
   Caption         =   "My Games!!"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14295
   LinkTopic       =   "Form1"
   ScaleHeight     =   8475
   ScaleWidth      =   14295
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMySetup 
      Caption         =   "My Setup"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   1
      Top             =   7800
      Width           =   2055
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12360
      TabIndex        =   0
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Image Image2 
      Height          =   3480
      Left            =   4920
      Picture         =   "frmMyGames.frx":0000
      Top             =   4080
      Width           =   2460
   End
   Begin VB.Image Image6 
      Height          =   5100
      Left            =   960
      Picture         =   "frmMyGames.frx":3573
      Top             =   3960
      Width           =   3600
   End
   Begin VB.Image Image1 
      Height          =   4245
      Left            =   0
      Picture         =   "frmMyGames.frx":F2B9
      Top             =   0
      Width           =   2925
   End
   Begin VB.Image Image3 
      Height          =   4350
      Left            =   2880
      Picture         =   "frmMyGames.frx":16A05
      Top             =   120
      Width           =   3060
   End
   Begin VB.Image Image4 
      Height          =   3855
      Left            =   7800
      Picture         =   "frmMyGames.frx":1B890
      Top             =   4320
      Width           =   2700
   End
   Begin VB.Image Image9 
      Height          =   4365
      Left            =   10920
      Picture         =   "frmMyGames.frx":1F34D
      Top             =   3840
      Width           =   3075
   End
   Begin VB.Image Image8 
      Height          =   3300
      Left            =   12000
      Picture         =   "frmMyGames.frx":2373D
      Top             =   0
      Width           =   2340
   End
   Begin VB.Image Image7 
      Height          =   4500
      Left            =   9000
      Picture         =   "frmMyGames.frx":25C48
      Top             =   0
      Width           =   3165
   End
   Begin VB.Image Image5 
      Height          =   4590
      Left            =   5880
      Picture         =   "frmMyGames.frx":297B9
      Top             =   360
      Width           =   3195
   End
End
Attribute VB_Name = "frmMyGames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form shows my game collection through use of Box Art
Option Explicit
Private Sub cmdBack_Click()
    frmMyGames.Hide     'Hides MyGames form
    frmAboutMe.Show     'Shows AboutMe form
End Sub
Private Sub cmdMySetup_Click()
    frmMyGames.Hide     'Hides MyGames form
    frmMySetup.Show     'Shows MySetup form
End Sub

