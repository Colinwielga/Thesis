VERSION 5.00
Begin VB.Form frmMatch 
   Caption         =   "Match"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   Picture         =   "frmMatch.frx":0000
   ScaleHeight     =   4485
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFacts 
      BackColor       =   &H00000080&
      Caption         =   "Click to play a matching game with facts"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CommandButton cmdPictures 
      BackColor       =   &H00000080&
      Caption         =   "Click to play a matching game with pictures"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Return to Main Menu"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3600
      Width           =   1455
   End
End
Attribute VB_Name = "frmMatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFacts_Click()
    'go to match facts form
    frmMatchFacts.Show
    frmMatch.Hide
    
End Sub

Private Sub cmdPictures_Click()
    'go to match pictures form
    frmMatchPictures.Show
    frmMatch.Hide
    
End Sub

Private Sub cmdReturn_Click()
    'return to main menu
    frmStart.Show
    frmMatch.Hide
End Sub
