VERSION 5.00
Begin VB.Form frmSoccer 
   Caption         =   "Soccer Shootout"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "frmSoccer.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Main Menu"
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   7800
      Width           =   1815
   End
   Begin VB.CommandButton cmdGame2 
      Caption         =   "France v. Germany"
      Height          =   2055
      Left            =   5160
      TabIndex        =   1
      Top             =   5400
      Width           =   2175
   End
   Begin VB.CommandButton cmdGame1 
      Caption         =   "England v. Argentina"
      Height          =   1815
      Left            =   5160
      TabIndex        =   0
      Top             =   1320
      Width           =   2175
   End
End
Attribute VB_Name = "frmSoccer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdGame1_Click()
    'Allows the user to access the particular game
    frmGame1.Show
    frmSoccer.Hide
End Sub

Private Sub cmdGame2_Click()
    'Allows the user to access the particular game
    frmGame2.Show
    frmSoccer.Hide
End Sub



Private Sub cmdReturn_Click()
    'Return to Main menu
    frmSoccer.Hide
    frmHome.Show
End Sub
