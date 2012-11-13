VERSION 5.00
Begin VB.Form frmStreet1 
   Caption         =   "Continuing on College Ave..."
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11130
   LinkTopic       =   "Form1"
   Picture         =   "frmStreet1.frx":0000
   ScaleHeight     =   8955
   ScaleWidth      =   11130
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4320
      TabIndex        =   0
      Top             =   6960
      Width           =   2535
   End
End
Attribute VB_Name = "frmStreet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdContinue_Click() 'enters you into the alley
    MsgBox ("You hear a noise down an alley.  You go to check it out and find...an alien!"), , ("Alien!")
    frmStreet1.Hide
    frmAlley.Show

End Sub
