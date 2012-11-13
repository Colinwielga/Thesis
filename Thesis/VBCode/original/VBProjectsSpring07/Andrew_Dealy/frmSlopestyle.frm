VERSION 5.00
Begin VB.Form frmSlopestyle 
   Caption         =   "Slopestyle"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   4515
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFinals 
      Caption         =   "Final Results"
      Height          =   615
      Left            =   2520
      TabIndex        =   2
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton cmdPrelims 
      BackColor       =   &H000000C0&
      Caption         =   "Preliminary Results"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin VB.PictureBox picSlope 
      Height          =   4575
      Left            =   0
      Picture         =   "frmSlopestyle.frx":0000
      ScaleHeight     =   4515
      ScaleWidth      =   4515
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmSlopestyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this page allows the user to choose either the preliminary runs or final competition rounds.

Private Sub cmdFinals_Click()
    frmslopestyleFinals.Show
    frmSlopestyle.Hide
End Sub

Private Sub cmdPrelims_Click()
    frmSlopestylePrelims.Show
    frmSlopestyle.Hide
End Sub
