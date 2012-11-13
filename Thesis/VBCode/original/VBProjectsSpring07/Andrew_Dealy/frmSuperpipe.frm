VERSION 5.00
Begin VB.Form frmSuperpipe 
   Caption         =   "Superpipe "
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   4470
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   6015
      Left            =   0
      Picture         =   "frmSuperpipe.frx":0000
      ScaleHeight     =   5955
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.CommandButton cmdFinals 
         Caption         =   "Finals"
         Height          =   495
         Left            =   2520
         TabIndex        =   2
         Top             =   5400
         Width           =   1695
      End
      Begin VB.CommandButton cmdPrelims 
         Caption         =   "Preliminaries"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   5400
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmSuperpipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'offers links to preliminary and final superpipe competitions
Private Sub cmdFinals_Click()
    frmSuperpipeFinals.Show
    frmSuperpipe.Hide
End Sub

Private Sub cmdPrelims_Click()
    frmSuperpipePrelims.Show
    frmSuperpipe.Hide
End Sub
