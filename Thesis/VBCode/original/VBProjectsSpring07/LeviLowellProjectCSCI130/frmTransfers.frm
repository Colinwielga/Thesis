VERSION 5.00
Begin VB.Form frmTransfers 
   Caption         =   "Account Transfers"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTransfers 
      Caption         =   "Transfer"
      Height          =   975
      Left            =   2400
      TabIndex        =   2
      Top             =   2040
      Width           =   2295
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Main Page"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label lblTransfers 
      Alignment       =   2  'Center
      Caption         =   "Here you can transfer money from your savings account to your checkings accouna and vice versa!"
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmTransfers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdReturn_Click()
frmTransfers.Hide
FrmMain.Show
End Sub

Private Sub cmdTransfers_Click()

End Sub
