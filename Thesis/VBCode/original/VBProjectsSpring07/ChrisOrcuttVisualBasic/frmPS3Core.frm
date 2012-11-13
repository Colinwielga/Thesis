VERSION 5.00
Begin VB.Form frmPS3Core 
   Caption         =   "Sony Playstation 3 (20 Gig) Unit"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   4200
      Width           =   1455
   End
End
Attribute VB_Name = "frmPS3Core"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdReturn_Click()
    frmPS3Core.Hide
    frmConsoleInfo.Show
End Sub
