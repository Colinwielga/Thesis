VERSION 5.00
Begin VB.Form frmorder 
   Caption         =   "Form1"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   10170
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdmainmenu 
      Caption         =   "main menu"
      Height          =   1455
      Left            =   7800
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frmorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdmainmenu_Click()
frmorder.Hide
frmMain.Show

End Sub
