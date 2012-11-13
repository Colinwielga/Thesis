VERSION 5.00
Begin VB.Form frmDailyDouble 
   Caption         =   "Daily Double"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDailyDouble 
      Height          =   3615
      Left            =   0
      Picture         =   "frmDailyDouble.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "frmDailyDouble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDailyDouble_Click()
    
    'Hides and shows the forms
    frmHis200.Show
    frmDailyDouble.Hide
    
End Sub
