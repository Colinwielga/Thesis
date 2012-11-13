VERSION 5.00
Begin VB.Form frmFRAMAP 
   Caption         =   "France Detail Map"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   Picture         =   "frmFRAMAP.frx":0000
   ScaleHeight     =   5280
   ScaleWidth      =   7740
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "Return"
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Top             =   2640
      Width           =   2295
   End
End
Attribute VB_Name = "frmFRAMAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Western Europe Travel Log
'Form Name: frmFRAMAP
'Author: Nate Burbeck and Brad Hagemeier
'Date Written: 26 March 2008
'Objective: to show the viewer a detailed map of France
Private Sub cmdback_Click()
    frmCountryInfo.Visible = True
    frmFRAMAP.Visible = False
End Sub

