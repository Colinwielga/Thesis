VERSION 5.00
Begin VB.Form frmIREMAP 
   Caption         =   "Ireland Detail Map"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7800
   LinkTopic       =   "Form3"
   Picture         =   "frmIREMAP.frx":0000
   ScaleHeight     =   5295
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "Return"
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Top             =   2520
      Width           =   2415
   End
End
Attribute VB_Name = "frmIREMAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Western Europe Travel Log
'Form Name: frmIREMAP
'Author: Nate Burbeck and Brad Hagemeier
'Date Written: 26 March 2008
'Objective: to show the viewer a detailed map of Ireland
Private Sub cmdback_Click()
    frmCountryInfo.Visible = True
    frmIREMAP.Visible = False
End Sub

