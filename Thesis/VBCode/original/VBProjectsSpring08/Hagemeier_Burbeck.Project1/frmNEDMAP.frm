VERSION 5.00
Begin VB.Form frmNEDMAP 
   Caption         =   "Netherlands Detail Map"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8565
   LinkTopic       =   "Form5"
   Picture         =   "frmNEDMAP.frx":0000
   ScaleHeight     =   5295
   ScaleWidth      =   8565
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "Return"
      Height          =   375
      Left            =   5400
      TabIndex        =   0
      Top             =   2520
      Width           =   2655
   End
End
Attribute VB_Name = "frmNEDMAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Western Europe Travel Log
'Form Name: frmNEDMAP
'Author: Nate Burbeck and Brad Hagemeier
'Date Written: 26 March 2008
'Objective: to show the viewer a detailed map of the Netherlands
Private Sub cmdback_Click()
    frmCountryInfo.Visible = True
    frmNEDMAP.Visible = False
End Sub

