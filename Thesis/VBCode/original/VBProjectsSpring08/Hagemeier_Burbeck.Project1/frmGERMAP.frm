VERSION 5.00
Begin VB.Form frmGERMAP 
   Caption         =   "Germany Detail Map"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7980
   LinkTopic       =   "Form2"
   Picture         =   "frmGERMAP.frx":0000
   ScaleHeight     =   5280
   ScaleWidth      =   7980
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "Return"
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Top             =   2760
      Width           =   2655
   End
End
Attribute VB_Name = "frmGERMAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Western Europe Travel Log
'Form Name: frmGERMAP
'Author: Nate Burbeck and Brad Hagemeier
'Date Written: 26 March 2008
'Objective: to show the viewer a detailed map of Germany
Private Sub cmdback_Click()
    frmCountryInfo.Visible = True
    frmGERMAP.Visible = False
End Sub

