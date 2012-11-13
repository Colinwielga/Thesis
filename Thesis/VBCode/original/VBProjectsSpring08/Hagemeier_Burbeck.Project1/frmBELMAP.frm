VERSION 5.00
Begin VB.Form frmBELMAP 
   Caption         =   "Belgium Detail Map"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   Picture         =   "frmBELMAP.frx":0000
   ScaleHeight     =   5280
   ScaleWidth      =   7860
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "Return"
      Height          =   375
      Left            =   5280
      TabIndex        =   0
      Top             =   2640
      Width           =   2295
   End
End
Attribute VB_Name = "frmBELMAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Western Europe Travel Log
'Form Name: frmBELMAP
'Author: Nate Burbeck and Brad Hagemeier
'Date Written: 26 March 2008
'Objective: to show the viewer a detailed map of Belgium
Private Sub cmdback_Click()
    frmCountryInfo.Visible = True
    frmBELMAP.Visible = False
End Sub

