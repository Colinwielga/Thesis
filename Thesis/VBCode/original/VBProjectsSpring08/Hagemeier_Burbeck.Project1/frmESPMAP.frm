VERSION 5.00
Begin VB.Form frmESPMAP 
   Caption         =   "Spain Detail Map"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7965
   LinkTopic       =   "Form7"
   Picture         =   "frmESPMAP.frx":0000
   ScaleHeight     =   5295
   ScaleWidth      =   7965
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "Return"
      Height          =   375
      Left            =   5280
      TabIndex        =   0
      Top             =   2400
      Width           =   2415
   End
End
Attribute VB_Name = "frmESPMAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Western Europe Travel Log
'Form Name: frmESPMAP
'Author: Nate Burbeck and Brad Hagemeier
'Date Written: 26 March 2008
'Objective: to show the viewer a detailed map of Spain
Private Sub cmdback_Click()
    frmCountryInfo.Visible = True
    frmESPMAP.Visible = False
End Sub

