VERSION 5.00
Begin VB.Form frmPORMAP 
   Caption         =   "Portugal Detail Map"
   ClientHeight    =   10680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8250
   LinkTopic       =   "Form6"
   Picture         =   "frmPORMAP.frx":0000
   ScaleHeight     =   10680
   ScaleWidth      =   8250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "Return"
      Height          =   375
      Left            =   5280
      TabIndex        =   0
      Top             =   5160
      Width           =   2655
   End
End
Attribute VB_Name = "frmPORMAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Western Europe Travel Log
'Form Name: frmPORMAP
'Author: Nate Burbeck and Brad Hagemeier
'Date Written: 26 March 2008
'Objective: to show the viewer a detailed map of Portugal
Private Sub cmdback_Click()
    frmCountryInfo.Visible = True
    frmPORMAP.Visible = False
End Sub

