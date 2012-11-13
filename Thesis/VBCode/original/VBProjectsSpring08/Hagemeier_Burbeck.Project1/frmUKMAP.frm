VERSION 5.00
Begin VB.Form frmUKMAP 
   Caption         =   "United Kingdom Detail Map"
   ClientHeight    =   10140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   Picture         =   "frmUKMAP.frx":0000
   ScaleHeight     =   10140
   ScaleWidth      =   8610
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "Return"
      Height          =   375
      Left            =   5280
      TabIndex        =   0
      Top             =   5160
      Width           =   2895
   End
End
Attribute VB_Name = "frmUKMAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Western Europe Travel Log
'Form Name: frmUKMAP
'Author: Nate Burbeck and Brad Hagemeier
'Date Written: 26 March 2008
'Objective: to show the viewer a detailed map of the United Kingdom
Private Sub cmdback_Click()
    frmCountryInfo.Visible = True
    frmUKMAP.Visible = False
End Sub

