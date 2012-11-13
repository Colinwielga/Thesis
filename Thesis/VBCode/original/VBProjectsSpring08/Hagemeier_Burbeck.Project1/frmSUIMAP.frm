VERSION 5.00
Begin VB.Form frmSUIMAP 
   Caption         =   "Switzerland Detail Map"
   ClientHeight    =   4860
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13665
   LinkTopic       =   "Form8"
   Picture         =   "frmSUIMAP.frx":0000
   ScaleHeight     =   4860
   ScaleWidth      =   13665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "Return"
      Height          =   375
      Left            =   9960
      TabIndex        =   0
      Top             =   2400
      Width           =   3135
   End
End
Attribute VB_Name = "frmSUIMAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Western Europe Travel Log
'Form Name: frmSUIMAP
'Author: Nate Burbeck and Brad Hagemeier
'Date Written: 26 March 2008
'Objective: to show the viewer a detailed map of Switzerland
Private Sub cmdback_Click()
    frmCountryInfo.Visible = True
    frmSUIMAP.Visible = False
End Sub

