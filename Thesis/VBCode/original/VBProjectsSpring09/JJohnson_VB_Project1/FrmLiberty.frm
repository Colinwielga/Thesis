VERSION 5.00
Begin VB.Form FrmLiberty 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   9270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   Picture         =   "FrmLiberty.frx":0000
   ScaleHeight     =   9270
   ScaleWidth      =   9195
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmdback1 
      BackColor       =   &H0080FF80&
      Caption         =   "Back"
      Height          =   735
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "FrmLiberty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Things to do in NYC
'Form Name: frmStart
'Author: Jake Johnson
'Date Written: 3/23/09
'Objective: Shows Statue of Liberty information

Private Sub Cmdback1_Click()
FrmLiberty.Hide
FrmSightseeing.Show
End Sub
