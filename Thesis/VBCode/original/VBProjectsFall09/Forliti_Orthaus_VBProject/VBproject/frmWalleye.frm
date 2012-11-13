VERSION 5.00
Begin VB.Form frmWalleye 
   Caption         =   "Walleye"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   Picture         =   "frmWalleye.frx":0000
   ScaleHeight     =   4785
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Fishing Page"
      CausesValidation=   0   'False
      Height          =   975
      Left            =   1440
      TabIndex        =   0
      Top             =   3720
      Width           =   3735
   End
End
Attribute VB_Name = "frmWalleye"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Minnesta Outdoors
'Walleye
'Andrew Forliti and Casey Orthaus
'October 19th, 2009
'this page shows the results of which color you picked

Private Sub cmdReturn_Click()

frmFishing.Show
frmWalleye.Hide

End Sub

