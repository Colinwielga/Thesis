VERSION 5.00
Begin VB.Form frmCrappie 
   Caption         =   "Black Crappie"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   Picture         =   "frmCrappie.frx":0000
   ScaleHeight     =   6705
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Fishing Page"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   5880
      Width           =   3495
   End
End
Attribute VB_Name = "frmCrappie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Minnesta Outdoors
'Crappie
'Andrew Forliti and Casey Orthaus
'October 19th, 2009
'this page shows the results of which color you picked

Private Sub cmdReturn_Click()

frmFishing.Show
frmCrappie.Hide

End Sub

