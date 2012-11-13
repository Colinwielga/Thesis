VERSION 5.00
Begin VB.Form frmNorthern 
   Caption         =   "Northern Pike"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   Picture         =   "frmNorthern.frx":0000
   ScaleHeight     =   5265
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Fishing Page"
      Height          =   735
      Left            =   2160
      TabIndex        =   0
      Top             =   4440
      Width           =   4695
   End
End
Attribute VB_Name = "frmNorthern"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Minnesta Outdoors
'Northern
'Andrew Forliti and Casey Orthaus
'October 19th, 2009
'this page shows the results of which color you picked

Private Sub cmdReturn_Click()

frmFishing.Show
frmNorthern.Hide

End Sub

