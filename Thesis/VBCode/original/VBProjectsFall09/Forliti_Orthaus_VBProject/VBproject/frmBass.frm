VERSION 5.00
Begin VB.Form frmBass 
   Caption         =   "Largemouth Bass"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   Picture         =   "frmBass.frx":0000
   ScaleHeight     =   6240
   ScaleWidth      =   11835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Fishing Page"
      Height          =   975
      Left            =   3960
      TabIndex        =   0
      Top             =   5160
      Width           =   3735
   End
End
Attribute VB_Name = "frmBass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Minnesta Outdoors
'Bass
'Andrew Forliti and Casey Orthaus
'October 19th, 2009
'this page shows the results of which color you picked


Private Sub cmdReturn_Click()

frmFishing.Show
frmBass.Hide

End Sub

