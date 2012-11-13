VERSION 5.00
Begin VB.Form frmCatFish 
   Caption         =   "Catfish"
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   Picture         =   "frmCatFish.frx":0000
   ScaleHeight     =   4995
   ScaleWidth      =   4785
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Fishing Page"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   3960
      Width           =   4575
   End
End
Attribute VB_Name = "frmCatFish"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Minnesta Outdoors
'CatFish
'Andrew Forliti and Casey Orthaus
'October 19th, 2009
'this page shows the results of which color you picked

Private Sub cmdReturn_Click()

frmFishing.Show
frmCatFish.Hide

End Sub

