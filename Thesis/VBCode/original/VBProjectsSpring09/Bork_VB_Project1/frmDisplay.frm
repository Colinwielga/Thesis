VERSION 5.00
Begin VB.Form frmDisplay 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Current Items"
   ClientHeight    =   9705
   ClientLeft      =   8955
   ClientTop       =   945
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   Picture         =   "frmDisplay.frx":0000
   ScaleHeight     =   9705
   ScaleWidth      =   6270
   Begin VB.PictureBox picResults 
      BackColor       =   &H80000009&
      Height          =   9495
      Left            =   240
      ScaleHeight     =   9435
      ScaleWidth      =   5715
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "frmDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Wilderness Outfitters Partial Outfitting
'Justin Bork
'March, 2009
'
'frmDisplay
'Displays items that have been chosen for rental.
'The items are seperated according to their section.
