VERSION 5.00
Begin VB.Form frmInventory 
   Caption         =   "Inventory"
   ClientHeight    =   8790
   ClientLeft      =   1560
   ClientTop       =   945
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   Picture         =   "frmInventory.frx":0000
   ScaleHeight     =   8790
   ScaleWidth      =   7350
   Begin VB.PictureBox picResults2 
      BackColor       =   &H8000000E&
      Height          =   8295
      Left            =   3720
      ScaleHeight     =   8235
      ScaleWidth      =   3435
      TabIndex        =   1
      Top             =   240
      Width           =   3495
   End
   Begin VB.PictureBox picResults 
      BackColor       =   &H8000000E&
      Height          =   8295
      Left            =   120
      ScaleHeight     =   8235
      ScaleWidth      =   3435
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "frmInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Wilderness Outfitters Partial Outfitting
'Justin Bork
'March, 2009
'
'frmInventory
'Displays current inventory
