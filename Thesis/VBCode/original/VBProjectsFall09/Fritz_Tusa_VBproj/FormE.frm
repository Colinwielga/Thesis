VERSION 5.00
Begin VB.Form HotelCost 
   Caption         =   "Hotel Cost"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdToTitleE 
      Caption         =   "To Title"
      Height          =   1095
      Left            =   1800
      TabIndex        =   0
      Top             =   1920
      Width           =   1935
   End
End
Attribute VB_Name = "HotelCost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdToTitleE_Click()
Title.Show
HotelCost.Hide
End Sub
