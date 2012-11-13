VERSION 5.00
Begin VB.Form SkiRental 
   Caption         =   "Ski/Snowboard Rental"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdToTitleD 
      Caption         =   "To Title"
      Height          =   1215
      Left            =   1440
      TabIndex        =   0
      Top             =   1800
      Width           =   2055
   End
End
Attribute VB_Name = "SkiRental"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdToTitleD_Click()
Title.Show
SkiRental.Hide
End Sub
