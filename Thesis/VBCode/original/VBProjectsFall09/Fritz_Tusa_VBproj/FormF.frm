VERSION 5.00
Begin VB.Form Runs 
   Caption         =   "What Ski Runs to do!"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdToTitleF 
      Caption         =   "To Title"
      Height          =   1215
      Left            =   1680
      TabIndex        =   0
      Top             =   1800
      Width           =   1815
   End
End
Attribute VB_Name = "Runs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdToTitleF_Click()
Title.Show
Runs.Hide
End Sub
