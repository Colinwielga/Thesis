VERSION 5.00
Begin VB.Form frmReferences 
   Caption         =   "References"
   ClientHeight    =   4305
   ClientLeft      =   3975
   ClientTop       =   3330
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   ScaleHeight     =   4305
   ScaleWidth      =   6645
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Click to return to the main menu"
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "frmReferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdReturn_Click()
    'return to the main menu
    frmStart.Show
    frmReferences.Hide
    
End Sub
