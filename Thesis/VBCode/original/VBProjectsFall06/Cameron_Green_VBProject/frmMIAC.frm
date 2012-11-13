VERSION 5.00
Begin VB.Form frmMIAC 
   BackColor       =   &H00008000&
   Caption         =   "MIAC Cross Country Championship Results"
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   ScaleHeight     =   8205
   ScaleWidth      =   8970
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Results Page"
      Height          =   1215
      Left            =   6360
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search for Past Results"
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   8460
      Left            =   0
      Picture         =   "frmMIAC.frx":0000
      Top             =   0
      Width           =   9000
   End
End
Attribute VB_Name = "frmMIAC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
    frmMIAC.Hide
    frmRaceResults.Show
End Sub

Private Sub Form_Load()

End Sub
