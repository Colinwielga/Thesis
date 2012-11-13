VERSION 5.00
Begin VB.Form frmImprove 
   BackColor       =   &H00008000&
   Caption         =   "How to Improve your VO2 Max"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13365
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   13365
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C00000&
      Caption         =   "Back to VO2 Max Page"
      Height          =   735
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6600
      Width           =   2175
   End
   Begin VB.Label lblImprove 
      BackColor       =   &H00008000&
      Caption         =   $"frmImprove.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   6375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   12975
   End
End
Attribute VB_Name = "frmImprove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'goes back to VO2 max page, text on page shows tips to improve VO2 max'
Private Sub cmdBack_Click()
    frmImprove.Hide
    frmVO2Max.Show
End Sub
