VERSION 5.00
Begin VB.Form frmCoaches 
   Caption         =   "Coaches"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   10500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back "
      Height          =   1215
      Left            =   840
      TabIndex        =   0
      Top             =   7200
      Width           =   3615
   End
   Begin VB.Label lblHead 
      Caption         =   "Head Coach: "
      Height          =   615
      Left            =   720
      TabIndex        =   1
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "frmCoaches"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    
Private Sub cmdBack_Click()
    frmCoaches.Visible = False
    frmPlayers.Visible = False
    frm1.Visible = True
End Sub
