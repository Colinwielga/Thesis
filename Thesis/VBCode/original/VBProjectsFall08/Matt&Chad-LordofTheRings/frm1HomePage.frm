VERSION 5.00
Begin VB.Form frm1HomePage 
   Caption         =   "Form1"
   ClientHeight    =   7860
   ClientLeft      =   2220
   ClientTop       =   1515
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   Picture         =   "frm1HomePage.frx":0000
   ScaleHeight     =   7860
   ScaleWidth      =   10470
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H80000000&
      Caption         =   "Start Your Journey"
      Height          =   975
      Left            =   8280
      MaskColor       =   &H80000000&
      TabIndex        =   0
      Top             =   2160
      Width           =   1815
   End
End
Attribute VB_Name = "frm1HomePage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdStart_Click()
    frm1HomePage.Hide
    frm2Characters.Show
End Sub

