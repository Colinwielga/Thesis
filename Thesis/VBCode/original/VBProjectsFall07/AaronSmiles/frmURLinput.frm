VERSION 5.00
Begin VB.Form frmURLinput 
   Caption         =   "Input your desired address."
   ClientHeight    =   2160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   ScaleHeight     =   2160
   ScaleWidth      =   6780
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   1200
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox txtURL 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Text            =   """google.com"""
      Top             =   600
      Width           =   4695
   End
   Begin VB.Label lblURL2 
      Caption         =   "click to add apple.com"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label lblURL1 
      Caption         =   "Click to add msn.com"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1935
   End
End
Attribute VB_Name = "frmURLinput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim txtURLinput As String
Private Sub Option1_Click()
txtURL = msn.com
End Sub

Private Sub Option2_Click()
txtURL = apple.com
End Sub
