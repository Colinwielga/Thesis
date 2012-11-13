VERSION 5.00
Begin VB.Form frmSTARTUP 
   BackColor       =   &H80000007&
   Caption         =   "::PICK A BRAND::"
   ClientHeight    =   6915
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   ScaleHeight     =   6915
   ScaleWidth      =   5610
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H000000FF&
      Caption         =   "EXIT"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   4800
      Width           =   1095
   End
   Begin VB.PictureBox picAdidas 
      Height          =   2895
      Left            =   1200
      Picture         =   "FRMSTA~1.frx":0000
      ScaleHeight     =   2835
      ScaleWidth      =   2955
      TabIndex        =   1
      Top             =   1440
      Width           =   3015
   End
   Begin VB.PictureBox picNike 
      Height          =   735
      Left            =   240
      Picture         =   "FRMSTA~1.frx":1769
      ScaleHeight     =   675
      ScaleWidth      =   4995
      TabIndex        =   0
      Top             =   240
      Width           =   5055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"FRMSTA~1.frx":2B8E
      ForeColor       =   &H8000000E&
      Height          =   1335
      Left            =   1680
      TabIndex        =   3
      Top             =   4800
      Width           =   3135
   End
End
Attribute VB_Name = "frmSTARTUP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExit_Click()
    End
End Sub

Private Sub Form_Load()
    MsgBox "::PLEASE CLICK AN ICON AFTER CLICKING OK::", , "::CLICK AND ICON::"
End Sub

Private Sub picAdidas_Click()
    frmSTARTUP.Hide
    frmAdidas.Show
End Sub

Private Sub picNike_Click()
    frmSTARTUP.Hide
    frmNike.Show
    
End Sub
