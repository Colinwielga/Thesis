VERSION 5.00
Begin VB.Form Winner 
   BackColor       =   &H00000000&
   Caption         =   "Winner"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   Picture         =   "Winner.frx":0000
   ScaleHeight     =   6150
   ScaleWidth      =   9315
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Quit"
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   $"Winner.frx":7D87
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2055
      Left            =   1200
      TabIndex        =   0
      Top             =   2400
      Width           =   7095
   End
End
Attribute VB_Name = "Winner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Defuse by Ryan Maurer and Joel Abel 11/06
Private Sub Command1_Click()
    End
End Sub

