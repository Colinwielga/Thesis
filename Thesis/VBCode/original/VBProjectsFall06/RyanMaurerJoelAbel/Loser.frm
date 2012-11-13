VERSION 5.00
Begin VB.Form Loser 
   Caption         =   "Loser"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11445
   LinkTopic       =   "Form1"
   Picture         =   "Loser.frx":0000
   ScaleHeight     =   8265
   ScaleWidth      =   11445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton quit 
      Caption         =   "Quit"
      Height          =   855
      Left            =   360
      TabIndex        =   1
      Top             =   6960
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "You failed to defuse the bomb, and now sadly your family must attend a closed casket funeral."
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   11535
   End
End
Attribute VB_Name = "Loser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Defuse by Ryan Maurer and Joel Abel 11/06
Private Sub quit_Click()
    End
End Sub
