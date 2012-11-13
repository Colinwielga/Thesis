VERSION 5.00
Begin VB.Form frmBox6 
   Caption         =   "Box Score From Game 6"
   ClientHeight    =   4170
   ClientLeft      =   3765
   ClientTop       =   3930
   ClientWidth     =   5880
   LinkTopic       =   "Form3"
   ScaleHeight     =   4170
   ScaleWidth      =   5880
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next Game"
      Height          =   615
      Left            =   4320
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "Previous Game"
      Height          =   615
      Left            =   4320
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return To Main Page"
      Height          =   615
      Left            =   4320
      TabIndex        =   0
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label lblGameNumber 
      BackColor       =   &H80000012&
      Caption         =   "Game 6"
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   615
   End
   Begin VB.Image imgBox6 
      Height          =   5760
      Left            =   -600
      Picture         =   "frmBox6.frx":0000
      Top             =   -120
      Width           =   7680
   End
End
Attribute VB_Name = "frmBox6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: 1987 World Series
'Form name: frmBox6
'Authors: Hans Paul and Cole Wuollet
'Date Written: Wednesday November 1, 2006
'Objective: To display the Box Score from Game 6 of the 1987 World Series
Option Explicit

Private Sub cmdNext_Click() 'Hides Current Form and Goes to Next Form
    frmBox6.Hide
    frmBox7.Show
End Sub

Private Sub cmdPrevious_Click() 'Hides Current Form and Goes to Previous Form
    frmBox6.Hide
    frmBox5.Show
End Sub

Private Sub cmdReturn_Click() 'Hides Current Form
    frmBox6.Hide
End Sub
