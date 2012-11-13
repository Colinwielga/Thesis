VERSION 5.00
Begin VB.Form frmBox4 
   Caption         =   "Box Score From Game 4"
   ClientHeight    =   4170
   ClientLeft      =   3765
   ClientTop       =   3930
   ClientWidth     =   5880
   LinkTopic       =   "Form5"
   ScaleHeight     =   4170
   ScaleWidth      =   5880
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next Game"
      Height          =   615
      Left            =   4320
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "Previous Game"
      Height          =   615
      Left            =   4320
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return To Main Page"
      Height          =   615
      Left            =   4320
      TabIndex        =   0
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label lblGameNumber 
      BackColor       =   &H80000012&
      Caption         =   "Game 4"
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   615
   End
   Begin VB.Image imgBox4 
      Height          =   5760
      Left            =   -360
      Picture         =   "frmBox4.frx":0000
      Top             =   -360
      Width           =   7680
   End
End
Attribute VB_Name = "frmBox4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: 1987 World Series
'Form name: frmBox4
'Authors: Hans Paul and Cole Wuollet
'Date Written: Wednesday November 1, 2006
'Objective: To display the Box Score from Game 4 of the 1987 World Series
Option Explicit

Private Sub cmdNext_Click() 'Hides Current Form and Goes to Next Form
    frmBox4.Hide
    frmBox5.Show
End Sub

Private Sub cmdPrevious_Click() 'Hides Current Form and Goes to Previous Form
    frmBox4.Hide
    frmBox3.Show
End Sub

Private Sub cmdReturn_Click() 'Hides Current Form
    frmBox4.Hide
End Sub
