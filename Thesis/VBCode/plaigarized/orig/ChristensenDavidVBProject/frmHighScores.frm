VERSION 5.00
Begin VB.Form frmHighScores 
   BackColor       =   &H00000000&
   Caption         =   "High Scores"
   ClientHeight    =   1095
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   ScaleHeight     =   1095
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGoBack 
      Caption         =   "Go Back"
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.ComboBox HighScores 
      Height          =   315
      ItemData        =   "frmHighScores.frx":0000
      Left            =   3120
      List            =   "frmHighScores.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblHighScore 
      BackColor       =   &H00000000&
      Caption         =   "Select which High Scores you would like to see, then press ok."
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmHighScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This takes you back
Private Sub cmdGoBack_Click()
    frmHighScores.Hide
    frmSudoku.Show
End Sub

'This checks if you made the high scores
Private Sub cmdOk_Click()
    If HighScores.Text = "Easy" Then
        If EasyTime = 0 Then
            MsgBox "Sorry you have not completed the easy puzzle"
        Else
            frmEasyHS.Show
        End If
    ElseIf HighScores.Text = "Medium" Then
        If MediumTime = 0 Then
            MsgBox "Sorry you have not completed the medium puzzle"
        Else
            frmMediumHS.Show
        End If
    ElseIf HighScores.Text = "Hard" Then
        If HardTime = 0 Then
            MsgBox "Sorry you have not completed the hard puzzle"
        Else
            frmHardHS.Show
        End If
    End If
End Sub

