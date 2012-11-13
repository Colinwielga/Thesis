VERSION 5.00
Begin VB.Form frmquestionresult 
   Caption         =   "Your score!"
   ClientHeight    =   2070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9495
   LinkTopic       =   "Form2"
   Picture         =   "frmquestionresult.frx":0000
   ScaleHeight     =   2070
   ScaleWidth      =   9495
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "Back to Main page"
      Height          =   1095
      Left            =   7680
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton cmdshow 
      Caption         =   "Show My Score"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.PictureBox picoutput 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   2040
      ScaleHeight     =   1035
      ScaleWidth      =   5235
      TabIndex        =   0
      Top             =   360
      Width           =   5295
   End
End
Attribute VB_Name = "frmquestionresult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdback_Click()
frmquestionresult.Visible = False
frmmain.Visible = True

End Sub

Private Sub cmdshow_Click()
picoutput.Print "Your Score is:"; Score
Select Case Score
Case Is = 0
picoutput.Print "Hey! You really don't know anything about Hello Kitty!"
Case 1 To 6
picoutput.Print "You know a little about Hello Kitty. Try to know more!~"
Case 7 To 9
picoutput.Print "You are really good."
Case Is = 10
picoutput.Print "Wow! Perfect Score. You are my hero!~"
End Select


Score = 0
End Sub
