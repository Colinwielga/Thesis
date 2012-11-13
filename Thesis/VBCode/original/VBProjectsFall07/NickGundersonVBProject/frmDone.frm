VERSION 5.00
Begin VB.Form frmDone 
   Caption         =   "Winner"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   Picture         =   "frmDone.frx":0000
   ScaleHeight     =   5985
   ScaleWidth      =   9705
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEnd 
      Caption         =   "End"
      Height          =   855
      Left            =   7680
      TabIndex        =   2
      Top             =   5160
      Width           =   2055
   End
   Begin VB.CommandButton cmdWinner 
      Appearance      =   0  'Flat
      Caption         =   "Click Here!!!!"
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   5280
      Width           =   2055
   End
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   9675
      TabIndex        =   0
      Top             =   0
      Width           =   9735
   End
End
Attribute VB_Name = "frmDone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this is your final screen and will tell you who won the matchup

Private Sub cmdEnd_Click()
End
End Sub

Private Sub cmdWinner_Click()
If Team1Points > Team2Points Then
    picResults.Print "Congats to "; Nteam1(1); "You have won in Fantasy football this week"
    'if team 1 scored more points then team two it will tell the user they won
Else
    If Team2Points > Team1Points Then
        picResults.Print "Congats to "; NTeam2(1); "You have won in Fantasy football this week"
    'if team 2 scored more points then team 1 it will tell the user team two won
    Else
        picResults.Print "There was a tie in fantasy this week,  There is no winner"
    'anything else would then be a tie
    End If
End If

End Sub


