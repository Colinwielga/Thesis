VERSION 5.00
Begin VB.Form frmRANKING 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13200
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   13200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0000FF00&
      Caption         =   "QUIT"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7440
      Width           =   1935
   End
   Begin VB.CommandButton cmdRanking 
      BackColor       =   &H00FFFF00&
      Caption         =   "Generate Your Ranking!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5880
      Width           =   6015
   End
   Begin VB.PictureBox picResults1 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   720
      ScaleHeight     =   4395
      ScaleWidth      =   11715
      TabIndex        =   0
      Top             =   360
      Width           =   11775
   End
End
Attribute VB_Name = "frmRANKING"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form gives the user their ranking


Private Sub cmdQuit_Click() 'ends the program

    End
End Sub

Private Sub cmdRanking_Click()

picResults1.Print playername; ", you answered "; ctr; "questions out of "; total; "correctly." 'this prints the # of correct answers and the total # of questions the user completed.

If ctr / total >= 0.9 Then          'this prints out a message about the quality of the work
    picResults1.Print "Could you be a music prodigy?"
ElseIf ctr / total >= 0.8 Then
    picResults1.Print "A little more practice please..."
ElseIf ctr / total >= 0.5 Then
    picResults1.Print "There might be hope for you... maybe."
ElseIf ctr / total < 0.5 Then
    picResults1.Print "Give up now."
End If

End Sub


