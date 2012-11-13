VERSION 5.00
Begin VB.Form trivia 
   BackColor       =   &H000000FF&
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10275
   LinkTopic       =   "Form2"
   ScaleHeight     =   7905
   ScaleWidth      =   10275
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton end 
      Caption         =   "Home"
      Height          =   1455
      Left            =   7320
      TabIndex        =   3
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton old 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Old"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Index           =   2
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   3135
   End
   Begin VB.CommandButton Basic 
      BackColor       =   &H00FF0000&
      Caption         =   "Basic"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Index           =   1
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   3135
   End
   Begin VB.CommandButton new 
      BackColor       =   &H00FF0000&
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Index           =   0
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1800
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "Twins Trivia"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   60
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1680
      TabIndex        =   4
      Top             =   240
      Width           =   7335
   End
End
Attribute VB_Name = "trivia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)

End Sub

Private Sub end_Click()
trivia.Hide
main.Show
End Sub
