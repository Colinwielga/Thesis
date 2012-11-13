VERSION 5.00
Begin VB.Form frmPublicOpinion 
   Caption         =   "What does Sports Nation think?"
   ClientHeight    =   10575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   10575
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdReturnMenu 
      BackColor       =   &H000000FF&
      DisabledPicture =   "frmPublicOpinion.frx":0000
      Height          =   1215
      Left            =   13200
      Picture         =   "frmPublicOpinion.frx":7DDC
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9120
      Width           =   1935
   End
   Begin VB.PictureBox picResultsPoll 
      Height          =   6615
      Left            =   0
      ScaleHeight     =   6555
      ScaleWidth      =   11835
      TabIndex        =   3
      Top             =   0
      Width           =   11895
   End
   Begin VB.CommandButton cmdQuestion3 
      Height          =   2055
      Left            =   12120
      Picture         =   "frmPublicOpinion.frx":F721
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4560
      Width           =   3015
   End
   Begin VB.CommandButton cmdQuestion2 
      Height          =   2055
      Left            =   12120
      Picture         =   "frmPublicOpinion.frx":1AB13
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2280
      Width           =   3015
   End
   Begin VB.CommandButton cmdQuestion1 
      Height          =   2055
      Left            =   12120
      Picture         =   "frmPublicOpinion.frx":25EDE
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   3015
   End
   Begin VB.Label lblTotalVotes 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Votes: 87,628 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   6720
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Image picRoseSlide 
      Height          =   10725
      Left            =   0
      Picture         =   "frmPublicOpinion.frx":310C7
      Top             =   0
      Width           =   15300
   End
End
Attribute VB_Name = "frmPublicOpinion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdQuestion1_Click()
    'displays public opinion results in a graph of the same questions the user just answered
    picResultsPoll.Cls
    picResultsPoll.Picture = LoadPicture(App.Path & "\" & "pollOne.jpg")
    lblTotalVotes.Visible = True
End Sub

Private Sub cmdQuestion2_Click()
    'displays public opinion results in a graph of the same questions the user just answered
    picResultsPoll.Cls
    picResultsPoll.Picture = LoadPicture(App.Path & "\" & "pollTwo.jpg")
    lblTotalVotes.Visible = True
End Sub

Private Sub cmdQuestion3_Click()
    'displays public opinion results in a graph of the same questions the user just answered
    picResultsPoll.Cls
    picResultsPoll.Picture = LoadPicture(App.Path & "\" & "pollThree.jpg")
    lblTotalVotes.Visible = True
End Sub

Private Sub cmdReturnMenu_Click()
    'returns user to previous screen (Polls) for further use
    picResultsPoll.Cls
    frmPublicOpinion.Hide
    lblTotalVotes.Visible = False
End Sub
