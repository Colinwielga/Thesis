VERSION 5.00
Begin VB.Form frmpics 
   BackColor       =   &H0080FF80&
   Caption         =   "Tennis"
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10365
   LinkTopic       =   "Form1"
   ScaleHeight     =   7740
   ScaleWidth      =   10365
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdforward 
      Caption         =   "Next Page"
      Height          =   615
      Left            =   5520
      TabIndex        =   6
      Top             =   5880
      Width           =   1575
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Go back"
      Height          =   615
      Left            =   3000
      TabIndex        =   5
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton cmdend 
      Caption         =   "Quit"
      Height          =   615
      Left            =   4320
      TabIndex        =   4
      Top             =   6840
      Width           =   1455
   End
   Begin VB.PictureBox picrog 
      Height          =   3135
      Left            =   5280
      Picture         =   "frmpics.frx":0000
      ScaleHeight     =   3075
      ScaleWidth      =   3195
      TabIndex        =   2
      Top             =   2040
      Width           =   3255
   End
   Begin VB.PictureBox picpete 
      Height          =   3255
      Left            =   1200
      Picture         =   "frmpics.frx":2895
      ScaleHeight     =   3195
      ScaleWidth      =   2835
      TabIndex        =   0
      Top             =   2040
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF80&
      Caption         =   "However current world #1 Roger Federer is on track to over take Pete's record. (he has 10)"
      Height          =   615
      Left            =   5400
      TabIndex        =   3
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "Pete Sampras has the most Grand Slam wins of any player in history. (14)"
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   1440
      Width           =   2895
   End
End
Attribute VB_Name = "frmpics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'navigate the pages
Private Sub cmdback_Click()
frmpics.Visible = False
frmtitles.Visible = True


End Sub

Private Sub cmdend_Click()
End
End Sub
'navigate the pages
Private Sub cmdforward_Click()
frmpics.Visible = False
frmsort.Visible = True

End Sub
