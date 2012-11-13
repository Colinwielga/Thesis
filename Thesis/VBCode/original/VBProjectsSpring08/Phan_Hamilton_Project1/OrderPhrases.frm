VERSION 5.00
Begin VB.Form Ordering_Phrases 
   Caption         =   "Form1"
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   ScaleHeight     =   5250
   ScaleWidth      =   7920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSameOrder 
      Caption         =   """I'll have the same."""
      Height          =   855
      Left            =   3360
      TabIndex        =   5
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton cmdPickyEater 
      Caption         =   """I'm kind of a picky eater."""
      Height          =   855
      Left            =   3480
      TabIndex        =   4
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton cmdNeverBeen 
      Caption         =   """I've never been here before."""
      Height          =   855
      Left            =   5400
      TabIndex        =   3
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton cmdHighRecommend 
      Caption         =   """I highly recommend it."""
      Height          =   855
      Left            =   5280
      TabIndex        =   2
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Comments"
      Height          =   855
      Left            =   1200
      TabIndex        =   1
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton cmdConfused 
      Caption         =   """I haven't made up my mind yet.  Everything looks so good."""
      Height          =   855
      Left            =   3240
      TabIndex        =   0
      Top             =   2280
      Width           =   1815
   End
End
Attribute VB_Name = "Ordering_Phrases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdReturn_Click()

Ordering_Phrases.Hide
Comment_Phrases.Show

End Sub

Private Sub Form_Load()

End Sub
