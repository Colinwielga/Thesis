VERSION 5.00
Begin VB.Form frmWinning 
   Caption         =   "THESE PEOPLE WON AND SO CAN YOU!"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   8445
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to Calculations!"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Label lblEngland 
      Caption         =   $"frmWinning.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   4080
      TabIndex        =   4
      Top             =   3240
      Width           =   4215
   End
   Begin VB.Label lblGood 
      Caption         =   """Sports Calculator is good!- Mary, Louisville"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   3
      Top             =   2400
      Width           =   4095
   End
   Begin VB.Label lblSteve 
      Caption         =   $"frmWinning.frx":015E
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   3615
   End
   Begin VB.Label lblJohn 
      Caption         =   $"frmWinning.frx":0296
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   7815
   End
End
Attribute VB_Name = "frmWinning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click() 'returns to frmSecond
    frmWinning.Hide
    frmSecond.Show
End Sub
