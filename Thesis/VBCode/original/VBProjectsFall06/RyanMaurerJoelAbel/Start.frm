VERSION 5.00
Begin VB.Form Start 
   Caption         =   "Start"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   Picture         =   "Start.frx":0000
   ScaleHeight     =   7995
   ScaleWidth      =   10845
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Back Out!"
      Height          =   975
      Left            =   6720
      TabIndex        =   2
      Top             =   1320
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Try to Save the Day!"
      Height          =   855
      Left            =   1080
      MaskColor       =   &H000000FF&
      TabIndex        =   1
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DEFUSE!"
      BeginProperty Font 
         Name            =   "Niagara Solid"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1575
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   9255
   End
End
Attribute VB_Name = "Start"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Defuse by Ryan Maurer and Joel Abel 11/06

Private Sub Command1_Click() 'loads the game form
    Defuse.Show
    Start.Hide
End Sub

Private Sub Command2_Click()
    End
End Sub


