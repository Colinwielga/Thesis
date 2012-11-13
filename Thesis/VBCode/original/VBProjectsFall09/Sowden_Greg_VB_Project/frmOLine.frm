VERSION 5.00
Begin VB.Form frmOLine 
   BackColor       =   &H0000C000&
   Caption         =   "O-Line"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Go Back to Positions"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4200
      Width           =   3015
   End
   Begin VB.CommandButton cmdLG 
      BackColor       =   &H0080FFFF&
      Caption         =   "Left Guard"
      Height          =   2055
      Left            =   3480
      Picture         =   "frmOLine.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton cmdRG 
      BackColor       =   &H0080FFFF&
      Caption         =   "Right Guard"
      Height          =   2055
      Left            =   8520
      Picture         =   "frmOLine.frx":117B
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   2055
   End
   Begin VB.CommandButton cmdLT 
      BackColor       =   &H0080FFFF&
      Caption         =   "Left Tackle"
      Height          =   2055
      Left            =   840
      Picture         =   "frmOLine.frx":22F6
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton cmdRT 
      BackColor       =   &H0080FFFF&
      Caption         =   "Right Tackle"
      Height          =   2055
      Left            =   11160
      Picture         =   "frmOLine.frx":3471
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      Width           =   2055
   End
   Begin VB.CommandButton cmdCenter 
      BackColor       =   &H0080FFFF&
      Caption         =   "Center"
      Height          =   2055
      Left            =   6000
      Picture         =   "frmOLine.frx":45EC
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1320
      Width           =   2055
   End
End
Attribute VB_Name = "frmOLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack9_Click()
    frmOLine.Hide
    frmLearn.Show

End Sub

Private Sub cmdCenter_Click()
    frmOLine.Hide
    frmCenter.Show
End Sub

Private Sub cmdLG_Click()
    frmOLine.Hide
    frmLG.Show
End Sub

Private Sub cmdLT_Click()
    frmOLine.Hide
    frmLT.Show
End Sub

Private Sub cmdRT_Click()
    frmOLine.Hide
    frmRT.Show
End Sub

Private Sub cmdRG_Click()
    frmOLine.Hide
    frmRG.Show
End Sub

