VERSION 5.00
Begin VB.Form frmFightStone 
   BackColor       =   &H00000000&
   Caption         =   "Fight the Dragon with Stones!"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   7710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmdquit 
      Caption         =   "Quit"
      Height          =   195
      Left            =   5760
      TabIndex        =   2
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton cmdFound 
      Caption         =   "Found them All!"
      Height          =   735
      Left            =   5520
      Picture         =   "frmFightStone.frx":0000
      TabIndex        =   1
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Image img5 
      Height          =   2115
      Index           =   1
      Left            =   4560
      Picture         =   "frmFightStone.frx":1A67
      Top             =   960
      Width           =   1755
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   2295
      Left            =   2400
      Shape           =   2  'Oval
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Image img3 
      Height          =   2115
      Index           =   1
      Left            =   2160
      Picture         =   "frmFightStone.frx":213B
      Top             =   2400
      Width           =   1755
   End
   Begin VB.Image img2 
      Height          =   2115
      Index           =   1
      Left            =   5280
      Picture         =   "frmFightStone.frx":280F
      Top             =   3120
      Width           =   1755
   End
   Begin VB.Image img1 
      Height          =   2115
      Index           =   0
      Left            =   600
      Picture         =   "frmFightStone.frx":2EE3
      Top             =   840
      Width           =   1755
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmFightStone.frx":35B7
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7575
   End
   Begin VB.Image img4 
      Height          =   2115
      Index           =   1
      Left            =   720
      Picture         =   "frmFightStone.frx":3699
      Top             =   1200
      Width           =   1755
   End
End
Attribute VB_Name = "frmFightStone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim J As Integer
'Katie Deeney & Elise Generex
'Create a Story
'Date Done: 10/10/2009
'This is a form where the user has to fight the dragon with stones
'You have to first find the stones by clicking around
'hit ok if you think you found all the stones
'If you did not find them all then you die
'If you found them all, you killed the dragon


Private Sub cmdFound_Click()
     If J = 5 Then
        MsgBox "You found them all and threw them at the dragon! You succed in killing the dragon! You saved the princess!!!!!!!", , "Nice Job!"
    Else
        MsgBox "Unfortunately, you didn't find all of the toads, so you did not slay the dragon, and you are now captured by the dragon annd have to wait forever and ever with him", , "Oh no!"
    End If
    img1(0).Visible = True
    img2(1).Visible = True
    img3(1).Visible = True
    img4(1).Visible = True
    img5(1).Visible = True
    MsgBox "This is where your story ends. Start Over.", , "Story Ends"
    frmFightStone.Hide
    frmWelcome.Show
End Sub

Private Sub Cmdquit_Click()
    End
End Sub

Private Sub img1_Click(Index As Integer)
    J = J + 1
    img1(0).Visible = False
End Sub

Private Sub img2_Click(Index As Integer)
    J = J + 1
    img2(1).Visible = False
End Sub

Private Sub img3_Click(Index As Integer)
    J = J + 1
    img3(1).Visible = False
End Sub

Private Sub img4_Click(Index As Integer)
    J = J + 1
    img4(1).Visible = False
End Sub

Private Sub img5_Click(Index As Integer)
    J = J + 1
    img5(1).Visible = False
End Sub
