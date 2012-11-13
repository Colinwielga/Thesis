VERSION 5.00
Begin VB.Form frmFightToad 
   BackColor       =   &H00000000&
   Caption         =   "Fight the Dragon With Toads!"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8775
   BeginProperty Font 
      Name            =   "Britannic Bold"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   8775
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      Height          =   180
      Left            =   6720
      TabIndex        =   2
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton cmdFound 
      Caption         =   "Found them All!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6480
      Picture         =   "frmFightToad.frx":0000
      TabIndex        =   1
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   1095
      Left            =   2400
      Shape           =   3  'Circle
      Top             =   1200
      Width           =   855
   End
   Begin VB.Image img2 
      Height          =   510
      Left            =   2760
      Picture         =   "frmFightToad.frx":1A67
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image img4 
      Height          =   1215
      Left            =   4200
      Picture         =   "frmFightToad.frx":1EC7
      Top             =   3840
      Width           =   1380
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderWidth     =   13
      Height          =   1335
      Left            =   6360
      Top             =   1320
      Width           =   855
   End
   Begin VB.Image img5 
      Height          =   1215
      Left            =   6480
      Picture         =   "frmFightToad.frx":2706
      Top             =   1440
      Width           =   1380
   End
   Begin VB.Image img1 
      Height          =   1215
      Left            =   240
      Picture         =   "frmFightToad.frx":2F45
      Top             =   960
      Width           =   1380
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmFightToad.frx":3784
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
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   7575
   End
   Begin VB.Image img3 
      Height          =   510
      Left            =   6120
      Picture         =   "frmFightToad.frx":387C
      Top             =   2160
      Width           =   615
   End
End
Attribute VB_Name = "frmFightToad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim J As Integer
'Katie Deeney & Elise Generex
'Create a Story
'Date Done: 10/10/2009
'This is a form where the user has to fight the dragon with toads
'You have to first find the toads by clicking around
'hit ok if you think you found all the toads
'If you did not find them all then you die
'If you found them all, you killed the dragon

Private Sub cmdFound_Click()
    If J = 5 Then
        MsgBox "You found them all and threw them at the dragon! You succed in killing the dragon! You saved the princess and are cured of your warts!", , "Nice Job!"
    Else
        MsgBox "Unfortunately, you didn't find all of the toads, so you did not slay the dragon, and you are now captured by the dragon annd have to wait forever and ever with him", , "Oh no!"
    End If
    img1.Visible = True
    img2.Visible = True
    img3.Visible = True
    img4.Visible = True
    img5.Visible = True
    MsgBox "This is where your story ends. Start Over.", , "Story Ends"
    frmFightToad.Hide
    frmWelcome.Show
    
End Sub

Private Sub cmdquit_Click()
    End
End Sub

Private Sub img1_Click()
    J = J + 1
    img1.Visible = False
End Sub

Private Sub img2_Click()
    J = J + 1
    img2.Visible = False
End Sub

Private Sub img3_Click()
    J = J + 1
    img3.Visible = False
End Sub

Private Sub img4_Click()
    J = J + 1
    img4.Visible = False
End Sub

Private Sub img5_Click()
    J = J + 1
    img5.Visible = False
End Sub
