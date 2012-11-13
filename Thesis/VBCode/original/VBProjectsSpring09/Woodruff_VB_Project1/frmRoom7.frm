VERSION 5.00
Begin VB.Form frmRoom7 
   BackColor       =   &H80000012&
   Caption         =   "Form1"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14880
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   14880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H80000015&
      Caption         =   "Head back to entrance."
      Height          =   800
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4320
      Width           =   2500
   End
   Begin VB.CommandButton cmdMain2 
      BackColor       =   &H80000015&
      Caption         =   "Try door."
      Height          =   800
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3360
      Width           =   2500
   End
   Begin VB.PictureBox picRoom7 
      Height          =   6015
      Left            =   3000
      ScaleHeight     =   5955
      ScaleWidth      =   11715
      TabIndex        =   0
      Top             =   600
      Width           =   11775
   End
   Begin VB.Label lblAction 
      BackColor       =   &H80000012&
      Caption         =   "Movement:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   240
      TabIndex        =   4
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label lblStoryRoom7 
      BackColor       =   &H80000012&
      Caption         =   $"frmRoom7.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   1815
      Left            =   3960
      TabIndex        =   1
      Top             =   6960
      Width           =   9135
   End
End
Attribute VB_Name = "frmRoom7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Name:  Super Awesome Cave Adventure Game
'Form Name:  frmRoom7
'Author:  Peter Woodruff
'Date Written:  3-15-09
'Purpose:  This is a 'room' of the game.  Nothing happens here, but it provides
'a secret to answering a puzzle.

Option Explicit

Private Sub cmdBack_Click()

    'Moves user to room 1
    frmRoom7.Visible = False
    frmRoom1.Visible = True
    
End Sub

Private Sub cmdMain2_Click()

    'If the user has the sword key, the door opens and he/she moves
    'to room 8
    If SwordKey = False Then
        MsgBox "Big surprise, it's locked.  The keyhole has a sword emblem right above it.", , ""

    Else
        MsgBox "You use the key with the sword icon.  It works!", , ""
        frmRoom7.Visible = False
        frmRoom8.Visible = True
        
    End If
    
End Sub

Private Sub Form_Load()


    picRoom7.Picture = LoadPicture(App.Path & "\creepyheads.jpg")

End Sub

