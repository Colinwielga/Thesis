VERSION 5.00
Begin VB.Form frmRoom1 
   BackColor       =   &H80000012&
   Caption         =   "Form1"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14880
   FontTransparent =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   14880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRight 
      BackColor       =   &H80000015&
      Caption         =   "Right Door"
      Height          =   800
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      UseMaskColor    =   -1  'True
      Width           =   2500
   End
   Begin VB.CommandButton cmdMain1 
      BackColor       =   &H80000015&
      Caption         =   "Bolted Door"
      Height          =   800
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4320
      Width           =   2500
   End
   Begin VB.CommandButton cmdLeft 
      BackColor       =   &H80000015&
      Caption         =   "Left Door"
      Height          =   800
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5520
      Width           =   2500
   End
   Begin VB.PictureBox picRoom1 
      Height          =   6855
      Left            =   2760
      ScaleHeight     =   6795
      ScaleWidth      =   11955
      TabIndex        =   0
      Top             =   480
      Width           =   12015
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
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label lblStoryRoom1 
      BackColor       =   &H80000012&
      Caption         =   $"frmRoom1.frx":0000
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
      Left            =   3480
      TabIndex        =   4
      Top             =   7440
      Width           =   9255
   End
End
Attribute VB_Name = "frmRoom1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Name:  Super Awesome Cave Adventure Game
'Form Name:  frmRoom1
'Author:  Peter Woodruff
'Date Written:  3-15-09
'Purpose:  This form is the first 'room' in the game.
Option Explicit

Private Sub cmdLeft_Click()

    'Enter Room 3
    frmRoom1.Visible = False
    frmRoom3.Visible = True
    
End Sub

Private Sub cmdMain1_Click()

    'User needs shieldkey to get to room 7
    
    If ShieldKey = False Then
        MsgBox "Nice try, it's locked.  The keyhole has a shield emblem right above it.", , ""

    Else
        MsgBox "You use the key with the shield icon.  It works!", , ""
        frmRoom1.Visible = False
        frmRoom7.Visible = True
        
    End If
    
End Sub

Private Sub cmdRight_Click()

    'Enter room 2
    
    frmRoom1.Visible = False
    frmRoom2.Visible = True
    
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Form_Load()

    'Load pic
    picRoom1.Picture = LoadPicture(App.Path & "\entranceroom.jpg")
    
End Sub

Private Sub Picture1_Click()

End Sub
