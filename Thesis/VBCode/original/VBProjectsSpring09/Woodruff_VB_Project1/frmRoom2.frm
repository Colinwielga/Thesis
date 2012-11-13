VERSION 5.00
Begin VB.Form frmRoom2 
   BackColor       =   &H80000017&
   Caption         =   "Form1"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14880
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   14880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDontPickUpKey 
      BackColor       =   &H80000015&
      Caption         =   "Don't Pick Up Key"
      Height          =   800
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4680
      Width           =   2500
   End
   Begin VB.CommandButton cmdPickUpKey 
      BackColor       =   &H80000015&
      Caption         =   "Pick Up Key"
      Height          =   800
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3720
      Width           =   2500
   End
   Begin VB.CommandButton cmdLadder 
      BackColor       =   &H80000015&
      Caption         =   "Try Ladder"
      Height          =   800
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6000
      Width           =   2500
   End
   Begin VB.CommandButton cmdRight 
      BackColor       =   &H80000015&
      Caption         =   "To Dark Room"
      Height          =   800
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3720
      Width           =   2500
   End
   Begin VB.CommandButton cmdLeft 
      BackColor       =   &H80000015&
      Caption         =   "To Entrance"
      Height          =   800
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4800
      Width           =   2500
   End
   Begin VB.PictureBox picRoom2 
      Height          =   7335
      Left            =   2760
      ScaleHeight     =   7275
      ScaleWidth      =   9195
      TabIndex        =   0
      Top             =   360
      Width           =   9255
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000017&
      Caption         =   "Action"
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
      Height          =   735
      Left            =   12240
      TabIndex        =   8
      Top             =   3000
      Width           =   2535
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
      TabIndex        =   7
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000017&
      Caption         =   $"frmRoom2.frx":0000
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
      Height          =   1455
      Left            =   2760
      TabIndex        =   1
      Top             =   7800
      Width           =   9255
   End
End
Attribute VB_Name = "frmRoom2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Name:  Super Awesome Cave Adventure Game
'Form Name:  frmRoom2
'Author:  Peter Woodruff
'Date Written:  3-15-09
'Purpose:  This is a 'room' of the game.  It is where the user can get the shield key.

Option Explicit

Private Sub cmdDontPickUpKey_Click()
    
    'Gives user bad advice
    MsgBox "Better not risk it, I guess.", , ""
    
End Sub

Private Sub cmdLadder_Click()
    
    'Keeps user from entering room 6 until he/she has been there
    If Ladder = False Then
        MsgBox "You can't see the bottom. Better not risk it.", , ""
    Else
        frmRoom2.Visible = False
        frmRoom6.Visible = True
        
    End If
    
End Sub

Private Sub cmdLeft_Click()
    
    'User leaves to room 1
    frmRoom2.Visible = False
    frmRoom1.Visible = True
    
End Sub

Private Sub cmdPickUpKey_Click()

    'User gets shieldkey
    ShieldKey = True
    
    'User gets hint
    MsgBox "You got the key.  It has a little shield icon on the base.", , ""
    
    'Hides useless options
    cmdPickUpKey.Visible = False
    cmdDontPickUpKey.Visible = False
    
    
End Sub

Private Sub cmdRight_Click()

    'Prevents user from entering room 9 until he/she has torch
    
    If Light = False Then
        MsgBox "It's really dark. You'd better get a light first.", , ""
    Else
        MsgBox "The torch lights up the dark room.", , ""
        frmRoom2.Visible = False
        frmRoom9.Visible = True
        
    End If
    
End Sub

Private Sub Form_Load()

    picRoom2.Picture = LoadPicture(App.Path & "\room2.bmp")
    
End Sub

