VERSION 5.00
Begin VB.Form frmRoom8 
   BackColor       =   &H80000017&
   Caption         =   "Form1"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14880
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   14880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOpen 
      BackColor       =   &H80000015&
      Caption         =   "Open it!"
      Height          =   800
      Left            =   240
      MaskColor       =   &H80000015&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5160
      Width           =   2500
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H80000015&
      Caption         =   "Go Back."
      Height          =   800
      Left            =   240
      MaskColor       =   &H80000015&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4080
      Width           =   2500
   End
   Begin VB.PictureBox picRoom8 
      Height          =   6255
      Left            =   3240
      ScaleHeight     =   6195
      ScaleWidth      =   10395
      TabIndex        =   0
      Top             =   720
      Width           =   10455
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
      Top             =   3240
      Width           =   2535
   End
   Begin VB.Label lblStoryRoom8 
      BackColor       =   &H80000017&
      Caption         =   $"frmRoom8.frx":0000
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
      Height          =   1935
      Left            =   3840
      TabIndex        =   1
      Top             =   7320
      Width           =   8535
   End
End
Attribute VB_Name = "frmRoom8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Name:  Super Awesome Cave Adventure Game
'Form Name:  frmRoom8
'Author:  Peter Woodruff
'Date Written:  3-15-09
'Purpose:  This is a 'room' of the game.  The user cannot pass unless he/she has the
'sword and secret spell.

Option Explicit

Private Sub cmdBack_Click()

    'Moves user to room 7
    frmRoom8.Visible = False
    frmRoom7.Visible = True
    
    
End Sub

Private Sub cmdOpen_Click()

    'If user has sword and sercet words then he/she can enter the final room
    If Sword = True And Secret = True Then
    
        frmRoom8.Visible = False
        frmRoom10.Visible = True
        
        MsgBox "You open the door and enter a giant chamber.  There's a huge dragon on the other side!  As the fire breathing cliché sets upon you, you notice and unlit torch, a glowy thing, and a rope.", , ""
        
    Else
        
        MsgBox "Sorry, you are not ready to enter this room.", , ""
    
    End If
    
End Sub

Private Sub Form_Load()

    picRoom8.Picture = LoadPicture(App.Path & "\bigdoor.jpg")
    
End Sub
