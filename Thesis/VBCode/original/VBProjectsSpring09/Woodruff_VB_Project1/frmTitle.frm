VERSION 5.00
Begin VB.Form frmTitle 
   BackColor       =   &H00800080&
   Caption         =   "Form1"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14880
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   14880
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTitle 
      Height          =   4575
      Left            =   9600
      ScaleHeight     =   4515
      ScaleWidth      =   3915
      TabIndex        =   3
      Top             =   2160
      Width           =   3975
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H80000015&
      Caption         =   "Quit"
      Height          =   800
      Left            =   8880
      MaskColor       =   &H80000015&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8280
      Width           =   2500
   End
   Begin VB.CommandButton cmdPlay 
      BackColor       =   &H80000015&
      Caption         =   "Play"
      Height          =   800
      Left            =   4200
      MaskColor       =   &H80000015&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8280
      Width           =   2500
   End
   Begin VB.Label lblGameTitle 
      BackColor       =   &H0000FF00&
      BackStyle       =   0  'Transparent
      Caption         =   "Super Awesome Cave Adventure Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9135
      Left            =   1560
      TabIndex        =   0
      Top             =   0
      Width           =   9255
   End
End
Attribute VB_Name = "frmTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Name:  Super Awesome Cave Adventure Game
'Form Name:  frmTitle
'Author:  Peter Woodruff
'Date Written:  3-15-09
'Purpose:  This is the title screen.  The user can play the game or
'quit.

Option Explicit
Dim C As Integer


Private Sub cmdPlay_Click()

    'Starts the game
    
    frmTitle.Visible = False
    frmCharacter.Visible = True
    
    

End Sub

Private Sub cmdQuit_Click()
    
    'End
    End
    
End Sub

Private Sub Command1_Click()
    frmTitle.Visible = False
    frmRoom9.Visible = True
    
End Sub

Private Sub Form_Load()
    
    'Load picture
    picTitle.Picture = LoadPicture(App.Path & "\swordshield.gif")
    
End Sub

