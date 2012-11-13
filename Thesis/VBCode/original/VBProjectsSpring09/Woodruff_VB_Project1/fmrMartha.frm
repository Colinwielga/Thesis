VERSION 5.00
Begin VB.Form frmMartha 
   BackColor       =   &H00800080&
   Caption         =   "Form1"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14880
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   14880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdConfirm 
      BackColor       =   &H80000015&
      Caption         =   "Enter Cave..."
      Height          =   800
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7800
      Width           =   2500
   End
   Begin VB.CommandButton cmdReselect 
      BackColor       =   &H80000015&
      Caption         =   "Reselect Character..."
      Height          =   800
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7800
      Width           =   2500
   End
   Begin VB.PictureBox picMartha 
      Height          =   2500
      Left            =   3720
      ScaleHeight     =   2445
      ScaleWidth      =   2940
      TabIndex        =   3
      Top             =   2520
      Width           =   3000
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800080&
      Caption         =   "You've selected:  Martha Steward"
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
      Height          =   615
      Index           =   0
      Left            =   7560
      TabIndex        =   2
      Top             =   2640
      Width           =   5655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800080&
      Caption         =   "Difficulty:  A little challenging..."
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
      Height          =   615
      Index           =   1
      Left            =   7560
      TabIndex        =   1
      Top             =   3600
      Width           =   6135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800080&
      Caption         =   $"fmrMartha.frx":0000
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
      Height          =   2655
      Index           =   2
      Left            =   7560
      TabIndex        =   0
      Top             =   4560
      Width           =   5655
   End
End
Attribute VB_Name = "frmMartha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Name:  Super Awesome Cave Adventure Game
'Form Name:  frmMartha
'Author:  Peter Woodruff
'Date Written:  3-23-09
'Purpose:  This allows the character to confirm his character selection or to return to the character select screen.  It also loads
'his or her life for the rest of the game.
Option Explicit

Private Sub cmdConfirm_Click()

    'Load Life
    Life = 5
    
    'Start game
    frmMartha.Visible = False
    frmRoom1.Visible = True
    
End Sub

Private Sub cmdReselect_Click()

    'Goes back to character select screen
    frmMartha.Visible = False
    frmCharacter.Visible = True
    
End Sub

Private Sub Form_Load()
    
    'load picture
    picMartha.Picture = LoadPicture(App.Path & "\MS.bmp")
    
End Sub

