VERSION 5.00
Begin VB.Form frmAlice 
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
   Begin VB.CommandButton cmdReselect 
      BackColor       =   &H80000015&
      Caption         =   "Reselect Character..."
      Height          =   800
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7320
      Width           =   2500
   End
   Begin VB.CommandButton cmdConfirm 
      BackColor       =   &H80000015&
      Caption         =   "Enter Cave..."
      Height          =   800
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7320
      Width           =   2500
   End
   Begin VB.PictureBox picAlice 
      Height          =   2500
      Left            =   3720
      ScaleHeight     =   2445
      ScaleWidth      =   2940
      TabIndex        =   0
      Top             =   2400
      Width           =   3000
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800080&
      Caption         =   $"fmrAlice.frx":0000
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
      Left            =   7200
      TabIndex        =   3
      Top             =   3840
      Width           =   5655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800080&
      Caption         =   "Difficulty:  Hard"
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
      Left            =   7200
      TabIndex        =   2
      Top             =   2880
      Width           =   6135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800080&
      Caption         =   "You've selected:  Alice"
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
      Left            =   7200
      TabIndex        =   1
      Top             =   1920
      Width           =   5655
   End
End
Attribute VB_Name = "frmAlice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Name:  Super Awesome Cave Adventure Game
'Form Name:  frmAlice
'Author:  Peter Woodruff
'Date Written:  3-23-09
'Purpose:  This allows the character to confirm his character selection or to return to the character select screen.  It also loads
'his or her life for the rest of the game.
Option Explicit

Private Sub cmdConfirm_Click()

    'Load Life
    Life = 4
    
    'Start game
    frmAlice.Visible = False
    frmRoom1.Visible = True
    
End Sub

Private Sub cmdReselect_Click()

    'Goes back to character select screen
    frmAlice.Visible = False
    frmCharacter.Visible = True
    
End Sub

Private Sub Form_Load()

    'Load picture
    picAlice.Picture = LoadPicture(App.Path & "\AW.jpg")
    
End Sub

