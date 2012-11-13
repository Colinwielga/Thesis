VERSION 5.00
Begin VB.Form frmmain 
   Caption         =   "Main Menu"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   Picture         =   "frmmain.frx":0000
   ScaleHeight     =   5835
   ScaleWidth      =   8805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H00C000C0&
      Caption         =   "Exit"
      Height          =   855
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "When you're ready to leave!"
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton cmdbuy 
      Caption         =   "Prizes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6960
      Picture         =   "frmmain.frx":0EE0
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Buy Buy Buy!"
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton cmddoor 
      Caption         =   "Pick a Door"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   2160
      Picture         =   "frmmain.frx":220E
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Try your luck!"
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton cmdtrivia 
      Caption         =   "Trivia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   480
      Picture         =   "frmmain.frx":2D3B
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Test your knowledge!"
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton cmdrace 
      Caption         =   "Off to the Races!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   4800
      Picture         =   "frmmain.frx":3447
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Bet on animal races!"
      Top             =   3600
      Width           =   1575
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    'A Day For Fun
    'Main
    'Stephanie Fiecke
    '10-25-06
    'This form is the main form for the user to be able to click into different forms
    'that contain different types of games.
    
Option Explicit

Private Sub cmdbuy_Click()
    'Hides the main form and shows the buy form
frmmain.Hide
frmbuy.Show
End Sub

Private Sub cmddoor_Click()
    'Hides the main form and shows the door form
frmmain.Hide
frmdoor.Show
End Sub

    
Private Sub cmdmatch_Click()
    'Hides the Main Form and brings up the matching game form
frmmain.Hide
frmmatch.Show
End Sub

Private Sub cmdquit_Click()
    'exits the program all together
End
End Sub

Private Sub cmdrace_Click()
    'hides the main form and shows the racing form
frmmain.Hide
frmrace.Show
End Sub

Private Sub cmdtrivia_Click()
    'hides the main form and shows the trivia game form
frmmain.Hide
frmtrivia.Show
End Sub

