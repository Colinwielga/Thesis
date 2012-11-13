VERSION 5.00
Begin VB.Form frmCharacter 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Pick your character's picture"
   ClientHeight    =   8055
   ClientLeft      =   5265
   ClientTop       =   2280
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   ScaleHeight     =   8055
   ScaleWidth      =   4830
   Begin VB.CommandButton cmdAriel 
      BackColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   3240
      Picture         =   "frmCharacter.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Ariel"
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdSnowWhite 
      BackColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   1680
      Picture         =   "frmCharacter.frx":08C9
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Snow White"
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton cmdCinderella 
      BackColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   3240
      Picture         =   "frmCharacter.frx":1197
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Cinderella"
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton cmdDonald 
      BackColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   3240
      Picture         =   "frmCharacter.frx":1C10
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Donald Duck"
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton cmdLaunchPad 
      BackColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   1680
      Picture         =   "frmCharacter.frx":2C80
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Launchpad"
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdBelle 
      BackColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   120
      Picture         =   "frmCharacter.frx":37ED
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Belle"
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton cmdScrooge 
      BackColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   120
      Picture         =   "frmCharacter.frx":417E
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Scrooge McDuck"
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton cmdGenie 
      BackColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   1680
      Picture         =   "frmCharacter.frx":4D54
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Genie"
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton cmdSimba 
      BackColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   120
      Picture         =   "frmCharacter.frx":56EF
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Simba"
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label lblCharacter 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Please click on the character you would like to be."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmCharacter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdAriel_Click()
    frmGameboard.picPlayer.Picture = LoadPicture(App.Path & "\Pictures\ariel.jpg")
    frmCharacter.Hide
    frmGameboard.Show
End Sub
Private Sub cmdBelle_Click()
    frmGameboard.picPlayer.Picture = LoadPicture(App.Path & "\Pictures\belle.jpg")
    frmCharacter.Hide
    frmGameboard.Show
End Sub
Private Sub cmdCinderella_Click()
    frmGameboard.picPlayer.Picture = LoadPicture(App.Path & "\Pictures\cind.jpg")
    frmCharacter.Hide
    frmGameboard.Show
End Sub
Private Sub cmdDonald_Click()
    frmGameboard.picPlayer.Picture = LoadPicture(App.Path & "\Pictures\donald.jpg")
    frmCharacter.Hide
    frmGameboard.Show
End Sub
Private Sub cmdGenie_Click()
    frmGameboard.picPlayer.Picture = LoadPicture(App.Path & "\Pictures\genie.jpg")
    frmCharacter.Hide
    frmGameboard.Show
End Sub
Private Sub cmdLaunchPad_Click()
    frmGameboard.picPlayer.Picture = LoadPicture(App.Path & "\Pictures\launch.jpg")
    frmCharacter.Hide
    frmGameboard.Show
End Sub
Private Sub cmdScrooge_Click()
    frmGameboard.picPlayer.Picture = LoadPicture(App.Path & "\Pictures\scrooge.jpg")
    frmCharacter.Hide
    frmGameboard.Show
End Sub
Private Sub cmdSimba_Click()
frmGameboard.picPlayer.Picture = LoadPicture(App.Path & "\Pictures\simba.jpg")
    frmCharacter.Hide
    frmGameboard.Show
End Sub
Private Sub cmdSnowWhite_Click()
    frmGameboard.picPlayer.Picture = LoadPicture(App.Path & "\Pictures\snowwhite.jpg")
    frmCharacter.Hide
    frmGameboard.Show
End Sub
