VERSION 5.00
Begin VB.Form frmCave32 
   BackColor       =   &H00000000&
   Caption         =   "Cave 3"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10020
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   10020
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdQuit 
      Caption         =   "Quit"
      Height          =   255
      Left            =   6840
      TabIndex        =   4
      Top             =   5880
      Width           =   2655
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next!"
      Height          =   735
      Left            =   6840
      TabIndex        =   2
      Top             =   5040
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.PictureBox pic1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   360
      ScaleHeight     =   2415
      ScaleWidth      =   1575
      TabIndex        =   1
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Oh No! It is too dark to see!  Search around and find the torch! (Click around)"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   1215
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6615
   End
   Begin VB.Image img2 
      Height          =   2655
      Left            =   6840
      Picture         =   "frmCave32.frx":0000
      Top             =   2280
      Visible         =   0   'False
      Width           =   1830
   End
   Begin VB.Image img1 
      Height          =   3405
      Left            =   2400
      Picture         =   "frmCave32.frx":16B5
      Top             =   1560
      Visible         =   0   'False
      Width           =   4500
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "You Found The Dragon and The Princess!!!!!"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6855
   End
End
Attribute VB_Name = "frmCave32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Katie Deeney & Elise Generex
'Create a Story
'Date Done: 10/10/2009
'Cave32
'In this cave, the user has to light up the cave
'The user found the princess and the dragon!
'The user is then told to find out what happens next! on the next form

Private Sub CmdNext_Click()
    pic1.Cls
    pic1.BackColor = &H0&
    frmCave32.BackColor = &H0&
    lbl1.Visible = True
    lbl2.Visible = False
    img1.Visible = False
    cmdNext.Visible = False
    img2.Visible = False
    frmCave32.Hide
    frmFight2.Show
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub pic1_Click()
    pic1.Picture = LoadPicture(App.Path & "\torch.jpg")
    pic1.BackColor = &HFFFF&
    frmCave32.BackColor = &HFFFF&
    lbl1.Visible = False
    lbl2.Visible = True
    img1.Visible = True
    img2.Visible = True
    cmdNext.Visible = True
End Sub
