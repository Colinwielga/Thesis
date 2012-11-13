VERSION 5.00
Begin VB.Form frmCave2 
   BackColor       =   &H00000000&
   Caption         =   "Cave 2"
   ClientHeight    =   3915
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   7740
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmdquit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Search Other Caves!"
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   3360
      Visible         =   0   'False
      Width           =   1935
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
      Top             =   960
      Width           =   1575
   End
   Begin VB.Image img1 
      Height          =   1560
      Left            =   4680
      Picture         =   "frmCave2.frx":0000
      Top             =   1320
      Visible         =   0   'False
      Width           =   1485
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
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   6615
   End
   Begin VB.Label lbl2 
      BackStyle       =   0  'Transparent
      Caption         =   "A sword.....That Could be Helpful!"
      Height          =   735
      Left            =   720
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   6495
   End
End
Attribute VB_Name = "frmCave2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Katie Deeney & Elise Generex
'Create a Story
'Date Done: 10/10/2009
'Cave2
'In this cave, the user has to light up the cave
'adds the sword into
'the inventory
Private Sub cmdBack_Click()
    pic1.Cls
    pic1.BackColor = &H0&
    frmCave2.BackColor = &H0&
    lbl1.Visible = True
    lbl2.Visible = False
    img1.Visible = False
    cmdBack.Visible = False
    frmCave2.Hide
    frmCaves1.Show
End Sub

Private Sub CmdQuit_Click()
    End
End Sub

Private Sub pic1_Click()
    pic1.Picture = LoadPicture(App.Path & "\torch.jpg")
    pic1.BackColor = &HFFFF&
    frmCave2.BackColor = &HFFFF&
    lbl1.Visible = False
    lbl2.Visible = True
    img1.Visible = True
    cmdBack.Visible = True
    Inventory = Inventory + " Sword,"
End Sub
