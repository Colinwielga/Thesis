VERSION 5.00
Begin VB.Form frmCave12 
   BackColor       =   &H00000000&
   Caption         =   "Cave 1"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   7185
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmdquit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Search Other Caves!"
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   4680
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.PictureBox pic1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   240
      ScaleHeight     =   2415
      ScaleWidth      =   1575
      TabIndex        =   2
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Oh Goody!  You found the Peanut Butter and Jelly Sandwich!"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   6495
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
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
   End
   Begin VB.Image img1 
      Height          =   2880
      Left            =   3120
      Picture         =   "frmCave12.frx":0000
      Top             =   1440
      Visible         =   0   'False
      Width           =   3750
   End
End
Attribute VB_Name = "frmCave12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Katie Deeney & Elise Generex
'Create a Story
'Date Done: 10/10/2009
'Cave1
'In this cave, the user has to light up the cave
'adds the peanut butter & jelly sandwhich into
'the inventory
Private Sub cmdBack_Click()
    pic1.Cls
    pic1.BackColor = &H0&
    frmCave12.BackColor = &H0&
    lbl1(0).Visible = True
    lbl2.Visible = False
    img1.Visible = False
    cmdBack.Visible = False
    frmCave12.Hide
    frmSavePrincess2.Show
End Sub

Private Sub Cmdquit_Click()
    End
End Sub

Private Sub pic1_Click()
    pic1.Picture = LoadPicture(App.Path & "\torch.jpg")
    pic1.BackColor = &HFFFF&
    frmCave12.BackColor = &HFFFF&
    lbl1(0).Visible = False
    lbl2.Visible = True
    img1.Visible = True
    cmdBack.Visible = True
    Inventory = Inventory + " Sandwich,"
End Sub
