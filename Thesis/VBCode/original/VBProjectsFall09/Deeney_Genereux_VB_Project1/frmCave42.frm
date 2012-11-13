VERSION 5.00
Begin VB.Form frmCave42 
   BackColor       =   &H00000000&
   Caption         =   "Cave 4"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   7005
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdEnd 
      Caption         =   "Quit"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00000000&
      Caption         =   "Search Other Caves!"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   5040
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.PictureBox pic1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   120
      ScaleHeight     =   2415
      ScaleWidth      =   1575
      TabIndex        =   2
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Image img1 
      Height          =   4980
      Left            =   3000
      Picture         =   "frmCave42.frx":0000
      Top             =   840
      Visible         =   0   'False
      Width           =   2535
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
      TabIndex        =   1
      Top             =   0
      Width           =   6615
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Oh Look! You found a Kung Fu Master!"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   6615
   End
End
Attribute VB_Name = "frmCave42"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Katie Deeney & Elise Generex
'Create a Story
'Date Done: 10/10/2009
'Cave4
'In this cave, the user has to light up the cave
'The user found the kung fu master
'The user then puts the kung fu master into their inventory
Option Explicit

Private Sub cmdBack_Click()
    pic1.Cls
    pic1.BackColor = &H0&
    frmCave42.BackColor = &H0&
    lbl1.Visible = True
    lbl2.Visible = False
    img1.Visible = False
    cmdBack.Visible = False
    frmCave42.Hide
    frmSavePrincess2.Show
End Sub

Private Sub CmdEnd_Click()
    End
End Sub

Private Sub pic1_Click()
    pic1.Picture = LoadPicture(App.Path & "\torch.jpg")
    pic1.BackColor = &HFFFF&
    frmCave42.BackColor = &HFFFF&
    lbl1.Visible = False
    lbl2.Visible = True
    img1.Visible = True
    cmdBack.Visible = True
    Inventory = Inventory + " Kung Fu,"
End Sub
