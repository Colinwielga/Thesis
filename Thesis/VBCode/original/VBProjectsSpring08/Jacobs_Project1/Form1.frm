VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H000000C0&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   FillColor       =   &H000000C0&
   ForeColor       =   &H000000C0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBegin 
      Caption         =   "Begin"
      Height          =   1575
      Left            =   1800
      TabIndex        =   2
      Top             =   5040
      Width           =   4095
   End
   Begin VB.PictureBox picFighter 
      BackColor       =   &H000000C0&
      Height          =   2655
      Left            =   2160
      ScaleHeight     =   2595
      ScaleWidth      =   2715
      TabIndex        =   1
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackColor       =   &H000000C0&
      Caption         =   $"Form1.frx":0000
      Height          =   1095
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   5055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBegin_Click()
'This gets the players ready for the game.'
MsgBox ("Welcome to the Dungeon")
Form1.Hide
Form2.Show
End Sub

Private Sub Form_Load()
'To give the players a look at their character'
picFighter.Picture = LoadPicture("fighter.jpg")
End Sub

