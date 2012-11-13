VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H000000FF&
   Caption         =   "Dyna Glides"
   ClientHeight    =   8745
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9720
   LinkTopic       =   "Form2"
   ScaleHeight     =   8745
   ScaleWidth      =   9720
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture4 
      Height          =   1935
      Left            =   360
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   1875
      ScaleWidth      =   3075
      TabIndex        =   8
      Top             =   840
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "Click to return to the main menu"
      Height          =   1575
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3600
      Width           =   1095
   End
   Begin VB.PictureBox Picture3 
      Height          =   2055
      Left            =   6000
      Picture         =   "Form2.frx":5ED0
      ScaleHeight     =   1995
      ScaleWidth      =   3075
      TabIndex        =   2
      Top             =   5520
      Width           =   3135
   End
   Begin VB.PictureBox Picture2 
      Height          =   1935
      Left            =   480
      Picture         =   "Form2.frx":B54B
      ScaleHeight     =   1875
      ScaleWidth      =   3195
      TabIndex        =   1
      Top             =   5640
      Width           =   3255
   End
   Begin VB.PictureBox Picture1 
      Height          =   2055
      Left            =   6000
      Picture         =   "Form2.frx":116B6
      ScaleHeight     =   1995
      ScaleWidth      =   3075
      TabIndex        =   0
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label Label4 
      Caption         =   "FXDWG/FXDWGI Dyna Wide Glide"
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   7560
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "FXDX/FXDXI Dyna Super Glide Sport"
      Height          =   375
      Left            =   6000
      TabIndex        =   5
      Top             =   7680
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "FXDL/FXDLI Dyna Low Rider"
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      Top             =   3000
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "FXD/FXDI Dyna Super Glide"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2880
      Width           =   3135
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Show
Form2.Hide
End Sub

