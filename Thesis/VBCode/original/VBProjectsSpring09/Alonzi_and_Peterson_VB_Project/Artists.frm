VERSION 5.00
Begin VB.Form Artists 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Alright, let's see some Artists!"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   9765
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton PianoGo 
      BackColor       =   &H00FF8080&
      Caption         =   "Go back to the Piano"
      Height          =   975
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5400
      Width           =   3255
   End
   Begin VB.CommandButton RHCP 
      BackColor       =   &H00FF8080&
      Caption         =   "Red Hot Chili Peppers"
      Height          =   975
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4200
      Width           =   3255
   End
   Begin VB.CommandButton Beethoven 
      BackColor       =   &H00FF8080&
      Caption         =   "Beethoven"
      Height          =   1095
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   3255
   End
   Begin VB.CommandButton Beatles 
      BackColor       =   &H00FF8080&
      Caption         =   "The Beatles"
      Height          =   1095
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   3255
   End
   Begin VB.CommandButton Bach 
      BackColor       =   &H00FF8080&
      Caption         =   "Bach"
      Height          =   1095
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   3255
   End
   Begin VB.PictureBox Art 
      BackColor       =   &H00FFFFFF&
      Height          =   6255
      Left            =   120
      ScaleHeight     =   6195
      ScaleWidth      =   5715
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "Artists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Palonzison Piano
'This is the Artists Form
'Matthew Peterson and Nicholas Alonzi are the authors of this Form
'This form was written in 2009 in the month of March
'This form allows the user to get a visual of some of the artists listed in the songs file.  It also allows for direct navigation back
    'to the piano.

Private Sub Bach_Click()
    Art.Picture = LoadPicture(App.Path & "\Picture\bach.jpg")
End Sub

Private Sub Beatles_Click()
    Art.Picture = LoadPicture(App.Path & "\Picture\beatles.jpg")
End Sub

Private Sub Beethoven_Click()
    Art.Picture = LoadPicture(App.Path & "\Picture\beethoven.jpg")
End Sub

Private Sub PianoGo_Click()
    Artists.Hide
    Piano.Show
End Sub

Private Sub RHCP_Click()
    Art.Picture = LoadPicture(App.Path & "\Picture\rhc.jpg")
End Sub

