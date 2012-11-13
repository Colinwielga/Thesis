VERSION 5.00
Begin VB.Form frmPictures 
   BackColor       =   &H0000FFFF&
   Caption         =   "Form1"
   ClientHeight    =   7815
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   10170
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd6 
      Caption         =   "See Picture 6"
      Height          =   495
      Left            =   8160
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmd5 
      Caption         =   "See Picture 5"
      Height          =   495
      Left            =   6600
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmd4 
      Caption         =   "See Picture 4"
      Height          =   495
      Left            =   5040
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "See Picture 3"
      Height          =   495
      Left            =   3480
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "See Picture 2"
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Back to Menu"
      Height          =   495
      Left            =   7800
      TabIndex        =   2
      Top             =   7080
      Width           =   1335
   End
   Begin VB.PictureBox picResults 
      Height          =   6375
      Left            =   360
      ScaleHeight     =   6315
      ScaleWidth      =   8835
      TabIndex        =   1
      Top             =   720
      Width           =   8895
   End
   Begin VB.CommandButton cmdSee 
      Caption         =   "See Picture 1"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmPictures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd2_Click()
picResults.Picture = LoadPicture(App.Path & "\Images\Me Halloween.jpg")
End Sub

Private Sub cmd3_Click()
picResults.Picture = LoadPicture(App.Path & "\Images\Me and Matthew(grad party).jpg")
End Sub

Private Sub cmd4_Click()
picResults.Picture = LoadPicture(App.Path & "\Images\Me and Bri(Halloween2).jpg")
End Sub

Private Sub cmd5_Click()
picResults.Picture = LoadPicture(App.Path & "\Images\Me Matthew and Branden camping.jpg")
End Sub

Private Sub cmd6_Click()
picResults.Picture = LoadPicture(App.Path & "\Images\Me and Daddy(Prom 05).jpg")
End Sub

Private Sub cmdMenu_Click()
frmPictures.Hide
frmMenu.Show
End Sub

Private Sub cmdSee_Click()
picResults.Picture = LoadPicture(App.Path & "\Images\Halloween.jpg")

End Sub


