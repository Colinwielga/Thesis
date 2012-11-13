VERSION 5.00
Begin VB.Form frmPictures 
   BackColor       =   &H0000FFFF&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Back to Menu"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.PictureBox picResults 
      Height          =   2055
      Left            =   360
      ScaleHeight     =   1995
      ScaleWidth      =   3675
      TabIndex        =   1
      Top             =   720
      Width           =   3735
   End
   Begin VB.CommandButton cmdSee 
      Caption         =   "See Next Picture"
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

Private Sub cmdMenu_Click()
frmPictures.Hide
frmMenu.Show
End Sub

Private Sub cmdSee_Click()
picResults.Picture = LoadPicture(App.Path & "\Images\Halloween.jpg")
End Sub




