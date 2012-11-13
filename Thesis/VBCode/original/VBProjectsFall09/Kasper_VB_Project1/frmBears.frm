VERSION 5.00
Begin VB.Form frmBears 
   Caption         =   "Bears"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   ScaleHeight     =   5835
   ScaleWidth      =   8715
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdrtn 
      Caption         =   "Return to Teams"
      Height          =   615
      Left            =   600
      TabIndex        =   3
      Top             =   3600
      Width           =   1335
   End
   Begin VB.PictureBox picResults 
      Height          =   4455
      Left            =   4320
      ScaleHeight     =   4395
      ScaleWidth      =   3675
      TabIndex        =   2
      Top             =   720
      Width           =   3735
   End
   Begin VB.CommandButton cmdDefense 
      Caption         =   "Coach Ditka Fun"
      Height          =   735
      Left            =   480
      TabIndex        =   1
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton cmdVisit 
      Caption         =   "DA Bears Stadium"
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   5880
      Left            =   0
      Picture         =   "frmBears.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8715
   End
End
Attribute VB_Name = "frmBears"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Author: Brandon Kasper
'Written 10/19/2009
'objective: load pictures of stadium and coach Ditka

Private Sub cmdDefense_Click()
    picResults.Cls 'cleas the picture box
    picResults.Picture = LoadPicture(App.Path & "\mike.jpg") 'loads the picture coach from file
End Sub

Private Sub cmdRtn_Click()
    frmBears.Hide 'hides the form from user
    frmTeams.Show 'shows form to user
End Sub

Private Sub cmdVisit_Click()
    picResults.Picture = LoadPicture(App.Path & "\stadium.jpg") 'loads picture of stadium for user
End Sub
