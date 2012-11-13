VERSION 5.00
Begin VB.Form Railings 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Form2"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7560
   LinkTopic       =   "Form2"
   ScaleHeight     =   6960
   ScaleWidth      =   7560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdForm 
      Caption         =   "Go Back"
      Height          =   735
      Left            =   2760
      TabIndex        =   6
      Top             =   6000
      Width           =   2295
   End
   Begin VB.PictureBox picDisplay 
      Height          =   3975
      Left            =   120
      Picture         =   "Railings.frx":0000
      ScaleHeight     =   3915
      ScaleWidth      =   7155
      TabIndex        =   5
      Top             =   840
      Width           =   7215
   End
   Begin VB.CommandButton cmdPlastic 
      Caption         =   "4: Vinyl Railing"
      Height          =   735
      Left            =   5880
      TabIndex        =   4
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton cmdPrivacy 
      Caption         =   "3: Privacy Wall"
      Height          =   735
      Left            =   4200
      TabIndex        =   3
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton cmdFlat 
      Caption         =   "2: Flat Railing"
      Height          =   735
      Left            =   2280
      TabIndex        =   2
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton cmdPost 
      Caption         =   "1: Post Railing"
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "David Schonhardt"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   6480
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Types of Railings"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   4935
   End
End
Attribute VB_Name = "Railings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Railings (Railing.frm), which is for the display of different types of railings.

Private Sub cmdFlat_Click()
picDisplay.Picture = LoadPicture(PublicModule.Path & "flatrail.jpg")
MsgBox "A flat railing is made up of cedar 2x2 spindles 40 inches high, with 5/4 cedar decking used as capping.", , "Flat Railing"
End Sub

Private Sub cmdForm_Click()
MainForm.Show
Railings.Hide
End Sub

Private Sub cmdPlastic_Click()
picDisplay.Picture = LoadPicture(PublicModule.Path & "plastic.jpg")
MsgBox "A vinyl railing is made out of vinyl spindles, with steel supports within the posts and rail.", , "Vinyl Railing"
End Sub

Private Sub cmdPost_Click()
picDisplay.Picture = LoadPicture(PublicModule.Path & "postbig.jpg")
MsgBox "A post railing is made of 2x2 cedar spindles and 4x4 cedar posts no more than eight feet apart, and 5/4 decking is used to cap.", , "Post Railing"
End Sub

Private Sub cmdPrivacy_Click()
picDisplay.Picture = LoadPicture(PublicModule.Path & "privacybig.jpg")
MsgBox "Privacy walls are eight feet tall and are made out of tongue & groove cedar with 4x4 posts and 2x4 supports to hold it in place.", , "Privacy Wall"
End Sub
