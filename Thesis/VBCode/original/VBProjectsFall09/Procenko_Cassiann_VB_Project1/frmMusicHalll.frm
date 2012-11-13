VERSION 5.00
Begin VB.Form frmMusicHall 
   BackColor       =   &H00000000&
   Caption         =   "Music Hall"
   ClientHeight    =   9315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8265
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9315
   ScaleWidth      =   8265
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLeave 
      Caption         =   "Leave the Music Hall."
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      TabIndex        =   3
      Top             =   7080
      Width           =   3015
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   2
      Top             =   8160
      Width           =   1575
   End
   Begin VB.CommandButton cmdLoadForm 
      Caption         =   "View Music Genres"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   1
      Top             =   7080
      Width           =   2655
   End
   Begin VB.CommandButton cmdWelcome 
      BackColor       =   &H00FF0000&
      Caption         =   $"frmMusicHalll.frx":0000
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      MaskColor       =   &H00FF0000&
      TabIndex        =   0
      Top             =   240
      Width           =   6855
   End
   Begin VB.Image Image1 
      Height          =   5220
      Left            =   1920
      Picture         =   "frmMusicHalll.frx":0092
      Top             =   1440
      Width           =   3810
   End
End
Attribute VB_Name = "frmMusicHall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Music VB Project by Cassiann Procenko
'Form Name is frmMusicHall
'Date written 10/16/2009
'Purpose of this form is to show the viewer the program and what it does, and also greet the viewer


Private Sub cmdLeave_Click()
'show and hide forms
    frmLeave.Show
    frmMusicHall.Hide
End Sub

Private Sub cmdQuit_Click()
'show and hide forms
    frmLeave.Show
    frmMusicHall.Hide
End Sub

Private Sub cmdWelcome_Click()
'define variables
Dim NameInput As String, FavoriteGenre As String

'create input box
NameInput = InputBox("Please enter your name.", , "Name")
MsgBox "Hello, " & NameInput & "!  Welcome to the Music Hall."
FavoriteGenre = InputBox("Please enter your favorite style of music.")
MsgBox "Wow, " & FavoriteGenre & " is a great style of music.  I hope you find what you are looking for!"
End Sub

Private Sub Image1_Click()
'show and hide forms
    frmMusicTypes.Show
    frmMusicHall.Hide
End Sub

Private Sub cmdLoadForm_Click()
'show and hide forms
    frmMusicTypes.Show
    frmMusicHall.Hide
End Sub
