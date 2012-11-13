VERSION 5.00
Begin VB.Form frmMenu 
   BackColor       =   &H0000FFFF&
   Caption         =   "Form1"
   ClientHeight    =   4650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5670
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4650
   ScaleWidth      =   5670
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Enter Your Name --->"
      Height          =   495
      Left            =   3360
      TabIndex        =   11
      Top             =   0
      Width           =   975
   End
   Begin VB.TextBox txtName 
      Height          =   495
      Left            =   4320
      TabIndex        =   10
      Top             =   0
      Width           =   1095
   End
   Begin VB.PictureBox picResults 
      Height          =   375
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   3315
      TabIndex        =   9
      Top             =   720
      Width           =   3375
   End
   Begin VB.CommandButton cmdPictures 
      BackColor       =   &H00FF00FF&
      Caption         =   "See Pictures of Kristie"
      Height          =   855
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FFFF00&
      Caption         =   "Quit"
      Height          =   855
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton cmdFulllist 
      BackColor       =   &H00FFFF00&
      Caption         =   "See Full List of Activities"
      Height          =   855
      Left            =   960
      MaskColor       =   &H8000000E&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton cmdOther 
      BackColor       =   &H00FF00FF&
      Caption         =   "Other Extracurriculars"
      Height          =   855
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton cmdFavorites 
      BackColor       =   &H00FFFF00&
      Caption         =   "Go to Favorites"
      Height          =   855
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton cmdTrivia 
      BackColor       =   &H00FF00FF&
      Caption         =   "Play Trivia"
      Height          =   855
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton cmdMusic 
      BackColor       =   &H00FFFF00&
      Caption         =   "Go to Music"
      Height          =   855
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton cmdSports 
      BackColor       =   &H00FF00FF&
      Caption         =   "Go to Sports"
      Height          =   855
      Left            =   960
      MaskColor       =   &H00800080&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblTitle 
      Caption         =   "Kristie's Activities"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdEnter_Click()
picResults.Print "Welcome "; txtName.Text; "!"
End Sub

Private Sub cmdFavorites_Click()
frmMenu.Hide
frmFavorites.Show
End Sub

Private Sub cmdFulllist_Click()
frmMenu.Hide
frmActivities.Show
End Sub

Private Sub cmdMusic_Click()
frmMenu.Hide
frmMusic.Show
End Sub

Private Sub cmdOther_Click()
frmMenu.Hide
frmOther.Show
End Sub

Private Sub cmdPictures_Click()
frmMenu.Hide
frmPictures.Show
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdSports_Click()
frmMenu.Hide
frmSports.Show
End Sub


Private Sub cmdTrivia_Click()
frmTrivia.Show
frmMenu.Hide
End Sub


