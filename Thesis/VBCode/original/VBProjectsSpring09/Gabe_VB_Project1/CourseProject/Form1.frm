VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "Menu"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   7230
   ScaleWidth      =   11580
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   12120
      TabIndex        =   4
      Top             =   6120
      Width           =   1935
   End
   Begin VB.CommandButton cmdStats 
      Caption         =   "Get to know the CSB/SJU presidents!"
      Height          =   1215
      Left            =   10440
      TabIndex        =   3
      Top             =   960
      Width           =   2295
   End
   Begin VB.CommandButton cmdSlideshow 
      Caption         =   "Slideshow"
      Height          =   1335
      Left            =   3000
      TabIndex        =   2
      Top             =   4680
      Width           =   2655
   End
   Begin VB.CommandButton cmdSJUslideshow 
      Caption         =   "What do you know about SJU?"
      Height          =   1695
      Left            =   7920
      TabIndex        =   1
      Top             =   3000
      Width           =   2295
   End
   Begin VB.CommandButton cmdCSBtrivia 
      Caption         =   "How well do you know Bennie history?"
      Height          =   1695
      Left            =   3000
      Picture         =   "Form1.frx":11B89
      TabIndex        =   0
      Top             =   960
      Width           =   3015
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Fun with CSB/SJU History!
'frmMenu
'Audrey Gabe
'Written 3/12/09
'This page gives the user the option of which other forms to look at


Private Sub cmdCSBtrivia_Click()
'brings user to BennieTrivia page
frmMenu.Hide
frmBennieTrivia.Show
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdSJUslideshow_Click()
frmMenu.Hide
frmJohnnieTrivia.Show
End Sub

Private Sub cmdSlideshow_Click()
frmMenu.Hide
frmSlideshow.Show
End Sub

Private Sub cmdStats_Click()
frmMenu.Hide
frmPresidents.Show
End Sub


