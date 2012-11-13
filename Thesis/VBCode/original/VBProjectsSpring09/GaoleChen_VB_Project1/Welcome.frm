VERSION 5.00
Begin VB.Form frmWelcome 
   Caption         =   "Form1"
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13665
   LinkTopic       =   "Form1"
   ScaleHeight     =   8205
   ScaleWidth      =   13665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMenu 
      BackColor       =   &H8000000D&
      Caption         =   "Menu"
      Height          =   1335
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6720
      Width           =   2535
   End
   Begin VB.CommandButton cmdSpecial 
      BackColor       =   &H8000000D&
      Caption         =   "Fast order-Special of the day!"
      Height          =   1335
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6720
      Width           =   2535
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H8000000D&
      Caption         =   "Quit"
      Height          =   1335
      Left            =   9360
      MaskColor       =   &H8000000D&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6720
      Width           =   2535
   End
   Begin VB.Image Image2 
      Height          =   5010
      Left            =   0
      Picture         =   "Welcome.frx":0000
      Top             =   1680
      Width           =   7500
   End
   Begin VB.Image Image1 
      Height          =   4920
      Left            =   6960
      Picture         =   "Welcome.frx":74BA
      Top             =   1680
      Width           =   6750
   End
   Begin VB.Label lblwelcome 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Welcome To Grand China Restaurant"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13695
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Digital Menu
'Form Name: frmWelcome
'Authors: Gaole Chen
'Date Written: 3/7/09
'Objective:This is the start up form for the project
'The user can select which type of food they like, and make a order
'on line.
'On the first form the user can choose to go through the menu, or
'pick up the Special Combo of the day.
Option Explicit

Private Sub cmdMenu_Click()

'The user will go to the menu form by clicking this button

frmWelcome.Hide
frmMenu.Show


End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdSpecial_Click()

'The user will go to the Special-combo form by clicking this button

frmWelcome.Hide
frmSpecial.Show

End Sub

