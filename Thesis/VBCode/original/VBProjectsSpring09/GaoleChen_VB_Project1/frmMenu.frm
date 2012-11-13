VERSION 5.00
Begin VB.Form frmMenu 
   BackColor       =   &H00404000&
   Caption         =   "Form1"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11010
   LinkTopic       =   "Form1"
   ScaleHeight     =   8475
   ScaleWidth      =   11010
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00404000&
      Caption         =   "Quit"
      Height          =   855
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7080
      Width           =   1335
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00404000&
      Caption         =   "Back"
      Height          =   855
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7080
      Width           =   1335
   End
   Begin VB.CommandButton cmdBeverage 
      BackColor       =   &H00404000&
      Caption         =   "Beverage"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   15.75
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5520
      Width           =   2775
   End
   Begin VB.CommandButton cmdDessert 
      BackColor       =   &H00404000&
      Caption         =   "Dessert"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   15.75
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4200
      Width           =   2775
   End
   Begin VB.CommandButton cmdDish 
      BackColor       =   &H00404000&
      Caption         =   "Main Dish"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   15.75
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2880
      Width           =   2775
   End
   Begin VB.CommandButton cmdSalad 
      BackColor       =   &H00404000&
      Caption         =   "Salad"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   15.75
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   9060
      Left            =   4920
      Picture         =   "frmMenu.frx":0000
      Top             =   0
      Width           =   9060
   End
   Begin VB.Label lblType 
      BackColor       =   &H00404000&
      Caption         =   "How may I help you?"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   9015
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Digital Menu
'Form Name: frmMenu
'Authors: Gaole Chen
'Date Written: 3/7/09
'Objective:This is the menu form for the project.
'The user can select which course they want.

Option Explicit

Private Sub cmdBack_Click()
'The user will go back to the previous screen
frmMenu.Hide
frmWelcome.Show
End Sub

Private Sub cmdBeverage_Click()
'The user will go to the Beverage menu by clicking this button
frmMenu.Hide
frmBeverage.Show
End Sub

Private Sub cmdDessert_Click()
'The user will go to the dessert menu by clicking this button
frmMenu.Hide
frmDessert.Show
End Sub

Private Sub cmdDish_Click()
'The user will go to the main dish menu by clicking this button
frmMenu.Hide
frmMain.Show
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdSalad_Click()
'The user will go to the salad menu by clicking this button
frmMenu.Hide
frmSalad.Show
End Sub

