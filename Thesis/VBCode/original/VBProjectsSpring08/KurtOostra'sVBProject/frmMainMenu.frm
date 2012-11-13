VERSION 5.00
Begin VB.Form frmMainMenu 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "frmMainMenu.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Credits"
      Height          =   855
      Left            =   2520
      TabIndex        =   8
      Top             =   9360
      Width           =   2415
   End
   Begin VB.CommandButton cmdCircle 
      BackColor       =   &H000000FF&
      Caption         =   "How big is the circle"
      Height          =   975
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CommandButton cmdPictures 
      BackColor       =   &H0000FF00&
      Caption         =   "Look at some picture of past events"
      Height          =   975
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CommandButton cmdBluePrints 
      BackColor       =   &H00FF00FF&
      Caption         =   "Learn about the college theater venues"
      Height          =   975
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CommandButton cmdInventory 
      BackColor       =   &H0000FFFF&
      Caption         =   "Inventory the light supply"
      Height          =   975
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CommandButton cmdFixtures 
      BackColor       =   &H00FF0000&
      Caption         =   "Learn about different light fixtures"
      Height          =   975
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quit"
      Height          =   855
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9360
      Width           =   2775
   End
   Begin VB.CommandButton cmdBalance 
      BackColor       =   &H00FFFF00&
      Caption         =   "Determine the weight to balance the fly rail"
      Height          =   975
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000C0&
      Caption         =   "Theater Lighting"
      BeginProperty Font 
         Name            =   "Rosewood Std Regular"
         Size            =   56.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   2760
      TabIndex        =   7
      Top             =   720
      Width           =   9735
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Theater Lighting
'Form Name: frmMainMenu
'Author: Kurt Oostra
'Date Written:3/11/08
'Objective: Starting point to exploring the project
Option Explicit
Private Sub cmdBalance_Click()
frmMainMenu.Hide
frmBalance.Show
End Sub

Private Sub cmdBluePrints_Click()
frmMainMenu.Hide
frmVenues.Show
End Sub

Private Sub cmdCircle_Click()
frmMainMenu.Hide
frmCircle.Show
End Sub

Private Sub cmdFixtures_Click()
frmMainMenu.Hide
frmFixtures.Show
End Sub

Private Sub cmdInventory_Click()
frmMainMenu.Hide
frmInventory.Show
End Sub

Private Sub cmdPictures_Click()
frmMainMenu.Hide
frmpictures.Show
End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
frmMainMenu.Hide
frmVenues.Show
End Sub
