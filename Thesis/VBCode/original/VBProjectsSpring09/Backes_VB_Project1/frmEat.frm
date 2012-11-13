VERSION 5.00
Begin VB.Form frmRestaurantsWarwick 
   BackColor       =   &H00008000&
   Caption         =   "places to eat at the Warwick"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   9405
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0000FFFF&
      Caption         =   "Quit"
      Height          =   495
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H0000FF00&
      Caption         =   "click to go back to the previous page"
      Height          =   735
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4080
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      Height          =   3375
      Left            =   5280
      Picture         =   "frmEat.frx":0000
      ScaleHeight     =   3315
      ScaleWidth      =   4035
      TabIndex        =   4
      Top             =   480
      Width           =   4095
   End
   Begin VB.PictureBox Picture1 
      Height          =   2655
      Left            =   1680
      Picture         =   "frmEat.frx":5E4D
      ScaleHeight     =   2595
      ScaleWidth      =   3555
      TabIndex        =   3
      Top             =   3360
      Width           =   3615
   End
   Begin VB.CommandButton cmdRbar 
      BackColor       =   &H00FFFF80&
      Caption         =   "Randolph's Bar"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   2535
   End
   Begin VB.CommandButton cmdMurals 
      BackColor       =   &H00FF80FF&
      Caption         =   "Join us at Murals on 54!"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label lblplacetoeat 
      BackColor       =   &H00404040&
      Caption         =   "Click on the buttons to the right to check out the different Restaurant options at the Warwick!!"
      BeginProperty Font 
         Name            =   "NancyBlue"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frmRestaurantsWarwick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name:New York And L.A
'Form Name: frmActivities
'Author: Emily Backes
'Date Written: 3-17-09
'Objective: This form tells the user some places to eat while
'staying at the Warwick in New York
'A message box comes up and tell the user what types of food is
'served at the restaurant

Option Explicit

Private Sub cmdBack_Click()
frmRestaurantsWarwick.Hide
frmRoomsWarwick.Show
End Sub

Private Sub cmdMurals_Click()
MsgBox ("Manhattan's newest historic restaurant, open for Breakfast, Lunch and Dinner we serve a variety of delicious meals!")

End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub cmdRbar_Click()
MsgBox ("Named 2007 best Hotel bar, Randolph's is open for Lunch and Dinner and servers a variety of drinks and wonderful food!")

End Sub
