VERSION 5.00
Begin VB.Form frmMapCities 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Form2"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7650
   LinkTopic       =   "Form2"
   ScaleHeight     =   7425
   ScaleWidth      =   7650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPacking 
      BackColor       =   &H00C0FFC0&
      Caption         =   "What Should I Pack?"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton cmdEnd 
      BackColor       =   &H00FFFFC0&
      Caption         =   "View Budget Summary"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4800
      Width           =   1335
   End
   Begin VB.PictureBox picResults 
      Height          =   5415
      Left            =   360
      Picture         =   "frmForm2.frx":0000
      ScaleHeight     =   5355
      ScaleWidth      =   5355
      TabIndex        =   0
      Top             =   1080
      Width           =   5415
      Begin VB.OptionButton Option6 
         Caption         =   "Option6"
         Height          =   195
         Left            =   3360
         TabIndex        =   6
         Top             =   3600
         Width           =   135
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Option3"
         Height          =   195
         Left            =   960
         TabIndex        =   5
         Top             =   4200
         Width           =   135
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   195
         Left            =   2040
         TabIndex        =   4
         Top             =   2040
         Width           =   135
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Left            =   2160
         TabIndex        =   3
         Top             =   2520
         Width           =   135
      End
   End
   Begin VB.Label lblCities 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Click on a city you wish to visit"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   480
      TabIndex        =   2
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "frmMapCities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FinalProject:Travel Europe
'frmMapCities
'Jessica Florek
'Written: 3/4/09
'Objective: have user select a city that they wish to travel to and show the
'form that allows them to choose from options in that city.

Option Explicit

'this form is a naviagtion tool to select different cities to visit
'each option bubble brings you to the form with the corresponding city
'citycounter keeps track of how many cities are 'visited' (clicked on) in order to later calculate the travel expense

Private Sub cmdEnd_Click()
frmMapCities.Hide
frmLAST.Show
End Sub

Private Sub cmdPacking_Click()
'this button brings you to a form that has lists of recommended items to pack and provides different sorting and searching options
frmMapCities.Hide
frmPacking.Show

End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub Option1_Click()
citycounter = citycounter + 1
frmMapCities.Hide
frmParis.Show
End Sub

Private Sub Option2_Click()
citycounter = citycounter + 1
frmMapCities.Hide
frmLondon.Show
End Sub

Private Sub Option3_Click()
citycounter = citycounter + 1
frmMapCities.Hide
frmMadrid.Show

End Sub

Private Sub Option6_Click()
citycounter = citycounter + 1
frmMapCities.Hide
frmVenice.Show
End Sub
