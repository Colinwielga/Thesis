VERSION 5.00
Begin VB.Form frmMainMenu 
   Caption         =   "Main Menu: Western Europe Travel Log"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10125
   LinkTopic       =   "Form2"
   ScaleHeight     =   6810
   ScaleWidth      =   10125
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGallery 
      Caption         =   "Picture Gallery"
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   5760
      Width           =   3135
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   8280
      TabIndex        =   4
      Top             =   5760
      Width           =   1575
   End
   Begin VB.CommandButton cmdCredits 
      Caption         =   "Credits"
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   5760
      Width           =   3135
   End
   Begin VB.CommandButton cmdLogEditor 
      Caption         =   "Take Surveys"
      Height          =   375
      Left            =   6720
      TabIndex        =   2
      Top             =   5280
      Width           =   3135
   End
   Begin VB.CommandButton cmdCountryInfo 
      Caption         =   "Discover Western Europe"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   5280
      Width           =   3135
   End
   Begin VB.CommandButton cmdTravelLog 
      Caption         =   "View Your Travel Information"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   5280
      Width           =   3135
   End
   Begin VB.Image picStation 
      Height          =   9000
      Left            =   -480
      Picture         =   "frmMainMenu.frx":0000
      Top             =   -1800
      Width           =   12000
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Project Name: Western Europe Travel Log
'Form Name: frmMainMenu
'Author: Nate Burbeck
'Date Written: 14 March 2008
'Objective: the main concourse of the project, buttons that get us to other forms in the program or to quit
Option Explicit
' this form acts as our main menu,
Private Sub cmdCountryInfo_Click()
frmMainMenu.Hide
frmCountryInfo.Show
End Sub

Private Sub cmdCredits_Click()
frmMainMenu.Hide
frmCredits.Show
End Sub

Private Sub cmdGallery_Click()
frmMainMenu.Hide
frmGallery.Show
End Sub

Private Sub cmdLogEditor_Click()
frmTravelLog.Show
frmMainMenu.Hide
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdTravelLog_Click()
frmViewData.Show
frmMainMenu.Hide
End Sub

