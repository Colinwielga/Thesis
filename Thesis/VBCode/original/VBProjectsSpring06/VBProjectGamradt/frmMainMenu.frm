VERSION 5.00
Begin VB.Form frmMainMenu 
   BackColor       =   &H00808000&
   Caption         =   "Main Menu"
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11040
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
   ScaleWidth      =   11040
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   9840
      TabIndex        =   2
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton cmdNavigateDemo 
      Caption         =   "Begin Individual Averages"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   4200
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      Top             =   3480
      Width           =   2535
   End
   Begin VB.CommandButton cmdNavigateAbout 
      Caption         =   "About Team Manager Pro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   4200
      TabIndex        =   0
      Top             =   5400
      Width           =   2535
   End
   Begin VB.Label lblPlayerMatchup 
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Main menu"
      BeginProperty Font 
         Name            =   "Trajan Pro"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   795
      Left            =   3720
      TabIndex        =   5
      Top             =   2520
      Width           =   3495
   End
   Begin VB.Label lblDesigner 
      BackColor       =   &H00808000&
      Caption         =   "By: Erik Gamradt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   7440
      Width           =   1575
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Team Manager Pro"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1335
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   10455
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Team Manager Pro (ErikGamradtVBProject.vbp)
'frmMainMenu (frmMainMenu.frm)
'Designed By: Erik Gamradt
'15 March 2006
'Team Manager Pro. was developed to relieve some of the laborious and stressful work of manually calculated statistics by hand through a program that accept player statistics, averages them out, and compares them in a variety of ways by a click of a button.  Additionally, the program allows the user to predict winners between various teams based on statistical data.
Private Sub cmdNavigateAbout_Click()
    frmMainMenu.Hide  'all navigation commands
    frmAbout.Show
End Sub

Private Sub cmdNavigateDemo_Click()
    frmIndividualAvg.Show
    frmMainMenu.Hide
End Sub

Private Sub cmdNavigateSignup_Click()
    frmSignUp.Show
    frmMainMenu.Hide
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

