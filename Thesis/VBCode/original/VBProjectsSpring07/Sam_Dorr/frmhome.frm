VERSION 5.00
Begin VB.Form frmhome 
   BackColor       =   &H000000C0&
   Caption         =   "Home"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10485
   LinkTopic       =   "Form1"
   Picture         =   "frmhome.frx":0000
   ScaleHeight     =   6930
   ScaleWidth      =   10485
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdregister1 
      Caption         =   "Back to Register"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2880
      TabIndex        =   8
      Top             =   5880
      Width           =   2295
   End
   Begin VB.CommandButton cmdcustom 
      Caption         =   "Customize Your Player"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   6
      Top             =   5880
      Width           =   2295
   End
   Begin VB.CommandButton cmdworkcited 
      Caption         =   "Work Cited"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6360
      TabIndex        =   5
      Top             =   4920
      Width           =   3975
   End
   Begin VB.CommandButton cmdtickets 
      Caption         =   "Tickets"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6360
      TabIndex        =   4
      Top             =   3720
      Width           =   3975
   End
   Begin VB.CommandButton cmdstadium 
      Caption         =   "Pictures of Rosenblatt Stadium"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6360
      TabIndex        =   3
      Top             =   2520
      Width           =   3975
   End
   Begin VB.CommandButton cmdoutlook 
      Caption         =   "2007 Team Outlook"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6360
      TabIndex        =   2
      Top             =   1320
      Width           =   3975
   End
   Begin VB.CommandButton cmdabout 
      Caption         =   "About College World Series "
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6360
      TabIndex        =   1
      Top             =   120
      Width           =   3975
   End
   Begin VB.PictureBox piclastname 
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000C0&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      Height          =   375
      Left            =   3000
      ScaleHeight     =   315
      ScaleWidth      =   1275
      TabIndex        =   0
      Top             =   2400
      Width           =   1335
   End
End
Attribute VB_Name = "frmhome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'College World Series.(NCAACollegeWorldSeries.vbp)

'Form name: frmiHome; Form caption: Home

'Author: Sam Dorr

'Date written: March 25, 2007

' Form Objective: The objective of frmhome is to be the central navigation
'                   point for the user.  All the forms are accessible from frmhome.
Option Explicit

Private Sub cmdabout_Click()
    frmhome.Hide
    frmabout.Show
End Sub

Private Sub cmdcustom_Click()
    piclastname.Print username 'print user name on back of pitchers jersey
End Sub

Private Sub cmdoutlook_Click()
    frmhome.Hide
    frmoutlook.Show
End Sub

Private Sub cmdquit_Click()
    End
End Sub

Private Sub cmdregister1_Click()
    frmhome.Hide
    frmIntro.Show
End Sub

Private Sub cmdstadium_Click()
    frmhome.Hide
    frmstadiumpictures.Show
End Sub

Private Sub cmdtickets_Click()
    frmhome.Hide
    frmtickets.Show
    
End Sub

Private Sub cmdworkcited_Click()
    frmhome.Hide
    frmworkcited.Show
End Sub
