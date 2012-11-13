VERSION 5.00
Begin VB.Form frmmainpage 
   BackColor       =   &H80000007&
   Caption         =   "Form3"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13035
   LinkTopic       =   "Form3"
   ScaleHeight     =   8220
   ScaleWidth      =   13035
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picresults 
      Height          =   5655
      Left            =   3600
      Picture         =   "Form3.frx":0000
      ScaleHeight     =   5595
      ScaleWidth      =   7155
      TabIndex        =   5
      Top             =   1800
      Width           =   7215
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   10440
      TabIndex        =   4
      Top             =   7680
      Width           =   2175
   End
   Begin VB.CommandButton cmdnicknames 
      Caption         =   "Take a Quiz on Player Nicknames"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   7560
      TabIndex        =   3
      Top             =   7680
      Width           =   2175
   End
   Begin VB.CommandButton cmdstats 
      Caption         =   "View Statistics"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4440
      TabIndex        =   2
      Top             =   7680
      Width           =   2415
   End
   Begin VB.CommandButton cmdmeetplayers 
      BackColor       =   &H000000FF&
      Caption         =   "Meet the Greatest Team in the History of Basketball"
      BeginProperty Font 
         Name            =   "Myriad Web Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1320
      MaskColor       =   &H000000FF&
      TabIndex        =   0
      Top             =   7680
      Width           =   2475
   End
   Begin VB.Label lblbulls 
      BackColor       =   &H000000FF&
      Caption         =   "The Greatest Team In NBA History!!!"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   615
      Left            =   3240
      TabIndex        =   1
      Top             =   1080
      Width           =   8055
   End
End
Attribute VB_Name = "frmmainpage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    'Chicago Bulls (Chicagobulls.vbp)
    'frmmainpage(frmmainpage.frm)
    'Written by: Brian Cullen
    'Written on: March 16, 2008
    'Objective: This form allows the user to meet the Chicago Bulls players by
    'viewing a photo of each player.

Private Sub cmdquit_Click()
'this page stops the program
End
End Sub
Private Sub cmdmeetplayers_Click()
'This hides the main page and shows the meet the players form.
frmmeetplayers.Show
frmmainpage.Hide
End Sub


Private Sub cmdstats_Click()
' This hides main page and goes to the statistics page
frmstats.Show
frmmainpage.Hide
End Sub

Private Sub cmdnicknames_Click()
'this page hides the mainpage and shows the nickname quiz form
frmquiz.Show
frmmainpage.Hide
End Sub




