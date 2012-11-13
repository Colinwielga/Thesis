VERSION 5.00
Begin VB.Form FrmIntro 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Veeder Performance Software"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   FillColor       =   &H00404040&
   LinkTopic       =   "Form1"
   Picture         =   "FrmIntro.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdMain 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Run Quarter Mile Conversion Software"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7800
      Width           =   3735
   End
   Begin VB.CommandButton CmdCredits1 
      BackColor       =   &H80000016&
      Caption         =   "Credits"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8400
      Width           =   1335
   End
   Begin VB.CommandButton CmdQuit2 
      BackColor       =   &H80000016&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9480
      Width           =   1335
   End
End
Attribute VB_Name = "FrmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'checks all variables
Option Explicit
'displays the credits form
Private Sub CmdCredits1_Click()
    FrmCredits.Visible = True
    FrmIntro.Visible = False
End Sub
'displays the main form
Private Sub CmdMain_Click()
    FrmIntro.Visible = False
    FrmMain.Visible = True
End Sub
'ends the program
Private Sub CmdQuit2_Click()
    End
End Sub



