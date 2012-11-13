VERSION 5.00
Begin VB.Form frmFightKungFu2 
   BackColor       =   &H0000FFFF&
   Caption         =   "Fight with the Kung Fu Master!"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   ScaleHeight     =   6300
   ScaleWidth      =   9705
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmdquit 
      Caption         =   "Quit"
      Height          =   255
      Left            =   7440
      TabIndex        =   2
      Top             =   6000
      Width           =   1935
   End
   Begin VB.CommandButton cmdRestart 
      Caption         =   "Now What?"
      Height          =   855
      Left            =   7320
      TabIndex        =   1
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   3405
      Left            =   4680
      Picture         =   "frmFightKungFu2.frx":0000
      Top             =   1560
      Width           =   4500
   End
   Begin VB.Image Image2 
      Height          =   4590
      Left            =   360
      Picture         =   "frmFightKungFu2.frx":4004
      Top             =   960
      Width           =   2985
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmFightKungFu2.frx":86D5
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8775
   End
End
Attribute VB_Name = "frmFightKungFu2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Katie Deeney & Elise Generex
'Create a Story
'Date Done: 10/10/2009
'If the user picks the kung fu master
'he defeats the dragon and the user can
'start from the beginning

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub cmdRestart_Click()
MsgBox "This is where your story happily ends.  Start Over", , "Story Ends"
Inventory = ""
frmFightKungFu2.Hide
frmWelcome.Show
End Sub
