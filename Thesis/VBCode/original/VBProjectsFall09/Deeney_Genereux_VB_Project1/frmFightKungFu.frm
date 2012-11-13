VERSION 5.00
Begin VB.Form frmFightKungFu 
   BackColor       =   &H0000FFFF&
   Caption         =   "Fight with the Kung Fu Master!"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10650
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   10650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdEnd 
      Caption         =   "Quit"
      Height          =   255
      Left            =   8280
      TabIndex        =   2
      Top             =   6480
      Width           =   2055
   End
   Begin VB.CommandButton cmdRestart 
      Caption         =   "Now What?"
      Height          =   855
      Left            =   8160
      TabIndex        =   1
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmFightKungFu.frx":0000
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
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   8775
   End
   Begin VB.Image Image2 
      Height          =   4590
      Left            =   720
      Picture         =   "frmFightKungFu.frx":00A8
      Top             =   1080
      Width           =   2985
   End
   Begin VB.Image Image1 
      Height          =   3405
      Left            =   5280
      Picture         =   "frmFightKungFu.frx":4779
      Top             =   2160
      Width           =   4500
   End
End
Attribute VB_Name = "frmFightKungFu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Katie Deeney & Elise Generex
'Create a Story
'Date Done: 10/10/2009
'If the user picks the kung fu master
'he defeats the dragon and the user can
'start from the beginning

Option Explicit

Private Sub CmdEnd_Click()
    End
End Sub

Private Sub cmdRestart_Click()
MsgBox "This is where your story happily ends.  Start Over", , "Story Ends"
Inventory = ""
frmFightKungFu.Hide
frmWelcome.Show

End Sub
