VERSION 5.00
Begin VB.Form FrmMain 
   BackColor       =   &H00800000&
   Caption         =   "The Minnesota Twins"
   ClientHeight    =   9165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   ScaleHeight     =   9165
   ScaleWidth      =   11625
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdAge 
      BackColor       =   &H000000C0&
      Caption         =   "Age and Availability"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7080
      Width           =   2295
   End
   Begin VB.CommandButton CmdQuit 
      BackColor       =   &H000000C0&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8160
      Width           =   2295
   End
   Begin VB.CommandButton CmdStats 
      BackColor       =   &H000000C0&
      Caption         =   "Team Statistics"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7080
      Width           =   2295
   End
   Begin VB.CommandButton CmdMeet 
      BackColor       =   &H000000C0&
      Caption         =   "Meet the Team"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7080
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Height          =   6975
      Left            =   1200
      Picture         =   "FrmMain.frx":0000
      ScaleHeight     =   6915
      ScaleWidth      =   8115
      TabIndex        =   0
      Top             =   0
      Width           =   8175
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'The Minnesota Twins
'FrmMain
'Sarah Mayer and Jake Klis
' Written on 03/21/09
' This is the main form and it is designed to help the user navigate through the program using buttons.
' This program is written to allow the user to look up various personal information and professional statistics
' on the starting Twins players.  The user is then able to see the players ranked according
'to a given stat, and find out their age and marital status.


Private Sub CmdAge_Click()
FrmAgeandAvailability.Show
FrmMain.Hide

End Sub

Private Sub CmdMeet_Click()
FrmMeet.Show
FrmMain.Hide
End Sub

Private Sub CmdQuit_Click()
MsgBox "You got " & TriviaCtr & " answers correct out of 5 possible", , "Good Job!"
End
End Sub

Private Sub CmdStats_Click()
frmStats.Show
FrmMain.Hide
End Sub

