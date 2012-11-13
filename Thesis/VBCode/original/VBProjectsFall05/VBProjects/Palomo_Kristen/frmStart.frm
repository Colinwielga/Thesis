VERSION 5.00
Begin VB.Form frmStart 
   BackColor       =   &H00C0C000&
   Caption         =   "Wheel of Fortune"
   ClientHeight    =   9240
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   ScaleHeight     =   9240
   ScaleWidth      =   11415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   8640
      TabIndex        =   3
      Top             =   5880
      Width           =   1935
   End
   Begin VB.CommandButton cmdInst 
      Caption         =   "Instructions"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   8640
      TabIndex        =   2
      Top             =   3840
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   7530
      Left            =   240
      Picture         =   "frmStart.frx":0000
      ScaleHeight     =   7470
      ScaleWidth      =   7500
      TabIndex        =   0
      Top             =   240
      Width           =   7560
      Begin VB.CommandButton cmdStart 
         Caption         =   "Play Game"
         BeginProperty Font 
            Name            =   "Kristen ITC"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   4680
         TabIndex        =   1
         Top             =   4080
         Width           =   1935
      End
   End
   Begin VB.Image Image1 
      Height          =   3000
      Left            =   8280
      Picture         =   "frmStart.frx":13EC3
      Top             =   360
      Width           =   2250
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Wheeltime.vbp
'Form Name: frmStart.frm
'Author: Kristen Palomo
'Date Written: October 31,2005
'Objective:  The code written allows a user to play my version of Wheel of Fortune by clicking the Play Game button which bring them to another form where the player can participate.
'Allows player to quit game
'Provides instructions for game

Private Sub cmdExit_Click()
End
End Sub



Private Sub cmdInst_Click()
MsgBox ("Make sure Caps Lock is on, Click on Play Game, Enter a number between 1 and 3, Click Show Puzzle, Click on Spin Wheel, Enter a Consonant in text box and click on Show Letters, or click on Buy Vowel and input A,E,I,O, or U, Click on Solve Puzzle when ready and input your answer"), , "How To Play"

End Sub

Private Sub cmdStart_Click()
frmStart.Visible = False
formGame.Visible = True
End Sub




