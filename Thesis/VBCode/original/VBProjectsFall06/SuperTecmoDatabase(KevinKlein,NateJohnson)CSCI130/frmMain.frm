VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00800000&
   Caption         =   "MAIN PAGE"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   10425
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture5 
      Height          =   3375
      Left            =   7800
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   3315
      ScaleWidth      =   2235
      TabIndex        =   9
      Top             =   4080
      Width           =   2295
   End
   Begin VB.PictureBox Picture4 
      Height          =   3735
      Left            =   6720
      Picture         =   "frmMain.frx":10A52
      ScaleHeight     =   3675
      ScaleWidth      =   2955
      TabIndex        =   8
      Top             =   240
      Width           =   3015
   End
   Begin VB.PictureBox Picture3 
      Height          =   3015
      Left            =   5160
      Picture         =   "frmMain.frx":12D88
      ScaleHeight     =   2955
      ScaleWidth      =   2235
      TabIndex        =   7
      Top             =   4200
      Width           =   2295
   End
   Begin VB.PictureBox Picture2 
      Height          =   3735
      Left            =   2040
      Picture         =   "frmMain.frx":155DB
      ScaleHeight     =   3675
      ScaleWidth      =   2715
      TabIndex        =   6
      Top             =   3720
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      Height          =   3375
      Left            =   2760
      Picture         =   "frmMain.frx":1CFB8
      ScaleHeight     =   3315
      ScaleWidth      =   3795
      TabIndex        =   5
      Top             =   240
      Width           =   3855
   End
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton cmdorderform 
      BackColor       =   &H000000FF&
      Caption         =   "Play Tecmo Super Bowl"
      Height          =   1455
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton cmdencyc 
      BackColor       =   &H000000FF&
      Caption         =   "TSB Hall of Fame"
      Height          =   1335
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton cmdtutorials 
      BackColor       =   &H000000FF&
      Caption         =   "Tutorials"
      Height          =   1455
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton CmdHistory 
      BackColor       =   &H000000FF&
      Caption         =   "History Lessons"
      Height          =   1335
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Super Tecmo Database
'Form name: frmMain
'Author: Nate Johnson & Kevin Klein
'Date Written: October 11th, 2006
'Objective of project: This project will allow its users to learn more about the game of football
'and will also allow them the oppurtunity to learn how to play the game of football with the Nintendo
'video game, Tecmo Super Bowl.
'Objective of form:  This form serves as a homepage for the user to manuever around the project on.
            
Option Explicit


Private Sub cmdencyc_Click()
frmMain.Hide 'hides the old form
frmhalloffame.Show 'shows the new form
End Sub

Private Sub CmdHistory_Click()
frmMain.Hide 'hides the old form
frmHistory.Show 'shows the new form
End Sub

Private Sub cmdorderform_Click()
MsgBox "When Nesten loads, press F1 and then scroll down until you see the Tecmo Super Bowl Rom File. Please Remember that possessing rom files for longer than 24 hours without owning the original cartridge is illegal. Nesten emulator can be found at Romnation.net. Please ignore the next message box as it is a warning from the emulator. It will not affect the game play."
Call Shell(App.Path & "\nesten\nesten.exe")
End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub cmdtutorials_Click()
frmMain.Hide 'hides the old form
frmtutorial.Show 'shows the new form
End Sub

