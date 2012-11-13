VERSION 5.00
Begin VB.Form frmtutorial 
   BackColor       =   &H00400000&
   Caption         =   "Choose Your Tutorial"
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   10320
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton Cmdcited 
      BackColor       =   &H000000FF&
      Caption         =   "Works Cited and Etc..."
      Height          =   1095
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6360
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Height          =   3375
      Left            =   2640
      Picture         =   "frmtutorial.frx":0000
      ScaleHeight     =   3315
      ScaleWidth      =   3795
      TabIndex        =   3
      Top             =   120
      Width           =   3855
   End
   Begin VB.CommandButton cmdplaytecmo 
      BackColor       =   &H000000FF&
      Caption         =   "TSB Tips and Hints"
      Height          =   2775
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   2295
   End
   Begin VB.CommandButton cmdplayfootball 
      BackColor       =   &H000000FF&
      Caption         =   "How to Play Football"
      Height          =   3255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton cmdmainmenu 
      BackColor       =   &H000000FF&
      Caption         =   "Main Menu"
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6360
      Width           =   2295
   End
End
Attribute VB_Name = "frmtutorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: Super Tecmo Database
'Form name: frmtutorial
'Author: Nate Johnson & Kevin Klein
'Date Written: October 11th, 2006
'Objective of project: This project will allow its users to learn more about the game of football
'and will also allow them the oppurtunity to learn how to play the game of football with the Nintendo
'video game, Tecmo Super Bowl.
'Objective of form: This form serves as a junction point for the user to choose their
'destination. From this form, the user can access, the howtoplayfootball form, the works cited form,
'and the tips and hints form.

Private Sub Cmdcited_Click()
frmcited.Show 'shows the new form
frmtutorial.Hide 'hides the old form
End Sub

Private Sub cmdmainmenu_Click()
frmtutorial.Hide 'hides the old form
frmMain.Show 'shows the new form

End Sub

Private Sub cmdplayfootball_Click()
frmhowtoplayfootball.Show 'shows the new form
frmtutorial.Hide 'hides the old form
End Sub

Private Sub cmdplaytecmo_Click()
frmtutorial.Hide 'hides the old form
frmtips.Show 'shows the new form
End Sub
