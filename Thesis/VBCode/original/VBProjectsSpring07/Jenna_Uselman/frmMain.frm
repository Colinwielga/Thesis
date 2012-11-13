VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000009&
   Caption         =   "ABC"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   4320
   ScaleWidth      =   7275
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H80000009&
      Caption         =   "exit program"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2880
      Width           =   3255
   End
   Begin VB.CommandButton cmdgame 
      BackColor       =   &H80000009&
      Caption         =   "trivia game"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   3255
   End
   Begin VB.CommandButton cmdABCschedule 
      BackColor       =   &H80000009&
      Caption         =   "abc schedule"
      BeginProperty Font 
         Name            =   "Bauhaus 93"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Project Name: ABC Television Channel
'Author: Jenna Uselman
'Date: March 25, 2007
'Purpose of Project: This project is designed for fans of the television channel ABC.
'                    The project allows users to see the ABC primetime schedule and
'                    play a trivia game about the shows. The trivia game also rates
'                    how much the user watches primetime ABC.
'Purpose of Form: This form is the introduction of the project, and works as the main
'                 menu of the project. Users can choose between three command buttons.
'                 They can either play the trivia game, view the schedule, or exit the
'                 program.

Private Sub cmdABCschedule_Click() 'This command button takes users to the ABC schedule form.
    frmABCschedule.Show
    frmMain.Hide
End Sub

Private Sub cmdGame_Click() 'This command button takes users to the trivia game form.
    frmgameintro.Show
    frmMain.Hide
End Sub

Private Sub cmdquit_Click() 'This command button allows the user to quit the program.
End
End Sub

