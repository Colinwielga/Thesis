VERSION 5.00
Begin VB.Form FrmCredits 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Credits"
   ClientHeight    =   10500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   Picture         =   "FrmCredits.frx":0000
   ScaleHeight     =   10500
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdQuit1 
      BackColor       =   &H00FF0000&
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
      Height          =   975
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9120
      Width           =   1815
   End
   Begin VB.CommandButton CmdMainForm 
      BackColor       =   &H00FF0000&
      Caption         =   "Back To Program"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8040
      Width           =   1815
   End
End
Attribute VB_Name = "FrmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'checks all the variables
Option Explicit
'displays the main form
Private Sub CmdMainForm_Click()
    FrmCredits.Visible = False
    FrmMain.Visible = True
End Sub
'Ends the program
Private Sub CmdQuit1_Click()
    End
End Sub
 

