VERSION 5.00
Begin VB.Form FrmHome 
   Caption         =   "Hang-Man"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   Picture         =   "FrmHome.frx":0000
   ScaleHeight     =   6945
   ScaleWidth      =   9435
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdQuit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quit"
      Height          =   975
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4680
      Width           =   2295
   End
   Begin VB.CommandButton CmdHelp 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Options"
      Height          =   975
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4680
      Width           =   2295
   End
   Begin VB.CommandButton CmdPlay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Play Hang-Man"
      Height          =   975
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4680
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hang-Man!"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   480
      TabIndex        =   4
      Top             =   2160
      Width           =   8295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome To"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      TabIndex        =   3
      Top             =   1320
      Width           =   4335
   End
End
Attribute VB_Name = "FrmHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Hang Man
'Form Name: FrmHome
'Authors: Breanna Riskey and Heidi Golliet
'Date Completed: Monday, March 31st
'Objective: The purpose of this form is to direct the user to the other forms available in the program.
'It is this form that the user will enter and exit the program from.


Option Explicit

Private Sub CmdHelp_Click()
    FrmHome.Visible = False
    FrmOptions.Visible = True
    FrmPlayGame.Visible = False
End Sub

Private Sub CmdPlay_Click()
    FrmPlayGame.Visible = True
    FrmHome.Visible = False
    FrmOptions.Visible = False
End Sub

Private Sub CmdQuit_Click()
End
End Sub

