VERSION 5.00
Begin VB.Form frmBostonStart 
   BackColor       =   &H00FF8080&
   Caption         =   "Form1"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   11175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdstart 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Press Here to Start Planning Your Dream Vacation in Boston!!!!!!!!!!!!!!!"
      Height          =   2055
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   3255
   End
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H0080FFFF&
      Caption         =   "Quit"
      Height          =   615
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton cmdback 
      BackColor       =   &H00FF80FF&
      Caption         =   "Back"
      Height          =   615
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label lbltitle 
      BackColor       =   &H0080C0FF&
      Caption         =   "Welcome to Boston!!!"
      BeginProperty Font 
         Name            =   "Bickham Script Pro Regular"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   855
      Left            =   840
      TabIndex        =   3
      Top             =   240
      Width           =   5655
   End
   Begin VB.Image Image1 
      Height          =   4635
      Left            =   480
      Picture         =   "frmBostonStart.frx":0000
      Top             =   1320
      Width           =   6000
   End
End
Attribute VB_Name = "frmBostonStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Travel Agency
'Form Name: frmBostonStart
'Authors: Cassie Scherer and Jordan Schmaltz
'Date Written: 3/5/08
'Objective: This page tells the user which vacation they've selected.
'They have the opportunity to go back, if they wish to select a different vacation
'The user goes from this page into the planning on their vacation


Private Sub cmdback_Click()

'Here we give the user the option to return to the previous screen.

frmBostonStart.Hide
frmBeginning.Show

End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub cmdstart_Click()

'Here the user continues on to choosing what hotel they will be staying at in Boston.

frmBostonStart.Hide
frmBostonHotels.Show

End Sub

Private Sub Form_Load()

'This code centers the form on computer screen upon loading

Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2

End Sub
