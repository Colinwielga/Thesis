VERSION 5.00
Begin VB.Form frmItalyStart 
   BackColor       =   &H00004080&
   Caption         =   "Form1"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   10530
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdstart 
      BackColor       =   &H0080C0FF&
      Caption         =   "Press Here to Start Planning Your Dream Vacation in Italy!!!!!!!!!!!!"
      Height          =   1935
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2280
      Width           =   2535
   End
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H0080C0FF&
      Caption         =   "Quit"
      Height          =   975
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H0080C0FF&
      Caption         =   "Back"
      Height          =   975
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Label lbltitle 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Welcome to Italy!!"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   3
      Top             =   360
      Width           =   6375
   End
   Begin VB.Image Image1 
      Height          =   4500
      Left            =   360
      Picture         =   "frmItalyStart.frx":0000
      Top             =   1320
      Width           =   6750
   End
End
Attribute VB_Name = "frmItalyStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Travel Agency
'Form Name: frmItalyStart
'Authors: Cassie Scherer and Jordan Schmaltz
'Date Written: 3/5/08
'Objective: This page tells the user which vacation they've selected.
'They have the opportunity to go back, if they wish to select a different vacation
'The user goes from this page into the planning on their vacation


Private Sub cmdback_Click()

'Here we allow the user to return to the previous screen

frmItalyStart.Hide
frmBeginning.Show

End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub cmdstart_Click()

'Here the user continues on in planning their vacation by selecting a hotel

frmItalyStart.Hide
frmItalyHotels.Show

End Sub

Private Sub Form_Load()

'This code centers the form on computer screen upon loading

Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2

End Sub
