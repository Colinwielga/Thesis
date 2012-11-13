VERSION 5.00
Begin VB.Form frmVegasStart 
   BackColor       =   &H000080FF&
   Caption         =   "Form1"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   10335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdstart 
      BackColor       =   &H00FF00FF&
      Caption         =   "Start Planning Your Dream Vacation in Las Vegas!!!!!!!!!"
      Height          =   1935
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2400
      Width           =   3375
   End
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H00C0C000&
      Caption         =   "Quit"
      Height          =   735
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdback 
      BackColor       =   &H00C0C000&
      Caption         =   "Back"
      Height          =   735
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label lbltitle 
      BackColor       =   &H00000000&
      Caption         =   "Welcome to Las Vegas!!"
      BeginProperty Font 
         Name            =   "Modern"
         Size            =   26.25
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   480
      TabIndex        =   3
      Top             =   480
      Width           =   5775
   End
   Begin VB.Image Image1 
      Height          =   3780
      Left            =   480
      Picture         =   "frmVegasStart.frx":0000
      Top             =   1560
      Width           =   5625
   End
End
Attribute VB_Name = "frmVegasStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Travel Agency
'Form Name: frmVegasStart
'Authors: Cassie Scherer and Jordan Schmaltz
'Date Written: 3/5/08
'Objective: This page tells the user which vacation they've selected.
'They have the opportunity to go back, if they wish to select a different vacation
'The user goes from this page into the planning on their vacation


Private Sub cmdback_Click()

'Here we are allowing the user to return to the previous screen if they decide they want to go back

frmVegasStart.Hide
frmBeginning.Show
End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub cmdstart_Click()

'Here we are allowing the user to start planning their hotel selection in Las Vegas by going to a different form

frmVegasStart.Hide
frmVegasHotels.Show

End Sub

Private Sub Form_Load()

'This code centers the form on computer screen upon loading

Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2

End Sub
