VERSION 5.00
Begin VB.Form frmJamaicaStart 
   BackColor       =   &H0000FFFF&
   Caption         =   "Form1"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   10920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdstart 
      BackColor       =   &H00C0C000&
      Caption         =   "Press Here to Start Planning Your Dream Jamaican Vacation!!!!!!!!!!!"
      Height          =   2055
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   3495
   End
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Quit"
      Height          =   855
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdback 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Back"
      Height          =   855
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label lbltitle 
      BackColor       =   &H0000FFFF&
      Caption         =   "Welcome to Jamaica!!"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      TabIndex        =   3
      Top             =   120
      Width           =   9495
   End
   Begin VB.Image Image1 
      Height          =   3930
      Left            =   480
      Picture         =   "frmJamaicaStart.frx":0000
      Top             =   960
      Width           =   6000
   End
End
Attribute VB_Name = "frmJamaicaStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name: Travel Agency
'Form Name: frmJamaicaStart
'Authors: Cassie Scherer and Jordan Schmaltz
'Date Written: 3/5/08
'Objective: This page tells the user which vacation they've selected.
'They have the opportunity to go back, if they wish to select a different vacation
'The user goes from this page into the planning on their vacation


Private Sub cmdback_Click()

'Here we are giving the user the option to return to the previous screen

frmJamaicaStart.Hide
frmBeginning.Show

End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub cmdstart_Click()

'Here we are allowing the user to start planning their hotel selection in Jamaica by going to a different form

frmJamaicaStart.Hide
frmJamaicaHotels.Show

End Sub

Private Sub Form_Load()

'This code centers the form on computer screen upon loading

Top = Screen.Height / 2 - Height / 2
Left = Screen.Width / 2 - Width / 2

End Sub
