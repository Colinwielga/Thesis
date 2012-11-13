VERSION 5.00
Begin VB.Form frmMarriott 
   BackColor       =   &H00FF80FF&
   Caption         =   "Marriott"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9240
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   9240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FFFF00&
      Caption         =   "Click to go back to start page"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3600
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0000FF00&
      Caption         =   "Quit"
      Height          =   495
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton cmdRoom 
      BackColor       =   &H0080FFFF&
      Caption         =   "Click here to see your room options"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      Height          =   3615
      Left            =   240
      Picture         =   "frmMarriott.frx":0000
      ScaleHeight     =   3555
      ScaleWidth      =   5715
      TabIndex        =   1
      Top             =   960
      Width           =   5775
   End
   Begin VB.Label lblwelcome 
      BackColor       =   &H00800080&
      Caption         =   "Welcome to The Marriott we are glad you pick to stay with us and hope you enjoy your vacation!"
      BeginProperty Font 
         Name            =   "Gulim"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   6855
   End
End
Attribute VB_Name = "frmMarriott"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name:Travel New York And L.A
'Form Name: frmActivities
'Author: Emily Backes
'Date Written: 3-17-09
'Objective: This form is the welcome page for the Marriott hotel
'The user can click on it to find out about rooms, restaurants and
'activities
Option Explicit


Private Sub cmdBack_Click()
'allows the user to go back to the start page
frmMarriott.Hide
frmOpen.Show

End Sub


Private Sub cmdquit_Click()
'quits the form
End
End Sub

Private Sub cmdroom_Click()
'allows the user to see the room selection form
frmMarriott.Hide
frmRoomMarriott.Show

End Sub
