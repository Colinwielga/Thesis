VERSION 5.00
Begin VB.Form frmParkCentral 
   BackColor       =   &H000080FF&
   Caption         =   "Park Central Hotel"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   10275
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRoom 
      BackColor       =   &H00FFFF80&
      Caption         =   "Click here to see the room options for the Park Central hotel"
      Height          =   1935
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2040
      Width           =   3735
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H008080FF&
      Caption         =   "Quit"
      Height          =   495
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H0080FFFF&
      Caption         =   "click to go back to start page"
      Height          =   1095
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4800
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      Height          =   3615
      Left            =   4440
      Picture         =   "frmParkCentral.frx":0000
      ScaleHeight     =   3555
      ScaleWidth      =   5235
      TabIndex        =   1
      Top             =   1680
      Width           =   5295
   End
   Begin VB.Label lblCentral 
      BackColor       =   &H00404040&
      Caption         =   $"frmParkCentral.frx":4E4E
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   7695
   End
End
Attribute VB_Name = "frmParkCentral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Name:Travel New York And L.A
'Form Name: frmActivities
'Author: Emily Backes
'Date Written: 3-17-09
'Objective: This form is the welcome page for the Park Centralhotel
'The user can click on it to find out about rooms, restaurants and
'activities
Option Explicit



Private Sub cmdBack_Click()
'allows the user to go back to the start page
frmParkCentral.Hide
frmOpen.Show
End Sub

Private Sub cmdquit_Click()
End
End Sub



Private Sub cmdroom_Click()
'allows the user to go see the room selection for the park central
frmParkCentral.Hide
frmRoomPC.Show
End Sub
