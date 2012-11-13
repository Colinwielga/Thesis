VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000080&
   Caption         =   "Main Page"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   10875
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDonLucia 
      BackColor       =   &H00000080&
      Height          =   3975
      Left            =   8400
      Picture         =   "Form.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3960
      Width           =   2055
   End
   Begin VB.CommandButton cmdHistory 
      Caption         =   "History of Gopher Hockey"
      Height          =   615
      Index           =   3
      Left            =   480
      TabIndex        =   6
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CommandButton cmdRoster 
      Caption         =   "Current Roster"
      Height          =   615
      Index           =   2
      Left            =   480
      TabIndex        =   5
      Top             =   1200
      Width           =   2295
   End
   Begin VB.CommandButton cmdSchedule 
      Caption         =   "Schedule"
      Height          =   615
      Index           =   0
      Left            =   480
      TabIndex        =   4
      Top             =   1920
      Width           =   2295
   End
   Begin VB.PictureBox picHome1 
      BackColor       =   &H00000080&
      Height          =   2775
      Left            =   3480
      Picture         =   "Form.frx":3EE0
      ScaleHeight     =   2715
      ScaleWidth      =   4275
      TabIndex        =   2
      Top             =   3960
      Width           =   4335
   End
   Begin VB.PictureBox picHome2 
      BackColor       =   &H00000080&
      Height          =   1695
      Left            =   4320
      Picture         =   "Form.frx":C814
      ScaleHeight     =   1635
      ScaleWidth      =   2835
      TabIndex        =   1
      Top             =   1560
      Width           =   2895
   End
   Begin VB.CommandButton cmdTickets 
      Caption         =   "Buy Tickets"
      Height          =   615
      Index           =   1
      Left            =   480
      TabIndex        =   0
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label lblWelcome 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "Welcome to Golden Gopher Hockey!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1560
      TabIndex        =   3
      Top             =   240
      Width           =   7815
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Gopher Hockey
'frmMain
'Cole and John
'10/30/06
'Objective: The objective of this form is set up a control page which acts as a home
'page and contains links to our entire project.  The user can (1) view current
'roster, (2) view the schedule, (3) view history of Gopher Hockey, and (4) purchase
'tickets.  The user can access these pages by clicking on the appropriate command
'button.  The overall purpose of our project is to provide information regarding
'the Minnesota Golden Gopher men's ice hockey team.  This information is vital to
'any hockey fan, especially in the state of Minnesota.

Option Explicit
Private Sub cmdDonLucia_Click()
    frmLucia.Visible = True     'accesses the Lucia form
    frmMain.Visible = False
End Sub

Private Sub cmdHistory_Click(Index As Integer)
    frmHistory.Visible = True   'accesses History form
    frmMain.Visible = False
End Sub

Private Sub cmdRoster_Click(Index As Integer)
    frmCurrentRoster.Visible = True
    frmMain.Visible = False
End Sub

Private Sub cmdSchedule_Click(Index As Integer)
    frmSchedule.Visible = True
    frmMain.Visible = False
End Sub

Private Sub cmdTickets_Click(Index As Integer)
    frmTickets.Visible = True
    frmMain.Visible = False
End Sub
