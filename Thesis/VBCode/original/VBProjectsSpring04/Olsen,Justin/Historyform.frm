VERSION 5.00
Begin VB.Form Historyform 
   BackColor       =   &H0000C000&
   Caption         =   " "
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8565
   LinkTopic       =   "Form2"
   ScaleHeight     =   5160
   ScaleWidth      =   8565
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdmylinkso 
      BackColor       =   &H00FF8080&
      Caption         =   "Go to my links!"
      Height          =   615
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton cmdstart1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Go back to the beginning."
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton cmddunbar 
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      Height          =   495
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000C0&
      Caption         =   "Click here to move on to building and buying tips."
      Height          =   495
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4440
      Width           =   3735
   End
   Begin VB.PictureBox Picture1 
      Height          =   2295
      Left            =   120
      Picture         =   "Historyform.frx":0000
      ScaleHeight     =   2235
      ScaleWidth      =   5955
      TabIndex        =   1
      Top             =   120
      Width           =   6015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Designed by: Justin Olsen"
      Height          =   255
      Left            =   8520
      TabIndex        =   5
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label History 
      Caption         =   $"Historyform.frx":5FF7
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   8175
   End
End
Attribute VB_Name = "Historyform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Purpose = This form is here to give the user a little background information on cedar strip canoes.
'Project Name= Visual Basic Canoe Project("M:\CS130\CanoeProject")
'Form Name= Form 1 ("M:\CS130\CanoeProject\Project1.vbp\Historyform.frm")
Private Sub cmddunbar_Click()
End
End Sub

Private Sub cmdmylinkso_Click()
Historyform.Hide
Form2.Show
End Sub

Private Sub cmdstart1_Click()
Historyform.Hide
Form1.Show
End Sub

Private Sub Command1_Click()
    Buildingform1.Show
    Historyform.Hide
End Sub

