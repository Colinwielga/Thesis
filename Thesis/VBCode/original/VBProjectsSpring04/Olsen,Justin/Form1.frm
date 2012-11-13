VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   8835
   ScaleWidth      =   11175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdquitnow 
      Caption         =   "Quit"
      Height          =   615
      Left            =   8760
      TabIndex        =   5
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdinfo 
      BackColor       =   &H000080FF&
      Caption         =   "Learn more about building your own canoe, or where to go to buy one."
      Height          =   1215
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3600
      Width           =   3255
   End
   Begin VB.CommandButton cmdhistory 
      BackColor       =   &H0000C000&
      Caption         =   "Learn about the history of Cedar Strip Canoes!"
      Height          =   1455
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   2895
   End
   Begin VB.PictureBox Picture1 
      Height          =   1815
      Left            =   4680
      Picture         =   "Form1.frx":83EF
      ScaleHeight     =   1755
      ScaleWidth      =   3675
      TabIndex        =   0
      Top             =   3360
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      Caption         =   "Designed by: Justin Olsen"
      Height          =   255
      Left            =   8760
      TabIndex        =   4
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":D9DB
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   2520
      Width           =   8175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Project Name= Visual Basic Canoe Project("M:\CS130\CanoeProject")
'Form Name= Form 1 ("M:\CS130\CanoeProject\Project1.vbp\Form1.frm")
'Author= Justin Olsen
'Date Written= 3-13-04
'Overall Purpose= To create a VB project that presents options for research and a general knowledge base of Cedar Strip Canoe Building.
'Purpose = This form is here to welcome the user and to give them a choice on what direction to go in the program, either to background information on cedar strip canoes, or to information on building and buying canoes.
Dim Path As String

Private Sub cmdhistory_Click()
    Historyform.Show
    Form1.Hide
End Sub

Private Sub cmdinfo_Click()
    Buildingform1.Show
    Form1.Hide
End Sub

Private Sub cmdquitnow_Click()
End
End Sub

Private Sub Form_Load()
Path = "M:\CS130\CanoeProject\"
End Sub


