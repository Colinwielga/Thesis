VERSION 5.00
Begin VB.Form frmCasino 
   Caption         =   "Casino"
   ClientHeight    =   6870
   ClientLeft      =   3450
   ClientTop       =   1935
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   9000
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00404080&
      Caption         =   "Go Home"
      Height          =   615
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdGetMoney 
      BackColor       =   &H00008000&
      Caption         =   "Get Money"
      Height          =   615
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdEnter 
      Height          =   1455
      Left            =   960
      Picture         =   "frmCasino.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Enter Casino"
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Image img1 
      Height          =   6900
      Left            =   0
      Picture         =   "frmCasino.frx":D043
      Top             =   0
      Width           =   9000
   End
End
Attribute VB_Name = "frmCasino"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Initial Form to get money to enter simulated Casino or exit

Private Sub cmdGetMoney_Click()
    'Go from Casino to get money page
    frmCasino.Hide
    frmLoanShark.Show
    End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdEnter_Click()
'You have to have money to enter casino
If balanceglobal > 0 Then
    frmCasino.Hide
    frmCop.Show
Else
    MsgBox "You need to get money from the ATM or find someone to loan you some.", , "Money????"
End If
End Sub

Private Sub cmdStats_Click()
frmCasino.Hide
frmStats.Show
End Sub

Private Sub Form_Initialize()
    'You need a name to later to take out money to get in
    nameglobal = InputBox("What is your name?", "Name")
End Sub

