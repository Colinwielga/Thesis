VERSION 5.00
Begin VB.Form frmNike1 
   Caption         =   "NikeHome"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14160
   LinkTopic       =   "Form1"
   Picture         =   "frmNike1.frx":0000
   ScaleHeight     =   8460
   ScaleWidth      =   14160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00FF0000&
      Caption         =   "Quit"
      Height          =   1335
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CommandButton cmdStoreHome 
      BackColor       =   &H000080FF&
      Caption         =   "Back To Store Home"
      Height          =   1335
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2160
      Width           =   2055
   End
   Begin VB.CommandButton cmdBaseball 
      BackColor       =   &H00FF0000&
      Caption         =   "Baseball"
      Height          =   1335
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton cmdFootball 
      BackColor       =   &H000080FF&
      Caption         =   "Football"
      Height          =   1335
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton cmdRunning 
      BackColor       =   &H00FF0000&
      Caption         =   "Running"
      Height          =   1335
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton cmdBasketball 
      BackColor       =   &H000080FF&
      Caption         =   "Basketball"
      Height          =   1335
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label lblNike 
      BackColor       =   &H80000009&
      Caption         =   "Nike"
      BeginProperty Font 
         Name            =   "Franklin Gothic Heavy"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1440
      TabIndex        =   6
      Top             =   2520
      Width           =   1215
   End
End
Attribute VB_Name = "frmNike1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'AthleticStore
'Nike1
'Andrew Eisinger
'3/18/09
'This Form lets the user pick a sport
'when they choose a sport they are taken to that particular part of the store

Private Sub cmdBaseball_Click()
frmNikeBaseball.Show
frmNike1.Hide
End Sub

Private Sub cmdBasketball_Click()
frmNikeBasketball.Show
frmNike1.Hide
End Sub

Private Sub cmdFootball_Click()
frmNikeFootball.Show
frmNike1.Hide
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdRunning_Click()
frmNikeRunning.Show
frmNike1.Hide
End Sub

Private Sub cmdStoreHome_Click()
frmStoreHome.Show
frmNike1.Hide
End Sub
