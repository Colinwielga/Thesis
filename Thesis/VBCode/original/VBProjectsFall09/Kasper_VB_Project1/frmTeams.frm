VERSION 5.00
Begin VB.Form frmTeams 
   Caption         =   "Teams"
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   Picture         =   "frmTeams.frx":0000
   ScaleHeight     =   4740
   ScaleWidth      =   4950
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H000000FF&
      Caption         =   "Return to Main Menu"
      Height          =   495
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H0000FFFF&
      Caption         =   "Exit"
      Height          =   495
      Left            =   2040
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdLions 
      BackColor       =   &H8000000D&
      Caption         =   "Visit Lions"
      Height          =   615
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdBears 
      BackColor       =   &H000040C0&
      Caption         =   "Visit Bears"
      Height          =   615
      Left            =   2880
      MaskColor       =   &H000040C0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdPackers 
      BackColor       =   &H00004000&
      Caption         =   "Visitr Packers"
      Height          =   615
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdVikings 
      BackColor       =   &H00400040&
      Caption         =   "Visit Vikings"
      Height          =   615
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3360
      Width           =   1335
   End
End
Attribute VB_Name = "frmTeams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Author: Brandon Kasper
'Written 10/19/2009
'this form allows the user to choose a team to view

Private Sub cmdBears_Click()
    frmTeams.Hide 'hides Team page
    frmBears.Show 'shows Bears Page
End Sub

Private Sub cmdLions_Click()
    frmTeams.Hide 'hides Teams page
    frmLions.Show 'shows lions page
End Sub

Private Sub cmdPackers_Click()
    frmTeams.Hide 'hides Team page
    frmPackers.Show 'shows Packers Page
End Sub

Private Sub cmdquit_Click()
    End
End Sub

Private Sub cmdReturn_Click()
    frmTeams.Hide 'hides Team page
    frmWelcome.Show 'shows Welcome Page
End Sub

Private Sub cmdVikings_Click()
    frmTeams.Hide 'hides Team page
    frmVikings.Show 'shows Vikings Page
End Sub


