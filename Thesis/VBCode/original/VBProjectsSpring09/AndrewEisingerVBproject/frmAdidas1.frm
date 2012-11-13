VERSION 5.00
Begin VB.Form frmAdidas1 
   Caption         =   "Adidas Home"
   ClientHeight    =   11565
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   16710
   LinkTopic       =   "Form1"
   Picture         =   "frmAdidas1.frx":0000
   ScaleHeight     =   11565
   ScaleWidth      =   16710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00404040&
      Caption         =   "Quit"
      Height          =   1575
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4560
      Width           =   2535
   End
   Begin VB.CommandButton cmdGoBack 
      BackColor       =   &H00404040&
      Caption         =   "Go Back to Store Home"
      Height          =   1575
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Width           =   2535
   End
   Begin VB.CommandButton cmdAdidasCycling 
      BackColor       =   &H00404040&
      Caption         =   "Cycling"
      Height          =   1575
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   2535
   End
   Begin VB.CommandButton cmdAdidasRunning 
      BackColor       =   &H00404040&
      Caption         =   "Running"
      Height          =   1575
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   2535
   End
   Begin VB.CommandButton cmdAdidasSoccer 
      BackColor       =   &H00404040&
      Caption         =   "Soccer"
      Height          =   1575
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2640
      Width           =   2535
   End
End
Attribute VB_Name = "frmAdidas1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' AthleticStore
' Adidas1
' Andrew Eisinger
' 3/17/09
'This program lets the user pick a sport
'This program then based on the selection sends to a different form

Private Sub cmdAdidasCycling_Click()
frmAdidasCycling.Show
frmAdidas1.Hide

End Sub

Private Sub cmdAdidasRunning_Click()
frmAdidasRunning.Show
frmAdidas1.Hide
End Sub

Private Sub cmdAdidasSoccer_Click()
frmAdidasSoccer.Show
frmAdidas1.Hide
End Sub

Private Sub cmdGoBack_Click()
frmStoreHome.Show
frmAdidas1.Hide
End Sub

Private Sub cmdQuit_Click()
End
End Sub
