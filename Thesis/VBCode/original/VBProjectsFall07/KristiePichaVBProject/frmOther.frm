VERSION 5.00
Begin VB.Form frmOther 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMenu 
      BackColor       =   &H0080C0FF&
      Caption         =   "Back to Menu"
      Height          =   735
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdSAAD 
      BackColor       =   &H00808080&
      Caption         =   "SAAD Officer"
      Height          =   735
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdNHS 
      BackColor       =   &H00808080&
      Caption         =   "NHS"
      Height          =   735
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton cmdStudentcouncil 
      BackColor       =   &H0080C0FF&
      Caption         =   "Student Council"
      Height          =   735
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "frmOther"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdMenu_Click()
frmOther.Hide
frmMenu.Show
End Sub

Private Sub cmdNHS_Click()
MsgBox "Kristie was in National Honor Society for four years in 9th-12th grades."
End Sub

Private Sub cmdSAAD_Click()
MsgBox "Kristie was a SAAD Officer for one year in 9th grade."
End Sub

Private Sub cmdStudentcouncil_Click()
MsgBox "Kristie was on Student Council for three years in 9th, 11th-12th grades."
End Sub
