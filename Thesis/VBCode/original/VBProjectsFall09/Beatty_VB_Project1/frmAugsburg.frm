VERSION 5.00
Begin VB.Form frmAugsburg 
   BackColor       =   &H80000003&
   Caption         =   "Form1"
   ClientHeight    =   3555
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7215
   FillColor       =   &H0000C000&
   ForeColor       =   &H80000013&
   LinkTopic       =   "Form1"
   Picture         =   "frmAugsburg.frx":0000
   ScaleHeight     =   3555
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtA4 
      BackColor       =   &H80000003&
      Height          =   495
      Left            =   4800
      TabIndex        =   5
      Text            =   "Stadium: Edor Nelson Field "
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox txtA3 
      BackColor       =   &H80000003&
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Text            =   "Colors: Maroon and Gray"
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox txtA2 
      BackColor       =   &H80000003&
      Height          =   495
      Left            =   4800
      TabIndex        =   3
      Text            =   "Nickname: Auggies"
      Top             =   720
      Width           =   2175
   End
   Begin VB.TextBox txtA1 
      BackColor       =   &H80000003&
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Text            =   "Located in Minneapolis, MN"
      Top             =   360
      Width           =   2175
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Main Page"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Top             =   2640
      Width           =   1095
   End
End
Attribute VB_Name = "frmAugsburg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  'Project name: quick Facts about the MIAC'
    'Form name:Augsburg
    'Author:Alec Beatty'
    'Written 10/18/2009'
    'Objective: to give basic info about Augsburg
Option Explicit


Private Sub cmdQuit_Click() 'Allows person to quit'
End
End Sub

Private Sub cmdReturn_Click() 'Allows person to Return to main menu'
    frmMIAC.Show
    frmAugsburg.Hide

End Sub
