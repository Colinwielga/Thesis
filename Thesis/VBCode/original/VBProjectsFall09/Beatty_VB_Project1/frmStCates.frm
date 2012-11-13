VERSION 5.00
Begin VB.Form frmStMary 
   BackColor       =   &H0000FF00&
   Caption         =   "Form1"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8505
   FillColor       =   &H0000FF00&
   LinkTopic       =   "Form1"
   Picture         =   "frmStCates.frx":0000
   ScaleHeight     =   6570
   ScaleWidth      =   8505
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClick 
      Caption         =   "Click for more information"
      Height          =   1455
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Main Page"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Label lblMary3 
      BackColor       =   &H0000FF00&
      Caption         =   "Colors: Red and White "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2520
      TabIndex        =   5
      Top             =   3720
      Width           =   2895
   End
   Begin VB.Label lblMary2 
      BackColor       =   &H0000FF00&
      Caption         =   "Nickname: The Cardnials"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   3240
      Width           =   3015
   End
   Begin VB.Label lblMary1 
      BackColor       =   &H0000FF00&
      Caption         =   "St. Mary's is located in Winona, MN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2520
      TabIndex        =   3
      Top             =   2400
      Width           =   2775
   End
End
Attribute VB_Name = "frmStMary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: quick Facts about the MIAC'
    'Form name:St.Mary
    'Author:Alec Beatty'
    'Written 10/18/2009'
    'Objective: to give basic info about St. Mary's'
Option Explicit

Private Sub cmdClick_Click()
MsgBox "St Mary's doesn't have a football team.", , "Fun Stuff"
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdReturn_Click()
    frmStCates.Hide
    frmMIAC.Show
End Sub

