VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H008080FF&
   Caption         =   "Form1"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   8655
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDuration 
      Height          =   495
      Left            =   4320
      TabIndex        =   3
      Top             =   2880
      Width           =   1935
   End
   Begin VB.TextBox txtBudget 
      Height          =   495
      Left            =   4320
      TabIndex        =   2
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   1215
      Left            =   5640
      TabIndex        =   1
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CommandButton cmdContinue 
      BackColor       =   &H0080FFFF&
      Caption         =   "Continue"
      Height          =   1215
      Left            =   5640
      MaskColor       =   &H0080FFFF&
      TabIndex        =   0
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label lblDuration 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Enter duration of trip"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   2880
      Width           =   3375
   End
   Begin VB.Label lblBudget 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Enter amount budgeted for trip"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   1680
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim budget As Single


