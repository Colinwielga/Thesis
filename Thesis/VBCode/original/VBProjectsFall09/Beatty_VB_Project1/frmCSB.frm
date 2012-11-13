VERSION 5.00
Begin VB.Form frmCSB 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   Picture         =   "frmCSB.frx":0000
   ScaleHeight     =   6450
   ScaleWidth      =   5580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClick 
      Caption         =   "Click for more information"
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Main Page"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "*St. Ben's nickname is the Blazers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2160
      TabIndex        =   5
      Top             =   4320
      Width           =   2775
   End
   Begin VB.Label lblCSB2 
      BackColor       =   &H000000FF&
      Caption         =   "*Colors: Red and White"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   4
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label lblCSB1 
      BackColor       =   &H000000FF&
      Caption         =   "*St. Bens is the only Benedictine college for women in the U.S."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2160
      TabIndex        =   3
      Top             =   3240
      Width           =   2775
   End
End
Attribute VB_Name = "frmCSB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: quick Facts about the MIAC'
    'Form name:CSB
    'Author:Alec Beatty'
    'Written 10/18/2009'
    'Objective: to give basic info about St. Ben's'
Option Explicit

Private Sub cmdClick_Click()
MsgBox "Did you know that St. Ben's is in St. Joseph, MN?", , "Woah!!"
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdReturn_Click()
    frmCSB.Hide
    frmMIAC.Show
    
End Sub

