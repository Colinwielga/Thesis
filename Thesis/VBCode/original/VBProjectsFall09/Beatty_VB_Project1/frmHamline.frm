VERSION 5.00
Begin VB.Form frmHamline 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form1"
   ClientHeight    =   6720
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   Picture         =   "frmHamline.frx":0000
   ScaleHeight     =   6720
   ScaleWidth      =   9060
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picResults 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   2160
      ScaleHeight     =   3315
      ScaleWidth      =   5115
      TabIndex        =   3
      Top             =   2280
      Width           =   5175
   End
   Begin VB.CommandButton cmdClick 
      Caption         =   "Click for information"
      Height          =   1575
      Left            =   360
      TabIndex        =   2
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Main Page"
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   4800
      Width           =   1695
   End
End
Attribute VB_Name = "frmHamline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

 'Project name: quick Facts about the MIAC'
    'Form name:Hamline
    'Author:Alec Beatty'
    'Written 10/18/2009'
    'Objective: to give basic info about Hamline'
Option Explicit


Private Sub cmdClick_Click()

    Dim Info As String
    
    Open App.Path & "\Hamline.txt" For Input As #4
    
    
    Do While Not EOF(4)
        Input #4, Info
        picResults.Print Info

    
    Loop
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdReturn_Click()
    frmHamline.Hide
    frmMIAC.Show
End Sub

