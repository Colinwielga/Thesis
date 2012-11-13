VERSION 5.00
Begin VB.Form frmCarleton 
   BackColor       =   &H0000FFFF&
   Caption         =   "Form1"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   Picture         =   "frmCarleton.frx":0000
   ScaleHeight     =   5775
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCA3 
      BackColor       =   &H0000FFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Text            =   "Stadium: Laird Stadium"
      Top             =   4680
      Width           =   1935
   End
   Begin VB.TextBox txtCa2 
      BackColor       =   &H0000FFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Text            =   "Nickname: Knights"
      Top             =   4320
      Width           =   1935
   End
   Begin VB.TextBox txtCa1 
      BackColor       =   &H0000FFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Text            =   "Colors: Maize and Blue"
      Top             =   3960
      Width           =   1935
   End
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
      Height          =   1335
      Left            =   2880
      ScaleHeight     =   1275
      ScaleWidth      =   2715
      TabIndex        =   3
      Top             =   2280
      Width           =   2775
   End
   Begin VB.CommandButton cmdClick 
      BackColor       =   &H0000FFFF&
      Caption         =   "Click for more information"
      Height          =   1215
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Main Page"
      Height          =   735
      Left            =   2280
      TabIndex        =   1
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0000FFFF&
      Caption         =   "Quit"
      Height          =   495
      Left            =   4080
      TabIndex        =   0
      Top             =   4560
      Width           =   1575
   End
End
Attribute VB_Name = "frmCarleton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: quick Facts about the MIAC'
    'Form name:Carleton
    'Author:Alec Beatty'
    'Written 10/18/2009'
    'Objective: to give basic info about Carleton'
Option Explicit



Private Sub cmdClick_Click()
    Dim Town As String
    
    Open App.Path & "\town.txt" For Input As #1 'File Path'
    
    picResults.Print "Carleton is located in"
    
    Do While Not EOF(1)
        Input #1, Town
        picResults.Print Town
    
    Loop
    
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdReturn_Click()
    frmCarleton.Hide
    frmMIAC.Show
    
End Sub
