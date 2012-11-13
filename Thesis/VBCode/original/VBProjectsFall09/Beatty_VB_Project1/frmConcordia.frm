VERSION 5.00
Begin VB.Form frmConcordia 
   BackColor       =   &H00000040&
   Caption         =   "Form1"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   Picture         =   "frmConcordia.frx":0000
   ScaleHeight     =   6345
   ScaleWidth      =   7950
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCo3 
      BackColor       =   &H00000040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Text            =   "Stadium: Jake Christiansen Stadium"
      Top             =   2040
      Width           =   3855
   End
   Begin VB.TextBox txtCo2 
      BackColor       =   &H00000040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3960
      TabIndex        =   5
      Text            =   "Colors: Maroon and Gold"
      Top             =   1680
      Width           =   3135
   End
   Begin VB.TextBox txtCo1 
      BackColor       =   &H00000040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   3960
      TabIndex        =   4
      Text            =   "Location: Moorhead, MN"
      Top             =   1320
      Width           =   3015
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
      Height          =   1695
      Left            =   3000
      ScaleHeight     =   1635
      ScaleWidth      =   2955
      TabIndex        =   3
      Top             =   2640
      Width           =   3015
   End
   Begin VB.CommandButton cmdClick 
      Caption         =   "Click for more information"
      Height          =   1335
      Left            =   480
      TabIndex        =   2
      Top             =   2400
      Width           =   2175
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Main Page"
      Height          =   735
      Left            =   480
      TabIndex        =   1
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   4920
      Width           =   1455
   End
End
Attribute VB_Name = "frmConcordia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: quick Facts about the MIAC'
    'Form name:Concordia
    'Author:Alec Beatty'
    'Written 10/18/2009'
    'Objective: to give basic info about Concordia'
Option Explicit

Private Sub cmdClick_Click()
 Dim Name As String
    
    Open App.Path & "\Cobbers.txt" For Input As #2
    
    picResults.Print "Concordia's nickname is:"
    
    Do While Not EOF(2)
        Input #2, Name
        picResults.Print Name
    
    Loop
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdReturn_Click()
    frmConcordia.Hide
    frmMIAC.Show
End Sub
