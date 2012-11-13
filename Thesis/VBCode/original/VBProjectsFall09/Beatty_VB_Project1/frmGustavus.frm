VERSION 5.00
Begin VB.Form frmGustavus 
   BackColor       =   &H0000C0C0&
   Caption         =   "Form1"
   ClientHeight    =   5820
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   Picture         =   "frmGustavus.frx":0000
   ScaleHeight     =   5820
   ScaleWidth      =   8925
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
      Height          =   3015
      Left            =   2640
      ScaleHeight     =   2955
      ScaleWidth      =   5715
      TabIndex        =   3
      Top             =   2400
      Width           =   5775
   End
   Begin VB.CommandButton cmdClick 
      Caption         =   "Click for information"
      Height          =   1335
      Left            =   360
      TabIndex        =   2
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Click to Return to Main Page"
      Height          =   975
      Left            =   360
      TabIndex        =   1
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   5040
      Width           =   1695
   End
End
Attribute VB_Name = "frmGustavus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: quick Facts about the MIAC'
    'Form name:Gustavus
    'Author:Alec Beatty'
    'Written 10/18/2009'
    'Objective: to give basic info about Gustavus'
Option Explicit

Private Sub cmdClick_Click()
 Dim Info As String
    
    Open App.Path & "\Gust.txt" For Input As #3
    
    Do While Not EOF(3)
        Input #3, Info
        picResults.Print Info
    
    Loop
    
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdReturn_Click()
    frmGustavus.Hide
    frmMIAC.Show
    
End Sub

