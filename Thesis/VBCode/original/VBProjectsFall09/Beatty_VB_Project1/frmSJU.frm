VERSION 5.00
Begin VB.Form frmSJU 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   Picture         =   "frmSJU.frx":0000
   ScaleHeight     =   5880
   ScaleWidth      =   7590
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
      Height          =   2535
      Left            =   2880
      ScaleHeight     =   2475
      ScaleWidth      =   3555
      TabIndex        =   3
      Top             =   2400
      Width           =   3615
   End
   Begin VB.CommandButton cmdClick 
      Caption         =   "Click for information"
      Height          =   1095
      Left            =   600
      TabIndex        =   2
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Main Page"
      Height          =   735
      Left            =   600
      TabIndex        =   1
      Top             =   3840
      Width           =   2055
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   735
      Left            =   600
      TabIndex        =   0
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Label lblSJU 
      BackColor       =   &H000000FF&
      Caption         =   "Did you know the Johnnies Colors are Cardnial and Blue?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      TabIndex        =   4
      Top             =   5040
      Width           =   3735
   End
End
Attribute VB_Name = "frmSJU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project name: quick Facts about the MIAC'
    'Form name:SJU
    'Author:Alec Beatty'
    'Written 10/18/2009'
    'Objective: to give basic info about St. John's
Option Explicit

Private Sub cmdClick_Click()

    Dim Info As String
    
    Open App.Path & "\SJU.txt" For Input As #6
    
    
    Do While Not EOF(6)
        Input #6, Info
        picResults.Print Info
    Loop
    
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdReturn_Click()
    frmSJU.Hide
    frmMIAC.Show
End Sub
